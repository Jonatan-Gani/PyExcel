using System;
using System.Collections.Generic;
using System.IO;
using Apache.Arrow;
using Apache.Arrow.Ipc;
using Apache.Arrow.Types;

namespace PyExcel.Excel;

/// <summary>
/// Arrow IPC marshalling — the C# half of the kernel data plane. Encodes
/// values originating from Excel ranges (object[,] / object[] / scalar)
/// into Arrow IPC streams that <c>pyexcel.kernel.arrow_io</c> decodes on
/// the kernel side, and inversely decodes streams produced by the kernel
/// back to shapes Excel-DNA can spill.
///
/// <para>Wire-format compatibility with <c>arrow_io.py</c> is the whole
/// point of this class. The shape and orientation hints live as Arrow
/// schema metadata under fixed keys:</para>
///
/// <code>
///   pyexcel-shape       = "table" | "vector" | "scalar"
///   pyexcel-orientation = "row" | "column"   (vectors only)
/// </code>
///
/// <para>Streams without a <c>pyexcel-shape</c> key decode as a table,
/// matching the Python side's behaviour for buffers produced by external
/// Arrow writers.</para>
///
/// <para>Type handling for <c>object?[,]</c> inputs is per-column: the
/// builder scans each column and picks the most specific Arrow type that
/// fits every value — double / boolean / string, with string acting as the
/// fallback for mixed-type or non-primitive columns. Nulls are preserved
/// throughout.</para>
/// </summary>
public static class ArrowMarshal
{
    // Schema-metadata keys (must match arrow_io.py).
    private const string MetaShapeKey = "pyexcel-shape";
    private const string MetaOrientKey = "pyexcel-orientation";

    // Field-level metadata key. Marks an Arrow string column whose cells
    // are formula source rather than literals — the host writes those via
    // Range.Formula instead of Range.Value2 so Excel evaluates them.
    private const string FieldMetaCellTypeKey = "pyexcel-cell-type";
    private const string CellTypeFormula = "formula";

    // Field-level metadata key for image payloads: the rendered format of
    // the binary column ("svg" or "png"). Only present on shape=image.
    private const string FieldMetaImageFormatKey = "pyexcel-image-format";

    // Wire values.
    private const string ShapeTable = "table";
    private const string ShapeVector = "vector";
    private const string ShapeScalar = "scalar";
    private const string ShapeChart = "chart";
    private const string ShapeImage = "image";
    private const string OrientRow = "row";
    private const string OrientColumn = "column";

    // Reusable formula field metadata — every Formula column gets the same
    // single-entry dictionary, so cache it instead of allocating per column.
    private static readonly IReadOnlyDictionary<string, string> FormulaFieldMetadata =
        new Dictionary<string, string>(StringComparer.Ordinal)
        {
            [FieldMetaCellTypeKey] = CellTypeFormula,
        };

    // -----------------------------------------------------------------------
    // Encode
    // -----------------------------------------------------------------------

    /// <summary>
    /// Encode a 2-D rectangular range as a table-shaped Arrow IPC stream.
    /// Column types are inferred per column from the cell values; mixed-type
    /// columns fall back to string with each value stringified via
    /// <see cref="object.ToString"/>.
    /// </summary>
    /// <param name="values">Row-major 2-D array. <c>values[r, c]</c> is the
    /// cell at row r, column c. Each cell may be double / string / bool /
    /// null; other types are coerced via <c>ToString()</c>.</param>
    /// <param name="columnNames">Optional column header names. If null or
    /// the wrong length, columns are named "0", "1", … positionally.</param>
    public static byte[] EncodeTable(object?[,] values, IReadOnlyList<string>? columnNames = null)
    {
        if (values is null) throw new ArgumentNullException(nameof(values));
        var rows = values.GetLength(0);
        var cols = values.GetLength(1);

        var fields = new Field[cols];
        var arrays = new IArrowArray[cols];
        for (var c = 0; c < cols; c++)
        {
            var column = new object?[rows];
            for (var r = 0; r < rows; r++) column[r] = values[r, c];
            arrays[c] = BuildColumn(column, out var type, out var fieldMetadata);
            var name = columnNames is { } cn && c < cn.Count && cn[c] is { } n ? n : c.ToString();
            fields[c] = new Field(name, type, nullable: true, fieldMetadata);
        }

        var schema = BuildSchema(fields, ShapeTable, orientation: null);
        return WriteStream(schema, arrays, rows);
    }

    /// <summary>
    /// Encode a 1-D sequence as a vector-shaped Arrow IPC stream. Type
    /// inference is the same as for a single column of
    /// <see cref="EncodeTable"/>.
    /// </summary>
    public static byte[] EncodeVector(
        object?[] values,
        ArrowOrientation orientation = ArrowOrientation.Column)
    {
        if (values is null) throw new ArgumentNullException(nameof(values));

        var array = BuildColumn(values, out var type, out var fieldMetadata);
        var fields = new[] { new Field("0", type, nullable: true, fieldMetadata) };
        var schema = BuildSchema(fields, ShapeVector, OrientationToWire(orientation));
        return WriteStream(schema, new[] { array }, values.Length);
    }

    /// <summary>
    /// Encode a single value as a scalar-shaped Arrow IPC stream. The
    /// resulting buffer is a 1×1 record batch carrying the value's inferred
    /// Arrow type.
    /// </summary>
    public static byte[] EncodeScalar(object? value)
    {
        var array = BuildColumn(new[] { value }, out var type, out var fieldMetadata);
        var fields = new[] { new Field("0", type, nullable: true, fieldMetadata) };
        var schema = BuildSchema(fields, ShapeScalar, orientation: null);
        return WriteStream(schema, new[] { array }, length: 1);
    }

    // -----------------------------------------------------------------------
    // Decode
    // -----------------------------------------------------------------------

    /// <summary>
    /// Peek at a buffer's shape metadata without materialising the data.
    /// Used by the host to decide spill geometry before allocating cells.
    /// </summary>
    public static (ArrowShape Shape, ArrowOrientation? Orientation) PeekShape(byte[] buffer)
    {
        if (buffer is null) throw new ArgumentNullException(nameof(buffer));
        using var ms = new MemoryStream(buffer, writable: false);
        using var reader = new ArrowStreamReader(ms);
        var schema = reader.Schema;
        return ReadShape(schema);
    }

    /// <summary>
    /// Decode any Arrow IPC stream produced by either <see cref="ArrowMarshal"/>
    /// or <c>arrow_io.py</c>. Returns one of:
    /// <list type="bullet">
    ///   <item><c>object?[,]</c> for shape=table</item>
    ///   <item><c>object?[]</c> for shape=vector</item>
    ///   <item>Unwrapped scalar value for shape=scalar</item>
    /// </list>
    /// Buffers without a <c>pyexcel-shape</c> metadata key decode as table,
    /// matching the Python side's defensive default for external Arrow
    /// writers.
    /// </summary>
    public static object? Decode(byte[] buffer)
    {
        if (buffer is null) throw new ArgumentNullException(nameof(buffer));
        using var ms = new MemoryStream(buffer, writable: false);
        using var reader = new ArrowStreamReader(ms);
        var schema = reader.Schema;
        var batches = ReadAllBatches(reader);
        try
        {
            var (shape, _) = ReadShape(schema);
            return shape switch
            {
                ArrowShape.Scalar => DecodeScalar(batches, schema),
                ArrowShape.Vector => DecodeVector(batches, schema),
                ArrowShape.Chart => DecodeChart(batches),
                ArrowShape.Image => DecodeImage(batches, schema),
                _ => DecodeTable(batches, schema),
            };
        }
        finally
        {
            foreach (var b in batches) b.Dispose();
        }
    }

    // -----------------------------------------------------------------------
    // Column type inference + array builders
    // -----------------------------------------------------------------------

    /// <summary>
    /// Build a typed Arrow array for one column of values. Sets
    /// <paramref name="arrowType"/> to the inferred Arrow type and
    /// <paramref name="fieldMetadata"/> to any field-level metadata the
    /// caller should attach (currently only the formula cell-type marker
    /// for columns whose values are all <see cref="Formula"/>), so the
    /// caller can construct the matching <see cref="Field"/>.
    /// </summary>
    private static IArrowArray BuildColumn(
        object?[] columnValues,
        out IArrowType arrowType,
        out IReadOnlyDictionary<string, string>? fieldMetadata)
    {
        var (canDouble, canBool, canFormula, sawNonNull) = ScanTypes(columnValues);

        fieldMetadata = null;

        if (!sawNonNull)
        {
            // All-null column. Use string + all nulls so pandas roundtrip is
            // well-defined (Arrow null type interacts poorly with pyarrow
            // → pandas in some versions).
            var b = new StringArray.Builder();
            for (var i = 0; i < columnValues.Length; i++) b.AppendNull();
            arrowType = StringType.Default;
            return b.Build();
        }

        if (canFormula)
        {
            // Formula column: encode as string, mark the field so the
            // decoder (and the eventual range writer) can recover the
            // intent.
            var b = new StringArray.Builder();
            foreach (var v in columnValues)
            {
                if (v is null) { b.AppendNull(); continue; }
                b.Append(((Formula)v).Text);
            }
            arrowType = StringType.Default;
            fieldMetadata = FormulaFieldMetadata;
            return b.Build();
        }

        if (canDouble)
        {
            var b = new DoubleArray.Builder();
            foreach (var v in columnValues)
            {
                if (v is null) { b.AppendNull(); continue; }
                b.Append(ToDouble(v));
            }
            arrowType = DoubleType.Default;
            return b.Build();
        }

        if (canBool)
        {
            var b = new BooleanArray.Builder();
            foreach (var v in columnValues)
            {
                if (v is null) { b.AppendNull(); continue; }
                b.Append((bool)v);
            }
            arrowType = BooleanType.Default;
            return b.Build();
        }

        // Mixed / string / unknown — fall back to string.
        {
            var b = new StringArray.Builder();
            foreach (var v in columnValues)
            {
                if (v is null) { b.AppendNull(); continue; }
                b.Append(v.ToString() ?? "");
            }
            arrowType = StringType.Default;
            return b.Build();
        }
    }

    /// <summary>
    /// Single pass to learn what's in the column. The result tells the
    /// builder which fast path applies (formula / double / bool) or that
    /// the column has to fall back to string. A column is the formula
    /// path only if every non-null entry is a <see cref="Formula"/>;
    /// mixing <see cref="Formula"/> with non-formula values has no clean
    /// wire representation today (the marker is per-column, not
    /// per-cell), so a mixed column falls through to the string path
    /// where each <see cref="Formula"/> stringifies to its <c>=…</c>
    /// text via <c>ToString()</c> — recognisable but not live.
    /// </summary>
    private static (bool canDouble, bool canBool, bool canFormula, bool sawNonNull) ScanTypes(object?[] values)
    {
        var allNumeric = true;
        var allBool = true;
        var allFormula = true;
        var sawNonNull = false;

        foreach (var v in values)
        {
            if (v is null) continue;
            sawNonNull = true;
            if (!IsNumeric(v)) allNumeric = false;
            if (v is not bool) allBool = false;
            if (v is not Formula) allFormula = false;
            if (!allNumeric && !allBool && !allFormula) break;
        }

        return (
            allNumeric && sawNonNull,
            allBool && sawNonNull,
            allFormula && sawNonNull,
            sawNonNull);
    }

    private static bool IsNumeric(object v) => v switch
    {
        double or float or decimal => true,
        int or long or short or sbyte => true,
        uint or ulong or ushort or byte => true,
        _ => false,
    };

    private static double ToDouble(object v) => v switch
    {
        double d => d,
        float f => f,
        decimal m => (double)m,
        int i => i,
        long l => l,
        short s => s,
        sbyte sb => sb,
        uint ui => ui,
        ulong ul => ul,
        ushort us => us,
        byte b => b,
        _ => throw new InvalidCastException(
            $"value of type {v.GetType().Name} reached the numeric path"),
    };

    // -----------------------------------------------------------------------
    // Schema + IPC writer/reader plumbing
    // -----------------------------------------------------------------------

    private static Schema BuildSchema(
        IReadOnlyList<Field> fields,
        string shape,
        string? orientation)
    {
        var metadata = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            [MetaShapeKey] = shape,
        };
        if (orientation is { })
            metadata[MetaOrientKey] = orientation;

        var builder = new Schema.Builder();
        foreach (var f in fields) builder.Field(f);
        builder.Metadata(metadata);
        return builder.Build();
    }

    private static byte[] WriteStream(Schema schema, IArrowArray[] arrays, int length)
    {
        var batch = new RecordBatch(schema, arrays, length);
        using var ms = new MemoryStream();
        using (var writer = new ArrowStreamWriter(ms, schema))
        {
            writer.WriteRecordBatch(batch);
            writer.WriteEnd();
        }
        return ms.ToArray();
    }

    private static (ArrowShape Shape, ArrowOrientation? Orientation) ReadShape(Schema schema)
    {
        var metadata = schema.Metadata;
        var shapeValue = metadata is { } m && m.TryGetValue(MetaShapeKey, out var sv) ? sv : ShapeTable;
        var orientationValue = metadata is { } m2 && m2.TryGetValue(MetaOrientKey, out var ov) ? ov : null;

        var shape = shapeValue switch
        {
            ShapeScalar => ArrowShape.Scalar,
            ShapeVector => ArrowShape.Vector,
            ShapeChart => ArrowShape.Chart,
            ShapeImage => ArrowShape.Image,
            _ => ArrowShape.Table,
        };
        var orientation = orientationValue switch
        {
            OrientRow => (ArrowOrientation?)ArrowOrientation.Row,
            OrientColumn => ArrowOrientation.Column,
            _ => null,
        };
        return (shape, orientation);
    }

    private static List<RecordBatch> ReadAllBatches(ArrowStreamReader reader)
    {
        var batches = new List<RecordBatch>();
        while (true)
        {
            var b = reader.ReadNextRecordBatch();
            if (b is null) break;
            batches.Add(b);
        }
        return batches;
    }

    // -----------------------------------------------------------------------
    // Decoders for each shape
    // -----------------------------------------------------------------------

    private static object?[,] DecodeTable(List<RecordBatch> batches, Schema schema)
    {
        // The producer side currently writes one batch per stream; we
        // concatenate if more arrive so external Arrow writers can be
        // multi-batch without surprising the host.
        var totalRows = 0;
        foreach (var b in batches) totalRows += b.Length;

        var cols = batches.Count > 0 ? batches[0].ColumnCount : 0;
        var result = new object?[totalRows, cols];

        // Pre-compute per-column formula-ness from the schema rather than
        // re-deriving it for each cell — keeps the hot loop tight.
        var formulaColumn = new bool[cols];
        for (var c = 0; c < cols; c++)
            formulaColumn[c] = c < schema.FieldsList.Count && IsFormulaField(schema.GetFieldByIndex(c));

        var rowOffset = 0;
        foreach (var batch in batches)
        {
            for (var c = 0; c < batch.ColumnCount; c++)
            {
                var arr = batch.Column(c);
                var isFormula = formulaColumn[c];
                for (var r = 0; r < batch.Length; r++)
                {
                    var cell = ReadCell(arr, r);
                    result[rowOffset + r, c] = isFormula
                        ? WrapFormula(cell)
                        : cell;
                }
            }
            rowOffset += batch.Length;
        }
        return result;
    }

    private static object?[] DecodeVector(List<RecordBatch> batches, Schema schema)
    {
        var totalRows = 0;
        foreach (var b in batches) totalRows += b.Length;
        var result = new object?[totalRows];

        var isFormula =
            schema.FieldsList.Count > 0 && IsFormulaField(schema.GetFieldByIndex(0));

        var rowOffset = 0;
        foreach (var batch in batches)
        {
            if (batch.ColumnCount == 0) continue;
            var arr = batch.Column(0);
            for (var r = 0; r < batch.Length; r++)
            {
                var cell = ReadCell(arr, r);
                result[rowOffset + r] = isFormula ? WrapFormula(cell) : cell;
            }
            rowOffset += batch.Length;
        }
        return result;
    }

    private static object? DecodeScalar(List<RecordBatch> batches, Schema schema)
    {
        if (batches.Count == 0 || batches[0].Length == 0 || batches[0].ColumnCount == 0)
            return null;
        var cell = ReadCell(batches[0].Column(0), 0);
        var isFormula =
            schema.FieldsList.Count > 0 && IsFormulaField(schema.GetFieldByIndex(0));
        return isFormula ? WrapFormula(cell) : cell;
    }

    /// <summary>Decode a chart-shaped buffer (1×1 string batch carrying
    /// the spec JSON) to a typed <see cref="ChartSpec"/>. A chart buffer
    /// with no cell is a wire-format violation, not a "None" — surface it.</summary>
    private static ChartSpec DecodeChart(List<RecordBatch> batches)
    {
        if (batches.Count == 0 || batches[0].Length == 0 || batches[0].ColumnCount == 0)
            throw new FormatException("chart-shaped buffer carries no spec cell");
        if (batches[0].Column(0) is not StringArray strings)
            throw new FormatException(
                $"chart-shaped buffer must carry a string column, got {batches[0].Column(0).GetType().Name}");
        var json = strings.GetString(0);
        if (json is null)
            throw new FormatException("chart-shaped buffer carries a null spec cell");
        return new ChartSpec(json);
    }

    /// <summary>Decode an image-shaped buffer (1×1 binary batch carrying
    /// the rendered bytes, format on the field metadata) to a typed
    /// <see cref="ChartImage"/>.</summary>
    private static ChartImage DecodeImage(List<RecordBatch> batches, Schema schema)
    {
        if (batches.Count == 0 || batches[0].Length == 0 || batches[0].ColumnCount == 0)
            throw new FormatException("image-shaped buffer carries no data cell");
        if (batches[0].Column(0) is not BinaryArray binary)
            throw new FormatException(
                $"image-shaped buffer must carry a binary column, got {batches[0].Column(0).GetType().Name}");
        if (binary.IsNull(0))
            throw new FormatException("image-shaped buffer carries a null data cell");
        var data = binary.GetBytes(0).ToArray();

        var format = ChartImage.FormatPng;
        if (schema.FieldsList.Count > 0)
        {
            var md = schema.GetFieldByIndex(0).Metadata;
            if (md is { } && md.TryGetValue(FieldMetaImageFormatKey, out var f))
                format = f;
        }
        return new ChartImage(data, format);
    }

    /// <summary>True iff <paramref name="field"/> carries the
    /// <c>pyexcel-cell-type = formula</c> field-level metadata marker.</summary>
    private static bool IsFormulaField(Field field)
    {
        var md = field.Metadata;
        return md is { } && md.TryGetValue(FieldMetaCellTypeKey, out var v)
            && string.Equals(v, CellTypeFormula, StringComparison.Ordinal);
    }

    /// <summary>Wrap a decoded cell as a <see cref="Formula"/>. Strings turn
    /// into formulas; nulls stay null; any other type is a wire-format
    /// violation we surface rather than silently lose.</summary>
    private static object? WrapFormula(object? cell) => cell switch
    {
        null => null,
        string s => new Formula(s),
        _ => throw new InvalidDataException(
            $"formula-marked field carried non-string cell of type {cell.GetType().Name}"),
    };

    /// <summary>
    /// Read one cell from an Arrow array, boxing it for transit through
    /// Excel's <c>object?[,]</c> grid. Returns null for null entries.
    ///
    /// <para>Numeric types are coerced to <see cref="double"/> on the way
    /// out — symmetric with the encode side, which sends every numeric
    /// (int / long / float / decimal / …) through the <see cref="DoubleArray"/>
    /// path. This matches Excel's "every number is a double" model, so a
    /// Python script that returns <c>42</c> spills as <c>42.0</c> rather
    /// than crashing the host's downstream <c>double</c> arithmetic.</para>
    ///
    /// <para>Arrow date / timestamp arrays decode to <see cref="DateTime"/>
    /// with <see cref="DateTimeKind.Unspecified"/>. The wire format does
    /// not carry timezone information (every Python-side path PyExcel uses
    /// produces naive <c>datetime</c> / <c>date</c> values, which pyarrow
    /// renders as <c>timestamp[us]</c> with no timezone), so attaching
    /// Utc / Local here would be a lie.</para>
    /// </summary>
    private static object? ReadCell(IArrowArray array, int index)
    {
        if (array.IsNull(index)) return null;
        // Use Values[index] for primitive numeric types instead of
        // GetValue(index), which returns Nullable<T>. We've already
        // ruled out the null case above, so the direct buffer access
        // is both safe and lets us cast to double without an unwrap dance.
        return array switch
        {
            DoubleArray d => (object?)d.Values[index],
            FloatArray f => (object?)(double)f.Values[index],
            Int64Array i64 => (object?)(double)i64.Values[index],
            Int32Array i32 => (object?)(double)i32.Values[index],
            Int16Array i16 => (object?)(double)i16.Values[index],
            Int8Array i8 => (object?)(double)i8.Values[index],
            UInt64Array u64 => (object?)(double)u64.Values[index],
            UInt32Array u32 => (object?)(double)u32.Values[index],
            UInt16Array u16 => (object?)(double)u16.Values[index],
            UInt8Array u8 => (object?)(double)u8.Values[index],
            BooleanArray b => b.GetValue(index),
            StringArray s => s.GetString(index),
            Date32Array d32 => Date32ToDateTime(d32.Values[index]),
            Date64Array d64 => Date64ToDateTime(d64.Values[index]),
            TimestampArray ts => TimestampToDateTime(
                ts.Values[index],
                ((TimestampType)ts.Data.DataType).Unit),
            _ => array.GetType().Name,  // last-ditch: surface the type name
        };
    }

    // -----------------------------------------------------------------------
    // Date / timestamp conversion helpers (Arrow → DateTime).
    //
    // The DateTime is always DateTimeKind.Unspecified — pyarrow's default
    // for a naive datetime is timestamp[us] with no timezone, and that is
    // every Python-side path PyExcel currently exercises. Attaching Utc /
    // Local here would invent information we don't have.
    // -----------------------------------------------------------------------

    private static readonly DateTime UnixEpoch =
        new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Unspecified);

    /// <summary>Date32 = days since 1970-01-01, no time component.</summary>
    private static DateTime Date32ToDateTime(int days) => UnixEpoch.AddDays(days);

    /// <summary>Date64 = milliseconds since 1970-01-01. Arrow's date64 is
    /// rare (pyarrow prefers date32 for <see cref="System.DateTime.Date"/>),
    /// but external writers may produce it, so we handle it.</summary>
    private static DateTime Date64ToDateTime(long ms) => UnixEpoch.AddMilliseconds(ms);

    /// <summary>
    /// Timestamp = a signed integer count of the unit (sec / ms / μs / ns)
    /// since the Unix epoch. .NET's DateTime ticks are 100 ns, so the
    /// conversion bottoms out at <see cref="DateTime.AddTicks"/>.
    /// </summary>
    private static DateTime TimestampToDateTime(long value, TimeUnit unit) => unit switch
    {
        TimeUnit.Second => UnixEpoch.AddSeconds(value),
        TimeUnit.Millisecond => UnixEpoch.AddMilliseconds(value),
        // 1 μs = 10 ticks (1 tick = 100 ns).
        TimeUnit.Microsecond => UnixEpoch.AddTicks(value * 10L),
        // 1 ns = 0.01 ticks. Integer-divide truncates sub-tick remainder —
        // DateTime can't represent sub-100-ns precision anyway.
        TimeUnit.Nanosecond => UnixEpoch.AddTicks(value / 100L),
        _ => throw new InvalidOperationException($"unknown TimeUnit: {unit}"),
    };

    private static string OrientationToWire(ArrowOrientation o) => o switch
    {
        ArrowOrientation.Row => OrientRow,
        ArrowOrientation.Column => OrientColumn,
        _ => throw new ArgumentOutOfRangeException(nameof(o)),
    };
}
