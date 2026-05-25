namespace PyExcel.Excel;

/// <summary>
/// The high-level shape of a value flowing across the kernel boundary.
/// Mirrors the Python <c>pyexcel.kernel.arrow_io.Shape</c> enum byte-for-byte
/// on the wire (the values are serialised as the Arrow schema-metadata
/// string under the <c>pyexcel-shape</c> key).
/// </summary>
public enum ArrowShape : byte
{
    /// <summary>2-D table — encoded as a multi-column Arrow record batch,
    /// decoded as <c>object?[,]</c>.</summary>
    Table,

    /// <summary>1-D vector — encoded as a single-column batch, decoded as
    /// <c>object?[]</c>. Carries an <see cref="ArrowOrientation"/> hint so
    /// the host can spill row-wise vs. column-wise.</summary>
    Vector,

    /// <summary>Single value — encoded as a 1×1 batch, decoded as a plain
    /// boxed scalar (<see cref="double"/>, <see cref="string"/>,
    /// <see cref="bool"/>, or <c>null</c>).</summary>
    Scalar,
}

/// <summary>Vector orientation hint — purely advisory metadata that lets
/// the host decide whether a 1-D result should spill across a row or down
/// a column. Mirrors <c>pyexcel.kernel.arrow_io.Orientation</c>.</summary>
public enum ArrowOrientation : byte
{
    Row,
    Column,
}
