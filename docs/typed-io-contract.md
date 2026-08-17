# PyExcel typed I/O contract

> **Status.** Design spec — the single source of truth for how Excel ranges become
> Python values and how Python return values become cells. Implementation tracked
> in `TASKS.md`. Supersedes the shape-inference rules described in
> `README.md`'s authoring section.

## Why this exists

Before this contract the kernel *guessed*. The host marshalled a range to Arrow,
tagged it `pyexcel-shape = table | vector | scalar`, and the Python side rebuilt
whatever that tag implied. Three things went wrong with that:

1. **The guess was wrong in practice.** Every multi-cell range reached
   `ArrowMarshal.EncodeTable`, because COM hands back a 2-D array for a 10×1
   selection exactly as it does for a 10×3 one. A single column documented as a
   `list` arrived as a `DataFrame`.
2. **Names were collected and discarded.** The ribbon parses `{name}=Range`
   bindings; `RangeRunner` then used only the range text, and the RUN_REQUEST
   meta carried no names at all. `transform` was called positionally while every
   document promised a dict.
3. **Failures were silent or late.** A wrong guess surfaced as an
   `AttributeError` deep inside user code rather than as a message about the
   binding that was misconfigured.

The fix is to stop inferring. The user declares a type per binding, the
declaration travels on the wire, and the kernel constructs and validates against
it.

## Division of labour

The host owns **geometry**; the kernel owns **Python types**.

- The host reads the range, resolves `Auto` to a concrete type (it is the side
  that knows the range is 10×1), and sends the cells as a **raw R×C grid** plus
  the resolved type declaration.
- The kernel receives the grid and the declaration and **constructs** the
  declared type from it.

This split matters. `set`, `tuple`, `pandas.Series` and `numpy.ndarray` have no
Arrow shape and no C# equivalent worth modelling; building them host-side would
mean teaching `ArrowMarshal` about Python's type system. Sending a grid and a
label keeps each side doing what it is good at, and it means adding a type later
touches the kernel only.

Inputs are therefore **constructed** (the declaration says what to build) and
outputs are **validated** (the declaration says what must come back). That
asymmetry is deliberate: Excel gives us nothing but cells, so an input type is
an instruction; Python gives us a real object, so an output type is an assertion.

## The type set

| Type | Wire name | Python result |
| --- | --- | --- |
| Auto | `auto` | resolved host-side by range shape; never reaches the kernel |
| DataFrame | `dataframe` | `pandas.DataFrame` |
| Series | `series` | `pandas.Series` |
| List | `list` | `list` |
| Tuple | `tuple` | `tuple` |
| Set | `set` | `set` |
| Dict | `dict` | `dict` |
| NDArray | `ndarray` | `numpy.ndarray` |
| Scalar | `scalar` | `int` / `float` / `bool` / `str` / `Timestamp` / `None` |

## Range shape classification

For a range of R rows × C columns:

| Shape | Condition |
| --- | --- |
| Cell | R == 1 and C == 1 |
| Row | R == 1 and C > 1 |
| Column | R > 1 and C == 1 |
| Grid | R > 1 and C > 1 |

## Auto defaults

`Auto` is the default for every new binding and resolves host-side:

| Shape | Resolves to |
| --- | --- |
| Cell | Scalar |
| Row | List |
| Column | List |
| Grid | DataFrame |

## The naming rule

One rule, applied everywhere:

- **DataFrame and Series take their names from the first row.** This is the
  pandas convention (`read_excel` defaults to `header=0`) and the reading most
  people expect from a selection that includes its headings.
- **List, Tuple, Set, NDArray and Scalar consume every cell as data.** They have
  no name to carry, so nothing is skipped.
- **Dict** takes keys from the first column (2-column form) or from the header
  row (column-oriented form) — see the matrix.

## Input coercion matrix

How each declared type is built from each range shape. `v` denotes a cell value.

| Declared | Cell (1×1) | Row (1×C) | Column (R×1) | Grid (R×C) |
| --- | --- | --- | --- | --- |
| **DataFrame** | 1 column named `v`, 0 rows | C columns named by the row, 0 rows | 1 column named by the top cell, R−1 rows | C columns named by row 1, R−1 rows |
| **Series** | empty Series named `v` | Series named by the first cell, C−1 values | Series named by the top cell, R−1 values | **error** |
| **List** | `[v]` | `[v₁…v_C]` | `[v₁…v_R]` | nested: `[[row], [row], …]` |
| **Tuple** | `(v,)` | `(v₁…v_C)` | `(v₁…v_R)` | nested: `((row), (row), …)` |
| **Set** | `{v}` | distinct values | distinct values | distinct values, flattened |
| **NDArray** | shape `(1,)` | shape `(C,)` | shape `(R,)` | shape `(R, C)` |
| **Dict** | **error** | C == 2 → `{k: v}`; C > 2 → column-oriented, 0 values each | **error** | C == 2 → `{k: v}` per row; C > 2 → column-oriented, header row = keys |
| **Scalar** | `v` | **error** | **error** | **error** |

### Notes on the error cells

- **Series × Grid** — a Series is one-dimensional; a multi-column range has no
  unambiguous reading. Use DataFrame, or narrow the selection.
- **Dict × Cell** and **Dict × Column** — one column supplies keys with nothing
  to pair them against. Both documented layouts need at least 2 columns.
- **Scalar × anything but Cell** — a scalar is one value; picking the top-left
  and discarding the rest would hide a misconfigured range.

### Error messages

Every rejected combination names the binding, the range, its actual dimensions,
the declared type, and the way out. Format:

```
input 'Sales' (Sheet1!A1:C10, 10x3): declared type Series requires a single
row or column, but the range is 10x3. Use DataFrame, or select one row or column.
```

```
input 'Rates' (Sheet1!A1:A10, 10x1): declared type Dict needs at least 2
columns — 2 columns read as key -> value pairs, 3 or more as column-oriented
lists keyed by the header row.
```

Input failures surface as the existing `BadInput` job-error code; no new wire
error codes are introduced.

## Output enforcement

Output bindings declare what `transform` must return for that key.

- **Auto** (default) — today's behaviour. Whatever comes back is rendered by its
  shape: DataFrame/Series as a table, list/tuple/set/ndarray as a spill range,
  scalar into one cell, Plotly figure as a chart, Matplotlib figure as an image.
- **Any explicit type** — the returned value is checked with a strict type test
  before encoding. A mismatch fails the run with:

```
output 'ProcessedSales': declared type DataFrame, but transform() returned list.
```

No silent conversion happens on the output side. Declaring a type is opting in
to strictness; leaving it `Auto` keeps the loose behaviour that makes iterating
on a script comfortable. Output failures surface as the existing `BadReturnType`
code.

## Wire representation

`RUN_REQUEST` meta gains two arrays. Both are ordered to match the Arrow payload
order, and both carry **resolved** types — `auto` is resolved host-side and never
appears on the wire:

```json
{
  "run_id": "…",
  "script": "…",
  "function": "transform",
  "inputs":  [ {"name": "Sales",          "type": "dataframe"},
               {"name": "TaxRate",        "type": "scalar"} ],
  "outputs": [ {"name": "ProcessedSales", "type": "dataframe"},
               {"name": "TotalRevenue",   "type": "auto"} ]
}
```

`outputs` may carry `auto` because it means "do not enforce", which is a real
instruction rather than an unresolved one.

Each input payload is a plain R×C Arrow table of raw cells. The
`pyexcel-shape` metadata key stays on the buffer for the UDF path and for
backward compatibility, but when `inputs` is present in the meta the declared
type wins and the shape tag is ignored.

## Calling convention

When the request carries an `inputs` array the worker builds the documented dict
and makes one call:

```python
transform(inputs, **kwargs)      # inputs: Dict[str, Any]
```

When it does not — the `=PYRUN(…)` UDF path, direct kernel-client callers, and
existing tests — the worker keeps today's positional dispatch:

```python
transform(*args, **kwargs)
```

Both paths stay supported. The presence of `inputs` in the meta is the switch.

### Auto-naming

A binding with no `{name}=` prefix is named by its resolved type and its ordinal
within that type, matching what `README.md` already advertises: `df1`, `df2`, …
for DataFrame; `list1`, … for List/Tuple/Set/NDArray; `value1`, … for Scalar;
`series1`, … for Series; `dict1`, … for Dict.

## Backward compatibility

- **Saved workbooks.** State persisted under the current schema has no type
  field. Every binding loads as `Auto`, which resolves to the shape-derived
  default — so an existing workbook behaves as it did, except that a single
  row or column now correctly becomes a `list` instead of a `DataFrame`. That
  change is the point of the exercise, and it is the one behavioural difference
  an untouched workbook will see.
- **The `{name}=Range` grammar.** Untyped binding text parses exactly as today.
  The type is an optional addition to the syntax.
- **The UDF path.** `=PYRUN(…)` sends no `inputs` array and is unaffected.
- **The kernel wire format.** Only additive meta keys; `Framing.cs` and
  `framing.py` are untouched, so the byte-for-byte mirror invariant holds.
