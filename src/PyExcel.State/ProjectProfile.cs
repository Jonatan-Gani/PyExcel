using System;

namespace PyExcel.State;

/// <summary>
/// The "profile of the workbook" — the single authoritative record of a PyExcel
/// project, embedded in the workbook itself as a human-readable
/// <c>CustomXMLPart</c> (see <c>WorkbookStatePersister</c>). It pairs the user's
/// <see cref="State"/> (enabled flag, actions, field bindings, selected
/// script/sheet) with <see cref="Metadata"/> describing the environment it was
/// created in, so the project's identity is intrinsic to the document: move,
/// rename, email, or round-trip the workbook through the cloud and everything
/// needed to recognise, load, run, and debug it travels inside the file — no
/// sidecar file and no per-machine app-data to keep in step.
/// </summary>
public sealed record ProjectProfile(WorkbookState State, ProjectMetadata Metadata);

/// <summary>
/// Environment / provenance metadata stored alongside the workbook state in the
/// project profile. Everything is optional (older or hand-edited profiles may
/// omit fields), and it's purely descriptive — the add-in never depends on it
/// to function, but it's what makes a project debuggable and reusable: which
/// Python, which machine, which add-in version produced it, and when.
/// </summary>
public sealed record ProjectMetadata(
    string? GeneratedBy = null,
    DateTimeOffset? CreatedUtc = null,
    DateTimeOffset? UpdatedUtc = null,
    string? Os = null,
    string? Machine = null,
    int? ProcessBits = null,
    string? Clr = null,
    string? PythonPath = null,
    string? PythonVersion = null,
    string? WorkbookName = null,
    string? WorkbookPath = null,
    string? ProjectDir = null);
