using System;

namespace PyExcel.State;

/// <summary>
/// The on-disk "profile of the workbook" — the single authoritative record of a
/// PyExcel project, written as a human-readable XML file in the project folder
/// (next to the workbook, the venv, the kernel, and userScripts). It pairs the
/// user's <see cref="State"/> (enabled flag, actions, field bindings, selected
/// script/sheet) with <see cref="Metadata"/> describing the environment it was
/// created in, so the project is self-contained: copy the folder and everything
/// needed to load, run, debug, or update it travels with it — no per-machine
/// app-data to also copy.
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
