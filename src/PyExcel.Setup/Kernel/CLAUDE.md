# Kernel

## Macro
Extracts the embedded Python kernel package and requirements from assembly
resources onto disk so the project venv can run it.

## Files
### KernelResourceExtractor.cs
Extracts the embedded `pyexcel` package and `requirements.txt` from assembly manifest
resources, preserving logical-name paths and skipping files whose content is unchanged.
Inputs: a target directory and an optional source assembly. Output: an `ExtractionResult`
(target dir, files written, files skipped); a missing kernel resource is flagged as a
build regression.

## Subdirectories
None.
