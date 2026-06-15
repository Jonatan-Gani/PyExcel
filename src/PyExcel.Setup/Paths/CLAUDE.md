# Paths

## Macro
Pure path normalisation and classification used by Setup to decide how a project
directory should be treated.

## Files
### ProjectPathResolver.cs
Normalises and classifies a project path as Local, UNC (`\\server\share`), or
OneDrive-synced (matched against OneDrive environment variables); pure logic with no disk
I/O. Inputs: a raw path string. Output: a `ProjectPathInfo` (original + normalised path,
`isUnc`, `isOneDriveSynced`, `oneDriveRoot`).

## Subdirectories
None.
