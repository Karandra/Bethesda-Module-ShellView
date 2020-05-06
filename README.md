# Bethesda Module ShellView
Shell property provider that allows to see some basic info about Bethesda module files on **Details** tab in the system file **Properties** dialog.

### Supported formats
- Morrowind
- Oblivion
- Skyrim LE
- Skyrim SE
- Fallout 3
- Fallout: New Vegas
- Fallout 4

### Supported properties
- Author
- Description
- Form version (where applicable)
- Flags (ESM, ESL, localized, etc)
- Master files list.

# Installation
Run `cmd.exe` as an administrator and use following comments. Use full paths to `regsvr32.exe` and the DLL if needed.

**Install**:
```ps
regsvr32 "Bethesda Module ShellView.dll"
```

**Uninstall**:
```ps
regsvr32 "Bethesda Module ShellView.dll" /u
```
