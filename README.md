# Bethesda Module ShellView
Shell property provider that allows to see some basic info about Bethesda module files on **Details** tab in the system file **Properties** dialog. The DLL registers itself to handle `.esp`, `.esm`, `.esl` and `.esu` file extensions with `BethesdaModule.Metadata` ProgID.

Property fields are marked as editable but actual file data edit is not supported. If you click **Apply** with any changed made to the new properties you will get **0x80004001: Not implemented** error. This is normal, property fields are made editable so you can copy metadata from there. If they're read-only there is no way to copy text.

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

# Build
Required [KxFramework](https://github.com/KerberX/KxFramework). You can easily get it using **VCPkg** package manager and provided portfile to build the **KxFramework** itself.

# Screenshots
![1](https://cdn.discordapp.com/attachments/511613474112274493/707517902047150080/1.png)
![2](https://cdn.discordapp.com/attachments/511613474112274493/707517906224939068/2.png)
![3](https://cdn.discordapp.com/attachments/511613474112274493/707517904618258452/3.png)
