---
name: vb6-build
description: Compile and register the VB6 projects in this repo (Interface.exe, ComSet.dll, HostForms_SA.dll, HostForms_RE.dll), including build order, output paths, binary compatibility, regsvr32 and missing-reference problems. Use when asked to build, compile, make, package or register components, or when a "Can't create object"/error 429/missing reference occurs.
---

# Building the Lab-I components

VB6 is installed at `C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE`.
All components are **32-bit**. Use `C:\Windows\SysWOW64\regsvr32.exe` to register them.

## Projects and outputs

| Project file | Output | Type | Path32 (output dir) |
|---|---|---|---|
| `ComSet/ComSet.vbp` | ComSet.dll | ActiveX DLL | `..\New Lab-I System v2` |
| `HostForms_SA v2/HostForms_SA.vbp` | HostForms_SA.dll | ActiveX DLL | `..\New Lab-I System v2` |
| `HostFomrs_RE/HostForms_RE.vbp` | HostForms_RE.dll | ActiveX DLL | its own folder |
| `New Lab-I System v2/HostInterface.vbp` | Interface.exe | Standard EXE | its own folder |

## Build order

Build the DLLs first, then the EXE. Interface.exe late-binds everything with
`CreateObject`, so it has no compile-time reference to the DLLs. The DLLs only need
to be registered at run time.

```powershell
$vb6 = "C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE"
& $vb6 /make "ComSet\ComSet.vbp" /out build.log
& $vb6 /make "HostForms_SA v2\HostForms_SA.vbp" /out build.log
& $vb6 /make "HostFomrs_RE\HostForms_RE.vbp" /out build.log
& $vb6 /make "New Lab-I System v2\HostInterface.vbp" /out build.log
Get-Content build.log
```

`/make` returns no console output. Always read the `/out` log file. A failed build
can still leave an older binary in place, so compare file timestamps too.

## Gotchas
- **Binary compatibility.** ComSet uses `CompatibleMode=1`. Keep the public interface
  of `ComSettings` and `ClassForms` stable. Changing a public method signature breaks
  late-bound callers in installed sites.
- **References outside this repo.** HostForms_SA pulls in
  `..\..\LIS Main Application\vbGraph\*` and
  `..\..\..\..\Laboratory Management System\Components\craxdrt.dll`. HostForms_RE
  references Crystal Reports 8 under `C:\Program Files (x86)\Seagate Software\`.
  If these paths are missing, the build fails with "missing reference" or "file not
  found".
- **Required OCXs** (must be registered on the build machine): MSCOMM32, MSCOMCTL,
  MSCOMCT2, MSHFLXGD, TABCTL32, COMDLG32, RICHTX32, MSWINSCK, crviewer.
- **Developer-machine project variants.** `*-L-WRFABROA.vbp` and `*-Luis-PC.vbp` hold
  other developers' paths. Build from the plain `.vbp`.
- **Version numbers.** `AutoIncrementVer=1` bumps `RevisionVer` in the `.vbp` on
  every make. Expect that diff.
- **Run-time plugins.** Run-time plugins (`Plugins/*.dll`), `SearchCodes.ClassSearch`
  and `BCSET.Barcode` are separate COM servers. Their source is not in this repo.

## Register

Run as administrator:

```powershell
$rs = "C:\Windows\SysWOW64\regsvr32.exe"
& $rs /s "New Lab-I System v2\ComSet.dll"
& $rs /s "New Lab-I System v2\HostForms_SA.dll"
Get-ChildItem "New Lab-I System v2\Plugins\*.dll" | ForEach-Object { & $rs /s $_.FullName }
```

## When you cannot build
If VB6 or a reference is not available in this environment, say so plainly. Do not
claim a change compiles. Instead, list the procedures you edited so the developer
can compile them in the IDE.
