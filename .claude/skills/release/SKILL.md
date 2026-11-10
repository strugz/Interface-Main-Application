---
name: release
description: Publish a Lab-I release on GitHub (strugz/Interface-Main-Application) so support staff can download Interface.exe, HostForms_SA.dll and ComSet.dll from the Releases page instead of getting files over Viber/chat. Use when the user says "release", "make a release", "publish", "release HostForms_SA", "new version for support", or "upload the build". A plain "build X" is NOT a release - use the vb6-build skill for that.
---

# Releasing Lab-I

A release is how support staff get new builds. They download them from
https://github.com/strugz/Interface-Main-Application/releases. **Build = local
only. Release = build + commit + tag + push + publish.** Do a release only when the
user asks for one.

The repo is private, so support staff need a GitHub account added as a
collaborator to see the Releases page. If someone can't see it, say so. Don't make
the repo public.

## 1. Check the starting point
```bash
git status --short          # must be clean: commit or stash source changes first
git switch main
git pull --ff-only
git describe --tags --abbrev=0   # previous release tag, used for the notes
```
- If there are uncommitted changes to source files, stop. Ask whether to commit
  them as part of this release or leave them out.
- Never release from a dirty tree. The binaries must match the tagged source.

## 2. Pick the components
- The default is `ComSet`, `HostForms_SA` and `Interface`. If the user names some,
  build only those. For example, "release HostForms_SA" means `-Projects HostForms_SA`.
- Add `HostForms_RE` only when the user asks for it.
- Instrument plugins (`Plugins\*.dll`) are **not** released from here. They come
  from `D:\1 Project\LIS Driver`.

## 3. Build
```powershell
.\tools\build.ps1 -Projects ComSet, HostForms_SA, Interface
```
- Every project must print `OK  v<version>`. If any fails, stop and show the log
  lines. Never publish an old binary.
- Record each built version from the output. VB6 has now bumped `RevisionVer` in
  each `.vbp`.

## 4. Commit and tag
```bash
git add -A -- '*.vbp'
git commit -m "Build Interface 3.1.95, HostForms_SA 2.0.156, ComSet 1.0.38"  # actual built versions
git tag -a v3.1.95 -m "Lab-I Interface 3.1.95 ..."                           # one line per component
git push origin main
git push origin v3.1.95
```

**Tag name:**
- If Interface.exe is in the release, the tag is `v<Interface version>`, for
  example `v3.1.95`.
- Otherwise use `<Component>-v<version>`, for example `HostForms_SA-v2.0.156` or
  `ComSet-v1.0.38`.
- If the tag already exists, stop. Don't move or overwrite tags.

End the commit message with the attribution lines from the system reminder.

## 5. Release notes (written for support staff)
Write the notes to a temp file. They contain:

1. **Version table.** One row per file: name, version, and SHA-256 hash. Get the
   hash with `(Get-FileHash <file> -Algorithm SHA256).Hash`.
2. **What changed.** Plain-language bullets taken from
   `git log <previous tag>..HEAD --no-merges`. Leave out version-bump-only
   commits. Describe each change the way a lab user would notice it.
3. **Testing status.** Say honestly what was tested and what wasn't. For example:
   "Built and installed on the dev PC; not yet tested with a live analyzer."
4. **Install / upgrade steps:**
   1. Close Lab-I.
   2. Back up the current files.
   3. Copy the new files into the Lab-I folder.
   4. From an admin prompt, register every DLL with
      `C:\Windows\SysWOW64\regsvr32 "<Lab-I folder>\<dll>"`.
   5. Check Help → About for the new version.
5. **Not included:** instrument plugins (LIS Driver), HostForms_RE (unless it was
   built), runtime prerequisites (Crystal Reports 8, OCXs, ADO, DAO), databases.

## 6. Publish
Pass only the files that were built:
```bash
gh release create v3.1.95 "New Lab-I System v2/Interface.exe" "New Lab-I System v2/HostForms_SA.dll" "New Lab-I System v2/ComSet.dll" --repo strugz/Interface-Main-Application --title "Lab-I Interface 3.1.95" --notes-file <notes.md> --verify-tag
```
Output locations:
- `Interface.exe`, `ComSet.dll` and `HostForms_SA.dll` are in `New Lab-I System v2\`.
- `HostForms_RE.dll` is in `HostFomrs_RE\`.

**Never attach:** `.mdb` files, logs, `Passwords.txt`, `.reg` exports, or anything
from `AppData\Data`, `ComTransLog`, `HostSpooler` or `Received`. These hold patient
data or secrets.

## 7. Verify and report
```bash
gh release view <tag> --repo strugz/Interface-Main-Application --json url,assets -q '.url, (.assets[] | .name+" "+(.size|tostring))'
```
- Each asset's size must equal the local file's size.
- Give the user the release link. Support staff use that link instead of chat
  attachments.
- List the versions in the release, and say what still needs testing at a site.
