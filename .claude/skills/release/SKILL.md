---
name: release
description: Publish a Lab-I release on GitHub (strugz/Interface-Main-Application) so support staff can download Interface.exe, HostForms_SA.dll and ComSet.dll from the Releases page instead of getting files over Viber/chat. Use when the user says "release", "make a release", "publish", "release HostForms_SA", "new version for support", or "upload the build". A plain "build X" is NOT a release - use the vb6-build skill for that.
---

# Releasing Lab-I

Support staff download builds from
https://github.com/strugz/Interface-Main-Application/releases.

- **Build** is local only.
- **Release** is build + commit the version bumps + tag + push + publish.

Do a release only when the user asks for one. The repo is private, so support
staff need to be added as collaborators to see it. Never make the repo public.

## How a release runs

Everything is in `tools/release.ps1`. The GitHub Actions workflow
`.github/workflows/release.yml` runs that same script on the self-hosted build
PC: runner `lab-i-vb6-build`, label `vb6`, in
`%USERPROFILE%\actions-runner\lab-i`.

The script does this:
1. Requires a clean `main` that is up to date with `origin/main`.
2. Downloads the last released `ComSet.dll` if `ComSet\ComSet.dll` is missing. It
   is the binary-compatibility reference and is not in git.
3. Builds the chosen projects with `tools\build.ps1`.
4. Restores the COM registrations the build moved. It runs `regsvr32` elevated,
   then checks each one.
5. Tags the release:
   - `v<Interface version>` when Interface.exe is included;
   - otherwise `<Component>-v<version>`, for example `HostForms_SA-v2.0.156`.

   The script stops if the tag already exists.
6. Commits the `.vbp` bumps (`Build ...`), tags, pushes, and creates the release.
   The release notes are written for support staff: version and hash table,
   changes since the last tag, testing status, install steps, and what's not
   included.
7. Checks that each uploaded file's size matches the local build.

## Default: run the workflow

```bash
gh workflow run release.yml --repo strugz/Interface-Main-Application -f projects="ComSet,HostForms_SA,Interface"
gh run list --repo strugz/Interface-Main-Application --workflow release.yml --limit 1
gh run watch <run id> --repo strugz/Interface-Main-Application --exit-status
```

- **Choosing components.** "Release HostForms_SA" means
  `-f projects="HostForms_SA"`. Add `HostForms_RE` only when asked.
- **Dry run.** To check a build without publishing, add `-f dry_run=true`.
- **Test release.** For "test release", "sample release" or "+test", add
  `-f test=true`. That publishes a GitHub **pre-release** tagged `v<version>-test`
  with a "do not install at sites" banner, never marked Latest. Locally, use
  `.\tools\release.ps1 -Test`.
- **Runner check.** Before triggering, confirm the runner is online:
  `gh api repos/strugz/Interface-Main-Application/actions/runners -q '.runners[] | .name + " " + .status'`.
  If it is offline, tell the user to start
  `%USERPROFILE%\actions-runner\lab-i\run.cmd` on the build PC, or use the local
  fallback below.
- **Source changes.** The workflow builds what is on GitHub `main`. Push the
  source changes first.

## Fallback: run it locally on the build PC

```powershell
.\tools\release.ps1 -DryRun                     # check first
.\tools\release.ps1                             # ComSet, HostForms_SA, Interface
.\tools\release.ps1 -Projects HostForms_SA      # one component
```

## After it finishes

- Give the user the release URL. The script prints it as `Published: ...`, and the
  workflow summary shows it too.
- List the versions, and say what still needs testing at a site.
- If the log shows `WARNING: Could not restore COM registration`, tell the user to
  run the printed `regsvr32` command as admin.

## Never

- Attach databases, logs, `Passwords.txt`, `.reg` exports, or anything from
  `AppData\Data`, `ComTransLog`, `HostSpooler` or `Received`. These hold patient
  data or secrets.
- Release instrument plugins from here. They come from `D:\1 Project\LIS Driver`.
- Move, delete or reuse an existing tag.
