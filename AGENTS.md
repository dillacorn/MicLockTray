# Repository instructions for MicLockTray

Applies to the entire repository.

## Project and behavior

- MicLockTray is a Windows .NET 8 / WinForms, x64, single-executable tray utility that enforces a user-selected **default recording-device microphone input** volume through Windows Core Audio callbacks. It is not an output-volume controller.
- Preserve event-driven enforcement, pause/resume, selection of the current default capture endpoint, 1–100% target range, single-instance behavior where implemented, and per-user autorun controls. Do not add polling, a service, administrator requirements, telemetry, or unrequested third-party utilities.
- `MicLockTray/Installer.cs` owns per-user startup registration; localization is embedded (see `MicLockTray/UiText.cs`). Preserve existing user settings and translation fallback when changing UI text.
- `MicLockTray/MicLockTray.csproj` is the Windows build source. The `.github/workflows/build.yml` workflow checks a Release build and win-x64 self-contained single-file publish output; **a passing build is not a published GitHub Release**.

## User-facing release notes

- Start each release with a short plain-language introduction explaining why the update matters.
- Follow with a compact **What's new** or **Changes** section containing only changes users will notice or need to act on. Prefer 3–6 concise bullets.
- Highlight user-facing fixes, settings changes, and any behavior or compatibility changes. Include upgrade instructions only when actually needed.
- Keep titles and descriptions readable and direct; avoid walls of text, internal implementation details, GitHub Actions logs, commit lists, test matrices, hashes, redundant headings, repeated boilerplate, and exhaustive verification narratives.
- Put necessary technical provenance or artifact verification in a separate developer-oriented file or concise collapsible details, not the primary release message. Do not claim testing or publication that has not happened.
- Preserve clear download/install links or directions when releasing binaries.

## Merge and release authorization

Operate autonomously within the user's approved scope; do not add needless approval checkpoints.

- Implementation, investigation, fixes, tests, branches, and ordinary review PRs are not implicit authorization to merge, tag, or publish.
- A clear request to merge, commit/apply directly to `main`, tag, or publish authorizes the named operation for the specified work. An approved release includes its necessary scoped merge unless the user specifies a different target or restriction; do not ask for the same approval twice.
- An authorized but **unfinished** merge/release remains authorized through related small fixes, adjustments, tests, and interruptions. If the user pauses the release to correct a defect, an unambiguous "good to go", "working, continue", "looks good", or equivalent can resume the **same previously approved** operation. A pause ("wait"/"hold off") must be respected until the user signals continuation.
- The approval is **consumed when that operation finishes**. After a release has been published, fixing another issue or hearing "perfect"/"that works" does not authorize a follow-up merge or release. A new release requires a new authorization, even for a fix to the previous one.
- Stop and clarify when authorization is genuinely unclear, explicitly withdrawn, or when the scope, release target, or version materially changes. Do not invent an approval merely because CI passed, a prior release happened, or publishing would be convenient.
- A user can expressly approve releasing without manual runtime testing. Run reasonable available static/CI checks and report what was actually validated; do not misrepresent missing runtime testing or ignore significant failing checks.
- Preparing release notes, identifying a proposed version, writing code, and opening a feature PR are not themselves publication. Creating or triggering release-publishing workflows, one-use GitHub Actions release bridges, Git tags, Releases, uploads, or other publication requires an active approval for that specific release.
- Preserve any stricter project-specific permissions for deployments, package publication, destructive changes, and external infrastructure. A GitHub Release authorization does not inherently approve deployment, package publication, or replacing production data.

## Validation and delivery

- Validate changes with `dotnet build MicLockTray/MicLockTray.csproj --configuration Release` and, for delivery changes, `dotnet publish MicLockTray/MicLockTray.csproj --configuration Release --runtime win-x64 --self-contained true -p:PublishSingleFile=true`, when a suitable SDK/runner is available.
- Distinguish source builds, CI artifacts, attached release executables, and published GitHub Releases. Do not upload or replace binaries, change a tag, or publish a release without authorization.
- State plainly when a real Windows microphone/autorun runtime check has not been performed. Keep all behavior scoped to the current user rather than machine-wide registry or audio policies.
