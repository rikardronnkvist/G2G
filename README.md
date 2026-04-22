# GPO-to-GIT (PowerShell)

Production-ready PowerShell solution for tracking Active Directory Group Policy Object (GPO) changes in Git using version-aware detection.

## What this repository tracks

This repo tracks **solution artifacts** only:

- `gpos/<GPO-GUID>/report.xml` (for new/changed GPOs)
- `gpos/<GPO-GUID>/meta.json`
- `links/gpo-links.json` (domain + OU link scope/order/enforced/disabled)
- `wmi-filters/wmi-filters.json`
- `state/gpo-state.json`

This repo intentionally does **not** store full `Backup-GPO` output.

## How it works

`scripts/Invoke-GpoGitSync.ps1` is designed for Scheduled Task execution on a domain-joined Windows Server with RSAT tools.

Change detection includes:

1. **Version-based detection** (no XML diffing):
   - User DS/SYSVOL versions
   - Computer DS/SYSVOL versions
2. **Metadata changes**:
   - `GpoStatus`
   - WMI filter assignment
3. **Scope/impact changes**:
   - Domain + OU `gPLink` snapshot including link order, enforced and disabled flags
   - Link changes are treated as change events even if GPO versions are unchanged
4. **Deleted GPO detection**:
   - Removed from Active Directory => removed from `gpos/<guid>/`

Per run, it creates a summary, writes log output, creates HTML report (`reports/run-<timestamp>.html`), and commits/pushes tracked artifacts to Git when changes are present.

## Requirements

- Windows Server (domain joined)
- PowerShell 5.1+ (PowerShell 7 also supported when RSAT cmdlets are available)
- RSAT modules:
  - `GroupPolicy`
  - `ActiveDirectory`
- Git CLI installed and available in PATH
- AD permissions to read GPOs, links, and WMI filter containers
- Git repo clone available locally and configured with `origin`

## Repository layout

```text
scripts/
  Invoke-GpoGitSync.ps1
  Invoke-GpoBackupToFork.ps1
gpos/
links/
wmi-filters/
state/
logs/          (ignored)
reports/       (ignored)
temp/          (ignored)
```

## Setup

1. Clone this private repository to your server:
   ```powershell
   git clone git@github.com:<org-or-user>/GPO-to-GIT.git C:\Ops\GPO-to-GIT
   ```
2. Validate prerequisites:
   ```powershell
   Get-Module -ListAvailable GroupPolicy,ActiveDirectory
   git --version
   ```
3. Perform a dry run:
   ```powershell
   powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\Ops\GPO-to-GIT\scripts\Invoke-GpoGitSync.ps1 -RepoPath C:\Ops\GPO-to-GIT -DryRun
   ```
4. Perform first real run:
   ```powershell
   powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\Ops\GPO-to-GIT\scripts\Invoke-GpoGitSync.ps1 -RepoPath C:\Ops\GPO-to-GIT
   ```

## Scheduled Task (recommended)

Sample action command:

```text
Program/script: powershell.exe
Arguments: -NoProfile -ExecutionPolicy Bypass -File "C:\Ops\GPO-to-GIT\scripts\Invoke-GpoGitSync.ps1" -RepoPath "C:\Ops\GPO-to-GIT"
Start in: C:\Ops\GPO-to-GIT
```

Recommended task settings:

- Run whether user is logged on or not
- Run with highest privileges (if your environment requires)
- Configure for your server OS version
- Trigger: every 15 minutes (or your preferred interval)
- **Task Scheduler > Settings > If the task is already running: "Do not start a new instance"**
- Stop task if it runs unexpectedly long (optional safety)

## Parameters (Invoke-GpoGitSync.ps1)

- `-RepoPath <path>`: Local Git repository root
- `-Branch <name>`: Branch to commit/push (defaults to current branch)
- `-StateFilePath <path>`: State file path (default `state/gpo-state.json`)
- `-LogDirectory <name>`: Log directory (default `logs`)
- `-TeamsWebhookUrl <url>`: Optional Teams incoming webhook URL
- `-TeamsTopCount <int>`: Number of item lines in webhook summary (default 10)
- `-DryRun`: No solution artifacts written, no delete, no commit/push

Examples:

```powershell
# Standard run
.\scripts\Invoke-GpoGitSync.ps1 -RepoPath C:\Ops\GPO-to-GIT

# Specific branch + Teams notification
.\scripts\Invoke-GpoGitSync.ps1 -RepoPath C:\Ops\GPO-to-GIT -Branch main -TeamsWebhookUrl "https://..."

# Dry run validation
.\scripts\Invoke-GpoGitSync.ps1 -RepoPath C:\Ops\GPO-to-GIT -DryRun
```

## Git authentication (documentation only)

Do not place credentials in scripts.

Recommended approach:

1. Create a dedicated service account for the scheduled task.
2. Generate an SSH key for that account (Ed25519 recommended).
3. Add the public key as a deploy key or user key with least required scope.
4. Ensure `origin` uses SSH URL (`git@github.com:...`).
5. Validate non-interactive access:
   ```powershell
   ssh -T git@github.com
   git -C C:\Ops\GPO-to-GIT pull --ff-only
   ```

## Teams webhook (optional)

- Provide webhook URL via `-TeamsWebhookUrl`.
- Webhook failures are **non-fatal** and logged as warnings.

## Exit codes

- `0` = success, no changes
- `1` = success, changes committed/pushed (or would have in dry run)
- `2` = soft failure (exports succeeded, push failed)
- `10` = prerequisites missing (`GroupPolicy`/required modules)
- `11` = AD query failed / permission issue
- `12` = filesystem/repo/git path issue or unhandled runtime error

## Troubleshooting

- **Module missing (code 10):** install RSAT Group Policy and AD tools.
- **AD query fails (code 11):** verify account read permissions and domain connectivity.
- **Push fails (code 2):** verify SSH key, repo permissions, branch protection rules.
- **No changes committed:** check if only ignored files changed (`logs/`, `reports/`, `temp/`).
- **Concurrent runs:** ensure Task Scheduler is set to **Do not start a new instance**.

## Optional Backups to Forked Repo

Use `scripts/Invoke-GpoBackupToFork.ps1` if you want full `Backup-GPO` artifacts in a **separate** repository (recommended: a private fork dedicated to backups).

### Backup workflow

1. Create/fork a separate private repo for backups.
2. Clone backup repo to a separate path (not inside this repo).
3. Run backup script:

```powershell
# Explicit changed GUID list from sync process
.\scripts\Invoke-GpoBackupToFork.ps1 -BackupRepoPath C:\Ops\GPO-Backups -ChangedGpoGuids "11111111-1111-1111-1111-111111111111","22222222-2222-2222-2222-222222222222"

# Or state-file method (backs up GUIDs from state data)
.\scripts\Invoke-GpoBackupToFork.ps1 -BackupRepoPath C:\Ops\GPO-Backups -MainRepoPath C:\Ops\GPO-to-GIT
```

### Important backup constraints

- Backups are written **only** under `-BackupRepoPath`.
- Backup repo path must differ from main repo path and must not be nested within it.
- No credentials/secrets in script code; use SSH configuration for backup repo authentication as documented above.
