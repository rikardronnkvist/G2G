# G2G - GPO-to-GIT

PowerShell solution for tracking Active Directory Group Policy Object (GPO) changes in Git using version-aware detection.

## How it works

`./Invoke-GpoGitSync.ps1` is designed for Scheduled Task execution on a domain-joined Windows Server with RSAT tools. The script scans for changes in in GPO's, WMI-Filters and links.

Per run, it writes console log output, creates a Markdown report when changes are present, and commits/pushes tracked artifacts to Git when changes are present.

If you only want to store artifacts locally you can use the `-DisableGIT` parameter.

If you want to track every domain in the current Active Directory forest, use `-CompleteForest` (the legacy spelling `-CompleteForrest` is also accepted).

In `-CompleteForest` mode, if a child domain contains a GPO with the same GUID as a parent domain GPO, the child-domain folder keeps `links.md` but the `README.md` points to the parent-domain GPO folder instead of exporting duplicate HTML/XML reports. This does not apply to the well-known default-policy GUIDs that legitimately exist in multiple domains.

## Why G2G instead of AGPM

With Advanced Group Policy Management (AGPM) now deprecated and no longer actively developed, organizations still need a reliable way to detect, audit, and understand changes to Group Policy.

This script provides a simple and transparent alternative by shifting GPO change tracking to Git-based version control, using data that already exists in Active Directory. Instead of relying on a proprietary workflow or additional infrastructure, it:

- Detects **what actually changed** (new, modified, deleted GPOs, links, and WMI filters) using version-aware comparison.
- Stores changes in **open, readable formats** (Markdown, HTML, XML) that are easy to review without special tools.
- Creates a **clear audit trail** with timestamps, diffs, and commit history, suitable for troubleshooting, audits, and compliance.
- Requires only **read access** to Active Directory and does not interfere with existing GPO management processes.
- Fits naturally into modern operational practices such as automation, peer review, and change transparency.

While it does not attempt to fully replace AGPM’s approval and editing workflow, it effectively covers the most critical requirement after AGPM: **knowing exactly when, where, and how Group Policy changed**, using tooling that is simple, vendor‑neutral, and future‑proof.

![GPO-to-GIT](./g2g.jpg?raw=true)
[![FOSSA Status](https://app.fossa.com/api/projects/git%2Bgithub.com%2Frikardronnkvist%2FG2G.svg?type=shield)](https://app.fossa.com/projects/git%2Bgithub.com%2Frikardronnkvist%2FG2G?ref=badge_shield)

# Requirements

- Domain joined Windows Server
- PowerShell 5.1+
- RSAT modules:
  - `GroupPolicy`
  - `ActiveDirectory`
- Git CLI installed and available in PATH
- AD permissions to **read** GPOs, links, and WMI filter containers
- Git repo clone available locally and configured with `origin`

# Setup

* Copy the script `Invoke-GpoGitSync.ps1` to your server
* Validate prerequisites:
   ```powershell
   Get-Module -ListAvailable GroupPolicy,ActiveDirectory
   git --version
   ```
* Create a empty Git repo (init with a simple README.md file)
* Create service account (in example `svcG2G`)
   * Grant the service account `Logon as a Batch Job`
   * Temporarily grant the service account access to logon
* Optional: Grant Read Access to the service account
   * To all GPO's
      ```powershell
      Get-GPO -All | ForEach-Object {
         Set-GPPermission `
            -Guid $_.Id `
            -TargetName "svcG2G" `
            -TargetType User `
            -PermissionLevel GpoRead
      }
      ```
   * To WMI filters
      * In Group Policy Management
      * Forest - Domains - your domain - WMI Filters - Delegation
      * Add your service account with Read access
* Start Powershell as the service account
* Clone the repo
   ```powershell
   New-Item -Path "C:\Ops\GPO-Backup" -ItemType Directory
   git clone https://github.com/myaccount/GPO-Backup.git C:\Ops\GPO-Backup
   ```
* Perform a dry run
   ```powershell
   . C:\Ops\Scripts\Invoke-GpoGitSync.ps1 -RepoPath "C:\Ops\GPO-Backup" -DryRun
   ```
* Configure Git authentication
   * Generate SSH key
      ```powershell
      ssh-keygen -t ed25519 -C "g2g@contoso.com" -f "C:\Users\svcG2G\.ssh\id_ed25519"
      Get-Content "C:\Users\svcG2G\.ssh\id_ed25519.pub"
      ```
   * Add the public key as a deploy key or user key with least required scope
   * Start the Windows SSH Agent (in a Powershell prompt as Administrator)
      ```powershell
      Get-Service ssh-agent | Set-Service -StartupType Automatic
      Start-Service ssh-agent
      ```
   * Store your passphrase for the SSH key
      ```powershell
      ssh-add C:\Users\svcG2G\.ssh\id_ed25519
      ```
   * (Depending on your setup you might need to fiddle around with Git config for `core.sshCommand` and/or the file `C:\Users\svcG2G\.ssh\config`)
   * Ensure `origin` uses SSH URL
      ```powershell
      git -C C:\Ops\GPO-Backup remote set-url origin git@github.com:myaccount/GPO-Backup.git
      ```
   * Validate non-interactive access (No input should be needed)
      ```powershell
      ssh -T git@github.com
      git -C C:\Ops\GPO-Backup pull --ff-only
      ```
   * Edit the README.md file (just add som random text to it)
   * Validate that you can push to Git
      ```powershell
      git -C C:\GIT\GPO-Backup add README.md
      git -C C:\GIT\GPO-Backup commit -m "Update README"
      git -C C:\GIT\GPO-Backup push
      ```
* Perform first real run
   ```powershell
   . "C:\Ops\Scripts\Invoke-GpoGitSync.ps1" -RepoPath "C:\Ops\GPO-Backup"
   ```
* Vaildate that you have all files in the target repo
* Remove service accounts temporarily granted access to local logon, only `Logon as a Batch Job` should be needed
* Automate it by creating a schedule task
  * Program: `powershell.exe`
  * Arguments: `-NoProfile -ExecutionPolicy RemoteSigned -File "C:\Ops\Scripts\Invoke-GpoGitSync.ps1" -RepoPath "C:\Ops\GPO-Backup"`
  * Start in: `C:\Ops\Scripts`
  * When running this task, use the following account: Previously created service account
  * Run whether user is logged on or not
  * Run with highest privileges (if your environment requires)
  * Configure for your server OS version
  * Trigger: choose an interval that matches your change rate
  * **Important!** Task Scheduler > Settings > If the task is already running: "Do not start a new instance"
  * Stop task if it runs unexpectedly long (optional safety)

# Workflow

* Check and validate all requirements
* Load the previous state from `gpo-state.json`
* Query the current AD domain, or every domain in the current forest when `-CompleteForest` is used, enumerate all GPOs, and build a normalized snapshot.
* Compare the current GPO snapshot with the previous state to detect:
   - new GPOs
   - changed GPOs by version, status, or assigned WMI filter
   - deleted GPOs
* Query AD containers in each scanned domain and collect GPO link information for the domain root and OUs.
* Export the current WMI filter snapshot, compare it with the previous run, and detect new, changed, or deleted filters.
* Compare stored per-GPO link data with the current link data to detect link-only changes.
* Export artifacts for changed targets
* Write the updated `gpo-state.json` with the latest snapshot data.
* Build a run summary, write a markdown report under `reports/`, and refresh the reports table in the repository `README.md`.
* Stop here when:
   - no changes were detected
   - `-DryRun` was used
   - `-DisableGIT` was used
   - the target path is not a git repository
* If git is enabled and changes exist, run `git checkout`, `git pull --ff-only`, `git add -A`, `git commit`, and `git push`.

## Output

* `README.md` summary and link to last 25 reports
* `gpos/<guid>/links.md` Markdown report of GPO links in single-domain mode
* `gpos/<guid>/report.html` complete GPO report in HTML format in single-domain mode
* `gpos/<guid>/report.xml` complete GPO report in XML format in single-domain mode
* `wmi-filters/<guid>.md` exported WMI filter snapshots in single-domain mode
* `gpos/<domain>/<guid>/...` namespaced GPO artifacts when `-CompleteForest` is used
* `wmi-filters/<domain>/<guid>.md` namespaced WMI filter snapshots when `-CompleteForest` is used
* `reports/g2g-YYYYMMDD-HHMMSS.md` Markdown report per run with detected changes
* `gpo-state.json` stores the last known state used for change detection

## License

This project is licensed under the [MIT License](LICENSE). Copyright (c) 2026 Rikard Rönnkvist.


[![FOSSA Status](https://app.fossa.com/api/projects/git%2Bgithub.com%2Frikardronnkvist%2FG2G.svg?type=large)](https://app.fossa.com/projects/git%2Bgithub.com%2Frikardronnkvist%2FG2G?ref=badge_large)