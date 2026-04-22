[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$BackupRepoPath,

    [Parameter(Mandatory = $false)]
    [string]$MainRepoPath = (Resolve-Path (Join-Path -Path $PSScriptRoot -ChildPath '..')).Path,

    [Parameter(Mandatory = $false)]
    [string[]]$ChangedGpoGuids,

    [Parameter(Mandatory = $false)]
    [string]$StateFilePath = 'state/gpo-state.json',

    [Parameter(Mandatory = $false)]
    [string]$Branch = '',

    [Parameter(Mandatory = $false)]
    [switch]$DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Log {
    param([string]$Message)
    Write-Host "[$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))] $Message"
}

function Get-ChangedGuidsFromState {
    param(
        [Parameter(Mandatory = $true)][string]$MainRepo,
        [Parameter(Mandatory = $true)][string]$StateRelativePath
    )

    $statePath = if ([System.IO.Path]::IsPathRooted($StateRelativePath)) { $StateRelativePath } else { Join-Path -Path $MainRepo -ChildPath $StateRelativePath }
    if (-not (Test-Path -LiteralPath $statePath)) {
        throw "State file not found: $statePath"
    }

    $raw = Get-Content -LiteralPath $statePath -Raw -Encoding UTF8
    $state = $raw | ConvertFrom-Json -Depth 20
    if (-not $state.Gpos) {
        return @()
    }

    return @($state.Gpos.PSObject.Properties.Name | Sort-Object)
}

try {
    Import-Module GroupPolicy -ErrorAction Stop
}
catch {
    throw 'GroupPolicy module is required (RSAT Group Policy Management tools).'
}

$BackupRepoPath = [System.IO.Path]::GetFullPath($BackupRepoPath)
$MainRepoPath = [System.IO.Path]::GetFullPath($MainRepoPath)

if (-not (Test-Path -LiteralPath $BackupRepoPath)) {
    throw "Backup repo path does not exist: $BackupRepoPath"
}
if (-not (Test-Path -LiteralPath (Join-Path -Path $BackupRepoPath -ChildPath '.git'))) {
    throw "Backup repo path is not a git repository: $BackupRepoPath"
}
if ($BackupRepoPath -eq $MainRepoPath) {
    throw 'Backup repo path must be different from main repo path.'
}
if ($BackupRepoPath.StartsWith($MainRepoPath, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw 'Backup repo path must not be inside the main repo path.'
}

if (-not $ChangedGpoGuids -or $ChangedGpoGuids.Count -eq 0) {
    Write-Log 'No ChangedGpoGuids supplied. Falling back to state-file method (backs up all known GUIDs from state).'
    $ChangedGpoGuids = Get-ChangedGuidsFromState -MainRepo $MainRepoPath -StateRelativePath $StateFilePath
}

if (-not $ChangedGpoGuids -or $ChangedGpoGuids.Count -eq 0) {
    Write-Log 'No GPO GUIDs to back up.'
    exit 0
}

$timestamp = (Get-Date).ToUniversalTime().ToString('yyyyMMdd-HHmmss')
$backupRunDir = Join-Path -Path $BackupRepoPath -ChildPath (Join-Path -Path 'backups' -ChildPath $timestamp)

if ($DryRun) {
    Write-Log "[DRYRUN] Would create backup folder: $backupRunDir"
}
else {
    New-Item -ItemType Directory -Path $backupRunDir -Force | Out-Null
}

$backedUp = New-Object System.Collections.Generic.List[string]
foreach ($guid in ($ChangedGpoGuids | Sort-Object -Unique)) {
    try {
        if ($DryRun) {
            Write-Log "[DRYRUN] Would backup GPO $guid to $backupRunDir"
        }
        else {
            Backup-GPO -Guid $guid -Path $backupRunDir -ErrorAction Stop | Out-Null
            $backedUp.Add($guid)
        }
    }
    catch {
        Write-Log "WARNING: Backup failed for $guid :: $($_.Exception.Message)"
    }
}

if ($DryRun) {
    Write-Log '[DRYRUN] Skipping git commit/push for backup repo.'
    exit 0
}

if ($backedUp.Count -eq 0) {
    Write-Log 'No successful backups were created.'
    exit 1
}

Push-Location $BackupRepoPath
try {
    if ([string]::IsNullOrWhiteSpace($Branch)) {
        $Branch = (git rev-parse --abbrev-ref HEAD).Trim()
    }
    if ([string]::IsNullOrWhiteSpace($Branch) -or $Branch -eq 'HEAD') {
        throw 'Could not determine git branch or repository is in detached HEAD state. Provide -Branch explicitly.'
    }

    & git checkout $Branch | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw "git checkout failed for branch $Branch"
    }

    & git pull --ff-only origin $Branch | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw "git pull failed for branch $Branch"
    }

    & git add -A
    & git diff --cached --quiet
    if ($LASTEXITCODE -eq 0) {
        Write-Log 'No backup changes to commit.'
        exit 0
    }

    $summary = "GPO backups: $($backedUp.Count) items ($timestamp)"
    & git commit -m $summary | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw 'git commit failed for backup repo'
    }

    & git push origin $Branch | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw 'git push failed for backup repo'
    }

    Write-Log "Backup commit pushed to '$Branch'."
    exit 0
}
finally {
    Pop-Location
}
