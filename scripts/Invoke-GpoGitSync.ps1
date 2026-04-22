[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$RepoPath = (Get-Location).Path,

    [Parameter(Mandatory = $false)]
    [string]$Branch = '',

    [Parameter(Mandatory = $false)]
    [string]$StateFilePath = 'state/gpo-state.json',

    [Parameter(Mandatory = $false)]
    [string]$LogDirectory = 'logs',

    [Parameter(Mandatory = $false)]
    [string]$TeamsWebhookUrl,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 50)]
    [int]$TeamsTopCount = 10,

    [Parameter(Mandatory = $false)]
    [switch]$DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:ExitCode = 0
$script:LogFilePath = $null

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string]$Level = 'INFO'
    )

    $timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $line = "[$timestamp][$Level] $Message"
    Write-Host $line
    if ($script:LogFilePath) {
        Add-Content -LiteralPath $script:LogFilePath -Value $line -Encoding UTF8
    }
}

function Fail {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $true)]
        [int]$Code
    )

    Write-Log -Message $Message -Level 'ERROR'
    $script:ExitCode = $Code
    exit $Code
}

function ConvertTo-StableJson {
    param(
        [Parameter(Mandatory = $true)]
        [object]$InputObject,
        [Parameter(Mandatory = $false)]
        [int]$Depth = 15
    )

    return ($InputObject | ConvertTo-Json -Depth $Depth)
}

function Get-GpoVersionString {
    param([Parameter(Mandatory = $true)][object]$GpoLike)
    return "U:$($GpoLike.UserDSVersion)/$($GpoLike.UserSysvolVersion);C:$($GpoLike.ComputerDSVersion)/$($GpoLike.ComputerSysvolVersion)"
}

function Get-GpoSnapshotEntry {
    param([Parameter(Mandatory = $true)][object]$Gpo)

    $wmi = $null
    if ($Gpo.WmiFilter) {
        $wmi = [ordered]@{
            Name = [string]$Gpo.WmiFilter.Name
            Path = [string]$Gpo.WmiFilter.Path
            Guid = [string]($Gpo.WmiFilter.Path -replace '.*\{([^\}]+)\}.*', '$1')
        }
    }

    return [ordered]@{
        Guid = [string]$Gpo.Id.Guid
        DisplayName = [string]$Gpo.DisplayName
        UserDSVersion = [int]$Gpo.User.DSVersion
        UserSysvolVersion = [int]$Gpo.User.SysVolVersion
        ComputerDSVersion = [int]$Gpo.Computer.DSVersion
        ComputerSysvolVersion = [int]$Gpo.Computer.SysVolVersion
        GpoStatus = [string]$Gpo.GpoStatus
        WmiFilter = $wmi
    }
}

function Export-GpoArtifacts {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,
        [Parameter(Mandatory = $true)]
        [object]$Gpo,
        [Parameter(Mandatory = $true)]
        [object]$Snapshot,
        [Parameter(Mandatory = $false)]
        [switch]$DryRunMode
    )

    $gpoGuid = [string]$Snapshot.Guid
    $gpoDir = Join-Path -Path $RepoRoot -ChildPath (Join-Path -Path 'gpos' -ChildPath $gpoGuid)
    $reportPath = Join-Path -Path $gpoDir -ChildPath 'report.xml'
    $metaPath = Join-Path -Path $gpoDir -ChildPath 'meta.json'

    if ($DryRunMode) {
        Write-Log -Message "[DRYRUN] Would export report/meta for '$($Snapshot.DisplayName)' ($gpoGuid)."
        return
    }

    if (-not (Test-Path -LiteralPath $gpoDir)) {
        New-Item -ItemType Directory -Path $gpoDir -Force | Out-Null
    }

    $reportXml = Get-GPOReport -Guid $gpoGuid -ReportType Xml
    Set-Content -LiteralPath $reportPath -Value $reportXml -Encoding UTF8

    $meta = [ordered]@{
        Guid = $Snapshot.Guid
        DisplayName = $Snapshot.DisplayName
        ExportedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
        Versions = [ordered]@{
            User = [ordered]@{
                DirectoryService = $Snapshot.UserDSVersion
                Sysvol = $Snapshot.UserSysvolVersion
            }
            Computer = [ordered]@{
                DirectoryService = $Snapshot.ComputerDSVersion
                Sysvol = $Snapshot.ComputerSysvolVersion
            }
        }
        GpoStatus = $Snapshot.GpoStatus
        WmiFilter = $Snapshot.WmiFilter
    }

    Set-Content -LiteralPath $metaPath -Value (ConvertTo-StableJson -InputObject $meta) -Encoding UTF8
}

function Parse-GpLinkValue {
    param([Parameter(Mandatory = $true)][string]$GpLink)

    $links = @()
    if ([string]::IsNullOrWhiteSpace($GpLink)) {
        return $links
    }

    $regex = [regex]'\[(?<path>LDAP://[^;\]]+);(?<options>\d+)\]'
    $linkMatches = $regex.Matches($GpLink)
    $order = 1
    foreach ($match in $linkMatches) {
        $rawPath = [string]$match.Groups['path'].Value
        $options = [int]$match.Groups['options'].Value
        $guid = [string]::Empty
        if ($rawPath -match '\{(?<guid>[0-9A-Fa-f\-]{36})\}') {
            $guid = $Matches['guid']
        }

        $links += [ordered]@{
            Order = $order
            GpoGuid = $guid.ToLowerInvariant()
            Path = $rawPath
            Enforced = [bool]($options -band 2)
            Disabled = [bool]($options -band 1)
            Options = $options
        }
        $order++
    }

    return $links
}

function Export-GpoLinksSnapshot {
    param(
        [Parameter(Mandatory = $true)][string]$DomainDistinguishedName,
        [Parameter(Mandatory = $true)][string]$DomainDnsRoot,
        [Parameter(Mandatory = $true)][switch]$DryRunMode,
        [Parameter(Mandatory = $true)][string]$OutputPath
    )

    $containers = New-Object System.Collections.Generic.List[object]

    $domainObject = Get-ADObject -Identity $DomainDistinguishedName -Properties distinguishedName, gPLink, gPOptions
    $ouObjects = Get-ADOrganizationalUnit -Filter * -Properties distinguishedName, gPLink, gPOptions | Sort-Object -Property DistinguishedName
    $allContainers = @($domainObject) + @($ouObjects)

    foreach ($container in $allContainers | Sort-Object -Property distinguishedName) {
        $dn = [string]$container.distinguishedName
        $links = Parse-GpLinkValue -GpLink ([string]$container.gPLink)

        $containers.Add([ordered]@{
                DistinguishedName = $dn
                IsDomainRoot = [bool]($dn -eq $DomainDistinguishedName)
                BlockInheritance = [bool](([int]$container.gPOptions) -band 1)
                Links = @($links)
            })
    }

    $snapshot = [ordered]@{
        GeneratedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
        DomainDnsRoot = $DomainDnsRoot
        DomainDistinguishedName = $DomainDistinguishedName
        Containers = @($containers)
    }

    if ($DryRunMode) {
        Write-Log -Message "[DRYRUN] Would write links snapshot to $OutputPath"
    }
    else {
        $parent = Split-Path -Path $OutputPath -Parent
        if (-not (Test-Path -LiteralPath $parent)) {
            New-Item -ItemType Directory -Path $parent -Force | Out-Null
        }
        Set-Content -LiteralPath $OutputPath -Value (ConvertTo-StableJson -InputObject $snapshot) -Encoding UTF8
    }

    return $snapshot
}

function Export-WmiFiltersSnapshot {
    param(
        [Parameter(Mandatory = $true)][string]$DomainDistinguishedName,
        [Parameter(Mandatory = $true)][array]$CurrentGpos,
        [Parameter(Mandatory = $true)][switch]$DryRunMode,
        [Parameter(Mandatory = $true)][string]$OutputPath
    )

    $wmiBase = "CN=SOM,CN=WMIPolicy,CN=System,$DomainDistinguishedName"
    $filters = @()

    try {
        $adFilters = Get-ADObject -SearchBase $wmiBase -LDAPFilter '(objectClass=msWMI-Som)' -Properties msWMI-Name, msWMI-ID, msWMI-Parm1, distinguishedName |
            Sort-Object -Property 'msWMI-Name'

        foreach ($item in $adFilters) {
            $guid = [string]$item.'msWMI-ID'
            if ($guid -and $guid.StartsWith('{') -and $guid.EndsWith('}')) {
                $guid = $guid.Trim('{}')
            }

            $filters += [ordered]@{
                Guid = [string]$guid
                Name = [string]$item.'msWMI-Name'
                Query = [string]$item.'msWMI-Parm1'
                DistinguishedName = [string]$item.DistinguishedName
            }
        }
    }
    catch {
        throw "Failed to query WMI filters from Active Directory path '$wmiBase'. $($_.Exception.Message)"
    }

    $gpoToFilterMap = @()
    foreach ($gpo in $CurrentGpos | Sort-Object -Property DisplayName) {
        if ($gpo.WmiFilter) {
            $filterGuid = [string]($gpo.WmiFilter.Path -replace '.*\{([^\}]+)\}.*', '$1')
            $gpoToFilterMap += [ordered]@{
                GpoGuid = [string]$gpo.Id.Guid
                GpoDisplayName = [string]$gpo.DisplayName
                WmiFilterGuid = $filterGuid
                WmiFilterName = [string]$gpo.WmiFilter.Name
            }
        }
    }

    $snapshot = [ordered]@{
        GeneratedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
        Filters = @($filters | Sort-Object -Property Name, Guid)
        GpoAssignments = @($gpoToFilterMap | Sort-Object -Property GpoDisplayName, GpoGuid)
    }

    if ($DryRunMode) {
        Write-Log -Message "[DRYRUN] Would write WMI filters snapshot to $OutputPath"
    }
    else {
        $parent = Split-Path -Path $OutputPath -Parent
        if (-not (Test-Path -LiteralPath $parent)) {
            New-Item -ItemType Directory -Path $parent -Force | Out-Null
        }
        Set-Content -LiteralPath $OutputPath -Value (ConvertTo-StableJson -InputObject $snapshot) -Encoding UTF8
    }

    return $snapshot
}

function Get-LinkDiffs {
    param(
        [Parameter(Mandatory = $true)][object]$PreviousSnapshot,
        [Parameter(Mandatory = $true)][object]$CurrentSnapshot
    )

    $previousJson = ConvertTo-StableJson -InputObject $PreviousSnapshot
    $currentJson = ConvertTo-StableJson -InputObject $CurrentSnapshot

    $memory1 = [System.IO.MemoryStream]::new([System.Text.Encoding]::UTF8.GetBytes($previousJson))
    $memory2 = [System.IO.MemoryStream]::new([System.Text.Encoding]::UTF8.GetBytes($currentJson))
    try {
        $previousHash = (Get-FileHash -InputStream $memory1 -Algorithm SHA256).Hash
        $currentHash = (Get-FileHash -InputStream $memory2 -Algorithm SHA256).Hash
    }
    finally {
        $memory1.Dispose()
        $memory2.Dispose()
    }

    if ($previousJson -eq $currentJson) {
        return [ordered]@{
            HasChanges = $false
            ChangedGpoGuids = @()
            PreviousHash = $previousHash
            CurrentHash = $currentHash
        }
    }

    $previousByDn = @{}
    foreach ($container in @($PreviousSnapshot.Containers)) {
        $previousByDn[[string]$container.DistinguishedName] = $container
    }

    $currentByDn = @{}
    foreach ($container in @($CurrentSnapshot.Containers)) {
        $currentByDn[[string]$container.DistinguishedName] = $container
    }

    $allDns = @($previousByDn.Keys + $currentByDn.Keys | Sort-Object -Unique)
    $affected = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($dn in $allDns) {
        $previousContainer = $previousByDn[$dn]
        $currentContainer = $currentByDn[$dn]
        $prevLinks = if ($previousContainer) { @($previousContainer.Links) } else { @() }
        $currLinks = if ($currentContainer) { @($currentContainer.Links) } else { @() }
        $prevLinksJson = ConvertTo-StableJson -InputObject $prevLinks
        $currLinksJson = ConvertTo-StableJson -InputObject $currLinks
        $prevBlock = if ($previousContainer) { [bool]$previousContainer.BlockInheritance } else { $false }
        $currBlock = if ($currentContainer) { [bool]$currentContainer.BlockInheritance } else { $false }

        if (($prevLinksJson -ne $currLinksJson) -or ($prevBlock -ne $currBlock)) {
            foreach ($link in ($prevLinks + $currLinks)) {
                if ($link.GpoGuid) {
                    [void]$affected.Add(([string]$link.GpoGuid).ToLowerInvariant())
                }
            }
        }
    }

    return [ordered]@{
        HasChanges = $true
        ChangedGpoGuids = @($affected.ToArray() | Sort-Object)
        PreviousHash = $previousHash
        CurrentHash = $currentHash
    }
}

function Write-ReportHtml {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [hashtable]$Summary,
        [Parameter(Mandatory = $false)]
        [switch]$DryRunMode
    )

    $style = @"
<style>
body { font-family: Segoe UI, Arial, sans-serif; margin: 20px; }
table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }
th, td { border: 1px solid #cccccc; padding: 8px; text-align: left; }
th { background: #f2f2f2; }
code { background: #f8f8f8; padding: 2px 4px; }
</style>
"@

    $html = New-Object System.Text.StringBuilder
    [void]$html.AppendLine('<!DOCTYPE html>')
    [void]$html.AppendLine('<html><head><meta charset="utf-8" />')
    [void]$html.AppendLine('<title>GPO Sync Report</title>')
    [void]$html.AppendLine($style)
    [void]$html.AppendLine('</head><body>')
    [void]$html.AppendLine('<h1>GPO Sync Report</h1>')
    [void]$html.AppendLine("<p>Generated: $(Get-Date -Format u)</p>")
    [void]$html.AppendLine("<p>DryRun: <code>$($DryRunMode.IsPresent)</code></p>")
    [void]$html.AppendLine("<p>Changed: <code>$($Summary.Changed.Count)</code>, New: <code>$($Summary.New.Count)</code>, Deleted: <code>$($Summary.Deleted.Count)</code>, Link-changed: <code>$($Summary.LinkChanged.Count)</code></p>")

    $tables = @(
        @{ Name = 'Changed'; Items = $Summary.Changed },
        @{ Name = 'New'; Items = $Summary.New },
        @{ Name = 'Deleted'; Items = $Summary.Deleted },
        @{ Name = 'Link-changed'; Items = $Summary.LinkChanged }
    )

    foreach ($table in $tables) {
        [void]$html.AppendLine("<h2>$($table.Name)</h2>")
        [void]$html.AppendLine('<table><thead><tr><th>Name</th><th>Guid</th><th>Detail</th></tr></thead><tbody>')
        foreach ($item in $table.Items) {
            $name = [System.Net.WebUtility]::HtmlEncode([string]$item.DisplayName)
            $guid = [System.Net.WebUtility]::HtmlEncode([string]$item.Guid)
            $detail = [System.Net.WebUtility]::HtmlEncode([string]$item.Detail)
            [void]$html.AppendLine("<tr><td>$name</td><td>$guid</td><td>$detail</td></tr>")
        }
        if (-not $table.Items -or $table.Items.Count -eq 0) {
            [void]$html.AppendLine('<tr><td colspan="3"><em>None</em></td></tr>')
        }
        [void]$html.AppendLine('</tbody></table>')
    }

    [void]$html.AppendLine('</body></html>')

    if ($DryRunMode) {
        Write-Log -Message "[DRYRUN] Would generate HTML report at $Path"
        return
    }

    $parent = Split-Path -Path $Path -Parent
    if (-not (Test-Path -LiteralPath $parent)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }

    Set-Content -LiteralPath $Path -Value $html.ToString() -Encoding UTF8
}

function Send-TeamsWebhook {
    param(
        [Parameter(Mandatory = $true)]
        [string]$WebhookUrl,
        [Parameter(Mandatory = $true)]
        [hashtable]$Summary,
        [Parameter(Mandatory = $false)]
        [switch]$DryRunMode
    )

    if ([string]::IsNullOrWhiteSpace($WebhookUrl)) {
        return
    }

    $topItems = @($Summary.Changed + $Summary.New + $Summary.Deleted + $Summary.LinkChanged | Select-Object -First $TeamsTopCount)
    $lines = @(
        'GPO sync result:'
        "Changed: $($Summary.Changed.Count)"
        "New: $($Summary.New.Count)"
        "Deleted: $($Summary.Deleted.Count)"
        "Link-changed: $($Summary.LinkChanged.Count)"
    )

    if ($topItems.Count -gt 0) {
        $lines += 'Top items:'
        $lines += ($topItems | ForEach-Object { "- $($_.DisplayName) [$($_.Guid)]" })
    }

    $payload = [ordered]@{
        text = ($lines -join "`n")
    }

    if ($DryRunMode) {
        Write-Log -Message '[DRYRUN] Would send Teams webhook notification.'
        return
    }

    try {
        Invoke-RestMethod -Method Post -Uri $WebhookUrl -Body (ConvertTo-StableJson -InputObject $payload) -ContentType 'application/json' | Out-Null
        Write-Log -Message 'Teams webhook notification sent.'
    }
    catch {
        Write-Log -Message "Teams webhook failed: $($_.Exception.Message)" -Level 'WARN'
    }
}

function Ensure-RepositoryStructure {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Root,
        [Parameter(Mandatory = $false)]
        [switch]$DryRunMode
    )

    $requiredDirs = @('state', 'gpos', 'links', 'wmi-filters', 'reports', 'logs', 'temp')
    foreach ($dir in $requiredDirs) {
        $path = Join-Path -Path $Root -ChildPath $dir
        if (-not (Test-Path -LiteralPath $path)) {
            if ($DryRunMode) {
                Write-Log -Message "[DRYRUN] Would create directory $path"
            }
            else {
                New-Item -ItemType Directory -Path $path -Force | Out-Null
            }
        }
    }
}

try {
    $RepoPath = [System.IO.Path]::GetFullPath($RepoPath)
    if (-not (Test-Path -LiteralPath $RepoPath)) {
        Fail -Message "Repository path does not exist: $RepoPath" -Code 12
    }

    Ensure-RepositoryStructure -Root $RepoPath -DryRunMode:$DryRun

    $logDirPath = Join-Path -Path $RepoPath -ChildPath $LogDirectory
    if (-not (Test-Path -LiteralPath $logDirPath)) {
        New-Item -ItemType Directory -Path $logDirPath -Force | Out-Null
    }

    $runTimestamp = (Get-Date).ToUniversalTime().ToString('yyyyMMdd-HHmmss')
    $script:LogFilePath = Join-Path -Path $logDirPath -ChildPath "run-$runTimestamp.log"
    New-Item -ItemType File -Path $script:LogFilePath -Force | Out-Null
    Write-Log -Message "Starting GPO sync. RepoPath=$RepoPath DryRun=$($DryRun.IsPresent)"

    try {
        Import-Module GroupPolicy -ErrorAction Stop
    }
    catch {
        Fail -Message 'GroupPolicy module is required but not available. Install RSAT Group Policy Management tools.' -Code 10
    }

    try {
        Import-Module ActiveDirectory -ErrorAction Stop
    }
    catch {
        Fail -Message 'ActiveDirectory module is required but not available.' -Code 10
    }

    $gitCmd = Get-Command git -ErrorAction SilentlyContinue
    if (-not $gitCmd) {
        Fail -Message 'git executable was not found in PATH.' -Code 12
    }

    $stateFileAbsPath = if ([System.IO.Path]::IsPathRooted($StateFilePath)) { $StateFilePath } else { Join-Path -Path $RepoPath -ChildPath $StateFilePath }
    $linksFilePath = Join-Path -Path $RepoPath -ChildPath 'links/gpo-links.json'
    $wmiFilePath = Join-Path -Path $RepoPath -ChildPath 'wmi-filters/wmi-filters.json'
    $reportsPath = Join-Path -Path $RepoPath -ChildPath 'reports'
    $htmlReportPath = Join-Path -Path $reportsPath -ChildPath "run-$runTimestamp.html"

    $previousState = [ordered]@{ Gpos = [ordered]@{}; LinksSnapshot = $null; LinksHash = '' }
    if (Test-Path -LiteralPath $stateFileAbsPath) {
        try {
            $rawState = Get-Content -LiteralPath $stateFileAbsPath -Raw -Encoding UTF8
            if (-not [string]::IsNullOrWhiteSpace($rawState)) {
                $parsed = $rawState | ConvertFrom-Json -Depth 20
                if ($parsed.Gpos) {
                    $gposOrdered = [ordered]@{}
                    foreach ($prop in $parsed.Gpos.PSObject.Properties | Sort-Object -Property Name) {
                        $gposOrdered[$prop.Name] = $prop.Value
                    }
                    $previousState.Gpos = $gposOrdered
                }
                if ($parsed.LinksSnapshot) {
                    $previousState.LinksSnapshot = $parsed.LinksSnapshot
                }
                if ($parsed.LinksHash) {
                    $previousState.LinksHash = [string]$parsed.LinksHash
                }
            }
        }
        catch {
            Fail -Message "Failed to parse state file '$stateFileAbsPath': $($_.Exception.Message)" -Code 12
        }
    }

    try {
        $domain = Get-ADDomain -ErrorAction Stop
    }
    catch {
        Fail -Message "Failed to query Active Directory domain details. $($_.Exception.Message)" -Code 11
    }

    try {
        $currentGpos = Get-GPO -All -ErrorAction Stop | Sort-Object -Property DisplayName
    }
    catch {
        Fail -Message "Failed to enumerate GPOs. Verify permissions. $($_.Exception.Message)" -Code 11
    }

    $currentSnapshotByGuid = [ordered]@{}
    foreach ($gpo in $currentGpos) {
        $entry = Get-GpoSnapshotEntry -Gpo $gpo
        $currentSnapshotByGuid[[string]$entry.Guid] = $entry
    }

    $changed = New-Object System.Collections.Generic.List[object]
    $newItems = New-Object System.Collections.Generic.List[object]
    $deleted = New-Object System.Collections.Generic.List[object]
    $linkChanged = New-Object System.Collections.Generic.List[object]
    $exportTargets = New-Object 'System.Collections.Generic.HashSet[string]'

    foreach ($guid in $currentSnapshotByGuid.Keys) {
        $current = $currentSnapshotByGuid[$guid]
        $previous = $null
        if ($previousState.Gpos.Contains($guid)) {
            $previous = $previousState.Gpos[$guid]
        }

        if (-not $previous) {
            $newItems.Add([ordered]@{ DisplayName = $current.DisplayName; Guid = $guid; Detail = 'New GPO' })
            [void]$exportTargets.Add($guid)
            continue
        }

        $versionChanged = (
            ([int]$previous.UserDSVersion -ne [int]$current.UserDSVersion) -or
            ([int]$previous.UserSysvolVersion -ne [int]$current.UserSysvolVersion) -or
            ([int]$previous.ComputerDSVersion -ne [int]$current.ComputerDSVersion) -or
            ([int]$previous.ComputerSysvolVersion -ne [int]$current.ComputerSysvolVersion)
        )
        $statusChanged = ([string]$previous.GpoStatus -ne [string]$current.GpoStatus)

        $previousWmiPath = ''
        $currentWmiPath = ''
        if ($previous.WmiFilter) { $previousWmiPath = [string]$previous.WmiFilter.Path }
        if ($current.WmiFilter) { $currentWmiPath = [string]$current.WmiFilter.Path }
        $wmiChanged = ($previousWmiPath -ne $currentWmiPath)

        if ($versionChanged -or $statusChanged -or $wmiChanged) {
            $detailParts = @()
            if ($versionChanged) {
                $detailParts += "Versions $(Get-GpoVersionString -GpoLike $previous) -> $(Get-GpoVersionString -GpoLike $current)"
            }
            if ($statusChanged) {
                $detailParts += "Status $($previous.GpoStatus) -> $($current.GpoStatus)"
            }
            if ($wmiChanged) {
                $detailParts += 'WMI filter changed'
            }

            $changed.Add([ordered]@{
                    DisplayName = $current.DisplayName
                    Guid = $guid
                    Detail = ($detailParts -join '; ')
                })
            [void]$exportTargets.Add($guid)
        }
    }

    foreach ($guid in $previousState.Gpos.Keys) {
        if (-not $currentSnapshotByGuid.Contains($guid)) {
            $deletedDisplay = [string]$previousState.Gpos[$guid].DisplayName
            $deleted.Add([ordered]@{ DisplayName = $deletedDisplay; Guid = $guid; Detail = 'Deleted GPO' })
            $gpoDir = Join-Path -Path $RepoPath -ChildPath (Join-Path -Path 'gpos' -ChildPath $guid)
            if (Test-Path -LiteralPath $gpoDir) {
                if ($DryRun) {
                    Write-Log -Message "[DRYRUN] Would remove folder: $gpoDir"
                }
                else {
                    Remove-Item -LiteralPath $gpoDir -Recurse -Force
                    Write-Log -Message "Removed deleted GPO folder: $gpoDir"
                }
            }
        }
    }

    try {
        $linksSnapshot = Export-GpoLinksSnapshot -DomainDistinguishedName $domain.DistinguishedName -DomainDnsRoot $domain.DNSRoot -DryRunMode:$DryRun -OutputPath $linksFilePath
        $wmiSnapshot = Export-WmiFiltersSnapshot -DomainDistinguishedName $domain.DistinguishedName -CurrentGpos $currentGpos -DryRunMode:$DryRun -OutputPath $wmiFilePath
    }
    catch {
        Fail -Message "Failed AD snapshot export: $($_.Exception.Message)" -Code 11
    }

    $previousLinksSnapshot = if ($previousState.LinksSnapshot) { $previousState.LinksSnapshot } else { [ordered]@{ Containers = @() } }
    $linkDiffs = Get-LinkDiffs -PreviousSnapshot $previousLinksSnapshot -CurrentSnapshot $linksSnapshot
    if ($linkDiffs.HasChanges) {
        foreach ($guid in $linkDiffs.ChangedGpoGuids) {
            if ($currentSnapshotByGuid.Contains($guid)) {
                $alreadyChanged = @($changed | Where-Object { $_.Guid -eq $guid }).Count -gt 0
                $alreadyNew = @($newItems | Where-Object { $_.Guid -eq $guid }).Count -gt 0
                if (-not $alreadyChanged -and -not $alreadyNew) {
                    $linkChanged.Add([ordered]@{
                            DisplayName = $currentSnapshotByGuid[$guid].DisplayName
                            Guid = $guid
                            Detail = 'Container link/order/enforced change'
                        })
                }
            }
        }
    }

    foreach ($guid in $exportTargets.ToArray() | Sort-Object) {
        $gpo = $currentGpos | Where-Object { [string]$_.Id.Guid -eq $guid } | Select-Object -First 1
        if ($null -ne $gpo) {
            Export-GpoArtifacts -RepoRoot $RepoPath -Gpo $gpo -Snapshot $currentSnapshotByGuid[$guid] -DryRunMode:$DryRun
        }
    }

    $wmiJson = ConvertTo-StableJson -InputObject $wmiSnapshot
    $wmiHashStream = [System.IO.MemoryStream]::new([System.Text.Encoding]::UTF8.GetBytes($wmiJson))
    try {
        $wmiHash = (Get-FileHash -InputStream $wmiHashStream -Algorithm SHA256).Hash
    }
    finally {
        $wmiHashStream.Dispose()
    }

    $newState = [ordered]@{
        UpdatedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
        LinksHash = $linkDiffs.CurrentHash
        LinksSnapshot = $linksSnapshot
        WmiSnapshotHash = $wmiHash
        Gpos = [ordered]@{}
    }
    foreach ($guid in $currentSnapshotByGuid.Keys | Sort-Object) {
        $newState.Gpos[$guid] = $currentSnapshotByGuid[$guid]
    }

    if ($DryRun) {
        Write-Log -Message "[DRYRUN] Would write updated state to $stateFileAbsPath"
    }
    else {
        $stateParent = Split-Path -Path $stateFileAbsPath -Parent
        if (-not (Test-Path -LiteralPath $stateParent)) {
            New-Item -ItemType Directory -Path $stateParent -Force | Out-Null
        }
        Set-Content -LiteralPath $stateFileAbsPath -Value (ConvertTo-StableJson -InputObject $newState) -Encoding UTF8
    }

    $summary = @{
        Changed = @($changed | Sort-Object -Property DisplayName, Guid)
        New = @($newItems | Sort-Object -Property DisplayName, Guid)
        Deleted = @($deleted | Sort-Object -Property DisplayName, Guid)
        LinkChanged = @($linkChanged | Sort-Object -Property DisplayName, Guid)
    }

    Write-ReportHtml -Path (Join-Path -Path $reportsPath -ChildPath "run-$runTimestamp.html") -Summary $summary -DryRunMode:$DryRun

    $changeCount = $summary.Changed.Count + $summary.New.Count + $summary.Deleted.Count + $summary.LinkChanged.Count
    $summaryLine = "GPO sync: $($summary.Changed.Count) changed, $($summary.New.Count) new, $($summary.Deleted.Count) deleted, $($summary.LinkChanged.Count) link-changed"
    Write-Log -Message $summaryLine

    $commitMessageBuilder = New-Object System.Text.StringBuilder
    [void]$commitMessageBuilder.AppendLine($summaryLine)
    $sections = @(
        @{ Name = 'Changed'; Items = $summary.Changed },
        @{ Name = 'New'; Items = $summary.New },
        @{ Name = 'Deleted'; Items = $summary.Deleted },
        @{ Name = 'Link-changed'; Items = $summary.LinkChanged }
    )

    foreach ($section in $sections) {
        [void]$commitMessageBuilder.AppendLine()
        [void]$commitMessageBuilder.AppendLine("$($section.Name):")
        if ($section.Items.Count -eq 0) {
            [void]$commitMessageBuilder.AppendLine('- none')
        }
        else {
            foreach ($item in $section.Items) {
                [void]$commitMessageBuilder.AppendLine("- $($item.DisplayName) [$($item.Guid)] :: $($item.Detail)")
            }
        }
    }
    $commitMessage = $commitMessageBuilder.ToString().TrimEnd()

    if ($changeCount -eq 0) {
        Send-TeamsWebhook -WebhookUrl $TeamsWebhookUrl -Summary $summary -DryRunMode:$DryRun
        Write-Log -Message 'No changes detected.'
        exit 0
    }

    if ($DryRun) {
        Write-Log -Message '[DRYRUN] Skipping git add/commit/push.'
        Send-TeamsWebhook -WebhookUrl $TeamsWebhookUrl -Summary $summary -DryRunMode:$DryRun
        exit 1
    }

    Push-Location $RepoPath
    try {
        if ([string]::IsNullOrWhiteSpace($Branch)) {
            $Branch = (git rev-parse --abbrev-ref HEAD).Trim()
        }
        if ([string]::IsNullOrWhiteSpace($Branch) -or $Branch -eq 'HEAD') {
            Fail -Message 'Could not determine git branch. Provide -Branch explicitly.' -Code 12
        }

        Write-Log -Message "Running git checkout $Branch"
        & git checkout $Branch | Out-Null
        if ($LASTEXITCODE -ne 0) {
            Fail -Message "git checkout $Branch failed." -Code 12
        }

        Write-Log -Message "Running git pull --ff-only origin $Branch"
        & git pull --ff-only origin $Branch | Out-Null
        if ($LASTEXITCODE -ne 0) {
            Fail -Message "git pull failed for branch '$Branch'." -Code 12
        }

        & git add -A
        if ($LASTEXITCODE -ne 0) {
            Fail -Message 'git add failed.' -Code 12
        }

        & git diff --cached --quiet
        $hasStagedChanges = ($LASTEXITCODE -ne 0)
        if (-not $hasStagedChanges) {
            Write-Log -Message 'No staged git changes found after export.'
            Send-TeamsWebhook -WebhookUrl $TeamsWebhookUrl -Summary $summary -DryRunMode:$false
            exit 0
        }

        $commitTempPath = Join-Path -Path $RepoPath -ChildPath ("temp/commit-$runTimestamp.txt")
        Set-Content -LiteralPath $commitTempPath -Value $commitMessage -Encoding UTF8
        & git commit -F $commitTempPath | Out-Null
        if ($LASTEXITCODE -ne 0) {
            Fail -Message 'git commit failed.' -Code 12
        }
        if (Test-Path -LiteralPath $commitTempPath) {
            Remove-Item -LiteralPath $commitTempPath -Force
        }

        & git push origin $Branch | Out-Null
        if ($LASTEXITCODE -ne 0) {
            Write-Log -Message 'git push failed after successful export/commit.' -Level 'ERROR'
            Send-TeamsWebhook -WebhookUrl $TeamsWebhookUrl -Summary $summary -DryRunMode:$false
            exit 2
        }

        Write-Log -Message 'git push completed successfully.'
        Send-TeamsWebhook -WebhookUrl $TeamsWebhookUrl -Summary $summary -DryRunMode:$false
        exit 1
    }
    finally {
        Pop-Location
    }
}
catch {
    if ($script:ExitCode -gt 0) {
        exit $script:ExitCode
    }

    Write-Log -Message "Unhandled failure: $($_.Exception.Message)" -Level 'ERROR'
    exit 12
}
