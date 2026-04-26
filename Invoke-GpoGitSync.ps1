[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = 'Path to the local git repository where GPO artifacts are stored.')]
    [string]$RepoPath,

    [Parameter(Mandatory = $false, HelpMessage = 'Git branch to use for pull/commit/push. Defaults to current branch if omitted.')]
    [string]$Branch = '',

    [Parameter(Mandatory = $false, HelpMessage = 'Path to the state file (absolute or relative to RepoPath).')]
    [string]$StateFilePath = 'gpo-state.json',

    [Parameter(Mandatory = $false, HelpMessage = 'Run in simulation mode without writing files or pushing git changes.')]
    [switch]$DryRun,

    [Parameter(Mandatory = $false, HelpMessage = 'Skip all git operations (checkout, pull, commit, push) and git installation checks. Artifacts are stored locally only.')]
    [switch]$DisableGIT
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:ExitCode = 0

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [ValidateSet('INFO', 'WARN', 'ERROR', 'DEBUG')]
        [string]$Level = 'INFO'
    )

    $timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $line = "[$timestamp][$Level] $Message"
    Write-Host $line
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
        [Parameter(Mandatory = $false)]
        [object]$InputObject,
        [Parameter(Mandatory = $false)]
        [int]$Depth = 15
    )

    function ConvertTo-DeterministicObject {
        param([Parameter(Mandatory = $false)][object]$Value)

        try {
            if ($null -eq $Value) {
                return $null
            }

            if ($Value -is [string] -or $Value -is [ValueType]) {
                return $Value
            }

            if ($Value -is [System.Collections.IDictionary]) {
                $ordered = [ordered]@{}
                $entries = @($Value.GetEnumerator())
                foreach ($entry in ($entries | Sort-Object -Property @{ Expression = { [string]$_.Key } })) {
                    $keyString = ''
                    $entryValue = $null
                    try { $keyString = [string]$entry.Key } catch { $keyString = '(key)' }
                    try { $entryValue = $entry.Value } catch { $entryValue = [string]$entry }

                    try {
                        $ordered[$keyString] = ConvertTo-DeterministicObject -Value $entryValue
                    }
                    catch {
                        $ordered[$keyString] = [string]$entryValue
                    }
                }
                return $ordered
            }

            if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
                $list = New-Object System.Collections.ArrayList
                foreach ($item in $Value) {
                    try {
                        [void]$list.Add([object](ConvertTo-DeterministicObject -Value $item))
                    }
                    catch {
                        [void]$list.Add([string]$item)
                    }
                }
                Write-Output -NoEnumerate ([object[]]$list.ToArray())
                return
            }

            $props = $Value.PSObject.Properties
            if ($props.Count -gt 0) {
                $ordered = [ordered]@{}
                foreach ($prop in ($props | Sort-Object -Property Name)) {
                    $propValue = $null
                    try { $propValue = $prop.Value } catch { $propValue = $null }

                    try {
                        $ordered[$prop.Name] = ConvertTo-DeterministicObject -Value $propValue
                    }
                    catch {
                        $ordered[$prop.Name] = [string]$propValue
                    }
                }
                return $ordered
            }

            return $Value
        }
        catch {
            try {
                return [string]$Value
            }
            catch {
                return $null
            }
        }
    }

    try {
        $deterministic = ConvertTo-DeterministicObject -Value $InputObject
        return ($deterministic | ConvertTo-Json -Depth $Depth)
    }
    catch {
        return ($InputObject | ConvertTo-Json -Depth $Depth)
    }
}

function Get-GpoVersionString {
    param([Parameter(Mandatory = $true)][object]$GpoLike)
    return "U:$($GpoLike.UserDSVersion)/$($GpoLike.UserSysvolVersion);C:$($GpoLike.ComputerDSVersion)/$($GpoLike.ComputerSysvolVersion)"
}

function Get-GpoSnapshotEntry {
    param([Parameter(Mandatory = $true)][object]$Gpo)

    $wmi = $null
    if ($Gpo.WmiFilter) {
        $wmiGuid = Normalize-GuidString -GuidValue ($Gpo.WmiFilter.Path -replace '.*\{([^\}]+)\}.*', '$1')
        $wmi = [ordered]@{
            Name = [string]$Gpo.WmiFilter.Name
            Path = [string]$Gpo.WmiFilter.Path
            Guid = $wmiGuid
        }
    }

    # Handle both nested structure (User.DSVersion) and flat structure (UserVersion) from different PS environments
    $userDSVersion = 0
    $userSysvolVersion = 0
    $computerDSVersion = 0
    $computerSysvolVersion = 0

    # Try nested User/Computer objects first (native PowerShell 7)
    try {
        if ($null -ne $Gpo.User -and $Gpo.User.PSObject.Properties.Name -contains 'DSVersion') {
            $userDSVersion = [int]$Gpo.User.DSVersion
            $userSysvolVersion = [int]$Gpo.User.SysVolVersion
        }
    }
    catch { }

    try {
        if ($null -ne $Gpo.Computer -and $Gpo.Computer.PSObject.Properties.Name -contains 'DSVersion') {
            $computerDSVersion = [int]$Gpo.Computer.DSVersion
            $computerSysvolVersion = [int]$Gpo.Computer.SysVolVersion
        }
    }
    catch { }

    # Fallback for WinPSCompatSession: try UserVersion/ComputerVersion strings
    if ($userDSVersion -eq 0 -and $userSysvolVersion -eq 0) {
        try {
            if ($Gpo.UserVersion -match ':') {
                $userParts = [string]$Gpo.UserVersion -split ':'
                $userDSVersion = if ($userParts.Count -gt 0 -and $userParts[0]) { [int]$userParts[0] } else { 0 }
                $userSysvolVersion = if ($userParts.Count -gt 1 -and $userParts[1]) { [int]$userParts[1] } else { 0 }
            }
        }
        catch { }
    }

    if ($computerDSVersion -eq 0 -and $computerSysvolVersion -eq 0) {
        try {
            if ($Gpo.ComputerVersion -match ':') {
                $computerParts = [string]$Gpo.ComputerVersion -split ':'
                $computerDSVersion = if ($computerParts.Count -gt 0 -and $computerParts[0]) { [int]$computerParts[0] } else { 0 }
                $computerSysvolVersion = if ($computerParts.Count -gt 1 -and $computerParts[1]) { [int]$computerParts[1] } else { 0 }
            }
        }
        catch { }
    }

    return [ordered]@{
        Guid = Normalize-GuidString -GuidValue $Gpo.Id.Guid
        DisplayName = [string]$Gpo.DisplayName
        UserDSVersion = $userDSVersion
        UserSysvolVersion = $userSysvolVersion
        ComputerDSVersion = $computerDSVersion
        ComputerSysvolVersion = $computerSysvolVersion
        GpoStatus = [string]$Gpo.GpoStatus
        WmiFilter = $wmi
    }
}

function Export-GpoArtifacts {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,
        [Parameter(Mandatory = $false)]
        [object]$Gpo,
        [Parameter(Mandatory = $true)]
        [object]$Snapshot,
        [Parameter(Mandatory = $false)]
        [switch]$Deleted,
        [Parameter(Mandatory = $false)]
        [switch]$DryRunMode
    )

    $gpoGuid = Normalize-GuidString -GuidValue (Get-PropertyValue -Object $Snapshot -Name 'Guid')
    $displayName = [string](Get-PropertyValue -Object $Snapshot -Name 'DisplayName')
    if ([string]::IsNullOrWhiteSpace($displayName)) {
        $displayName = $gpoGuid
    }

    $gpoDir = Join-Path -Path $RepoRoot -ChildPath (Join-Path -Path 'gpos' -ChildPath $gpoGuid)
    $reportPathXml = Join-Path -Path $gpoDir -ChildPath 'report.xml'
    $reportPathHtml = Join-Path -Path $gpoDir -ChildPath 'report.html'
    $readmePath = Join-Path -Path $gpoDir -ChildPath 'README.md'
    $linksPath = Join-Path -Path $gpoDir -ChildPath 'links.md'

    if ($Deleted) {
        $deletedAt = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss zzz')

        if ($DryRunMode) {
            Write-Log -Message "[DRYRUN] Would update deleted GPO README for '$displayName' ($gpoGuid)."
            return
        }

        if (-not (Test-Path -LiteralPath $gpoDir)) {
            New-Item -ItemType Directory -Path $gpoDir -Force | Out-Null
        }

        if (-not (Test-Path -LiteralPath $readmePath)) {
            $status = [string](Get-PropertyValue -Object $Snapshot -Name 'GpoStatus')
            $readme = New-Object System.Text.StringBuilder
            [void]$readme.AppendLine("# $displayName")
            [void]$readme.AppendLine()
            [void]$readme.AppendLine('## Metadata')
            [void]$readme.AppendLine()
            [void]$readme.AppendLine('| Property | Value |')
            [void]$readme.AppendLine('|----------|-------|')
            [void]$readme.AppendLine("| **GUID** | $gpoGuid |")
            if (-not [string]::IsNullOrWhiteSpace($status)) {
                [void]$readme.AppendLine("| **Status** | $status |")
            }
            [void]$readme.AppendLine("| **Deleted** | $deletedAt |")
            [void]$readme.AppendLine()
            Set-Content -LiteralPath $readmePath -Value $readme.ToString() -Encoding UTF8
            return
        }

        $lines = New-Object 'System.Collections.Generic.List[string]'
        foreach ($line in Get-Content -LiteralPath $readmePath -Encoding UTF8) {
            $lines.Add([string]$line)
        }

        $deletedRow = "| **Deleted** | $deletedAt |"
        $metadataIndex = -1
        for ($i = 0; $i -lt $lines.Count; $i++) {
            if ($lines[$i] -eq '## Metadata') {
                $metadataIndex = $i
                break
            }
        }

        if ($metadataIndex -ge 0) {
            $sectionEnd = $lines.Count
            for ($i = $metadataIndex + 1; $i -lt $lines.Count; $i++) {
                if ($lines[$i] -match '^## ') {
                    $sectionEnd = $i
                    break
                }
            }

            $updated = $false
            for ($i = $metadataIndex + 1; $i -lt $sectionEnd; $i++) {
                if ($lines[$i] -like '| **Deleted** |*') {
                    $lines[$i] = $deletedRow
                    $updated = $true
                    break
                }
            }

            if (-not $updated) {
                $insertAt = $sectionEnd
                for ($i = $metadataIndex + 1; $i -lt $sectionEnd; $i++) {
                    if ([string]::IsNullOrWhiteSpace($lines[$i])) {
                        $insertAt = $i
                        break
                    }
                }
                $lines.Insert($insertAt, $deletedRow)
            }
        }
        else {
            if ($lines.Count -gt 0 -and -not [string]::IsNullOrWhiteSpace($lines[$lines.Count - 1])) {
                $lines.Add('')
            }
            $lines.Add('## Metadata')
            $lines.Add('')
            $lines.Add('| Property | Value |')
            $lines.Add('|----------|-------|')
            $lines.Add("| **GUID** | $gpoGuid |")
            $lines.Add($deletedRow)
            $lines.Add('')
        }

        Set-Content -LiteralPath $readmePath -Value $lines -Encoding UTF8
        Write-Log -Message "Updated deleted GPO README: $readmePath"
        return
    }

    if ($DryRunMode) {
        Write-Log -Message "[DRYRUN] Would export report/README for '$displayName' ($gpoGuid)."
        return
    }

    if (-not (Test-Path -LiteralPath $gpoDir)) {
        New-Item -ItemType Directory -Path $gpoDir -Force | Out-Null
    }

    $reportXml = Get-GPOReport -Guid $gpoGuid -ReportType Xml
    Set-Content -LiteralPath $reportPathXml -Value $reportXml -Encoding UTF8

    $reportHtml = Get-GPOReport -Guid $gpoGuid -ReportType Html
    Set-Content -LiteralPath $reportPathHtml -Value $reportHtml -Encoding UTF8

    $readme = New-Object System.Text.StringBuilder
    [void]$readme.AppendLine("# $displayName")
    [void]$readme.AppendLine()
    [void]$readme.AppendLine("## Metadata")
    [void]$readme.AppendLine()
    [void]$readme.AppendLine("| Property | Value |")
    [void]$readme.AppendLine("|----------|-------|")
    [void]$readme.AppendLine("| **GUID** | $gpoGuid |")
    [void]$readme.AppendLine("| **Status** | $($Snapshot.GpoStatus) |")
    [void]$readme.AppendLine("| **Exported** | $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz') |")
    [void]$readme.AppendLine()
    [void]$readme.AppendLine("## Versions")
    [void]$readme.AppendLine()
    [void]$readme.AppendLine("| Part | File/DS | Version |")
    [void]$readme.AppendLine("|----------|-------|-------|")
    [void]$readme.AppendLine("| User | Directory Service | $($Snapshot.UserDSVersion) |")
    [void]$readme.AppendLine("| User | Sysvol | $($Snapshot.UserSysvolVersion) |")
    [void]$readme.AppendLine("| Computer | Directory Service | $($Snapshot.ComputerDSVersion) |")
    [void]$readme.AppendLine("| Computer | Sysvol | $($Snapshot.ComputerSysvolVersion) |")
    [void]$readme.AppendLine()
    
    if ($Snapshot.WmiFilter) {
        $wmiGuidForLink = [string]$Snapshot.WmiFilter.Guid
        if (-not [string]::IsNullOrWhiteSpace($wmiGuidForLink)) {
            $wmiGuidForLink = $wmiGuidForLink.ToLowerInvariant()
        }
        [void]$readme.AppendLine("## WMI Filter")
        [void]$readme.AppendLine()
        [void]$readme.AppendLine("| Property | Value |")
        [void]$readme.AppendLine("|----------|-------|")
        [void]$readme.AppendLine("| **Name** | $($Snapshot.WmiFilter.Name) |")
        [void]$readme.AppendLine("| **GUID** | [$($Snapshot.WmiFilter.Guid)](../../wmi-filters/$wmiGuidForLink.md) |")
        [void]$readme.AppendLine()
    }
    
    [void]$readme.AppendLine("## Files")
    [void]$readme.AppendLine()
    [void]$readme.AppendLine("| File | Content |")
    [void]$readme.AppendLine("|----------|-------|")
    [void]$readme.AppendLine("| [report.html](report.html) | Full GPO report in HTML format |")
    [void]$readme.AppendLine("| [report.xml](report.xml) | Full GPO report in XML format |")
    if (Test-Path -LiteralPath $linksPath) {
        [void]$readme.AppendLine("| [links.md](links.md) | GPO link information (containers and OU hierarchy) |")
    }
    [void]$readme.AppendLine()

    Set-Content -LiteralPath $readmePath -Value $readme.ToString() -Encoding UTF8
}

function ConvertFrom-GpLinkValue {
    param([Parameter(Mandatory = $false)][string]$GpLink)

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

function Get-ContainersWithLinks {
    param(
        [Parameter(Mandatory = $true)][string]$DomainDistinguishedName
    )

    $containers = @()

    $domainObject = $null
    try {
        $domainObject = Get-ADObject -Identity $DomainDistinguishedName -Properties distinguishedName, gPLink, gPOptions
    }
    catch {
        Write-Log -Message "Warning: Failed to query domain object: $($_.Exception.Message)" -Level 'WARN'
    }

    $ouObjects = @()
    try {
        $ouObjects = @(Get-ADOrganizationalUnit -Filter * -Properties distinguishedName, gPLink, gPOptions -ErrorAction Stop)
    }
    catch {
        Write-Log -Message "Warning: Failed to query OUs: $($_.Exception.Message)" -Level 'WARN'
    }

    $allContainers = @()
    if ($null -ne $domainObject) {
        $allContainers += $domainObject
    }
    if ($ouObjects.Count -gt 0) {
        $allContainers += $ouObjects
    }

    foreach ($container in @($allContainers)) {
        $dn = ''
        $gpLinkValue = ''
        $gpOptions = 0

        try {
            if ($container.PSObject.Properties.Match('distinguishedName').Count -gt 0) {
                $dn = [string]$container.PSObject.Properties['distinguishedName'].Value
            }
        }
        catch { }

        if ([string]::IsNullOrWhiteSpace($dn)) {
            continue
        }

        try {
            if ($container.PSObject.Properties.Match('gPLink').Count -gt 0) {
                $gpLinkValue = [string]$container.PSObject.Properties['gPLink'].Value
            }
        }
        catch {
            $gpLinkValue = ''
        }

        try {
            if ($container.PSObject.Properties.Match('gPOptions').Count -gt 0 -and $null -ne $container.PSObject.Properties['gPOptions'].Value) {
                $gpOptions = [int]$container.PSObject.Properties['gPOptions'].Value
            }
        }
        catch {
            $gpOptions = 0
        }

        $links = ConvertFrom-GpLinkValue -GpLink $gpLinkValue

        $containers += [ordered]@{
            DistinguishedName = $dn
            IsDomainRoot = [bool]($dn -eq $DomainDistinguishedName)
            BlockInheritance = [bool]($gpOptions -band 1)
            Links = @($links)
        }
    }

    return @($containers | Sort-Object -Property @{ Expression = { [string]$_.DistinguishedName } })
}

function Export-GpoLinksForEachGpo {
    param(
        [Parameter(Mandatory = $true)][object[]]$Containers,
        [Parameter(Mandatory = $true)][string]$RepoRoot,
        [Parameter(Mandatory = $true)][string[]]$CurrentGpoGuids,
        [Parameter(Mandatory = $false)][switch]$DryRunMode
    )

    $gpoLinksMap = @{}
    foreach ($container in @($Containers)) {
        $rawLinks = Get-PropertyValue -Object $container -Name 'Links'
        if ($null -ne $rawLinks) {
            foreach ($link in @($rawLinks)) {
                if ($null -eq $link) { continue }
                $gpoGuid = Get-PropertyValue -Object $link -Name 'GpoGuid'
                if ($gpoGuid) {
                    $guidStr = ([string]$gpoGuid).ToLowerInvariant()
                    if (-not $gpoLinksMap.ContainsKey($guidStr)) {
                        $gpoLinksMap[$guidStr] = @()
                    }
                    $linkObj = [ordered]@{
                        Order = [int](Get-PropertyValue -Object $link -Name 'Order')
                        GpoGuid = [string](Get-PropertyValue -Object $link -Name 'GpoGuid')
                        Path = [string](Get-PropertyValue -Object $link -Name 'Path')
                        Enforced = [bool](Get-PropertyValue -Object $link -Name 'Enforced')
                        Disabled = [bool](Get-PropertyValue -Object $link -Name 'Disabled')
                        Options = [int](Get-PropertyValue -Object $link -Name 'Options')
                    }
                    $gpoLinksMap[$guidStr] += [ordered]@{
                        DistinguishedName = Get-PropertyValue -Object $container -Name 'DistinguishedName'
                        IsDomainRoot = [bool](Get-PropertyValue -Object $container -Name 'IsDomainRoot')
                        BlockInheritance = [bool](Get-PropertyValue -Object $container -Name 'BlockInheritance')
                        Links = @($linkObj)
                    }
                }
            }
        }
    }

    foreach ($gpoGuid in $gpoLinksMap.Keys) {
        $gpoDir = Join-Path -Path $RepoRoot -ChildPath (Join-Path -Path 'gpos' -ChildPath $gpoGuid)
        $linksPath = Join-Path -Path $gpoDir -ChildPath 'links.md'

        $linksData = [ordered]@{
            Containers = @($gpoLinksMap[$gpoGuid])
        }

        if ($DryRunMode) {
            Write-Log -Message "[DRYRUN] Would write links for GPO $gpoGuid to $linksPath"
        }
        else {
            if (-not (Test-Path -LiteralPath $gpoDir)) {
                New-Item -ItemType Directory -Path $gpoDir -Force | Out-Null
            }
            $jsonContent = ConvertTo-StableJson -InputObject $linksData
            $markdown = New-Object System.Text.StringBuilder
            [void]$markdown.AppendLine("# GPO Links")
            [void]$markdown.AppendLine()
            [void]$markdown.AppendLine("| Container | DomainRoot | BlockInheritance |")
            [void]$markdown.AppendLine("| --- | --- | --- |")
            foreach ($container in @($linksData.Containers | Sort-Object -Property DistinguishedName)) {
                $dn = [string](Get-PropertyValue -Object $container -Name 'DistinguishedName')
                $isDomainRoot = [bool](Get-PropertyValue -Object $container -Name 'IsDomainRoot')
                $blockInheritance = [bool](Get-PropertyValue -Object $container -Name 'BlockInheritance')
                [void]$markdown.AppendLine("| $($dn.Replace('|', '\\|')) | $isDomainRoot | $blockInheritance |")
            }
            [void]$markdown.AppendLine()
            [void]$markdown.AppendLine("## Raw Data")
            [void]$markdown.AppendLine()
            [void]$markdown.AppendLine("~~~json")
            [void]$markdown.AppendLine($jsonContent)
            [void]$markdown.AppendLine("~~~")

            Set-Content -LiteralPath $linksPath -Value $markdown.ToString() -Encoding UTF8

            $legacyJsonPath = Join-Path -Path $gpoDir -ChildPath 'links.json'
            if (Test-Path -LiteralPath $legacyJsonPath) {
                Remove-Item -LiteralPath $legacyJsonPath -Force
            }
        }
    }

    foreach ($gpoGuid in @($CurrentGpoGuids | ForEach-Object { Normalize-GuidString -GuidValue $_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)) {
        if ($gpoLinksMap.ContainsKey($gpoGuid)) {
            continue
        }

        $gpoDir = Join-Path -Path $RepoRoot -ChildPath (Join-Path -Path 'gpos' -ChildPath $gpoGuid)
        $linksPath = Join-Path -Path $gpoDir -ChildPath 'links.md'
        if (-not (Test-Path -LiteralPath $linksPath)) {
            continue
        }

        if ($DryRunMode) {
            Write-Log -Message "[DRYRUN] Would remove stale links file for GPO $gpoGuid at $linksPath"
        }
        else {
            Remove-Item -LiteralPath $linksPath -Force
        }
    }
}

function Get-GpoLinksFromFile {
    param(
        [Parameter(Mandatory = $true)][string]$GpoDir
    )

    $linksPath = Join-Path -Path $GpoDir -ChildPath 'links.md'
    if (-not (Test-Path -LiteralPath $linksPath)) {
        return $null
    }

    try {
        $rawContent = Get-Content -LiteralPath $linksPath -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($rawContent)) {
            return $null
        }

        $jsonMatch = [regex]::Match($rawContent, '(?s)~~~json\s*(?<json>.*?)\s*~~~')
        if (-not $jsonMatch.Success) {
            Write-Log -Message "Warning: No JSON code block found in links file '$linksPath'." -Level 'WARN'
            return $null
        }

        $jsonContent = [string]$jsonMatch.Groups['json'].Value
        if ([string]::IsNullOrWhiteSpace($jsonContent)) {
            return $null
        }

        $parsed = $jsonContent | ConvertFrom-Json -Depth 20
        return $parsed
    }
    catch {
        Write-Log -Message "Warning: Failed to parse links file '$linksPath': $($_.Exception.Message)" -Level 'WARN'
        return $null
    }
}

function Get-WmiQueryFromAdObject {
    param(
        [Parameter(Mandatory = $true)][object]$AdObject
    )

    $parmNames = @('msWMI-Parm1', 'msWMI-Parm2', 'msWMI-Parm3', 'msWMI-Parm4')
    $values = New-Object 'System.Collections.Generic.List[string]'

    foreach ($parmName in $parmNames) {
        $parmValue = $null
        try {
            if ($AdObject.PSObject.Properties[$parmName]) {
                $parmValue = $AdObject.PSObject.Properties[$parmName].Value
            }
        }
        catch { }

        if ($null -eq $parmValue) {
            continue
        }

        foreach ($valuePart in @($parmValue)) {
            if ($null -eq $valuePart) { continue }
            $valueString = [string]$valuePart
            if (-not [string]::IsNullOrWhiteSpace($valueString)) {
                $values.Add($valueString)
            }
        }
    }

    if ($values.Count -eq 0) {
        return ''
    }

    # Prefer a real WQL query even when Parm1 contains display/description text.
    $wqlRegex = [regex]'(?is)\bselect\b.+?\bfrom\b.+?(?=(?:\r?\n\r?\n)|\z)'
    foreach ($value in $values) {
        $valueTrimmed = $value.Trim()
        $queryMatch = $wqlRegex.Match($valueTrimmed)
        if ($queryMatch.Success) {
            return [string]$queryMatch.Value.Trim()
        }
    }

    # Fallback: use the last non-empty parm value because query data is commonly stored after metadata/description.
    return [string]$values[$values.Count - 1].Trim()
}

function ConvertTo-NormalizedWmiRawData {
    param(
        [Parameter(Mandatory = $false)][object]$RawData,
        [Parameter(Mandatory = $false)][string]$FallbackQuery = ''
    )

    $normalized = [ordered]@{}
    foreach ($parmName in @('msWMI-Parm1', 'msWMI-Parm2', 'msWMI-Parm3', 'msWMI-Parm4')) {
        $value = $null
        if ($null -ne $RawData) {
            $value = Get-PropertyValue -Object $RawData -Name $parmName
        }

        $parts = @()
        if ($null -ne $value) {
            foreach ($part in @($value)) {
                if ($null -eq $part) { continue }
                $text = [string]$part
                $text = $text -replace "`r`n", "`n"
                $text = $text -replace "`r", "`n"
                $text = $text.Trim()
                if (-not [string]::IsNullOrWhiteSpace($text)) {
                    $parts += $text
                }
            }
        }

        $normalized[$parmName] = @($parts)
    }

    $query = ''
    if ($null -ne $RawData) {
        $query = [string](Get-PropertyValue -Object $RawData -Name 'ExtractedQuery')
    }
    if ([string]::IsNullOrWhiteSpace($query)) {
        $query = [string]$FallbackQuery
    }
    $query = $query -replace "`r`n", "`n"
    $query = $query -replace "`r", "`n"
    $query = $query.Trim()
    $normalized['ExtractedQuery'] = $query

    return $normalized
}

function Get-WmiRawDataFromAdObject {
    param(
        [Parameter(Mandatory = $true)][object]$AdObject
    )

    $rawData = [ordered]@{}
    foreach ($parmName in @('msWMI-Parm1', 'msWMI-Parm2', 'msWMI-Parm3', 'msWMI-Parm4')) {
        $parmValue = $null
        try {
            if ($AdObject.PSObject.Properties[$parmName]) {
                $parmValue = $AdObject.PSObject.Properties[$parmName].Value
            }
        }
        catch { }
        $rawData[$parmName] = $parmValue
    }

    $query = Get-WmiQueryFromAdObject -AdObject $AdObject
    $rawData['ExtractedQuery'] = $query
    return (ConvertTo-NormalizedWmiRawData -RawData $rawData -FallbackQuery $query)
}

function Export-WmiFiltersSnapshot {
    param(
        [Parameter(Mandatory = $true)][string]$DomainDistinguishedName,
        [Parameter(Mandatory = $true)][array]$CurrentGpos,
        [Parameter(Mandatory = $true)][switch]$DryRunMode,
        [Parameter(Mandatory = $true)][string]$OutputDirectory,
        [Parameter(Mandatory = $false)][array]$PreviousFilters = @(),
        [Parameter(Mandatory = $false)][switch]$MarkDeletedFromPrevious
    )

    $wmiBase = "CN=SOM,CN=WMIPolicy,CN=System,$DomainDistinguishedName"
    $filters = @()

    try {
        # Query all msWMI-Parm* properties to find query data (may be in Parm1, Parm2, Parm3, Parm4)
        $adFilters = Get-ADObject -SearchBase $wmiBase -LDAPFilter '(objectClass=msWMI-Som)' -Properties 'msWMI-Name', 'msWMI-ID', 'msWMI-Parm1', 'msWMI-Parm2', 'msWMI-Parm3', 'msWMI-Parm4', 'distinguishedName' |
            Sort-Object -Property @{ Expression = { [string]$_.'msWMI-Name' } }

        foreach ($item in $adFilters) {
            $guid = Normalize-GuidString -GuidValue $item.'msWMI-ID'
            $rawData = Get-WmiRawDataFromAdObject -AdObject $item
            $query = [string](Get-PropertyValue -Object $rawData -Name 'ExtractedQuery')

            $filters += [ordered]@{
                Guid = [string]$guid
                Name = [string]$item.'msWMI-Name'
                Query = $query
                DistinguishedName = [string]$item.DistinguishedName
                RawData = $rawData
            }
        }
    }
    catch {
        throw "Failed to query WMI filters from Active Directory path '$wmiBase'. $($_.Exception.Message)"
    }

    $gpoToFilterMap = @()
    $gpoArray = @($CurrentGpos)
    foreach ($gpo in ($gpoArray | Sort-Object -Property @{ Expression = { [string]$_.DisplayName } })) {
        try {
            if ($gpo.WmiFilter -and $gpo.WmiFilter.Path) {
                $filterGuid = Normalize-GuidString -GuidValue ($gpo.WmiFilter.Path -replace '.*\{([^\}]+)\}.*', '$1')
                $gpoToFilterMap += [ordered]@{
                    GpoGuid = Normalize-GuidString -GuidValue $gpo.Id.Guid
                    GpoDisplayName = [string]$gpo.DisplayName
                    WmiFilterGuid = $filterGuid
                    WmiFilterName = [string]$gpo.WmiFilter.Name
                }
            }
        }
        catch {
            # Silently skip GPOs with inaccessible WmiFilter properties
        }
    }

    $snapshot = [ordered]@{
        Filters = @($filters | Sort-Object -Property Name, Guid)
        GpoAssignments = @($gpoToFilterMap | Sort-Object -Property GpoDisplayName, GpoGuid)
    }

    if ($DryRunMode) {
        Write-Log -Message "[DRYRUN] Would write WMI filters to $OutputDirectory"
    }
    else {
        if (-not (Test-Path -LiteralPath $OutputDirectory)) {
            New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
        }
        
        # Write individual filter files
        foreach ($filter in $filters) {
            $filterPath = Join-Path -Path $OutputDirectory -ChildPath "$($filter.Guid).md"
            $readme = New-Object System.Text.StringBuilder
            [void]$readme.AppendLine("# $($filter.Name)")
            [void]$readme.AppendLine()
            [void]$readme.AppendLine('| Property | Value |')
            [void]$readme.AppendLine('| --- | --- |')
            [void]$readme.AppendLine("| Guid | $($filter.Guid) |")
            [void]$readme.AppendLine("| Name | $($filter.Name) |")
            [void]$readme.AppendLine("| DistinguishedName | $($filter.DistinguishedName) |")
            [void]$readme.AppendLine()
            [void]$readme.AppendLine('## Raw Data')
            [void]$readme.AppendLine()
            [void]$readme.AppendLine('~~~json')
            [void]$readme.AppendLine((ConvertTo-StableJson -InputObject $filter.RawData))
            [void]$readme.AppendLine('~~~')
            Set-Content -LiteralPath $filterPath -Value $readme.ToString() -Encoding UTF8
        }
    }

    if ($MarkDeletedFromPrevious) {
        $currentGuids = New-Object 'System.Collections.Generic.HashSet[string]'
        foreach ($filter in @($filters)) {
            if ($null -eq $filter) { continue }
            $guid = Normalize-GuidString -GuidValue (Get-PropertyValue -Object $filter -Name 'Guid')
            if (-not [string]::IsNullOrWhiteSpace($guid)) {
                [void]$currentGuids.Add($guid)
            }
        }

        foreach ($prevFilter in @($PreviousFilters)) {
            if ($null -eq $prevFilter) { continue }
            $prevGuid = Normalize-GuidString -GuidValue (Get-PropertyValue -Object $prevFilter -Name 'Guid')
            if ([string]::IsNullOrWhiteSpace($prevGuid)) { continue }
            if (-not $currentGuids.Contains($prevGuid)) {
                Mark-WmiFilterAsDeleted -OutputDirectory $OutputDirectory -FilterSnapshot $prevFilter -DryRunMode:$DryRunMode
            }
        }
    }

    return $snapshot
}

function Mark-WmiFilterAsDeleted {
    param(
        [Parameter(Mandatory = $true)][string]$OutputDirectory,
        [Parameter(Mandatory = $true)][object]$FilterSnapshot,
        [Parameter(Mandatory = $false)][switch]$DryRunMode
    )

    $filterGuid = Normalize-GuidString -GuidValue (Get-PropertyValue -Object $FilterSnapshot -Name 'Guid')
    if ([string]::IsNullOrWhiteSpace($filterGuid)) {
        return
    }

    $filterName = [string](Get-PropertyValue -Object $FilterSnapshot -Name 'Name')
    if ([string]::IsNullOrWhiteSpace($filterName)) {
        $filterName = $filterGuid
    }

    $filterPath = Join-Path -Path $OutputDirectory -ChildPath "$filterGuid.md"
    $deletedAt = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss zzz')

    if ($DryRunMode) {
        Write-Log -Message "[DRYRUN] Would mark deleted WMI filter '$filterName' ($filterGuid)."
        return
    }

    if (-not (Test-Path -LiteralPath $OutputDirectory)) {
        New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
    }

    if (-not (Test-Path -LiteralPath $filterPath)) {
        $readme = New-Object System.Text.StringBuilder
        [void]$readme.AppendLine("# $filterName")
        [void]$readme.AppendLine()
        [void]$readme.AppendLine('## Metadata')
        [void]$readme.AppendLine()
        [void]$readme.AppendLine('| Property | Value |')
        [void]$readme.AppendLine('|----------|-------|')
        [void]$readme.AppendLine("| **GUID** | $filterGuid |")
        [void]$readme.AppendLine("| **Name** | $filterName |")
        [void]$readme.AppendLine("| **Deleted** | $deletedAt |")
        [void]$readme.AppendLine()
        Set-Content -LiteralPath $filterPath -Value $readme.ToString() -Encoding UTF8
        return
    }

    $lines = New-Object 'System.Collections.Generic.List[string]'
    foreach ($line in Get-Content -LiteralPath $filterPath -Encoding UTF8) {
        $lines.Add([string]$line)
    }

    $deletedRow = "| **Deleted** | $deletedAt |"
    $metadataIndex = -1
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -eq '## Metadata') {
            $metadataIndex = $i
            break
        }
    }

    if ($metadataIndex -ge 0) {
        $sectionEnd = $lines.Count
        for ($i = $metadataIndex + 1; $i -lt $lines.Count; $i++) {
            if ($lines[$i] -match '^## ') {
                $sectionEnd = $i
                break
            }
        }

        $updated = $false
        for ($i = $metadataIndex + 1; $i -lt $sectionEnd; $i++) {
            if ($lines[$i] -like '| **Deleted** |*') {
                $lines[$i] = $deletedRow
                $updated = $true
                break
            }
        }

        if (-not $updated) {
            $insertAt = $sectionEnd
            for ($i = $metadataIndex + 1; $i -lt $sectionEnd; $i++) {
                if ([string]::IsNullOrWhiteSpace($lines[$i])) {
                    $insertAt = $i
                    break
                }
            }
            $lines.Insert($insertAt, $deletedRow)
        }
    }
    else {
        if ($lines.Count -gt 0 -and -not [string]::IsNullOrWhiteSpace($lines[$lines.Count - 1])) {
            $lines.Add('')
        }
        $lines.Add('## Metadata')
        $lines.Add('')
        $lines.Add('| Property | Value |')
        $lines.Add('|----------|-------|')
        $lines.Add("| **GUID** | $filterGuid |")
        $lines.Add("| **Name** | $filterName |")
        $lines.Add($deletedRow)
        $lines.Add('')
    }

    Set-Content -LiteralPath $filterPath -Value $lines -Encoding UTF8
    Write-Log -Message "Marked deleted WMI filter: $filterPath"
}

function Normalize-GuidString {
    param([Parameter(Mandatory = $false)][object]$GuidValue)

    if ($null -eq $GuidValue) { return '' }

    $guid = [string]$GuidValue
    if ([string]::IsNullOrWhiteSpace($guid)) { return '' }

    $guid = $guid.Trim()
    if ($guid.StartsWith('{') -and $guid.EndsWith('}')) {
        $guid = $guid.Trim('{}')
    }

    return $guid.ToLowerInvariant()
}

function Get-PropertyValue {
    param(
        [Parameter(Mandatory = $false)][object]$Object,
        [Parameter(Mandatory = $true)][string]$Name
    )
    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) { return $Object[$Name] }
    $prop = $Object.PSObject.Properties[$Name]
    if ($null -ne $prop) { return $prop.Value }
    return $null
}

function ConvertTo-NormalizedLinksSnapshot {
    param([Parameter(Mandatory = $false)][object]$Snapshot)

    $normalizedContainers = @()
    if ($null -ne $Snapshot) {
        $rawContainers = Get-PropertyValue -Object $Snapshot -Name 'Containers'
        if ($null -ne $rawContainers) {
            foreach ($c in @($rawContainers)) {
                if ($null -eq $c) { continue }
                $normalizedLinks = @()
                $rawLinks = Get-PropertyValue -Object $c -Name 'Links'
                if ($null -ne $rawLinks) {
                    foreach ($l in @($rawLinks)) {
                        if ($null -eq $l) { continue }
                        $normalizedLinks += [ordered]@{
                            Disabled = [bool](Get-PropertyValue -Object $l -Name 'Disabled')
                            Enforced = [bool](Get-PropertyValue -Object $l -Name 'Enforced')
                            GpoGuid  = [string](Get-PropertyValue -Object $l -Name 'GpoGuid')
                            Options  = [int](Get-PropertyValue -Object $l -Name 'Options')
                            Order    = [int](Get-PropertyValue -Object $l -Name 'Order')
                            Path     = [string](Get-PropertyValue -Object $l -Name 'Path')
                        }
                    }
                }
                $normalizedContainers += [ordered]@{
                    BlockInheritance  = [bool](Get-PropertyValue -Object $c -Name 'BlockInheritance')
                    DistinguishedName = [string](Get-PropertyValue -Object $c -Name 'DistinguishedName')
                    IsDomainRoot      = [bool](Get-PropertyValue -Object $c -Name 'IsDomainRoot')
                    Links             = @($normalizedLinks)
                }
            }
        }
    }

    return [ordered]@{
        Containers              = @($normalizedContainers | Sort-Object -Property @{ Expression = { [string]$_.DistinguishedName } })
        DomainDistinguishedName = [string](Get-PropertyValue -Object $Snapshot -Name 'DomainDistinguishedName')
        DomainDnsRoot           = [string](Get-PropertyValue -Object $Snapshot -Name 'DomainDnsRoot')
    }
}

function Get-LinkDiffs {
    param(
        [Parameter(Mandatory = $true)][object]$PreviousSnapshot,
        [Parameter(Mandatory = $true)][object]$CurrentSnapshot
    )

    if ($null -eq $PreviousSnapshot) {
        $PreviousSnapshot = [ordered]@{ Containers = @() }
    }
    if ($null -eq $CurrentSnapshot) {
        $CurrentSnapshot = [ordered]@{ Containers = @() }
    }
    if ($null -eq $PreviousSnapshot.Containers) {
        $PreviousSnapshot.Containers = @()
    }
    if ($null -eq $CurrentSnapshot.Containers) {
        $CurrentSnapshot.Containers = @()
    }

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
        if ($null -eq $container) { continue }
        $dn = [string](Get-PropertyValue -Object $container -Name 'DistinguishedName')
        if (-not [string]::IsNullOrEmpty($dn)) { $previousByDn[$dn] = $container }
    }

    $currentByDn = @{}
    foreach ($container in @($CurrentSnapshot.Containers)) {
        if ($null -eq $container) { continue }
        $dn = [string](Get-PropertyValue -Object $container -Name 'DistinguishedName')
        if (-not [string]::IsNullOrEmpty($dn)) { $currentByDn[$dn] = $container }
    }

    $allDns = @(
        @($previousByDn.Keys) + @($currentByDn.Keys) |
            ForEach-Object { [string]$_ } |
            Sort-Object -Unique
    )
    $affected = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($dn in $allDns) {
        $previousContainer = $previousByDn[$dn]
        $currentContainer = $currentByDn[$dn]
        $prevLinksRaw = Get-PropertyValue -Object $previousContainer -Name 'Links'
        $currLinksRaw = Get-PropertyValue -Object $currentContainer -Name 'Links'
        $prevLinks = if ($null -ne $prevLinksRaw) { @($prevLinksRaw) } else { @() }
        $currLinks = if ($null -ne $currLinksRaw) { @($currLinksRaw) } else { @() }
        $prevLinksJson = ConvertTo-StableJson -InputObject $prevLinks
        $currLinksJson = ConvertTo-StableJson -InputObject $currLinks
        $prevBlock = [bool](Get-PropertyValue -Object $previousContainer -Name 'BlockInheritance')
        $currBlock = [bool](Get-PropertyValue -Object $currentContainer -Name 'BlockInheritance')

        if (($prevLinksJson -ne $currLinksJson) -or ($prevBlock -ne $currBlock)) {
            foreach ($link in (@($prevLinks) + @($currLinks))) {
                if ($null -eq $link) { continue }
                $gpoGuid = Get-PropertyValue -Object $link -Name 'GpoGuid'
                if ($gpoGuid) {
                    [void]$affected.Add(([string]$gpoGuid).ToLowerInvariant())
                }
            }
        }
    }

    $changedGuids = @($affected | ForEach-Object { [string]$_ } | Sort-Object -Unique)

    return [ordered]@{
        HasChanges = ($affected.Count -gt 0)
        ChangedGpoGuids = $changedGuids
        PreviousHash = $previousHash
        CurrentHash = $currentHash
    }
}

function Write-ReportMarkdown {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [hashtable]$Summary,
        [Parameter(Mandatory = $false)]
        [switch]$DryRunMode
    )

    $markdown = New-Object System.Text.StringBuilder
    [void]$markdown.AppendLine('# GPO Sync Report')
    [void]$markdown.AppendLine()
    [void]$markdown.AppendLine("Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz')")
    [void]$markdown.AppendLine()
    [void]$markdown.AppendLine("Summary: changed=$($Summary.Changed.Count), new=$($Summary.New.Count), deleted=$($Summary.Deleted.Count), link-changed=$($Summary.LinkChanged.Count)")
    [void]$markdown.AppendLine()

    # Combine all items into a single list and sort by DisplayName, then Detail
    $allItems = @(
        $Summary.Changed +
        $Summary.New +
        $Summary.Deleted +
        $Summary.LinkChanged |
        Sort-Object -Property @{ Expression = { [string]$_.DisplayName } }, @{ Expression = { [string]$_.Detail } }
    )

    if ($allItems.Count -gt 0) {
        [void]$markdown.AppendLine('| Type | Name | Guid | Detail |')
        [void]$markdown.AppendLine('| --- | --- | --- | --- |')
        foreach ($item in $allItems) {
            $type = ([string]$item.Type).Replace('|', '\|')
            $name = ([string]$item.DisplayName).Replace('|', '\|')
            $guidRaw = [string]$item.Guid
            $guid = $guidRaw.Replace('|', '\|')
            $guidForPath = Normalize-GuidString -GuidValue $guidRaw
            $guidLink = ''
            if (-not [string]::IsNullOrWhiteSpace($guidForPath)) {
                $itemType = ([string]$item.Type).ToLowerInvariant()
                if ($itemType -eq 'filter') {
                    $guidLink = "[$guid](../wmi-filters/$guidForPath.md)"
                }
                else {
                    $guidLink = "[$guid](../gpos/$guidForPath)"
                }
            }
            $detail = ([string]$item.Detail).Replace('|', '\|')
            [void]$markdown.AppendLine("| $type | $name | $guidLink | $detail |")
        }
        [void]$markdown.AppendLine()
    }

    if ($DryRunMode) {
        Write-Log -Message "[DRYRUN] Would generate Markdown report at $Path"
        return
    }

    $parent = Split-Path -Path $Path -Parent
    if (-not (Test-Path -LiteralPath $parent)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }

    Set-Content -LiteralPath $Path -Value $markdown.ToString() -Encoding UTF8
}

function Update-RepoReadmeReportsTable {
    param(
        [Parameter(Mandatory = $true)][string]$RepoRoot,
        [Parameter(Mandatory = $true)][string]$ReportsPath,
        [Parameter(Mandatory = $false)][int]$MaxReports = 25,
        [Parameter(Mandatory = $false)][switch]$DryRunMode
    )

    $readmePath = Join-Path -Path $RepoRoot -ChildPath 'README.md'

    $reportFiles = @()
    if (Test-Path -LiteralPath $ReportsPath) {
        $reportFiles = @(
            Get-ChildItem -LiteralPath $ReportsPath -File -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -like 'g2g-*.md' -or $_.Name -like 'run-*.md' } |
                Sort-Object -Property Name -Descending |
                Select-Object -First $MaxReports
        )
    }

    function Get-ReportSummaryFromFile {
        param([Parameter(Mandatory = $true)][string]$FilePath)

        try {
            $summaryLine = Get-Content -LiteralPath $FilePath -Encoding UTF8 | Where-Object { $_ -like 'Summary:*' } | Select-Object -First 1
            if ([string]::IsNullOrWhiteSpace($summaryLine)) {
                return '_No summary found_'
            }

            $summary = [string]$summaryLine.Substring(8).Trim()
            if ([string]::IsNullOrWhiteSpace($summary)) {
                return '_No summary found_'
            }

            return $summary.Replace('|', '\|')
        }
        catch {
            return '_Summary unavailable_'
        }
    }

    $content = New-Object System.Text.StringBuilder
    [void]$content.AppendLine('# GPO Sync Reports')
    [void]$content.AppendLine()
    [void]$content.AppendLine("Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz')")
    [void]$content.AppendLine()
    [void]$content.AppendLine('| Run (Local) | Report | Summary |')
    [void]$content.AppendLine('| --- | --- | --- |')

    if ($reportFiles.Count -eq 0) {
        [void]$content.AppendLine('| - | _No reports yet_ | - |')
    }
    else {
        foreach ($file in $reportFiles) {
            $timestamp = [string]::Empty
            if ($file.BaseName -match '^(?:g2g|run)-(\d{8})-(\d{6})$') {
                $localTimestampString = "{0}-{1}-{2} {3}:{4}:{5}" -f $Matches[1].Substring(0, 4), $Matches[1].Substring(4, 2), $Matches[1].Substring(6, 2), $Matches[2].Substring(0, 2), $Matches[2].Substring(2, 2), $Matches[2].Substring(4, 2)
                $localTimestamp = [datetime]::ParseExact($localTimestampString, "yyyy-MM-dd HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture)
                $localTimestamp = [datetime]::SpecifyKind($localTimestamp, [System.DateTimeKind]::Local)
                $timestamp = $localTimestamp.ToString('yyyy-MM-dd HH:mm:ss')
            }
            else {
                $timestamp = $file.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')
            }

            $relativePath = "reports/$($file.Name)"
            $summary = Get-ReportSummaryFromFile -FilePath $file.FullName
            [void]$content.AppendLine("| $timestamp | [$($file.Name)]($relativePath) | $summary |")
        }
    }

    [void]$content.AppendLine()

    if ($DryRunMode) {
        Write-Log -Message "[DRYRUN] Would update README report table at $readmePath"
        return
    }

    Set-Content -LiteralPath $readmePath -Value $content -Encoding UTF8
}

function Initialize-RepositoryStructure {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Root,
        [Parameter(Mandatory = $false)]
        [switch]$DryRunMode
    )

    $requiredDirs = @('gpos', 'wmi-filters', 'reports')
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

    Initialize-RepositoryStructure -Root $RepoPath -DryRunMode:$DryRun

    $runTimestamp = (Get-Date).ToString('yyyyMMdd-HHmmss')
    Write-Log -Message "Starting GPO sync. RepoPath=$RepoPath DryRun=$($DryRun.IsPresent)"

    try {
        Import-Module GroupPolicy -SkipEditionCheck -ErrorAction Stop
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

    if (-not $DisableGIT) {
        $gitCmd = Get-Command git -ErrorAction SilentlyContinue
        if (-not $gitCmd) {
            Fail -Message 'git executable was not found in PATH.' -Code 12
        }
    }

    $stateFileAbsPath = if ([System.IO.Path]::IsPathRooted($StateFilePath)) { $StateFilePath } else { Join-Path -Path $RepoPath -ChildPath $StateFilePath }
    $wmiOutputDir = Join-Path -Path $RepoPath -ChildPath 'wmi-filters'
    $reportsPath = Join-Path -Path $RepoPath -ChildPath 'reports'

    $previousState = [ordered]@{ Gpos = [ordered]@{}; WmiFilters = @() }
    if (Test-Path -LiteralPath $stateFileAbsPath) {
        try {
            $rawState = Get-Content -LiteralPath $stateFileAbsPath -Raw -Encoding UTF8
            if (-not [string]::IsNullOrWhiteSpace($rawState)) {
                $parsed = $rawState | ConvertFrom-Json -Depth 20
                if ($parsed.Gpos) {
                    $gposOrdered = [ordered]@{}
                    foreach ($prop in $parsed.Gpos.PSObject.Properties | Sort-Object -Property Name) {
                        $normalizedGuid = Normalize-GuidString -GuidValue $prop.Name
                        if ([string]::IsNullOrWhiteSpace($normalizedGuid)) {
                            $normalizedGuid = ([string]$prop.Name).ToLowerInvariant()
                        }
                        $gposOrdered[$normalizedGuid] = $prop.Value
                    }
                    $previousState.Gpos = $gposOrdered
                }
                if ($parsed.WmiFilters) {
                    $previousState.WmiFilters = @($parsed.WmiFilters)
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
        $entryGuid = Normalize-GuidString -GuidValue $entry.Guid
        $currentSnapshotByGuid[$entryGuid] = $entry
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
            Write-Log -Message "New: $guid, $($current.DisplayName)"
            $newItems.Add([ordered]@{ Type = 'GPO'; DisplayName = $current.DisplayName; Guid = $guid; Detail = 'New GPO' })
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

            Write-Log -Message "Changed: $guid, $($current.DisplayName) - $($detailParts -join '; ')"
            $changed.Add([ordered]@{
                    Type = 'GPO'
                    DisplayName = $current.DisplayName
                    Guid = $guid
                    Detail = ($detailParts -join '; ')
                })
            [void]$exportTargets.Add($guid)
        }
    }

    foreach ($guid in $previousState.Gpos.Keys) {
        $normalizedGuid = Normalize-GuidString -GuidValue $guid
        if (-not $currentSnapshotByGuid.Contains($normalizedGuid)) {
            $deletedDisplay = [string]$previousState.Gpos[$guid].DisplayName
            Write-Log -Message "Deleted: $normalizedGuid, $deletedDisplay"
            $deleted.Add([ordered]@{ Type = 'GPO'; DisplayName = $deletedDisplay; Guid = $normalizedGuid; Detail = 'Deleted GPO' })
            $previousSnapshot = $previousState.Gpos[$guid]
            Export-GpoArtifacts -RepoRoot $RepoPath -Snapshot $previousSnapshot -Deleted -DryRunMode:$DryRun
        }
    }

    $containers = @()
    try {
        $containers = Get-ContainersWithLinks -DomainDistinguishedName $domain.DistinguishedName
    }
    catch {
        $exType = $_.Exception.GetType().FullName
        Fail -Message "Failed to query AD containers: $($_.Exception.Message) [$exType]" -Code 11
    }

    try {
        $wmiSnapshot = Export-WmiFiltersSnapshot -DomainDistinguishedName $domain.DistinguishedName -CurrentGpos $currentGpos -DryRunMode:$DryRun -OutputDirectory $wmiOutputDir -PreviousFilters $previousState.WmiFilters -MarkDeletedFromPrevious
    }
    catch {
        $exType = $_.Exception.GetType().FullName
        Fail -Message "Failed WMI snapshot export: $($_.Exception.Message) [$exType]" -Code 11
    }

    # Track WMI filter changes
    $previousWmiFilters = Get-PropertyValue -Object $previousState -Name 'WmiFilters'
    if ($null -eq $previousWmiFilters) { $previousWmiFilters = @() }
    $currentWmiFilters = @($wmiSnapshot.Filters)
    
    $previousFiltersByGuid = @{}
    foreach ($filter in @($previousWmiFilters)) {
        if ($null -ne $filter) {
            $filterGuid = Normalize-GuidString -GuidValue (Get-PropertyValue -Object $filter -Name 'Guid')
            if ($filterGuid) { $previousFiltersByGuid[$filterGuid] = $filter }
        }
    }
    
    $currentFiltersByGuid = @{}
    foreach ($filter in $currentWmiFilters) {
        if ($null -ne $filter) {
            $filterGuid = Normalize-GuidString -GuidValue (Get-PropertyValue -Object $filter -Name 'Guid')
            if ($filterGuid) { $currentFiltersByGuid[$filterGuid] = $filter }
        }
    }
    
    # Detect new and changed filters
    foreach ($filterGuid in $currentFiltersByGuid.Keys) {
        $current = $currentFiltersByGuid[$filterGuid]
        $previous = $previousFiltersByGuid[$filterGuid]
        
        if (-not $previous) {
            $filterName = [string](Get-PropertyValue -Object $current -Name 'Name')
            Write-Log -Message "New: $filterGuid, $filterName (WMI Filter)"
            $newItems.Add([ordered]@{ Type = 'Filter'; DisplayName = $filterName; Guid = $filterGuid; Detail = 'New WMI Filter' })
        }
        else {
            $prevQuery = [string](Get-PropertyValue -Object $previous -Name 'Query')
            $currQuery = [string](Get-PropertyValue -Object $current -Name 'Query')
            $prevName = [string](Get-PropertyValue -Object $previous -Name 'Name')
            $currName = [string](Get-PropertyValue -Object $current -Name 'Name')
            $prevRawData = ConvertTo-NormalizedWmiRawData -RawData (Get-PropertyValue -Object $previous -Name 'RawData') -FallbackQuery $prevQuery
            $currRawData = ConvertTo-NormalizedWmiRawData -RawData (Get-PropertyValue -Object $current -Name 'RawData') -FallbackQuery $currQuery
            $prevRawDataJson = ConvertTo-StableJson -InputObject $prevRawData
            $currRawDataJson = ConvertTo-StableJson -InputObject $currRawData
            $rawDataChanged = ($prevRawDataJson -ne $currRawDataJson)

            if ($prevQuery -ne $currQuery -or $prevName -ne $currName -or $rawDataChanged) {
                $detailParts = @()
                if ($prevName -ne $currName) { $detailParts += "Name: $prevName -> $currName" }
                if ($prevQuery -ne $currQuery) { $detailParts += "Query changed" }
                if ($rawDataChanged) { $detailParts += "Raw data changed" }
                Write-Log -Message "Changed: $filterGuid, $currName (WMI Filter) - $($detailParts -join '; ')"
                $changed.Add([ordered]@{
                    Type = 'Filter'
                    DisplayName = $currName
                    Guid = $filterGuid
                    Detail = ($detailParts -join '; ')
                })
            }
        }
    }
    
    # Detect deleted filters
    foreach ($filterGuid in $previousFiltersByGuid.Keys) {
        if (-not $currentFiltersByGuid.Contains($filterGuid)) {
            $filterName = [string](Get-PropertyValue -Object $previousFiltersByGuid[$filterGuid] -Name 'Name')
            Write-Log -Message "Deleted: $filterGuid, $filterName (WMI Filter)"
            $deleted.Add([ordered]@{ Type = 'Filter'; DisplayName = $filterName; Guid = $filterGuid; Detail = 'Deleted WMI Filter' })
            # Fallback: explicitly stamp deleted metadata for filters detected as deleted in summary logic.
            Mark-WmiFilterAsDeleted -OutputDirectory $wmiOutputDir -FilterSnapshot $previousFiltersByGuid[$filterGuid] -DryRunMode:$DryRun
        }
    }

    foreach ($guid in $currentSnapshotByGuid.Keys) {
        $gpoDir = Join-Path -Path $RepoPath -ChildPath (Join-Path -Path 'gpos' -ChildPath $guid)

        $currentGpoLinksData = @()
        foreach ($container in $containers) {
            $rawLinks = Get-PropertyValue -Object $container -Name 'Links'
            if ($null -eq $rawLinks) { continue }
            $matchingLink = @($rawLinks | Where-Object {
                ([string](Get-PropertyValue -Object $_ -Name 'GpoGuid')).ToLowerInvariant() -eq $guid
            }) | Select-Object -First 1
            if ($null -eq $matchingLink) { continue }
            $linkObj = [ordered]@{
                Order    = [int](Get-PropertyValue -Object $matchingLink -Name 'Order')
                GpoGuid  = [string](Get-PropertyValue -Object $matchingLink -Name 'GpoGuid')
                Path     = [string](Get-PropertyValue -Object $matchingLink -Name 'Path')
                Enforced = [bool](Get-PropertyValue -Object $matchingLink -Name 'Enforced')
                Disabled = [bool](Get-PropertyValue -Object $matchingLink -Name 'Disabled')
                Options  = [int](Get-PropertyValue -Object $matchingLink -Name 'Options')
            }
            $currentGpoLinksData += [ordered]@{
                DistinguishedName = [string](Get-PropertyValue -Object $container -Name 'DistinguishedName')
                IsDomainRoot      = [bool](Get-PropertyValue -Object $container -Name 'IsDomainRoot')
                BlockInheritance  = [bool](Get-PropertyValue -Object $container -Name 'BlockInheritance')
                Links             = @($linkObj)
            }
        }
        $currentGpoLinksData = @($currentGpoLinksData | Sort-Object -Property @{ Expression = { [string]$_.DistinguishedName } })
        
        $previousLinks = Get-GpoLinksFromFile -GpoDir $gpoDir
        $previousContainers = @()
        if ($null -ne $previousLinks) {
            $previousContainers = @(Get-PropertyValue -Object $previousLinks -Name 'Containers')
        }

        # Normalize both to ensure consistent comparison (ConvertFrom-Json can deserialize differently than native objects)
        $prevNormalized = @()
        foreach ($c in @($previousContainers)) {
            if ($null -eq $c) { continue }
            $normalizedLinks = @()
            $rawLinks = Get-PropertyValue -Object $c -Name 'Links'
            if ($null -ne $rawLinks) {
                foreach ($l in @($rawLinks)) {
                    if ($null -eq $l) { continue }
                    $normalizedLinks += [ordered]@{
                        Disabled = [bool](Get-PropertyValue -Object $l -Name 'Disabled')
                        Enforced = [bool](Get-PropertyValue -Object $l -Name 'Enforced')
                        GpoGuid  = [string](Get-PropertyValue -Object $l -Name 'GpoGuid')
                        Options  = [int](Get-PropertyValue -Object $l -Name 'Options')
                        Order    = [int](Get-PropertyValue -Object $l -Name 'Order')
                        Path     = [string](Get-PropertyValue -Object $l -Name 'Path')
                    }
                }
            }
            $prevNormalized += [ordered]@{
                BlockInheritance  = [bool](Get-PropertyValue -Object $c -Name 'BlockInheritance')
                DistinguishedName = [string](Get-PropertyValue -Object $c -Name 'DistinguishedName')
                IsDomainRoot      = [bool](Get-PropertyValue -Object $c -Name 'IsDomainRoot')
                Links             = @($normalizedLinks)
            }
        }

        $currNormalized = @()
        foreach ($c in @($currentGpoLinksData)) {
            if ($null -eq $c) { continue }
            $normalizedLinks = @()
            $rawLinks = Get-PropertyValue -Object $c -Name 'Links'
            if ($null -ne $rawLinks) {
                foreach ($l in @($rawLinks)) {
                    if ($null -eq $l) { continue }
                    $normalizedLinks += [ordered]@{
                        Disabled = [bool](Get-PropertyValue -Object $l -Name 'Disabled')
                        Enforced = [bool](Get-PropertyValue -Object $l -Name 'Enforced')
                        GpoGuid  = [string](Get-PropertyValue -Object $l -Name 'GpoGuid')
                        Options  = [int](Get-PropertyValue -Object $l -Name 'Options')
                        Order    = [int](Get-PropertyValue -Object $l -Name 'Order')
                        Path     = [string](Get-PropertyValue -Object $l -Name 'Path')
                    }
                }
            }
            $currNormalized += [ordered]@{
                BlockInheritance  = [bool](Get-PropertyValue -Object $c -Name 'BlockInheritance')
                DistinguishedName = [string](Get-PropertyValue -Object $c -Name 'DistinguishedName')
                IsDomainRoot      = [bool](Get-PropertyValue -Object $c -Name 'IsDomainRoot')
                Links             = @($normalizedLinks)
            }
        }

        $prevJson = ConvertTo-StableJson -InputObject $prevNormalized
        $currJson = ConvertTo-StableJson -InputObject $currNormalized
        if ($prevJson -ne $currJson) {
            [void]$exportTargets.Add($guid)

            $alreadyChanged = @($changed | Where-Object { $_.Guid -eq $guid }).Count -gt 0
            $alreadyNew = @($newItems | Where-Object { $_.Guid -eq $guid }).Count -gt 0
            if (-not $alreadyChanged -and -not $alreadyNew) {
                Write-Log -Message "Link-changed: $guid, $($currentSnapshotByGuid[$guid].DisplayName)"
                $linkChanged.Add([ordered]@{
                        Type = 'GPO'
                        DisplayName = $currentSnapshotByGuid[$guid].DisplayName
                        Guid = $guid
                        Detail = 'Container link/order/enforced change'
                    })
            }
        }
    }

    Export-GpoLinksForEachGpo -Containers $containers -RepoRoot $RepoPath -CurrentGpoGuids $currentSnapshotByGuid.Keys -DryRunMode:$DryRun

    foreach ($guid in @($exportTargets) | Sort-Object) {
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
        WmiSnapshotHash = $wmiHash
        WmiFilters = @($currentWmiFilters | Sort-Object -Property Guid)
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
        Update-RepoReadmeReportsTable -RepoRoot $RepoPath -ReportsPath $reportsPath -MaxReports 25 -DryRunMode:$DryRun
        Write-Log -Message 'No changes detected.'
        exit 0
    }

    Write-ReportMarkdown -Path (Join-Path -Path $reportsPath -ChildPath "g2g-$runTimestamp.md") -Summary $summary -DryRunMode:$DryRun
    Update-RepoReadmeReportsTable -RepoRoot $RepoPath -ReportsPath $reportsPath -MaxReports 25 -DryRunMode:$DryRun

    if ($DryRun) {
        Write-Log -Message '[DRYRUN] Skipping git add/commit/push.'
        exit 1
    }

    if ($DisableGIT) {
        Write-Log -Message 'DisableGIT switch is set. Skipping git operations. Artifacts stored locally only.'
        exit 1
    }

    $isGitRepo = $false
    try {
        $null = & git -C $RepoPath rev-parse --git-dir 2>&1
        $isGitRepo = ($LASTEXITCODE -eq 0)
    }
    catch { }

    if (-not $isGitRepo) {
        Write-Log -Message 'RepoPath is not a git repository. Files exported locally; skipping git operations.'
        exit 1
    }

    Push-Location $RepoPath
    try {
        if ([string]::IsNullOrWhiteSpace($Branch)) {
            $Branch = (git rev-parse --abbrev-ref HEAD).Trim()
        }
        if ([string]::IsNullOrWhiteSpace($Branch) -or $Branch -eq 'HEAD') {
            Fail -Message 'Could not determine git branch or repository is in detached HEAD state. Provide -Branch explicitly.' -Code 12
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
            exit 0
        }

        $commitMessagePath = Join-Path -Path $RepoPath -ChildPath "commit-$runTimestamp.txt"
        Set-Content -LiteralPath $commitMessagePath -Value $commitMessage -Encoding UTF8
        & git commit -F $commitMessagePath | Out-Null
        if ($LASTEXITCODE -ne 0) {
            Fail -Message 'git commit failed.' -Code 12
        }
        if (Test-Path -LiteralPath $commitMessagePath) {
            Remove-Item -LiteralPath $commitMessagePath -Force
        }

        & git push origin $Branch | Out-Null
        if ($LASTEXITCODE -ne 0) {
            Write-Log -Message 'git push failed after successful export/commit.' -Level 'ERROR'
            exit 2
        }

        Write-Log -Message 'git push completed successfully.'
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

    $position = $_.InvocationInfo.PositionMessage
    $stack = $_.ScriptStackTrace
    Write-Log -Message "Unhandled failure: $($_.Exception.Message) | $position | $stack" -Level 'ERROR'
    exit 12
}
