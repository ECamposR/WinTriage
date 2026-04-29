[CmdletBinding()]
param(
    [switch]$Quick,
    [switch]$Full,
    [ValidateRange(1, 90)]
    [int]$Days = 7,
    [string]$OutputPath = "C:\ProgramData\WinTriage\Reports",
    [switch]$JsonOnly,
    [switch]$NoColor,
    [switch]$UseExitCode,
    [switch]$OpenReport,
    [switch]$DebugErrors
)

# WinTriage is read-only by design.
# It collects diagnostic data and generates reports without modifying system configuration.

$script:WTVersion = '0.2.2'
$script:WTIsJsonOnly = $JsonOnly.IsPresent
$script:WTNoColor = $NoColor.IsPresent
$script:WTDebugErrors = $DebugErrors.IsPresent

function Test-WTAdministrator {
    [CmdletBinding()]
    param()

    try {
        $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        return $false
    }
}

function ConvertTo-WTYesNo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [bool]$Value
    )

    if ($Value) {
        return 'Yes'
    }

    return 'No'
}

function ConvertTo-WTAbsolutePath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $expandedPath = [System.Environment]::ExpandEnvironmentVariables($Path)
    return [System.IO.Path]::GetFullPath($expandedPath)
}

function Resolve-WTOutputBasePath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RequestedPath
    )

    $requestedAbsolutePath = ConvertTo-WTAbsolutePath -Path $RequestedPath
    $tempRoot = if ($env:TEMP) { $env:TEMP } else { [System.IO.Path]::GetTempPath() }
    $fallbackAbsolutePath = ConvertTo-WTAbsolutePath -Path (Join-Path -Path $tempRoot -ChildPath 'WinTriage\Reports')
    $lastError = $null

    foreach ($candidate in @(
        [pscustomobject]@{ Path = $requestedAbsolutePath; IsFallback = $false },
        [pscustomobject]@{ Path = $fallbackAbsolutePath; IsFallback = $true }
    )) {
        try {
            [void][System.IO.Directory]::CreateDirectory($candidate.Path)
            return [pscustomobject]@{
                BasePath         = $candidate.Path
                RequestedPath    = $requestedAbsolutePath
                FallbackPath     = $fallbackAbsolutePath
                UsedFallback     = $candidate.IsFallback
                FallbackReason   = if ($candidate.IsFallback) { $lastError } else { $null }
                FallbackUsedText = if ($candidate.IsFallback) { 'Yes' } else { 'No' }
            }
        }
        catch {
            $lastError = $_.Exception.Message
        }
    }

    return [pscustomobject]@{
        BasePath         = $null
        RequestedPath    = $requestedAbsolutePath
        FallbackPath     = $fallbackAbsolutePath
        UsedFallback     = $false
        FallbackReason   = $lastError
        FallbackUsedText = 'No'
    }
}

function New-WTReportDirectory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$BasePath,

        [Parameter(Mandatory = $true)]
        [string]$Hostname
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd_HHmmss_fff'
    $root = Join-Path -Path $BasePath -ChildPath $Hostname
    [void][System.IO.Directory]::CreateDirectory($root)

    for ($attempt = 0; $attempt -lt 100; $attempt++) {
        $candidateName = if ($attempt -eq 0) {
            $timestamp
        }
        else {
            '{0}_{1:00}' -f $timestamp, $attempt
        }

        $reportDir = Join-Path -Path $root -ChildPath $candidateName
        if (Test-Path -LiteralPath $reportDir) {
            continue
        }

        try {
            [void][System.IO.Directory]::CreateDirectory($reportDir)
            return (ConvertTo-WTAbsolutePath -Path $reportDir)
        }
        catch {
            if (Test-Path -LiteralPath $reportDir) {
                continue
            }
        }
    }

    throw 'Unable to create a unique report directory.'
}

function New-WTReportObject {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Mode,

        [Parameter(Mandatory = $true)]
        [int]$Days,

        [Parameter(Mandatory = $true)]
        [string]$OutputDirectory,

        [Parameter(Mandatory = $true)]
        [bool]$IsAdmin,

        [Parameter(Mandatory = $true)]
        [string]$OutputBasePath,

        [Parameter(Mandatory = $true)]
        [bool]$UsedOutputFallback
    )

    $startedAt = Get-Date

    $report = [ordered]@{
        Metadata = [ordered]@{
            ToolName         = 'WinTriage'
            Version          = $script:WTVersion
            RunId            = [guid]::NewGuid().ToString()
            Hostname         = $env:COMPUTERNAME
            User             = '{0}\{1}' -f $env:USERDOMAIN, $env:USERNAME
            IsAdmin          = $IsAdmin
            Mode             = $Mode
            Days             = $Days
            StartedAt        = $startedAt.ToString('o')
            FinishedAt       = $null
            OutputPath       = $OutputBasePath
            OutputBasePath   = $OutputBasePath
            OutputFallbackUsed = $UsedOutputFallback
            ReportDirectory    = $OutputDirectory
            JsonPath           = $null
            MarkdownPath       = $null
            JsonGenerated      = $false
            MarkdownGenerated  = $false
            ExitCode           = $null
            PowerShellVersion = $PSVersionTable.PSVersion.ToString()
            IsReadOnly       = $true
        }
        Context = [ordered]@{
            IsDomainJoined = $null
            DomainName     = $null
            Manufacturer   = $null
            Model          = $null
            SerialNumber   = $null
        }
        Raw = [ordered]@{
            System      = $null
            Updates     = $null
            Defender    = $null
            Domain      = $null
            Disk        = $null
            Services    = $null
            Performance = $null
            Events      = $null
        }
        Normalized = [ordered]@{
            System      = $null
            Updates     = $null
            Defender    = $null
            Domain      = $null
            Disk        = $null
            Services    = $null
            Performance = $null
            Events      = $null
        }
        Findings = @()
        Summary = [ordered]@{
            Critical = 0
            High     = 0
            Medium   = 0
            Low      = 0
            Info     = 0
            Skipped  = 0
        }
        Execution = [ordered]@{
            Errors   = @()
            Skipped  = @()
            Timings  = @()
            Warnings = @()
        }
    }

    return [pscustomobject]$report
}

function Add-WTExecutionWarning {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$Report,

        [Parameter(Mandatory = $true)]
        [string]$Scope,

        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $entry = [ordered]@{
        Scope     = $Scope
        Message   = $Message
        Timestamp = (Get-Date).ToString('o')
    }

    $Report.Execution.Warnings += [pscustomobject]$entry
    return $Report
}

function Add-WTFinding {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$Report,

        [Parameter(Mandatory = $true)]
        [string]$Id,

        [Parameter(Mandatory = $true)]
        [string]$Category,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Critical', 'High', 'Medium', 'Low', 'Info', 'Skipped')]
        [string]$Severity,

        [Parameter(Mandatory = $true)]
        [string]$Title,

        [Parameter(Mandatory = $true)]
        [string]$Description,

        [Parameter(Mandatory = $true)]
        [object[]]$Evidence,

        [Parameter(Mandatory = $true)]
        [string]$Recommendation,

        [Parameter(Mandatory = $true)]
        [string]$Source,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Pass', 'Fail', 'Warning', 'Info', 'Skipped', 'Error')]
        [string]$Status,

        [bool]$RequiresAdmin = $false
    )

    if (Test-WTFindingExists -Report $Report -Id $Id -Category $Category -Source $Source) {
        return $Report
    }

    $finding = [ordered]@{
        Id            = $Id
        Category      = $Category
        Severity      = $Severity
        Title         = $Title
        Description   = $Description
        Evidence      = @($Evidence)
        Recommendation = $Recommendation
        Source        = $Source
        Timestamp     = (Get-Date).ToString('o')
        Status        = $Status
        RequiresAdmin = $RequiresAdmin
    }

    $Report.Findings += [pscustomobject]$finding
    return $Report
}

function Test-WTFindingExists {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$Report,

        [Parameter(Mandatory = $true)]
        [string]$Id,

        [Parameter(Mandatory = $true)]
        [string]$Category,

        [Parameter(Mandatory = $true)]
        [string]$Source
    )

    foreach ($finding in @($Report.Findings)) {
        if ($finding.Id -eq $Id -and $finding.Category -eq $Category -and $finding.Source -eq $Source) {
            return $true
        }
    }

    return $false
}

function Add-WTExecutionError {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$Report,

        [Parameter(Mandatory = $true)]
        [string]$Scope,

        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $entry = [ordered]@{
        Scope     = $Scope
        Message   = $Message
        Timestamp = (Get-Date).ToString('o')
    }

    $Report.Execution.Errors += [pscustomobject]$entry
    return $Report
}

function Add-WTSkippedCheck {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$Report,

        [Parameter(Mandatory = $true)]
        [string]$Check,

        [Parameter(Mandatory = $true)]
        [string]$Reason
    )

    $entry = [ordered]@{
        Check     = $Check
        Reason    = $Reason
        Timestamp = (Get-Date).ToString('o')
    }

    $Report.Execution.Skipped += [pscustomobject]$entry
    return $Report
}

function Invoke-WTSafeCollector {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$Report,

        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock,

        [bool]$RequiresAdmin = $false
    )

    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    if ($RequiresAdmin -and -not $Report.Metadata.IsAdmin) {
        Add-WTSkippedCheck -Report $Report -Check $Name -Reason 'Requires administrator privileges.'
        $sw.Stop()
        $Report.Execution.Timings += [pscustomobject]@{
            Name      = $Name
            Status    = 'Skipped'
            DurationMs = [math]::Round($sw.Elapsed.TotalMilliseconds, 2)
            Timestamp  = (Get-Date).ToString('o')
        }
        return $null
    }

    try {
        $result = & $ScriptBlock
        $sw.Stop()
        $Report.Execution.Timings += [pscustomobject]@{
            Name      = $Name
            Status    = 'Success'
            DurationMs = [math]::Round($sw.Elapsed.TotalMilliseconds, 2)
            Timestamp  = (Get-Date).ToString('o')
        }
        return $result
    }
    catch {
        $sw.Stop()
        Add-WTExecutionError -Report $Report -Scope $Name -Message $_.Exception.Message
        $Report.Execution.Timings += [pscustomobject]@{
            Name      = $Name
            Status    = 'Error'
            DurationMs = [math]::Round($sw.Elapsed.TotalMilliseconds, 2)
            Timestamp  = (Get-Date).ToString('o')
        }
        return $null
    }
}

function Invoke-WTSafeStep {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$Report,

        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock
    )

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $result = & $ScriptBlock
        $sw.Stop()
        $Report.Execution.Timings += [pscustomobject]@{
            Name       = $Name
            Status     = 'Success'
            DurationMs = [math]::Round($sw.Elapsed.TotalMilliseconds, 2)
            Timestamp  = (Get-Date).ToString('o')
        }
        return $result
    }
    catch {
        $sw.Stop()
        if ($Report) {
            Add-WTExecutionError -Report $Report -Scope $Name -Message $_.Exception.Message
            $Report.Execution.Timings += [pscustomobject]@{
                Name       = $Name
                Status     = 'Error'
                DurationMs = [math]::Round($sw.Elapsed.TotalMilliseconds, 2)
                Timestamp  = (Get-Date).ToString('o')
            }
        }
        return $null
    }
}

function Write-WTFatalErrorFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter()]
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )

    $tempRoot = if ($env:TEMP) { $env:TEMP } else { [System.IO.Path]::GetTempPath() }
    $basePaths = @(
        'C:\ProgramData\WinTriage\fatal-error.txt',
        (Join-Path -Path $tempRoot -ChildPath 'WinTriage\fatal-error.txt')
    )

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add(('Timestamp: {0}' -f (Get-Date).ToString('o'))) | Out-Null
    $lines.Add(('Message: {0}' -f $Message)) | Out-Null
    if ($ErrorRecord -and $ErrorRecord.InvocationInfo) {
        if ($ErrorRecord.Exception -and $ErrorRecord.Exception.Message) {
            $lines.Add(('Exception: {0}' -f $ErrorRecord.Exception.Message)) | Out-Null
        }
        if ($ErrorRecord.InvocationInfo.ScriptLineNumber) {
            $lines.Add(('ScriptLineNumber: {0}' -f $ErrorRecord.InvocationInfo.ScriptLineNumber)) | Out-Null
        }
        if ($ErrorRecord.InvocationInfo.Line) {
            $lines.Add(('Line: {0}' -f $ErrorRecord.InvocationInfo.Line.Trim())) | Out-Null
        }
    }

    foreach ($candidate in $basePaths) {
        try {
            $parent = Split-Path -Path $candidate -Parent
            if ($parent) {
                [void][System.IO.Directory]::CreateDirectory($parent)
            }
            $content = $lines.ToArray()
            [System.IO.File]::WriteAllLines($candidate, $content, [System.Text.Encoding]::UTF8)
            return $candidate
        }
        catch {
        }
    }

    return $null
}

function Update-WTSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$Report
    )

    $summary = [ordered]@{
        Critical = 0
        High     = 0
        Medium   = 0
        Low      = 0
        Info     = 0
        Skipped  = 0
    }

    foreach ($finding in $Report.Findings) {
        switch ($finding.Severity) {
            'Critical' { $summary.Critical++ }
            'High'     { $summary.High++ }
            'Medium'   { $summary.Medium++ }
            'Low'      { $summary.Low++ }
            'Info'     { $summary.Info++ }
            'Skipped'  { $summary.Skipped++ }
        }
    }

    $Report.Summary = [pscustomobject]$summary
    return $Report
}

function Write-WTConsoleSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$Report,

        [bool]$NoColor = $false
    )

    if ($script:WTIsJsonOnly) {
        return
    }

    $adminText = ConvertTo-WTYesNo -Value ([bool]$Report.Metadata.IsAdmin)
    $readOnlyText = ConvertTo-WTYesNo -Value ([bool]$Report.Metadata.IsReadOnly)
    $fallbackText = ConvertTo-WTYesNo -Value ([bool]$Report.Metadata.OutputFallbackUsed)
    $osSummary = 'Unknown'
    $uptimeText = 'Unknown'
    $diskText = 'Unknown'
    $ramText = 'Unknown'
    $systemDriveLabel = 'C:'

    if ($Report.Normalized.System -and $Report.Normalized.System.OSCaption) {
        $osSummary = '{0} build {1}' -f $Report.Normalized.System.OSCaption, (Get-WTSystemBuildText -SystemInfo $Report.Normalized.System)
    }
    if ($Report.Normalized.System -and $Report.Normalized.System.UptimeDays -ne $null) {
        $uptimeText = '{0} days' -f $Report.Normalized.System.UptimeDays
    }
    if ($Report.Normalized.System -and $Report.Normalized.System.SystemDrive) {
        $systemDriveLabel = $Report.Normalized.System.SystemDrive
    }
    if ($Report.Normalized.Disk) {
        $systemDisk = $null
        foreach ($disk in @($Report.Normalized.Disk)) {
            if ($disk.IsSystemDrive) {
                $systemDisk = $disk
                break
            }
        }
        if ($systemDisk -and $systemDisk.FreePercent -ne $null) {
            $diskText = '{0} {1}% free' -f (ConvertTo-WTDisplayValue -Value $systemDisk.DriveLetter -Fallback $systemDriveLabel), $systemDisk.FreePercent
        }
    }
    if ($Report.Normalized.Performance -and $Report.Normalized.Performance.FreeRamPercent -ne $null) {
        $ramText = '{0}% free' -f $Report.Normalized.Performance.FreeRamPercent
    }
    if ($Report.Metadata.MarkdownGenerated) {
        $markdownPathText = $Report.Metadata.MarkdownPath
    }
    elseif ($script:WTIsJsonOnly) {
        $markdownPathText = '(not generated -JsonOnly)'
    }
    else {
        $markdownPathText = '(not generated)'
    }
    $jsonPathText = if ($Report.Metadata.JsonGenerated) { $Report.Metadata.JsonPath } else { '(not available)' }

    $lines = @(
        ('WinTriage {0} - {1}' -f $Report.Metadata.Version, $Report.Metadata.Hostname),
        ('Mode: {0} | Days: {1} | Admin: {2} | ReadOnly: {3}' -f $Report.Metadata.Mode, $Report.Metadata.Days, $adminText, $readOnlyText),
        ('Critical: {0}  High: {1}  Medium: {2}  Low: {3}  Info: {4}  Skipped: {5}' -f $Report.Summary.Critical, $Report.Summary.High, $Report.Summary.Medium, $Report.Summary.Low, $Report.Summary.Info, $Report.Summary.Skipped),
        ('Findings: {0}  Errors: {1}  Warnings: {2}' -f @($Report.Findings).Count, @($Report.Execution.Errors).Count, @($Report.Execution.Warnings).Count),
        ('OS: {0}' -f $osSummary),
        ('Uptime: {0}' -f $uptimeText),
        ('Disk: {0}' -f $diskText),
        ('RAM: {0}' -f $ramText),
        ('Report directory: {0}' -f $Report.Metadata.ReportDirectory),
        ('JSON: {0}' -f $jsonPathText),
        ('Markdown: {0}' -f $markdownPathText),
        ('Output fallback used: {0}' -f $fallbackText),
        'Read-only diagnostic completed.'
    )

    if ($NoColor) {
        foreach ($line in $lines) {
            Write-Host $line
        }
        return
    }

    Write-Host $lines[0] -ForegroundColor Cyan
    Write-Host $lines[1] -ForegroundColor Gray
    Write-Host $lines[2] -ForegroundColor White
    Write-Host $lines[3] -ForegroundColor DarkGray
    Write-Host $lines[4] -ForegroundColor Gray
    Write-Host $lines[5] -ForegroundColor Gray
    Write-Host $lines[6] -ForegroundColor Gray
    Write-Host $lines[7] -ForegroundColor Gray
    Write-Host $lines[8] -ForegroundColor DarkGray
    Write-Host $lines[9] -ForegroundColor Gray
    Write-Host $lines[10] -ForegroundColor Gray
    Write-Host $lines[11] -ForegroundColor DarkGray
    Write-Host $lines[12] -ForegroundColor Green
}

function Export-WTJsonReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$Report,

        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $json = $Report | ConvertTo-Json -Depth 10
    [System.IO.File]::WriteAllText($Path, $json, [System.Text.Encoding]::UTF8)
    return $Path
}

function Export-WTMarkdownReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$Report,

        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('# WinTriage Report')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('## Summary')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine(('| Severity | Count |'))
    [void]$sb.AppendLine(('| --- | ---: |'))
    [void]$sb.AppendLine(('| Critical | {0} |' -f $Report.Summary.Critical))
    [void]$sb.AppendLine(('| High | {0} |' -f $Report.Summary.High))
    [void]$sb.AppendLine(('| Medium | {0} |' -f $Report.Summary.Medium))
    [void]$sb.AppendLine(('| Low | {0} |' -f $Report.Summary.Low))
    [void]$sb.AppendLine(('| Info | {0} |' -f $Report.Summary.Info))
    [void]$sb.AppendLine(('| Skipped | {0} |' -f $Report.Summary.Skipped))
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('## System Context')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine(('* Hostname: {0}' -f $Report.Metadata.Hostname))
    [void]$sb.AppendLine(('* User: {0}' -f $Report.Metadata.User))
    [void]$sb.AppendLine(('* Admin: {0}' -f (ConvertTo-WTYesNo -Value ([bool]$Report.Metadata.IsAdmin))))
    [void]$sb.AppendLine(('* Mode: {0}' -f $Report.Metadata.Mode))
    [void]$sb.AppendLine(('* Days: {0}' -f $Report.Metadata.Days))
    [void]$sb.AppendLine(('* Read-only: {0}' -f (ConvertTo-WTYesNo -Value ([bool]$Report.Metadata.IsReadOnly))))
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('## System Overview')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine(('* OS: {0}' -f (ConvertTo-WTDisplayValue -Value $Report.Normalized.System.OSCaption)))
    [void]$sb.AppendLine(('* Version: {0}' -f (ConvertTo-WTDisplayValue -Value $Report.Normalized.System.OSVersion)))
    $systemBuildText = Get-WTSystemBuildText -SystemInfo $Report.Normalized.System
    $systemInstallDateText = ConvertTo-WTDisplayValue -Value (ConvertTo-WTDateTimeString -Value $Report.Normalized.System.InstallDate) -Fallback 'Not available'
    $systemLastBootText = ConvertTo-WTDisplayValue -Value (ConvertTo-WTDateTimeString -Value $Report.Normalized.System.LastBootUpTime) -Fallback 'Not available'
    $systemUptimeText = if ($Report.Normalized.System.UptimeDays -ne $null) { '{0} days' -f $Report.Normalized.System.UptimeDays } else { 'Unknown' }
    [void]$sb.AppendLine(('* BuildNumber: {0}' -f (ConvertTo-WTDisplayValue -Value $systemBuildText)))
    [void]$sb.AppendLine(('* Architecture: {0}' -f (ConvertTo-WTDisplayValue -Value $Report.Normalized.System.OSArchitecture)))
    [void]$sb.AppendLine(('* Hostname: {0}' -f (ConvertTo-WTDisplayValue -Value $Report.Normalized.System.Hostname)))
    [void]$sb.AppendLine(('* User: {0}' -f (ConvertTo-WTDisplayValue -Value $Report.Normalized.System.User)))
    [void]$sb.AppendLine(('* Manufacturer/Model: {0} / {1}' -f (ConvertTo-WTDisplayValue -Value $Report.Normalized.System.Manufacturer), (ConvertTo-WTDisplayValue -Value $Report.Normalized.System.Model)))
    [void]$sb.AppendLine(('* Serial Number: {0}' -f (ConvertTo-WTDisplayValue -Value $Report.Normalized.System.SerialNumber)))
    [void]$sb.AppendLine(('* Domain/Workgroup: {0}' -f (ConvertTo-WTDisplayValue -Value $Report.Normalized.System.DomainOrWorkgroup)))
    [void]$sb.AppendLine(('* InstallDate: {0}' -f $systemInstallDateText))
    [void]$sb.AppendLine(('* LastBootUpTime: {0}' -f $systemLastBootText))
    [void]$sb.AppendLine(('* Uptime: {0}' -f $systemUptimeText))
    [void]$sb.AppendLine(('* Virtual machine: {0}' -f (ConvertTo-WTYesNoUnknown -Value $Report.Normalized.System.IsVirtualMachine)))
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('## Disk Overview')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine(('| Drive | FileSystem | SizeGB | FreeGB | FreePercent | Status | SystemDrive |'))
    [void]$sb.AppendLine(('| --- | --- | ---: | ---: | ---: | --- | --- |'))
    $diskRows = @($Report.Normalized.Disk | Where-Object { $_ })
    if ($diskRows.Count -eq 0) {
        [void]$sb.AppendLine('| Unknown | Unknown | Unknown | Unknown | Unknown | Unknown | Unknown |')
    }
    else {
        foreach ($disk in $diskRows) {
            [void]$sb.AppendLine(('| {0} | {1} | {2} | {3} | {4} | {5} | {6} |' -f (ConvertTo-WTDisplayValue -Value $disk.DriveLetter), (ConvertTo-WTDisplayValue -Value $disk.FileSystem), (ConvertTo-WTDisplayValue -Value $disk.SizeGB), (ConvertTo-WTDisplayValue -Value $disk.FreeGB), (ConvertTo-WTDisplayValue -Value $disk.FreePercent), (ConvertTo-WTDisplayValue -Value $disk.Status), (ConvertTo-WTYesNo -Value ([bool]$disk.IsSystemDrive))))
        }
    }
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('## Performance Overview')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine(('* RAM Total/Free/Used: {0} / {1} / {2} GB' -f (ConvertTo-WTDisplayValue -Value $Report.Normalized.Performance.TotalRamGB), (ConvertTo-WTDisplayValue -Value $Report.Normalized.Performance.FreeRamGB), (ConvertTo-WTDisplayValue -Value $Report.Normalized.Performance.UsedRamGB)))
    [void]$sb.AppendLine(('* RAM Free/Used Percent: {0}% / {1}%' -f (ConvertTo-WTDisplayValue -Value $Report.Normalized.Performance.FreeRamPercent), (ConvertTo-WTDisplayValue -Value $Report.Normalized.Performance.UsedRamPercent)))
    [void]$sb.AppendLine(('* CPU: {0}' -f (ConvertTo-WTDisplayValue -Value $Report.Normalized.Performance.ProcessorName)))
    [void]$sb.AppendLine(('* Logical processors: {0}' -f (ConvertTo-WTDisplayValue -Value $Report.Normalized.Performance.NumberOfLogicalProcessors)))
    [void]$sb.AppendLine(('* Physical processors: {0}' -f (ConvertTo-WTDisplayValue -Value $Report.Normalized.Performance.NumberOfProcessors)))
    [void]$sb.AppendLine(('* CPU load: {0}' -f (ConvertTo-WTDisplayValue -Value $Report.Normalized.Performance.CpuLoadPercent)))
    [void]$sb.AppendLine(('* Uptime: {0} days' -f (ConvertTo-WTDisplayValue -Value $Report.Normalized.Performance.UptimeDays)))
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('Top processes by memory:')
    $memoryRows = @($Report.Normalized.Performance.TopProcessesByMemory | Where-Object { $_ })
    if ($memoryRows.Count -eq 0) {
        [void]$sb.AppendLine('- Not available')
    }
    else {
        foreach ($process in $memoryRows) {
            [void]$sb.AppendLine(('- {0} (Id {1}, {2} MB, CPU {3})' -f $process.ProcessName, $process.Id, $process.WorkingSetMB, (ConvertTo-WTDisplayValue -Value $process.CPU)))
        }
    }
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('## Findings')
    [void]$sb.AppendLine('')

    if (@($Report.Findings).Count -eq 0) {
        [void]$sb.AppendLine('No findings were generated.')
    }
    else {
        foreach ($finding in $Report.Findings) {
            [void]$sb.AppendLine(('### [{0}] {1}' -f $finding.Severity, $finding.Title))
            [void]$sb.AppendLine(('* Id: {0}' -f $finding.Id))
            [void]$sb.AppendLine(('* Category: {0}' -f $finding.Category))
            [void]$sb.AppendLine(('* Status: {0}' -f $finding.Status))
            [void]$sb.AppendLine(('* Source: {0}' -f $finding.Source))
            [void]$sb.AppendLine(('* Description: {0}' -f $finding.Description))
            [void]$sb.AppendLine(('* Recommendation: {0}' -f $finding.Recommendation))
            [void]$sb.AppendLine(('* Timestamp: {0}' -f $finding.Timestamp))
            [void]$sb.AppendLine('')
            if ($finding.Evidence -and $finding.Evidence.Count -gt 0) {
                [void]$sb.AppendLine('Evidence:')
                foreach ($item in $finding.Evidence) {
                    [void]$sb.AppendLine(('- {0}' -f $item))
                }
                [void]$sb.AppendLine('')
            }
        }
    }

    [void]$sb.AppendLine('## Execution')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine(('* Report directory: {0}' -f $Report.Metadata.ReportDirectory))
    [void]$sb.AppendLine(('* JSON: {0}' -f $Report.Metadata.JsonPath))
    [void]$sb.AppendLine(('* Markdown: {0}' -f $Report.Metadata.MarkdownPath))
    [void]$sb.AppendLine(('* Exit code: {0}' -f $Report.Metadata.ExitCode))
    [void]$sb.AppendLine(('* Output fallback used: {0}' -f (ConvertTo-WTYesNo -Value ([bool]$Report.Metadata.OutputFallbackUsed))))
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine(('* Errors: {0}' -f @($Report.Execution.Errors).Count))
    [void]$sb.AppendLine(('* Warnings: {0}' -f @($Report.Execution.Warnings).Count))
    [void]$sb.AppendLine(('* Skipped checks: {0}' -f @($Report.Execution.Skipped).Count))

    [System.IO.File]::WriteAllText($Path, $sb.ToString(), [System.Text.Encoding]::UTF8)
    return $Path
}

function ConvertTo-WTDisplayValue {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object]$Value,

        [AllowNull()]
        [AllowEmptyString()]
        [string]$Fallback = 'Unknown'
    )

    if ($null -eq $Value -or $Value -eq '') {
        return $Fallback
    }

    return $Value
}

function ConvertTo-WTYesNoUnknown {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($Value -eq $true -or $Value -eq 'True') {
        return 'Yes'
    }

    if ($Value -eq $false -or $Value -eq 'False') {
        return 'No'
    }

    return 'Unknown'
}

function Get-WTSystemBuildText {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [psobject]$SystemInfo
    )

    if (-not $SystemInfo) {
        return 'Unknown'
    }

    foreach ($propertyName in @('BuildNumber', 'OSBuildNumber', 'OSBuild', 'Build')) {
        $value = $null
        try {
            $value = $SystemInfo.$propertyName
        }
        catch {
            $value = $null
        }

        if ($null -ne $value -and $value -ne '') {
            return $value
        }
    }

    return 'Unknown'
}

function ConvertTo-WTDateTimeString {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value -or $Value -eq '') {
        return $null
    }

    try {
        return ([datetime]$Value).ToString('o')
    }
    catch {
        return $null
    }
}

function ConvertTo-WTGB {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object]$Bytes
    )

    try {
        if ($null -eq $Bytes -or $Bytes -eq '' -or [double]$Bytes -le 0) {
            return $null
        }

        return [math]::Round(([double]$Bytes / 1GB), 2)
    }
    catch {
        return $null
    }
}

function ConvertTo-WTUptimeDays {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object]$LastBootUpTime
    )

    if ($null -eq $LastBootUpTime -or $LastBootUpTime -eq '') {
        return $null
    }

    try {
        return [math]::Round(((Get-Date) - ([datetime]$LastBootUpTime)).TotalDays, 2)
    }
    catch {
        return $null
    }
}

function Get-WTSystemInfo {
    [CmdletBinding()]
    param()

    $os = $null
    $computer = $null
    $bios = $null
    $processor = $null
    $timezone = $null

    try { $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop } catch { }
    try { $computer = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop } catch { }
    try { $bios = Get-CimInstance -ClassName Win32_BIOS -ErrorAction Stop } catch { }
    try { $processor = Get-CimInstance -ClassName Win32_Processor -ErrorAction Stop } catch { }
    try { $timezone = Get-CimInstance -ClassName Win32_TimeZone -ErrorAction Stop } catch { }

    $bootDevice = $null
    $systemDrive = $null
    $windowsDirectory = $null
    $installDate = $null
    $lastBoot = $null
    $uptimeDays = $null
    $partOfDomain = $null
    $domain = $null
    $workgroup = $null
    $totalPhysicalMemoryGB = $null
    $logicalProcessors = $null
    $physicalProcessors = $null
    $systemType = $null
    $manufacturer = $null
    $model = $null
    $serialNumber = $null
    $biosVersion = $null
    $biosReleaseDate = $null

    if ($os) {
        $bootDevice = $os.BootDevice
        $systemDrive = $os.SystemDrive
        $windowsDirectory = $os.WindowsDirectory
        $installDate = ConvertTo-WTDateTimeString -Value $os.InstallDate
        $lastBoot = ConvertTo-WTDateTimeString -Value $os.LastBootUpTime
        $uptimeDays = ConvertTo-WTUptimeDays -LastBootUpTime $os.LastBootUpTime
    }

    if ($computer) {
        $partOfDomain = [bool]$computer.PartOfDomain
        $domain = $computer.Domain
        $workgroup = $computer.Workgroup
        $totalPhysicalMemoryGB = ConvertTo-WTGB -Bytes $computer.TotalPhysicalMemory
        $logicalProcessors = $computer.NumberOfLogicalProcessors
        $physicalProcessors = $computer.NumberOfProcessors
        $systemType = $computer.SystemType
        $manufacturer = $computer.Manufacturer
        $model = $computer.Model
    }

    if ($bios) {
        $serialNumber = $bios.SerialNumber
        $biosVersion = $bios.SMBIOSBIOSVersion
        if (-not $biosVersion) {
            $biosVersion = $bios.Version
        }
        $biosReleaseDate = ConvertTo-WTDateTimeString -Value $bios.ReleaseDate
    }

    $processorName = $null
    if ($processor) {
        $processorName = ($processor | Select-Object -First 1).Name
        if (-not $physicalProcessors) {
            $physicalProcessors = @($processor).Count
        }
        if (-not $logicalProcessors) {
            $logicalProcessors = ($processor | Measure-Object -Property NumberOfLogicalProcessors -Sum).Sum
        }
        if (-not $totalPhysicalMemoryGB -and $os) {
            $totalPhysicalMemoryGB = ConvertTo-WTGB -Bytes ([double]$os.TotalVisibleMemorySize * 1KB)
        }
    }

    return [pscustomobject]@{
        Hostname              = ConvertTo-WTDisplayValue -Value $env:COMPUTERNAME
        User                  = ConvertTo-WTDisplayValue -Value ('{0}\{1}' -f $env:USERDOMAIN, $env:USERNAME)
        OSCaption             = if ($os) { ConvertTo-WTDisplayValue -Value $os.Caption } else { $null }
        OSVersion             = if ($os) { ConvertTo-WTDisplayValue -Value $os.Version } else { $null }
        OSBuildNumber         = if ($os) { ConvertTo-WTDisplayValue -Value $os.BuildNumber } else { $null }
        BuildNumber           = if ($os) { ConvertTo-WTDisplayValue -Value $os.BuildNumber } else { $null }
        OSBuild               = if ($os) { ConvertTo-WTDisplayValue -Value $os.BuildNumber } else { $null }
        Build                 = if ($os) { ConvertTo-WTDisplayValue -Value $os.BuildNumber } else { $null }
        Version               = if ($os) { ConvertTo-WTDisplayValue -Value $os.Version } else { $null }
        OSArchitecture        = if ($os) { ConvertTo-WTDisplayValue -Value $os.OSArchitecture } else { $null }
        InstallDate           = $installDate
        LastBootUpTime        = $lastBoot
        UptimeDays            = $uptimeDays
        Manufacturer          = if ($manufacturer) { $manufacturer } else { $null }
        Model                 = if ($model) { $model } else { $null }
        SerialNumber          = if ($serialNumber) { $serialNumber } else { $null }
        BIOSVersion           = if ($biosVersion) { $biosVersion } else { $null }
        BIOSReleaseDate       = $biosReleaseDate
        PartOfDomain          = $partOfDomain
        Domain                = if ($domain) { $domain } else { $null }
        Workgroup             = if ($workgroup) { $workgroup } else { $null }
        TotalPhysicalMemoryGB = $totalPhysicalMemoryGB
        NumberOfLogicalProcessors = $logicalProcessors
        NumberOfProcessors    = $physicalProcessors
        SystemType            = if ($systemType) { $systemType } else { $null }
        BootDevice            = if ($bootDevice) { $bootDevice } else { $null }
        SystemDrive           = if ($systemDrive) { $systemDrive } else { $null }
        WindowsDirectory      = if ($windowsDirectory) { $windowsDirectory } else { $null }
        TimeZone              = if ($timezone) { ConvertTo-WTDisplayValue -Value $timezone.Caption } else { $null }
        PowerShellVersion     = $PSVersionTable.PSVersion.ToString()
        IsVirtualMachine      = 'Unknown'
        ProcessorName         = if ($processorName) { $processorName } else { $null }
    }
}

function ConvertTo-WTNormalizedSystemInfo {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [psobject]$SystemInfo
    )

    if (-not $SystemInfo) {
        return [pscustomobject]@{
            Hostname                  = 'Unknown'
            User                      = 'Unknown'
            OSCaption                 = 'Unknown'
            OSVersion                 = 'Unknown'
            OSBuildNumber             = 'Unknown'
            BuildNumber               = 'Unknown'
            OSBuild                   = 'Unknown'
            Build                     = 'Unknown'
            Version                   = 'Unknown'
            OSArchitecture            = 'Unknown'
            InstallDate               = $null
            LastBootUpTime            = $null
            UptimeDays                = $null
            Manufacturer              = 'Unknown'
            Model                     = 'Unknown'
            SerialNumber              = 'Unknown'
            BIOSVersion               = 'Unknown'
            BIOSReleaseDate           = $null
            PartOfDomain              = $null
            DomainOrWorkgroup         = 'Unknown'
            Domain                    = 'Unknown'
            Workgroup                 = 'Unknown'
            TotalPhysicalMemoryGB     = $null
            NumberOfLogicalProcessors = $null
            NumberOfProcessors        = $null
            SystemType                = 'Unknown'
            BootDevice                = 'Unknown'
            SystemDrive               = 'Unknown'
            WindowsDirectory          = 'Unknown'
            TimeZone                  = 'Unknown'
            PowerShellVersion         = 'Unknown'
            IsVirtualMachine          = 'Unknown'
            IsDomainJoined            = $null
        }
    }

    $manufacturer = ConvertTo-WTDisplayValue -Value $SystemInfo.Manufacturer
    $model = ConvertTo-WTDisplayValue -Value $SystemInfo.Model
    $isVirtualMachine = 'Unknown'
    $vmString = ('{0} {1}' -f $manufacturer, $model).ToLowerInvariant()

    if ($vmString -match 'virtual machine|vmware|virtualbox|qemu|kvm|hyper-v' -or $manufacturer -match 'Microsoft Corporation' -and $model -match 'Virtual Machine') {
        $isVirtualMachine = 'True'
    }
    elseif ($manufacturer -ne 'Unknown' -or $model -ne 'Unknown') {
        $isVirtualMachine = 'False'
    }

    $partOfDomain = $SystemInfo.PartOfDomain
    $domainState = 'Unknown'
    if ($partOfDomain -eq $true) {
        $domainState = 'Domain'
    }
    elseif ($partOfDomain -eq $false) {
        $domainState = 'Workgroup'
    }

    $domainOrWorkgroup = 'Unknown'
    if ($domainState -eq 'Domain' -and $SystemInfo.Domain) {
        $domainOrWorkgroup = $SystemInfo.Domain
    }
    elseif ($domainState -eq 'Workgroup' -and $SystemInfo.Workgroup) {
        $domainOrWorkgroup = $SystemInfo.Workgroup
    }

    return [pscustomobject]@{
        Hostname                  = ConvertTo-WTDisplayValue -Value $SystemInfo.Hostname
        User                      = ConvertTo-WTDisplayValue -Value $SystemInfo.User
        OSCaption                 = ConvertTo-WTDisplayValue -Value $SystemInfo.OSCaption
        OSVersion                 = ConvertTo-WTDisplayValue -Value $SystemInfo.OSVersion
        OSBuildNumber             = ConvertTo-WTDisplayValue -Value $SystemInfo.OSBuildNumber
        BuildNumber               = ConvertTo-WTDisplayValue -Value $SystemInfo.BuildNumber
        OSBuild                   = ConvertTo-WTDisplayValue -Value $SystemInfo.OSBuild
        Build                     = ConvertTo-WTDisplayValue -Value $SystemInfo.Build
        Version                   = ConvertTo-WTDisplayValue -Value $SystemInfo.Version
        OSArchitecture            = ConvertTo-WTDisplayValue -Value $SystemInfo.OSArchitecture
        InstallDate               = $SystemInfo.InstallDate
        LastBootUpTime            = $SystemInfo.LastBootUpTime
        UptimeDays                = $SystemInfo.UptimeDays
        Manufacturer              = ConvertTo-WTDisplayValue -Value $SystemInfo.Manufacturer
        Model                     = ConvertTo-WTDisplayValue -Value $SystemInfo.Model
        SerialNumber              = ConvertTo-WTDisplayValue -Value $SystemInfo.SerialNumber
        BIOSVersion               = ConvertTo-WTDisplayValue -Value $SystemInfo.BIOSVersion
        BIOSReleaseDate           = $SystemInfo.BIOSReleaseDate
        PartOfDomain              = $partOfDomain
        DomainOrWorkgroup         = $domainOrWorkgroup
        Domain                    = ConvertTo-WTDisplayValue -Value $SystemInfo.Domain
        Workgroup                 = ConvertTo-WTDisplayValue -Value $SystemInfo.Workgroup
        TotalPhysicalMemoryGB     = $SystemInfo.TotalPhysicalMemoryGB
        NumberOfLogicalProcessors = $SystemInfo.NumberOfLogicalProcessors
        NumberOfProcessors        = $SystemInfo.NumberOfProcessors
        SystemType                = ConvertTo-WTDisplayValue -Value $SystemInfo.SystemType
        BootDevice                = ConvertTo-WTDisplayValue -Value $SystemInfo.BootDevice
        SystemDrive               = ConvertTo-WTDisplayValue -Value $SystemInfo.SystemDrive
        WindowsDirectory          = ConvertTo-WTDisplayValue -Value $SystemInfo.WindowsDirectory
        TimeZone                  = ConvertTo-WTDisplayValue -Value $SystemInfo.TimeZone
        PowerShellVersion         = ConvertTo-WTDisplayValue -Value $SystemInfo.PowerShellVersion
        IsVirtualMachine          = $isVirtualMachine
        IsDomainJoined            = $partOfDomain
    }
}

function ConvertTo-WTNormalizedPerformanceInfo {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [psobject]$PerformanceInfo
    )

    if (-not $PerformanceInfo) {
        return [pscustomobject]@{
            TotalRamGB               = $null
            FreeRamGB                = $null
            UsedRamGB                = $null
            FreeRamPercent           = $null
            UsedRamPercent           = $null
            UptimeDays               = $null
            ProcessorName            = 'Unknown'
            NumberOfLogicalProcessors = $null
            NumberOfProcessors       = $null
            CpuLoadPercent           = $null
            TopProcessesByCPU        = @()
            TopProcessesByMemory     = @()
        }
    }

    $total = $PerformanceInfo.TotalRamGB
    $free = $PerformanceInfo.FreeRamGB
    $used = $null
    $freePercent = $null
    $usedPercent = $null

    if ($total -ne $null -and $free -ne $null) {
        $used = [math]::Round(($total - $free), 2)
        if ($total -gt 0) {
            $freePercent = [math]::Round(($free / $total) * 100, 2)
            $usedPercent = [math]::Round(($used / $total) * 100, 2)
        }
    }

    return [pscustomobject]@{
        TotalRamGB             = $total
        FreeRamGB              = $free
        UsedRamGB              = $used
        FreeRamPercent         = $freePercent
        UsedRamPercent         = $usedPercent
        UptimeDays             = $PerformanceInfo.UptimeDays
        ProcessorName          = ConvertTo-WTDisplayValue -Value $PerformanceInfo.ProcessorName
        NumberOfLogicalProcessors = $PerformanceInfo.NumberOfLogicalProcessors
        NumberOfProcessors     = $PerformanceInfo.NumberOfProcessors
        CpuLoadPercent         = $PerformanceInfo.CpuLoadPercent
        TopProcessesByCPU      = @($PerformanceInfo.TopProcessesByCPU | Where-Object { $_ })
        TopProcessesByMemory   = @($PerformanceInfo.TopProcessesByMemory | Where-Object { $_ })
    }
}

function Get-WTDiskInfo {
    [CmdletBinding()]
    param()

    $systemDrive = $env:SystemDrive
    if (-not $systemDrive) {
        $systemDrive = 'C:'
    }

    $drives = @()
    try {
        $drives = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction Stop
    }
    catch {
        $drives = @()
    }

    $result = @()
    foreach ($drive in $drives) {
        $sizeBytes = $drive.Size
        $freeBytes = $drive.FreeSpace
        $sizeKnown = $false
        $sizeGB = $null
        $freeGB = $null
        $usedGB = $null
        $freePercent = $null
        $usedPercent = $null
        $status = 'OK'

        if ($sizeBytes -ne $null -and [double]$sizeBytes -gt 0) {
            $sizeKnown = $true
            $sizeGB = ConvertTo-WTGB -Bytes $sizeBytes
            if ($freeBytes -ne $null) {
                $freeGB = ConvertTo-WTGB -Bytes $freeBytes
                $usedGB = if ($sizeGB -ne $null -and $freeGB -ne $null) { [math]::Round(($sizeGB - $freeGB), 2) } else { $null }
                if ($sizeGB -gt 0) {
                    $freePercent = [math]::Round((([double]$freeBytes / [double]$sizeBytes) * 100), 2)
                    $usedPercent = [math]::Round((100 - $freePercent), 2)
                }
            }
        }
        else {
            $status = 'UnknownSize'
        }

        $isSystemDrive = $false
        if ($drive.DeviceID -and $systemDrive -and ($drive.DeviceID.TrimEnd('\') -eq $systemDrive.TrimEnd('\'))) {
            $isSystemDrive = $true
        }

        $isLowSpaceCandidate = $sizeKnown

        $result += [pscustomobject]@{
            DriveLetter        = $drive.DeviceID
            VolumeName         = $drive.VolumeName
            FileSystem         = $drive.FileSystem
            SizeKnown          = $sizeKnown
            SizeGB             = $sizeGB
            FreeGB             = $freeGB
            UsedGB             = $usedGB
            FreePercent        = $freePercent
            UsedPercent        = $usedPercent
            Status             = $status
            IsSystemDrive      = $isSystemDrive
            IsLowSpaceCandidate = $isLowSpaceCandidate
        }
    }

    return $result
}

function ConvertTo-WTNormalizedDiskInfo {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object[]]$DiskInfo
    )

    if (-not $DiskInfo) {
        return @()
    }

    $normalized = @()
    foreach ($disk in $DiskInfo) {
        if (-not $disk) {
            continue
        }
        $normalized += [pscustomobject]@{
            DriveLetter         = ConvertTo-WTDisplayValue -Value $disk.DriveLetter
            VolumeName          = ConvertTo-WTDisplayValue -Value $disk.VolumeName
            FileSystem          = ConvertTo-WTDisplayValue -Value $disk.FileSystem
            SizeKnown           = [bool]$disk.SizeKnown
            SizeGB              = $disk.SizeGB
            FreeGB              = $disk.FreeGB
            UsedGB              = $disk.UsedGB
            FreePercent         = $disk.FreePercent
            UsedPercent         = $disk.UsedPercent
            Status              = ConvertTo-WTDisplayValue -Value $disk.Status
            IsSystemDrive       = [bool]$disk.IsSystemDrive
            IsLowSpaceCandidate = [bool]$disk.IsLowSpaceCandidate
        }
    }

    $sorted = $normalized | Sort-Object -Property @{ Expression = { if ($_.IsSystemDrive) { 0 } else { 1 } } }, @{ Expression = { $_.DriveLetter } }
    return @($sorted)
}

function Get-WTPerformanceInfo {
    [CmdletBinding()]
    param()

    $os = $null
    $computer = $null
    $processors = @()
    try { $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop } catch { }
    try { $computer = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop } catch { }
    try { $processors = Get-CimInstance -ClassName Win32_Processor -ErrorAction Stop } catch { }

    $totalRamGB = $null
    $freeRamGB = $null
    $usedRamGB = $null
    $freeRamPercent = $null
    $usedRamPercent = $null
    $uptimeDays = $null
    $processorName = $null
    $logicalProcessors = $null
    $physicalProcessors = $null
    $cpuLoadPercent = $null

    if ($os) {
        $totalRamGB = ConvertTo-WTGB -Bytes ([double]$os.TotalVisibleMemorySize * 1KB)
        $freeRamGB = ConvertTo-WTGB -Bytes ([double]$os.FreePhysicalMemory * 1KB)
        if ($totalRamGB -ne $null -and $freeRamGB -ne $null) {
            $usedRamGB = [math]::Round(($totalRamGB - $freeRamGB), 2)
            if ($totalRamGB -gt 0) {
                $freeRamPercent = [math]::Round((($freeRamGB / $totalRamGB) * 100), 2)
                $usedRamPercent = [math]::Round((100 - $freeRamPercent), 2)
            }
        }
        $uptimeDays = ConvertTo-WTUptimeDays -LastBootUpTime $os.LastBootUpTime
    }

    if ($computer) {
        $logicalProcessors = $computer.NumberOfLogicalProcessors
        $physicalProcessors = $computer.NumberOfProcessors
    }

    if ($processors -and @($processors).Count -gt 0) {
        $processorName = ($processors | Select-Object -First 1).Name
        if (-not $logicalProcessors) {
            $logicalProcessors = ($processors | Measure-Object -Property NumberOfLogicalProcessors -Sum).Sum
        }
        if (-not $physicalProcessors) {
            $physicalProcessors = @($processors).Count
        }
        $loadValues = @()
        foreach ($p in $processors) {
            if ($p.LoadPercentage -ne $null) {
                $loadValues += [double]$p.LoadPercentage
            }
        }
        if ($loadValues.Count -gt 0) {
            $cpuLoadPercent = [math]::Round((($loadValues | Measure-Object -Average).Average), 2)
        }
    }

    $topProcessesByCpu = @()
    try {
        $processes = Get-Process -ErrorAction Stop | Sort-Object -Property CPU -Descending | Select-Object -First 10
        foreach ($process in $processes) {
            $topProcessesByCpu += [pscustomobject]@{
                ProcessName = $process.ProcessName
                Id          = $process.Id
                CPU         = if ($process.CPU -ne $null) { [math]::Round([double]$process.CPU, 2) } else { $null }
                WorkingSetMB = [math]::Round(($process.WorkingSet64 / 1MB), 2)
            }
        }
    }
    catch {
        $topProcessesByCpu = @()
    }

    $topProcessesByMemory = @()
    try {
        $processesByMem = Get-Process -ErrorAction Stop | Sort-Object -Property WorkingSet64 -Descending | Select-Object -First 10
        foreach ($process in $processesByMem) {
            $topProcessesByMemory += [pscustomobject]@{
                ProcessName  = $process.ProcessName
                Id           = $process.Id
                WorkingSetMB = [math]::Round(($process.WorkingSet64 / 1MB), 2)
                CPU          = if ($process.CPU -ne $null) { [math]::Round([double]$process.CPU, 2) } else { $null }
            }
        }
    }
    catch {
        $topProcessesByMemory = @()
    }

    return [pscustomobject]@{
        TotalRamGB             = $totalRamGB
        FreeRamGB              = $freeRamGB
        UsedRamGB              = $usedRamGB
        FreeRamPercent         = $freeRamPercent
        UsedRamPercent         = $usedRamPercent
        UptimeDays             = $uptimeDays
        ProcessorName          = if ($processorName) { $processorName } else { $null }
        NumberOfLogicalProcessors = $logicalProcessors
        NumberOfProcessors     = $physicalProcessors
        CpuLoadPercent         = $cpuLoadPercent
        TopProcessesByCPU      = $topProcessesByCpu
        TopProcessesByMemory   = $topProcessesByMemory
    }
}

function Invoke-WTSystemRules {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$Report
    )

    $system = $Report.Normalized.System
    if (-not $system) {
        Add-WTFinding -Report $Report -Id 'WT-SYS-OS-UNKNOWN' -Category 'System' -Severity 'Medium' -Title 'System information could not be determined' -Description 'Essential operating system details are unavailable.' -Evidence @('System normalization returned no data.') -Recommendation 'Run again with the same account or verify WMI/CIM health.' -Source 'Invoke-WTSystemRules' -Status 'Warning'
        return
    }

    $osCaptionKnown = ($system.OSCaption -and $system.OSCaption -ne 'Unknown')
    $osVersionKnown = ($system.OSVersion -and $system.OSVersion -ne 'Unknown')
    $osBuildKnown = ($system.OSBuildNumber -and $system.OSBuildNumber -ne 'Unknown')

    if ($osCaptionKnown -and $osVersionKnown -and $osBuildKnown) {
        Add-WTFinding -Report $Report -Id 'WT-SYS-OS-INFO' -Category 'System' -Severity 'Info' -Title 'Operating system detected' -Description 'The operating system was successfully detected.' -Evidence @("OS=$($system.OSCaption)", "Version=$($system.OSVersion)", "Build=$($system.OSBuildNumber)") -Recommendation 'No action required.' -Source 'Invoke-WTSystemRules' -Status 'Info'
    }
    else {
        Add-WTFinding -Report $Report -Id 'WT-SYS-OS-UNKNOWN' -Category 'System' -Severity 'Medium' -Title 'Essential system information is incomplete' -Description 'One or more critical operating system values could not be determined.' -Evidence @("OSCaption=$($system.OSCaption)", "Version=$($system.OSVersion)", "Build=$($system.OSBuildNumber)") -Recommendation 'Verify WMI/CIM access and rerun.' -Source 'Invoke-WTSystemRules' -Status 'Warning'
    }

    $domainEvidence = @()
    if ($system.IsDomainJoined -eq $true) {
        $domainEvidence += ('Joined to domain: {0}' -f $system.DomainOrWorkgroup)
    }
    elseif ($system.IsDomainJoined -eq $false) {
        $domainEvidence += ('Workgroup: {0}' -f $system.DomainOrWorkgroup)
    }
    else {
        $domainEvidence += 'Domain membership unknown'
    }

    Add-WTFinding -Report $Report -Id 'WT-SYS-DOMAIN-INFO' -Category 'System' -Severity 'Info' -Title 'Domain or workgroup membership detected' -Description 'The machine membership state was identified.' -Evidence $domainEvidence -Recommendation 'No action required.' -Source 'Invoke-WTSystemRules' -Status 'Info'

    if ($system.IsVirtualMachine -eq 'True') {
        Add-WTFinding -Report $Report -Id 'WT-SYS-VM-INFO' -Category 'System' -Severity 'Info' -Title 'Virtual machine detected' -Description 'The machine appears to be virtualized based on model and manufacturer heuristics.' -Evidence @("Manufacturer=$($system.Manufacturer)", "Model=$($system.Model)") -Recommendation 'No action required.' -Source 'Invoke-WTSystemRules' -Status 'Info'
    }

    if ($system.UptimeDays -ne $null) {
        if ($system.UptimeDays -ge 30) {
            Add-WTFinding -Report $Report -Id 'WT-SYS-UPTIME-VERY-LONG' -Category 'System' -Severity 'Medium' -Title 'System uptime is very long' -Description 'The system has been running for 30 days or more.' -Evidence @("UptimeDays=$($system.UptimeDays)") -Recommendation 'Confirm whether a reboot is pending or scheduled maintenance is overdue.' -Source 'Invoke-WTSystemRules' -Status 'Warning'
        }
        elseif ($system.UptimeDays -ge 14) {
            Add-WTFinding -Report $Report -Id 'WT-SYS-UPTIME-LONG' -Category 'System' -Severity 'Low' -Title 'System uptime is long' -Description 'The system has been running for 14 days or more.' -Evidence @("UptimeDays=$($system.UptimeDays)") -Recommendation 'No immediate action required, but verify update and maintenance posture.' -Source 'Invoke-WTSystemRules' -Status 'Warning'
        }
    }
}

function Invoke-WTDiskRules {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$Report
    )

    $disks = @($Report.Normalized.Disk)
    foreach ($disk in $disks) {
        if (-not $disk) {
            continue
        }

        if ($disk.SizeKnown -eq $false) {
            Add-WTFinding -Report $Report -Id 'WT-DISK-UNKNOWN-SIZE' -Category 'Disk' -Severity 'Info' -Title ('Disk size unknown for {0}' -f $disk.DriveLetter) -Description 'The logical disk size could not be determined, so low-space evaluation is skipped.' -Evidence @("Drive=$($disk.DriveLetter)", "Status=$($disk.Status)") -Recommendation 'No action required unless this is unexpected.' -Source 'Invoke-WTDiskRules' -Status 'Info'
            continue
        }

        if ($disk.IsSystemDrive -eq $true -and $disk.FreePercent -ne $null) {
            if ($disk.FreePercent -lt 5) {
                Add-WTFinding -Report $Report -Id 'WT-DISK-SYSTEM-CRITICAL' -Category 'Disk' -Severity 'Critical' -Title ('System drive low on space: {0}' -f $disk.DriveLetter) -Description 'The system drive has less than 5 percent free space.' -Evidence @("Drive=$($disk.DriveLetter)", "FreePercent=$($disk.FreePercent)", "FreeGB=$($disk.FreeGB)") -Recommendation 'Free up disk space or move data off the system drive.' -Source 'Invoke-WTDiskRules' -Status 'Fail'
            }
            elseif ($disk.FreePercent -lt 10) {
                Add-WTFinding -Report $Report -Id 'WT-DISK-SYSTEM-HIGH' -Category 'Disk' -Severity 'High' -Title ('System drive has low free space: {0}' -f $disk.DriveLetter) -Description 'The system drive has between 5 and 10 percent free space.' -Evidence @("Drive=$($disk.DriveLetter)", "FreePercent=$($disk.FreePercent)", "FreeGB=$($disk.FreeGB)") -Recommendation 'Monitor closely and plan cleanup or expansion.' -Source 'Invoke-WTDiskRules' -Status 'Warning'
            }
            elseif ($disk.FreePercent -lt 15) {
                Add-WTFinding -Report $Report -Id 'WT-DISK-SYSTEM-MEDIUM' -Category 'Disk' -Severity 'Medium' -Title ('System drive free space is getting low: {0}' -f $disk.DriveLetter) -Description 'The system drive has between 10 and 15 percent free space.' -Evidence @("Drive=$($disk.DriveLetter)", "FreePercent=$($disk.FreePercent)", "FreeGB=$($disk.FreeGB)") -Recommendation 'Plan cleanup before the drive reaches a critical threshold.' -Source 'Invoke-WTDiskRules' -Status 'Warning'
            }
        }
        elseif ($disk.IsLowSpaceCandidate -eq $true -and $disk.FreePercent -ne $null -and $disk.FreePercent -lt 10) {
            Add-WTFinding -Report $Report -Id 'WT-DISK-DATA-LOW' -Category 'Disk' -Severity 'Medium' -Title ('Data drive low on space: {0}' -f $disk.DriveLetter) -Description 'A non-system drive has less than 10 percent free space.' -Evidence @("Drive=$($disk.DriveLetter)", "FreePercent=$($disk.FreePercent)", "FreeGB=$($disk.FreeGB)") -Recommendation 'Review data growth and move or archive files if needed.' -Source 'Invoke-WTDiskRules' -Status 'Warning'
        }
    }
}

function Invoke-WTPerformanceRules {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$Report
    )

    $perf = $Report.Normalized.Performance
    if (-not $perf) {
        return
    }

    if ($perf.TotalRamGB -ne $null -and $perf.FreeRamGB -ne $null) {
        Add-WTFinding -Report $Report -Id 'WT-PERF-RAM-INFO' -Category 'Performance' -Severity 'Info' -Title 'Memory summary collected' -Description 'Basic RAM metrics were collected.' -Evidence @("TotalRamGB=$($perf.TotalRamGB)", "FreeRamGB=$($perf.FreeRamGB)", "UsedRamGB=$($perf.UsedRamGB)") -Recommendation 'No action required.' -Source 'Invoke-WTPerformanceRules' -Status 'Info'

        if ($perf.FreeRamPercent -ne $null) {
            if ($perf.FreeRamPercent -lt 10) {
                Add-WTFinding -Report $Report -Id 'WT-PERF-RAM-LOW' -Category 'Performance' -Severity 'Medium' -Title 'Available memory is low' -Description 'Free RAM is below 10 percent.' -Evidence @("FreeRamPercent=$($perf.FreeRamPercent)", "FreeRamGB=$($perf.FreeRamGB)") -Recommendation 'Identify memory-heavy workloads or consider adding RAM.' -Source 'Invoke-WTPerformanceRules' -Status 'Warning'
            }
            elseif ($perf.FreeRamPercent -lt 20) {
                Add-WTFinding -Report $Report -Id 'WT-PERF-RAM-WARNING' -Category 'Performance' -Severity 'Low' -Title 'Available memory is moderate' -Description 'Free RAM is between 10 and 20 percent.' -Evidence @("FreeRamPercent=$($perf.FreeRamPercent)", "FreeRamGB=$($perf.FreeRamGB)") -Recommendation 'Monitor for memory pressure if the issue is user-facing.' -Source 'Invoke-WTPerformanceRules' -Status 'Info'
            }
        }
    }

    if ($perf.CpuLoadPercent -ne $null) {
        Add-WTFinding -Report $Report -Id 'WT-PERF-CPU-INFO' -Category 'Performance' -Severity 'Info' -Title 'CPU load collected' -Description 'A basic CPU load sample was collected.' -Evidence @("CpuLoadPercent=$($perf.CpuLoadPercent)", "Processor=$($perf.ProcessorName)") -Recommendation 'No action required.' -Source 'Invoke-WTPerformanceRules' -Status 'Info'

        if ($perf.CpuLoadPercent -ge 90) {
            Add-WTFinding -Report $Report -Id 'WT-PERF-CPU-HIGH' -Category 'Performance' -Severity 'Medium' -Title 'CPU load is high' -Description 'CPU load is at or above 90 percent.' -Evidence @("CpuLoadPercent=$($perf.CpuLoadPercent)", "Processor=$($perf.ProcessorName)") -Recommendation 'Identify the process or workload causing CPU saturation.' -Source 'Invoke-WTPerformanceRules' -Status 'Warning'
        }
    }

    if ($perf.TopProcessesByMemory -and @($perf.TopProcessesByMemory).Count -gt 0) {
        Add-WTFinding -Report $Report -Id 'WT-PERF-PROC-MEM-INFO' -Category 'Performance' -Severity 'Info' -Title 'Top memory processes collected' -Description 'The highest memory consumers were captured for quick triage.' -Evidence @("Count=$(@($perf.TopProcessesByMemory).Count)") -Recommendation 'Review the Markdown report for the full list.' -Source 'Invoke-WTPerformanceRules' -Status 'Info'
    }
}

function Get-WTBasicSystemInfo {
    [CmdletBinding()]
    param()
    return Get-WTSystemInfo
}

function Invoke-WinTriage {
    [CmdletBinding()]
    param()

    $mode = 'Standard'
    $exitCode = 3

    try {
        if ($Quick.IsPresent -and $Full.IsPresent) {
            throw "Parameters -Quick and -Full cannot be used together."
        }

        if ($Quick.IsPresent) {
            $mode = 'Quick'
        }
        elseif ($Full.IsPresent) {
            $mode = 'Full'
        }
        else {
            $mode = 'Standard'
        }

        $isAdmin = Test-WTAdministrator
        $hostname = if ($env:COMPUTERNAME) { $env:COMPUTERNAME } else { 'UNKNOWN' }
        $resolvedOutput = Resolve-WTOutputBasePath -RequestedPath $OutputPath
        if (-not $resolvedOutput.BasePath) {
            throw ('Unable to create output path. Requested: {0}. Fallback: {1}. Error: {2}' -f $resolvedOutput.RequestedPath, $resolvedOutput.FallbackPath, $resolvedOutput.FallbackReason)
        }

        $reportDirectory = New-WTReportDirectory -BasePath $resolvedOutput.BasePath -Hostname $hostname
        $jsonPath = ConvertTo-WTAbsolutePath -Path (Join-Path -Path $reportDirectory -ChildPath 'WinTriage.json')
        $markdownPath = ConvertTo-WTAbsolutePath -Path (Join-Path -Path $reportDirectory -ChildPath 'WinTriage.md')

        $report = New-WTReportObject -Mode $mode -Days $Days -OutputDirectory $reportDirectory -IsAdmin $isAdmin -OutputBasePath $resolvedOutput.BasePath -UsedOutputFallback $resolvedOutput.UsedFallback
        $report.Metadata.JsonPath = $jsonPath
        $report.Metadata.MarkdownPath = $markdownPath

        if ($resolvedOutput.UsedFallback) {
            Add-WTExecutionWarning -Report $report -Scope 'OutputPath' -Message ('Fallback output path used because the requested path could not be created. Requested: {0}. Fallback: {1}. Reason: {2}' -f $resolvedOutput.RequestedPath, $resolvedOutput.FallbackPath, $resolvedOutput.FallbackReason)
        }

        $systemRaw = Invoke-WTSafeCollector -Report $report -Name 'SystemInfo' -ScriptBlock {
            Get-WTSystemInfo
        }

        if ($systemRaw) {
            $report.Raw.System = $systemRaw
            Invoke-WTSafeStep -Report $report -Name 'NormalizeSystem' -ScriptBlock {
                $report.Normalized.System = ConvertTo-WTNormalizedSystemInfo -SystemInfo $report.Raw.System
                $report.Context.Manufacturer = $report.Normalized.System.Manufacturer
                $report.Context.Model = $report.Normalized.System.Model
                $report.Context.SerialNumber = $report.Normalized.System.SerialNumber
                $report.Context.IsDomainJoined = $report.Normalized.System.IsDomainJoined
                $report.Context.DomainName = $report.Normalized.System.DomainOrWorkgroup
            } | Out-Null
        }

        $diskRaw = Invoke-WTSafeCollector -Report $report -Name 'DiskInfo' -ScriptBlock {
            Get-WTDiskInfo
        }
        if ($diskRaw) {
            $report.Raw.Disk = $diskRaw
            Invoke-WTSafeStep -Report $report -Name 'NormalizeDisk' -ScriptBlock {
                $report.Normalized.Disk = ConvertTo-WTNormalizedDiskInfo -DiskInfo $report.Raw.Disk
            } | Out-Null
        }

        $perfRaw = Invoke-WTSafeCollector -Report $report -Name 'PerformanceInfo' -ScriptBlock {
            Get-WTPerformanceInfo
        }
        if ($perfRaw) {
            $report.Raw.Performance = $perfRaw
            Invoke-WTSafeStep -Report $report -Name 'NormalizePerformance' -ScriptBlock {
                $report.Normalized.Performance = ConvertTo-WTNormalizedPerformanceInfo -PerformanceInfo $report.Raw.Performance
            } | Out-Null
        }

        $report.Raw.Updates = [pscustomobject]@{
            Status = 'NotImplemented'
            Mode   = $mode
        }
        $report.Normalized.Updates = $report.Raw.Updates

        $report.Raw.Defender = [pscustomobject]@{
            Status = 'NotImplemented'
            RequiresAdmin = $true
        }
        $report.Normalized.Defender = $report.Raw.Defender

        $report.Raw.Domain = [pscustomobject]@{
            IsDomainJoined = $report.Context.IsDomainJoined
            DomainName     = $report.Context.DomainName
        }
        $report.Normalized.Domain = $report.Raw.Domain

        $report.Raw.Services = [pscustomobject]@{
            Status = 'NotImplemented'
        }
        $report.Normalized.Services = $report.Raw.Services

        $report.Raw.Events = [pscustomobject]@{
            Status = 'NotImplemented'
        }
        $report.Normalized.Events = $report.Raw.Events

        Add-WTFinding -Report $report `
            -Id 'WT-RO-001' `
            -Category 'Framework' `
            -Severity 'Info' `
            -Title 'Tool executed in read-only mode' `
            -Description 'WinTriage is running in a non-remediating, read-only diagnostic mode.' `
            -Evidence @('IsReadOnly=true') `
            -Recommendation 'No action required.' `
            -Source 'Invoke-WinTriage' `
            -Status 'Info' `
            -RequiresAdmin $false

        Add-WTFinding -Report $report `
            -Id 'WT-MODE-001' `
            -Category 'Framework' `
            -Severity 'Info' `
            -Title ('Execution mode: {0}' -f $mode) `
            -Description 'The selected operating mode controls the breadth of future collectors.' `
            -Evidence @("Mode=$mode", "Days=$Days") `
            -Recommendation 'No action required.' `
            -Source 'Invoke-WinTriage' `
            -Status 'Info' `
            -RequiresAdmin $false

        Invoke-WTSafeStep -Report $report -Name 'SystemRules' -ScriptBlock {
            Invoke-WTSystemRules -Report $report
        } | Out-Null
        Invoke-WTSafeStep -Report $report -Name 'DiskRules' -ScriptBlock {
            Invoke-WTDiskRules -Report $report
        } | Out-Null
        Invoke-WTSafeStep -Report $report -Name 'PerformanceRules' -ScriptBlock {
            Invoke-WTPerformanceRules -Report $report
        } | Out-Null

        if (-not $isAdmin) {
            Add-WTSkippedCheck -Report $report -Check 'AdminExample' -Reason 'Skipped example because the script is not running elevated.'
            Add-WTFinding -Report $report `
                -Id 'WT-SKIP-001' `
                -Category 'Privilege' `
                -Severity 'Skipped' `
                -Title 'Example admin-only diagnostic skipped' `
                -Description 'This is a placeholder for a future admin-only check.' `
                -Evidence @('Not running as administrator') `
                -Recommendation 'Rerun elevated only if a future admin-only collector requires it.' `
                -Source 'Invoke-WinTriage' `
                -Status 'Skipped' `
                -RequiresAdmin $true
        }
        else {
            Invoke-WTSafeCollector -Report $report -Name 'AdminExample' -RequiresAdmin $true -ScriptBlock {
                return [pscustomobject]@{
                    Status = 'Example only'
                }
            } | Out-Null
        }

        Invoke-WTSafeStep -Report $report -Name 'SummaryUpdate' -ScriptBlock {
            Update-WTSummary -Report $report | Out-Null
        } | Out-Null

        if ($report.Summary.Critical -gt 0 -or $report.Summary.High -gt 0) {
            $exitCode = 1
        }
        elseif (@($report.Execution.Errors).Count -gt 0) {
            $exitCode = 2
        }
        else {
            $exitCode = 0
        }

        $report.Metadata.FinishedAt = (Get-Date).ToString('o')

        if (-not $script:WTIsJsonOnly) {
            try {
                Export-WTMarkdownReport -Report $report -Path $markdownPath | Out-Null
                $report.Metadata.MarkdownGenerated = $true
            }
            catch {
                Add-WTExecutionWarning -Report $report -Scope 'MarkdownExport' -Message ('Unable to write Markdown report. {0}' -f $_.Exception.Message)
                $report.Metadata.MarkdownGenerated = $false
            }
        }

        if ($OpenReport.IsPresent -and -not $script:WTIsJsonOnly -and $report.Metadata.MarkdownGenerated) {
            try {
                Invoke-Item -LiteralPath $markdownPath
            }
            catch {
                Add-WTExecutionWarning -Report $report -Scope 'OpenReport' -Message ('Unable to open Markdown report. {0}' -f $_.Exception.Message)
            }
        }

        $report.Metadata.ExitCode = $exitCode
        $report.Metadata.JsonGenerated = $true
        Export-WTJsonReport -Report $report -Path $jsonPath | Out-Null
        Write-WTConsoleSummary -Report $report -NoColor:$script:WTNoColor

        return $exitCode
    }
    catch {
        if ($report) {
            Add-WTExecutionError -Report $report -Scope 'Fatal' -Message $_.Exception.Message
            if (-not $report.Metadata.FinishedAt) {
                $report.Metadata.FinishedAt = (Get-Date).ToString('o')
            }
            $report.Metadata.ExitCode = 2
            if (-not $script:WTIsJsonOnly -and -not $report.Metadata.MarkdownGenerated) {
                try {
                    Export-WTMarkdownReport -Report $report -Path $markdownPath | Out-Null
                    $report.Metadata.MarkdownGenerated = $true
                }
                catch {
                    Add-WTExecutionWarning -Report $report -Scope 'MarkdownExport' -Message ('Unable to write Markdown report after fatal error. {0}' -f $_.Exception.Message)
                }
            }
            try {
                $report.Metadata.JsonGenerated = $true
                Export-WTJsonReport -Report $report -Path $jsonPath | Out-Null
            }
            catch {
                $report.Metadata.JsonGenerated = $false
                Add-WTExecutionWarning -Report $report -Scope 'JsonExport' -Message ('Unable to write JSON report after fatal error. {0}' -f $_.Exception.Message)
            }
            try {
                Write-WTConsoleSummary -Report $report -NoColor:$script:WTNoColor
            }
            catch {
            }
            return 2
        }

        $fatalFile = Write-WTFatalErrorFile -Message $_.Exception.Message -ErrorRecord $_
        try {
            if ($script:WTDebugErrors) {
                Write-Host 'WinTriage failed during initialization. No report was generated.' -ForegroundColor Red
                if ($_.Exception -and $_.Exception.Message) {
                    Write-Host ('Exception: {0}' -f $_.Exception.Message) -ForegroundColor Red
                }
                if ($_.InvocationInfo) {
                    if ($_.InvocationInfo.ScriptLineNumber) {
                        Write-Host ('ScriptLineNumber: {0}' -f $_.InvocationInfo.ScriptLineNumber) -ForegroundColor DarkRed
                    }
                    if ($_.InvocationInfo.Line) {
                        Write-Host ('Line: {0}' -f $_.InvocationInfo.Line.Trim()) -ForegroundColor DarkRed
                    }
                }
                if ($fatalFile) {
                    Write-Host ('Fatal error details written to: {0}' -f $fatalFile) -ForegroundColor DarkRed
                }
            }
            else {
                Write-Host 'WinTriage failed during initialization. No report was generated.' -ForegroundColor Red
                if ($fatalFile) {
                    Write-Host ('Fatal error details written to: {0}' -f $fatalFile) -ForegroundColor DarkRed
                }
            }
        }
        catch {
        }
        return 3
    }
}

$script:WTExitCode = Invoke-WinTriage
if ($UseExitCode.IsPresent) {
    exit $script:WTExitCode
}
