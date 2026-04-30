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
    [switch]$DebugErrors,
    [switch]$SelfTestEventParser
)

# WinTriage is read-only by design.
# It collects diagnostic data and generates reports without modifying system configuration.

$script:WTVersion = '0.4.2'
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
            RawDirectory       = $null
            JsonPath           = $null
            MarkdownPath       = $null
            EventCsvSystemPath = $null
            EventCsvApplicationPath = $null
            EventCsvAllPath    = $null
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
            PowerBoot   = $null
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
            PowerBoot   = $null
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

function Get-WTExecutionErrorClassification {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Scope,

        [AllowNull()]
        [string]$Message
    )

    $nonFatalScopes = @(
        'OutputPath',
        'MarkdownExport',
        'JsonExport',
        'OpenReport',
        'EventProcessParsing',
        'AdminExample',
        'EventCsv',
        'UpdateCsv'
    )

    if ($Scope -in $nonFatalScopes) {
        return [pscustomobject]@{
            Impact                  = 'NonFatal'
            AffectsReportIntegrity  = $false
        }
    }

    if ($Scope -match '^(SystemInfo|PowerBootInfo|DiskInfo|PerformanceInfo|UpdateInfo|EventInfo|SystemRules|PowerBootRules|DiskRules|PerformanceRules|UpdateRules|EventRules|EventCorrelationRules)$') {
        return [pscustomobject]@{
            Impact                  = 'Partial'
            AffectsReportIntegrity  = $true
        }
    }

    if ($Scope -match '^(Fatal)$') {
        return [pscustomobject]@{
            Impact                  = 'Fatal'
            AffectsReportIntegrity  = $true
        }
    }

    return [pscustomobject]@{
        Impact                  = 'Partial'
        AffectsReportIntegrity  = $true
    }
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
        [string]$Message,

        [ValidateSet('NonFatal', 'Partial', 'Fatal')]
        [string]$Impact = 'Partial',

        [bool]$AffectsReportIntegrity = $true
    )

    $entry = [ordered]@{
        Scope     = $Scope
        Message   = $Message
        Timestamp = (Get-Date).ToString('o')
        Impact    = $Impact
        AffectsReportIntegrity = $AffectsReportIntegrity
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
        [void](Add-WTSkippedCheck -Report $Report -Check $Name -Reason 'Requires administrator privileges.')
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
        $classification = Get-WTExecutionErrorClassification -Scope $Name -Message $_.Exception.Message
        [void](Add-WTExecutionError -Report $Report -Scope $Name -Message $_.Exception.Message -Impact $classification.Impact -AffectsReportIntegrity $classification.AffectsReportIntegrity)
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
            $classification = Get-WTExecutionErrorClassification -Scope $Name -Message $_.Exception.Message
            [void](Add-WTExecutionError -Report $Report -Scope $Name -Message $_.Exception.Message -Impact $classification.Impact -AffectsReportIntegrity $classification.AffectsReportIntegrity)
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
    $kernelUptimeText = 'Unknown'
    $diskText = 'Unknown'
    $ramText = 'Unknown'
    $powerText = 'Fast Startup Unknown, recent shutdowns Unknown, unexpected Unknown'
    $updatesText = 'Unknown'
    $eventsText = 'Unknown'
    $systemDriveLabel = 'C:'
    $powerBoot = $Report.Normalized.PowerBoot
    $updates = $Report.Normalized.Updates

    if ($Report.Normalized.System -and $Report.Normalized.System.OSCaption) {
        $osSummary = '{0} build {1}' -f $Report.Normalized.System.OSCaption, (Get-WTSystemBuildText -SystemInfo $Report.Normalized.System)
    }
    if ($powerBoot -and $powerBoot.KernelUptimeDays -ne $null) {
        $kernelUptimeText = '{0} days' -f $powerBoot.KernelUptimeDays
    }
    elseif ($Report.Normalized.System -and $Report.Normalized.System.KernelUptimeDays -ne $null) {
        $kernelUptimeText = '{0} days' -f $Report.Normalized.System.KernelUptimeDays
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
    if ($powerBoot) {
        $fastStartupText = ConvertTo-WTEnabledDisabledUnknown -Value $powerBoot.FastStartupEnabled
        $recentShutdownCount = if ($powerBoot.RecentShutdownEventsCount -ne $null) { $powerBoot.RecentShutdownEventsCount } else { 'Unknown' }
        $recentUnexpectedCount = if ($powerBoot.RecentUnexpectedShutdownEventsCount -ne $null) { $powerBoot.RecentUnexpectedShutdownEventsCount } else { 'Unknown' }
        $powerText = 'Fast Startup {0}, recent shutdowns {1}, unexpected {2}' -f $fastStartupText, $recentShutdownCount, $recentUnexpectedCount
    }
    if ($updates) {
        $updateRebootText = ConvertTo-WTYesNoUnknown -Value $updates.PendingReboot
        $restartSignalCount = 0
        $restartSignalKnown = $false
        if ($updates.PendingReboot -ne $null) {
            $restartSignalKnown = $true
            if ($updates.PendingReboot -eq $true) {
                $restartSignalCount++
            }
        }
        if ($updates.PendingFileRenameOperationsPresent -ne $null) {
            $restartSignalKnown = $true
            if ($updates.PendingFileRenameOperationsPresent -eq $true) {
                $restartSignalCount++
            }
        }
        $restartSignalText = if ($restartSignalKnown) { $restartSignalCount } else { 'Unknown' }
        $servicingIndicatorCount = @($updates.ServicingWerEvents).Count + @($updates.StoreWerEvents).Count + @($updates.EdgeUpdateWerEvents).Count
        $updatesText = '{0} WU errors, update reboot {1}, restart signals {2}, servicing indicators {3}' -f $updates.WindowsUpdateFailureCount, $updateRebootText, $restartSignalText, $servicingIndicatorCount
    }
    if ($Report.Normalized.Events) {
        if (@($Report.Normalized.Events.LogsUnavailable).Count -ge 2 -and @($Report.Normalized.Events.AllEvents).Count -eq 0) {
            $eventsText = 'Unknown'
        }
        else {
            $eventsText = '{0} unexpected shutdowns, {1} bugchecks, {2} app crashes' -f @($Report.Normalized.Events.UnexpectedShutdownEvents).Count, @($Report.Normalized.Events.BugCheckEvents).Count, @($Report.Normalized.Events.ApplicationCrashEvents).Count
        }
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
        ('Kernel uptime: {0}' -f $kernelUptimeText),
        ('Disk: {0}' -f $diskText),
        ('RAM: {0}' -f $ramText),
        ('Power: {0}' -f $powerText),
        ('Updates: {0}' -f $updatesText),
        ('Events: {0}' -f $eventsText),
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
    }
    else {
        Write-Host $lines[0] -ForegroundColor Cyan
        Write-Host $lines[1] -ForegroundColor Gray
        Write-Host $lines[2] -ForegroundColor White
        Write-Host $lines[3] -ForegroundColor DarkGray
        Write-Host $lines[4] -ForegroundColor Gray
        Write-Host $lines[5] -ForegroundColor Gray
        Write-Host $lines[6] -ForegroundColor Gray
        Write-Host $lines[7] -ForegroundColor Gray
        Write-Host $lines[8] -ForegroundColor Gray
        Write-Host $lines[9] -ForegroundColor Gray
        Write-Host $lines[10] -ForegroundColor Gray
        Write-Host $lines[11] -ForegroundColor DarkGray
        Write-Host $lines[12] -ForegroundColor Gray
        Write-Host $lines[13] -ForegroundColor Gray
        Write-Host $lines[14] -ForegroundColor DarkGray
    }

    $errorRows = @($Report.Execution.Errors)
    $warningRows = @($Report.Execution.Warnings)

    if ($errorRows.Count -gt 0) {
        if ($NoColor) {
            Write-Host 'Execution errors:'
        }
        else {
            Write-Host 'Execution errors:' -ForegroundColor Red
        }
        $errorDisplayRows = @($errorRows | Select-Object -First 3)
        foreach ($err in $errorDisplayRows) {
            $line = '- {0}: {1}' -f (ConvertTo-WTDisplayValue -Value $err.Scope -Fallback 'Unknown'), (ConvertTo-WTDisplayValue -Value $err.Message -Fallback 'Unknown')
            if ($NoColor) {
                Write-Host $line
            }
            else {
                Write-Host $line -ForegroundColor DarkRed
            }
        }
        if ($errorRows.Count -gt 3) {
            $moreCount = $errorRows.Count - 3
            $moreLine = '... and {0} more' -f $moreCount
            if ($NoColor) {
                Write-Host $moreLine
            }
            else {
                Write-Host $moreLine -ForegroundColor DarkRed
            }
        }
    }

    if ($warningRows.Count -gt 0) {
        if ($NoColor) {
            Write-Host 'Execution warnings:'
        }
        else {
            Write-Host 'Execution warnings:' -ForegroundColor Yellow
        }
        $warningDisplayRows = @($warningRows | Select-Object -First 3)
        foreach ($warn in $warningDisplayRows) {
            $line = '- {0}: {1}' -f (ConvertTo-WTDisplayValue -Value $warn.Scope -Fallback 'Unknown'), (ConvertTo-WTDisplayValue -Value $warn.Message -Fallback 'Unknown')
            if ($NoColor) {
                Write-Host $line
            }
            else {
                Write-Host $line -ForegroundColor DarkYellow
            }
        }
        if ($warningRows.Count -gt 3) {
            $moreCount = $warningRows.Count - 3
            $moreLine = '... and {0} more' -f $moreCount
            if ($NoColor) {
                Write-Host $moreLine
            }
            else {
                Write-Host $moreLine -ForegroundColor DarkYellow
            }
        }
    }

    if ($NoColor) {
        Write-Host $lines[15]
        return
    }

    Write-Host $lines[15] -ForegroundColor Green
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
    $powerBoot = $Report.Normalized.PowerBoot
    $kernelLastBootText = 'Not available'
    if ($powerBoot -and $powerBoot.KernelLastBootUpTime) {
        $kernelLastBootText = ConvertTo-WTDisplayValue -Value $powerBoot.KernelLastBootUpTime -Fallback 'Not available'
    }
    elseif ($Report.Normalized.System -and $Report.Normalized.System.KernelLastBootUpTime) {
        $kernelLastBootText = ConvertTo-WTDisplayValue -Value $Report.Normalized.System.KernelLastBootUpTime -Fallback 'Not available'
    }
    $kernelUptimeText = 'Unknown'
    if ($powerBoot -and $powerBoot.KernelUptimeDays -ne $null) {
        $kernelUptimeText = '{0} days' -f $powerBoot.KernelUptimeDays
    }
    elseif ($Report.Normalized.System -and $Report.Normalized.System.KernelUptimeDays -ne $null) {
        $kernelUptimeText = '{0} days' -f $Report.Normalized.System.KernelUptimeDays
    }
    [void]$sb.AppendLine(('* BuildNumber: {0}' -f (ConvertTo-WTDisplayValue -Value $systemBuildText)))
    [void]$sb.AppendLine(('* Architecture: {0}' -f (ConvertTo-WTDisplayValue -Value $Report.Normalized.System.OSArchitecture)))
    [void]$sb.AppendLine(('* Hostname: {0}' -f (ConvertTo-WTDisplayValue -Value $Report.Normalized.System.Hostname)))
    [void]$sb.AppendLine(('* User: {0}' -f (ConvertTo-WTDisplayValue -Value $Report.Normalized.System.User)))
    [void]$sb.AppendLine(('* Manufacturer/Model: {0} / {1}' -f (ConvertTo-WTDisplayValue -Value $Report.Normalized.System.Manufacturer), (ConvertTo-WTDisplayValue -Value $Report.Normalized.System.Model)))
    [void]$sb.AppendLine(('* Serial Number: {0}' -f (ConvertTo-WTDisplayValue -Value $Report.Normalized.System.SerialNumber)))
    [void]$sb.AppendLine(('* Domain/Workgroup: {0}' -f (ConvertTo-WTDisplayValue -Value $Report.Normalized.System.DomainOrWorkgroup)))
    [void]$sb.AppendLine(('* InstallDate: {0}' -f $systemInstallDateText))
    [void]$sb.AppendLine(('* KernelLastBootUpTime: {0}' -f $kernelLastBootText))
    [void]$sb.AppendLine(('* Kernel uptime: {0}' -f $kernelUptimeText))
    [void]$sb.AppendLine(('* Virtual machine: {0}' -f (ConvertTo-WTYesNoUnknown -Value $Report.Normalized.System.IsVirtualMachine)))
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('## Power / Boot Overview')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine(('* Kernel last boot time: {0}' -f $kernelLastBootText))
    [void]$sb.AppendLine(('* Kernel uptime: {0}' -f $kernelUptimeText))
    [void]$sb.AppendLine(('* Fast Startup enabled: {0}' -f (ConvertTo-WTEnabledDisabledUnknown -Value $powerBoot.FastStartupEnabled)))
    [void]$sb.AppendLine(('* Recent boot events: {0}' -f (Get-WTArrayCountSafe -Value $powerBoot.BootEvents)))
    [void]$sb.AppendLine(('* Recent shutdown events: {0}' -f (Get-WTArrayCountSafe -Value $powerBoot.ShutdownEvents)))
    [void]$sb.AppendLine(('* Planned shutdown/restart events: {0}' -f (Get-WTArrayCountSafe -Value $powerBoot.PlannedShutdownEvents)))
    [void]$sb.AppendLine(('* Unexpected shutdown events: {0}' -f (Get-WTArrayCountSafe -Value $powerBoot.UnexpectedShutdownEvents)))
    [void]$sb.AppendLine(('* Sleep/hibernate events: {0}' -f (Get-WTArrayCountSafe -Value $powerBoot.SleepHibernateEvents)))
    [void]$sb.AppendLine(('* Resume events: {0}' -f (Get-WTArrayCountSafe -Value $powerBoot.ResumeEvents)))
    [void]$sb.AppendLine(('* Last boot event: {0}' -f (ConvertTo-WTDisplayValue -Value (ConvertTo-WTDateTimeString -Value $powerBoot.LastBootEventTime) -Fallback 'Not available')))
    [void]$sb.AppendLine(('* Last shutdown event: {0}' -f (ConvertTo-WTDisplayValue -Value (ConvertTo-WTDateTimeString -Value $powerBoot.LastShutdownEventTime) -Fallback 'Not available')))
    [void]$sb.AppendLine(('* Last planned shutdown event: {0}' -f (ConvertTo-WTDisplayValue -Value (ConvertTo-WTDateTimeString -Value $powerBoot.LastPlannedShutdownEventTime) -Fallback 'Not available')))
    [void]$sb.AppendLine(('* Last unexpected shutdown event: {0}' -f (ConvertTo-WTDisplayValue -Value (ConvertTo-WTDateTimeString -Value $powerBoot.LastUnexpectedShutdownEventTime) -Fallback 'Not available')))
    [void]$sb.AppendLine(('* Last resume event: {0}' -f (ConvertTo-WTDisplayValue -Value (ConvertTo-WTDateTimeString -Value $powerBoot.LastResumeEventTime) -Fallback 'Not available')))
    [void]$sb.AppendLine(('* Interpretation: {0}' -f (ConvertTo-WTDisplayValue -Value $powerBoot.PowerCycleInterpretation -Fallback 'Unknown')))
    [void]$sb.AppendLine(('* Uptime interpretation: {0}' -f (ConvertTo-WTDisplayValue -Value $powerBoot.UptimeInterpretation -Fallback 'Unknown')))
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('## Windows Update / Servicing Overview')
    [void]$sb.AppendLine('')
    $updates = $Report.Normalized.Updates
    if ($updates) {
        $updateRebootText = ConvertTo-WTYesNoUnknown -Value $updates.PendingReboot
        $genericRestartText = ConvertTo-WTYesNoUnknown -Value $updates.GenericRestartSignal
        $pendingFileRenameText = ConvertTo-WTYesNoUnknown -Value $updates.PendingFileRenameOperationsPresent
        $pendingRebootSourceText = ConvertTo-WTDisplayValue -Value $updates.PendingRebootSource -Fallback 'None'
        $cbsExistsText = ConvertTo-WTYesNoUnknown -Value $updates.CbsLog.Exists
        $dismExistsText = ConvertTo-WTYesNoUnknown -Value $updates.DismLog.Exists
        $cbsLastWriteText = 'Not available'
        if ($updates.CbsLog -and $updates.CbsLog.LastWriteTime) {
            $cbsLastWriteText = ConvertTo-WTDisplayValue -Value (ConvertTo-WTDateTimeString -Value $updates.CbsLog.LastWriteTime) -Fallback 'Not available'
        }
        $dismLastWriteText = 'Not available'
        if ($updates.DismLog -and $updates.DismLog.LastWriteTime) {
            $dismLastWriteText = ConvertTo-WTDisplayValue -Value (ConvertTo-WTDateTimeString -Value $updates.DismLog.LastWriteTime) -Fallback 'Not available'
        }
        $cbsSizeText = 'Not available'
        if ($updates.CbsLog -and $updates.CbsLog.SizeMB -ne $null) {
            $cbsSizeText = '{0} MB' -f $updates.CbsLog.SizeMB
        }
        [void]$sb.AppendLine(('* Hotfix count: {0}' -f (ConvertTo-WTDisplayValue -Value $updates.HotfixCount)))
        [void]$sb.AppendLine(('* Last hotfix installed: {0}' -f (ConvertTo-WTDisplayValue -Value (ConvertTo-WTDateTimeString -Value $updates.LastHotfixInstalledOn) -Fallback 'Not available')))
        [void]$sb.AppendLine(('* Update pending reboot: {0}' -f $updateRebootText))
        [void]$sb.AppendLine(('* Generic restart signal: {0}' -f $genericRestartText))
        [void]$sb.AppendLine(('* Pending file rename operations: {0}' -f $pendingFileRenameText))
        [void]$sb.AppendLine(('* Pending reboot source: {0}' -f $pendingRebootSourceText))
        [void]$sb.AppendLine(('* CBS.log exists: {0}' -f $cbsExistsText))
        [void]$sb.AppendLine(('* CBS.log last write time: {0}' -f $cbsLastWriteText))
        [void]$sb.AppendLine(('* CBS.log size: {0}' -f $cbsSizeText))
        [void]$sb.AppendLine(('* DISM log exists: {0}' -f $dismExistsText))
        [void]$sb.AppendLine(('* DISM log last write time: {0}' -f $dismLastWriteText))
        [void]$sb.AppendLine(('* Windows Update events: {0}' -f (Get-WTArrayCountSafe -Value $updates.WindowsUpdateEvents)))
        [void]$sb.AppendLine(('* Windows Update errors: {0}' -f (Get-WTArrayCountSafe -Value $updates.WindowsUpdateErrorEvents)))
        [void]$sb.AppendLine(('* Windows Update warnings: {0}' -f (Get-WTArrayCountSafe -Value $updates.WindowsUpdateWarningEvents)))
        [void]$sb.AppendLine(('* Servicing WER events: {0}' -f (Get-WTArrayCountSafe -Value $updates.ServicingWerEvents)))
        [void]$sb.AppendLine(('* Store WER events: {0}' -f (Get-WTArrayCountSafe -Value $updates.StoreWerEvents)))
        [void]$sb.AppendLine(('* Edge/WebView WER events: {0}' -f (Get-WTArrayCountSafe -Value $updates.EdgeUpdateWerEvents)))
        [void]$sb.AppendLine(('* Update health summary: {0}' -f (ConvertTo-WTDisplayValue -Value $updates.UpdateHealthSummary -Fallback 'Unknown')))
        [void]$sb.AppendLine(('* Servicing health summary: {0}' -f (ConvertTo-WTDisplayValue -Value $updates.ServicingHealthSummary -Fallback 'Unknown')))
        [void]$sb.AppendLine(('* Store update summary: {0}' -f (ConvertTo-WTDisplayValue -Value $updates.StoreUpdateSummary -Fallback 'Unknown')))
        [void]$sb.AppendLine(('* Edge update summary: {0}' -f (ConvertTo-WTDisplayValue -Value $updates.EdgeUpdateSummary -Fallback 'Unknown')))
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('### Windows Update Services')
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('| Service | DisplayName | Status | StartType |')
        [void]$sb.AppendLine('| --- | --- | --- | --- |')
        $serviceRows = @($updates.Services)
        if ($serviceRows.Count -eq 0) {
            [void]$sb.AppendLine('| Unknown | Unknown | Unknown | Unknown |')
        }
        else {
            foreach ($svc in $serviceRows) {
                [void]$sb.AppendLine(('| {0} | {1} | {2} | {3} |' -f (ConvertTo-WTDisplayValue -Value $svc.Name), (ConvertTo-WTDisplayValue -Value $svc.DisplayName), (ConvertTo-WTDisplayValue -Value $svc.Status), (ConvertTo-WTDisplayValue -Value $svc.StartType)))
            }
        }
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('### Recent Windows Update Events')
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('| TimeCreated | Id | Level | MessageShort |')
        [void]$sb.AppendLine('| --- | ---: | --- | --- |')
        $wuRecentRows = @($updates.RecentWindowsUpdateEvents)
        if ($wuRecentRows.Count -eq 0) {
            [void]$sb.AppendLine('No relevant update events found.')
        }
        else {
            foreach ($evt in $wuRecentRows) {
                $timeText = ConvertTo-WTDisplayValue -Value (ConvertTo-WTDateTimeString -Value $evt.TimeCreated) -Fallback 'Unknown'
                $levelText = (ConvertTo-WTDisplayValue -Value $evt.LevelDisplayName) -replace '\|', '\|'
                $messageText = (ConvertTo-WTDisplayValue -Value $evt.MessageShort) -replace '\|', '\|'
                [void]$sb.AppendLine(('| {0} | {1} | {2} | {3} |' -f $timeText, (ConvertTo-WTDisplayValue -Value $evt.Id), $levelText, $messageText))
            }
        }
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('### Servicing / Store / Edge Update Indicators')
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('| TimeCreated | Type | WerEventName | MessageShort |')
        [void]$sb.AppendLine('| --- | --- | --- | --- |')
        $indicatorRows = @($updates.RecentUpdateIndicators)
        if ($indicatorRows.Count -eq 0) {
            [void]$sb.AppendLine('No update indicators detected.')
        }
        else {
            foreach ($evt in $indicatorRows) {
                $timeText = ConvertTo-WTDisplayValue -Value (ConvertTo-WTDateTimeString -Value $evt.TimeCreated) -Fallback 'Unknown'
                $typeText = ConvertTo-WTDisplayValue -Value $evt.Type
                $nameText = ConvertTo-WTDisplayValue -Value $evt.WerEventName
                $messageText = (ConvertTo-WTDisplayValue -Value $evt.MessageShort) -replace '\|', '\|'
                [void]$sb.AppendLine(('| {0} | {1} | {2} | {3} |' -f $timeText, $typeText, $nameText, $messageText))
            }
        }
        [void]$sb.AppendLine('')
    }
    else {
        [void]$sb.AppendLine('No Windows Update data available.')
        [void]$sb.AppendLine('')
    }
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
    $performanceKernelUptimeText = 'Unknown'
    if ($Report.Normalized.Performance -and $Report.Normalized.Performance.KernelUptimeDays -ne $null) {
        $performanceKernelUptimeText = '{0} days' -f $Report.Normalized.Performance.KernelUptimeDays
    }
    elseif ($Report.Normalized.Performance -and $Report.Normalized.Performance.UptimeDays -ne $null) {
        $performanceKernelUptimeText = '{0} days' -f $Report.Normalized.Performance.UptimeDays
    }
    [void]$sb.AppendLine(('* Kernel uptime: {0}' -f $performanceKernelUptimeText))
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
    [void]$sb.AppendLine('## Event Overview')
    [void]$sb.AppendLine('')
    $eventOverview = $Report.Normalized.Events
    if ($eventOverview) {
        $windowStartText = ConvertTo-WTDisplayValue -Value (ConvertTo-WTDateTimeString -Value $eventOverview.WindowStart) -Fallback 'Unknown'
        $windowEndText = ConvertTo-WTDisplayValue -Value (ConvertTo-WTDateTimeString -Value $eventOverview.WindowEnd) -Fallback 'Unknown'
        $logsUnavailableText = 'None'
        if ($eventOverview.LogsUnavailable -and @($eventOverview.LogsUnavailable).Count -gt 0) {
            $logsUnavailableText = @($eventOverview.LogsUnavailable | ForEach-Object {
                if ($_.LogName -and $_.Reason) {
                    '{0}: {1}' -f $_.LogName, $_.Reason
                }
                elseif ($_.LogName) {
                    $_.LogName
                }
                else {
                    'Unknown'
                }
            }) -join ' | '
        }

        [void]$sb.AppendLine(('* WindowStart: {0}' -f $windowStartText))
        [void]$sb.AppendLine(('* WindowEnd: {0}' -f $windowEndText))
        [void]$sb.AppendLine(('* Days: {0}' -f (ConvertTo-WTDisplayValue -Value $eventOverview.Days)))
        [void]$sb.AppendLine(('* Total events collected: {0}' -f (ConvertTo-WTDisplayValue -Value $eventOverview.TotalEvents)))
        [void]$sb.AppendLine(('* System events: {0}' -f (ConvertTo-WTDisplayValue -Value $eventOverview.SystemEventCount)))
        [void]$sb.AppendLine(('* Application events: {0}' -f (ConvertTo-WTDisplayValue -Value $eventOverview.ApplicationEventCount)))
        [void]$sb.AppendLine(('* Unexpected shutdown events: {0}' -f (Get-WTArrayCountSafe -Value $eventOverview.UnexpectedShutdownEvents)))
        [void]$sb.AppendLine(('* Kernel-Power events: {0}' -f (Get-WTArrayCountSafe -Value $eventOverview.KernelPowerEvents)))
        [void]$sb.AppendLine(('* BugCheck events: {0}' -f (Get-WTArrayCountSafe -Value $eventOverview.BugCheckEvents)))
        [void]$sb.AppendLine(('* Normal shutdown events: {0}' -f (Get-WTArrayCountSafe -Value $eventOverview.NormalShutdownEvents)))
        [void]$sb.AppendLine(('* Service failure events: {0}' -f (Get-WTArrayCountSafe -Value $eventOverview.ServiceFailureEvents)))
        [void]$sb.AppendLine(('* Service install events: {0}' -f (Get-WTArrayCountSafe -Value $eventOverview.ServiceInstallEvents)))
        [void]$sb.AppendLine(('* Application crash events: {0}' -f (Get-WTArrayCountSafe -Value $eventOverview.ApplicationCrashEvents)))
        [void]$sb.AppendLine(('* WER non-critical events: {0}' -f (Get-WTArrayCountSafe -Value $eventOverview.NonCriticalWerEvents)))
        [void]$sb.AppendLine(('* Logs unavailable: {0}' -f $logsUnavailableText))
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('### Recent Important Events')
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('| TimeCreated | Log | Provider | Id | Level | MessageShort |')
        [void]$sb.AppendLine('| --- | --- | --- | ---: | --- | --- |')
        $recentEvents = @($eventOverview.RecentImportantEvents)
        if ($recentEvents.Count -eq 0) {
            [void]$sb.AppendLine('No relevant events found.')
        }
        else {
            foreach ($evt in $recentEvents) {
                $timeText = ConvertTo-WTDisplayValue -Value (ConvertTo-WTDateTimeString -Value $evt.TimeCreated) -Fallback 'Unknown'
                $logText = (ConvertTo-WTDisplayValue -Value $evt.LogName) -replace '\|', '\|'
                $providerText = (ConvertTo-WTDisplayValue -Value $evt.ProviderName) -replace '\|', '\|'
                $levelText = (ConvertTo-WTDisplayValue -Value $evt.LevelDisplayName) -replace '\|', '\|'
                $messageText = (ConvertTo-WTDisplayValue -Value $evt.MessageShort) -replace '\|', '\|'
                [void]$sb.AppendLine(('| {0} | {1} | {2} | {3} | {4} | {5} |' -f $timeText, $logText, $providerText, (ConvertTo-WTDisplayValue -Value $evt.Id), $levelText, $messageText))
            }
        }
    }
    else {
        [void]$sb.AppendLine('No relevant events found.')
    }
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('### Application Crash Summary')
    [void]$sb.AppendLine('')
    $crashSummaryRows = @($eventOverview.ApplicationCrashSummaryByProcess)
    if ($crashSummaryRows.Count -eq 0) {
        [void]$sb.AppendLine('No application crashes detected.')
    }
    else {
        [void]$sb.AppendLine('| Process | Count | LastEvent | Example |')
        [void]$sb.AppendLine('| --- | ---: | --- | --- |')
        foreach ($row in @($crashSummaryRows | Select-Object -First 10)) {
            $rowProcess = ConvertTo-WTDisplayValue -Value $row.ProcessName
            $rowCount = if ($null -ne $row.Count) { $row.Count } else { 0 }
            $rowLast = ConvertTo-WTDisplayValue -Value (ConvertTo-WTDateTimeString -Value $row.LastEvent) -Fallback 'Unknown'
            $rowExample = ConvertTo-WTDisplayValue -Value $row.ExampleMessageShort -Fallback 'Unknown'
            [void]$sb.AppendLine(('| {0} | {1} | {2} | {3} |' -f $rowProcess, $rowCount, $rowLast, $rowExample))
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
    [void]$sb.AppendLine(('* Exit code: {0}' -f (ConvertTo-WTDisplayValue -Value $Report.Metadata.ExitCode)))
    [void]$sb.AppendLine(('* Output fallback used: {0}' -f (ConvertTo-WTYesNo -Value ([bool]$Report.Metadata.OutputFallbackUsed))))
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine(('* Errors: {0}' -f @($Report.Execution.Errors).Count))
    [void]$sb.AppendLine(('* Warnings: {0}' -f @($Report.Execution.Warnings).Count))
    [void]$sb.AppendLine(('* Skipped checks: {0}' -f @($Report.Execution.Skipped).Count))
    [void]$sb.AppendLine('')

    $executionErrors = @($Report.Execution.Errors)
    if ($executionErrors.Count -eq 0) {
        [void]$sb.AppendLine('### Execution Errors')
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('No execution errors recorded.')
        [void]$sb.AppendLine('')
    }
    else {
        [void]$sb.AppendLine('### Execution Errors')
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('| Scope | Message | Impact | AffectsReportIntegrity | Timestamp |')
        [void]$sb.AppendLine('| --- | --- | --- | --- | --- |')
        foreach ($err in @($executionErrors | Select-Object -First 20)) {
            [void]$sb.AppendLine(('| {0} | {1} | {2} | {3} | {4} |' -f (
                (ConvertTo-WTDisplayValue -Value $err.Scope -Fallback 'Unknown') -replace '\|', '\|',
                (ConvertTo-WTDisplayValue -Value $err.Message -Fallback 'Unknown') -replace '\|', '\|',
                (ConvertTo-WTDisplayValue -Value $err.Impact -Fallback 'Unknown') -replace '\|', '\|',
                (ConvertTo-WTYesNo -Value ([bool]$err.AffectsReportIntegrity)),
                (ConvertTo-WTDisplayValue -Value $err.Timestamp -Fallback 'Unknown') -replace '\|', '\|'
            )))
        }
        [void]$sb.AppendLine('')
    }

    $executionWarnings = @($Report.Execution.Warnings)
    if ($executionWarnings.Count -eq 0) {
        [void]$sb.AppendLine('### Execution Warnings')
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('No execution warnings recorded.')
        [void]$sb.AppendLine('')
    }
    else {
        [void]$sb.AppendLine('### Execution Warnings')
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('| Scope | Message | Timestamp |')
        [void]$sb.AppendLine('| --- | --- | --- |')
        foreach ($warn in @($executionWarnings | Select-Object -First 20)) {
            [void]$sb.AppendLine(('| {0} | {1} | {2} |' -f (
                (ConvertTo-WTDisplayValue -Value $warn.Scope -Fallback 'Unknown') -replace '\|', '\|',
                (ConvertTo-WTDisplayValue -Value $warn.Message -Fallback 'Unknown') -replace '\|', '\|',
                (ConvertTo-WTDisplayValue -Value $warn.Timestamp -Fallback 'Unknown') -replace '\|', '\|'
            )))
        }
        [void]$sb.AppendLine('')
    }

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

    if ($null -eq $Value) {
        return $Fallback
    }

    if ($Value -is [string] -and [string]::IsNullOrWhiteSpace($Value)) {
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

function ConvertTo-WTEnabledDisabledUnknown {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($Value -eq $true -or $Value -eq 'True') {
        return 'Enabled'
    }

    if ($Value -eq $false -or $Value -eq 'False') {
        return 'Disabled'
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

function ConvertTo-WTDateSortValue {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value -or $Value -eq '') {
        return [datetime]::MinValue
    }

    try {
        return [datetime]$Value
    }
    catch {
        return [datetime]::MinValue
    }
}

function Get-WTFileMetadataInfo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $result = [ordered]@{
        Exists            = $false
        Path              = $Path
        SizeMB            = $null
        LastWriteTime     = $null
        IsLarge           = $false
        IsRecentlyModified = $false
    }

    try {
        if (Test-Path -LiteralPath $Path) {
            $item = Get-Item -LiteralPath $Path -ErrorAction Stop
            $sizeMB = $null
            try {
                if ($item.Length -ne $null) {
                    $sizeMB = [math]::Round(($item.Length / 1MB), 2)
                }
            }
            catch {
                $sizeMB = $null
            }

            $lastWriteTime = $null
            try {
                $lastWriteTime = $item.LastWriteTime
            }
            catch {
                $lastWriteTime = $null
            }

            $result.Exists = $true
            $result.SizeMB = $sizeMB
            $result.LastWriteTime = $lastWriteTime
            $result.IsLarge = ($sizeMB -ne $null -and $sizeMB -ge 100)
            $result.IsRecentlyModified = ($lastWriteTime -ne $null -and ((Get-Date) - $lastWriteTime).TotalHours -le 24)
        }
    }
    catch {
        $result.Exists = $false
    }

    return [pscustomobject]$result
}

function Get-WTRegistryBooleanState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Name,

        [ValidateSet('Presence', 'NonZero', 'ValuePresence')]
        [string]$Mode = 'Presence'
    )

    try {
        switch ($Mode) {
            'Presence' {
                return [bool](Test-Path -LiteralPath $Path)
            }
            'NonZero' {
                $item = Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop
                $value = $item.$Name
                if ($null -eq $value -or $value -eq '') {
                    return $false
                }
                return ([int64]$value -gt 0)
            }
            'ValuePresence' {
                $item = Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop
                $value = $item.$Name
                if ($null -eq $value) {
                    return $false
                }
                if ($value -is [array]) {
                    return (@($value).Count -gt 0)
                }
                if ($value -is [string]) {
                    return -not [string]::IsNullOrWhiteSpace($value)
                }
                return $true
            }
        }
    }
    catch {
        return $null
    }
}

function Get-WTPendingFileRenameOperationsInfo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    try {
        $item = Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop
        $rawValue = $item.$Name
    }
    catch {
        return [pscustomobject]@{
            Present = $null
            Count   = $null
            Sample  = @()
        }
    }

    $entries = @()
    if ($null -ne $rawValue) {
        foreach ($entry in @($rawValue)) {
            if (-not [string]::IsNullOrWhiteSpace([string]$entry)) {
                $entries += ([string]$entry).Trim()
            }
        }
    }

    return [pscustomobject]@{
        Present = ($entries.Count -gt 0)
        Count   = $entries.Count
        Sample  = @($entries | Select-Object -First 10)
    }
}

function Get-WTPendingRebootSource {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [psobject]$PendingDetails
    )

    if (-not $PendingDetails) {
        return 'None'
    }

    $allPrimaryUnknown = (
        $null -eq $PendingDetails.CbsRebootPending -and
        $null -eq $PendingDetails.WuRebootRequired -and
        $null -eq $PendingDetails.UpdateExeVolatile -and
        $null -eq $PendingDetails.PendingFileRenameOperations
    )
    if ($allPrimaryUnknown) {
        return 'Unknown'
    }

    $sources = New-Object System.Collections.Generic.List[string]

    if ($PendingDetails.CbsRebootPending -eq $true) {
        [void]$sources.Add('CBS')
    }
    if ($PendingDetails.WuRebootRequired -eq $true) {
        [void]$sources.Add('WindowsUpdate')
    }
    if ($PendingDetails.UpdateExeVolatile -eq $true) {
        [void]$sources.Add('UpdateExeVolatile')
    }

    if ($PendingDetails.GenericPendingFileRename -eq $true) {
        if ($sources.Count -eq 0) {
            return 'PendingFileRenameOperationsOnly'
        }

        [void]$sources.Add('PendingFileRenameOperations')
    }

    if ($sources.Count -eq 0) {
        return 'None'
    }

    if ($sources.Count -eq 1) {
        return $sources[0]
    }

    return 'Mixed'
}

function Get-WTUpdateCollectionLimit {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [string]$Mode = 'Standard'
    )

    switch ($Mode) {
        'Quick' { return 100 }
        'Full' { return 1000 }
        default { return 300 }
    }
}

function Get-WTSetupCollectionLimit {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [string]$Mode = 'Standard'
    )

    switch ($Mode) {
        'Quick' { return 50 }
        'Full' { return 300 }
        default { return 150 }
    }
}

function Test-WTTextContainsAny {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [string]$Text,

        [Parameter(Mandatory = $true)]
        [string[]]$Patterns
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $false
    }

    foreach ($pattern in $Patterns) {
        if ($Text -match $pattern) {
            return $true
        }
    }

    return $false
}

function Get-WTServiceSnapshot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $service = $null
    $cimService = $null
    $exists = $true
    $status = 'Unknown'
    $startType = 'Unknown'
    $displayName = $Name
    $canStop = $null

    try {
        $service = Get-Service -Name $Name -ErrorAction Stop
    }
    catch {
        $exists = $false
    }

    if ($service) {
        $status = [string]$service.Status
        $displayName = $service.DisplayName
        try {
            $canStop = [bool]$service.CanStop
        }
        catch {
            $canStop = $null
        }
    }

    try {
        $escapedName = $Name.Replace("'", "''")
        $cimService = Get-CimInstance -ClassName Win32_Service -Filter ("Name='{0}'" -f $escapedName) -ErrorAction Stop | Select-Object -First 1
    }
    catch {
        $cimService = $null
    }

    if ($cimService -and $cimService.StartMode) {
        $startType = [string]$cimService.StartMode
    }

    return [pscustomobject]@{
        Name        = $Name
        DisplayName = $displayName
        Status      = $status
        StartType   = $startType
        CanStop     = $canStop
        Exists      = $exists
    }
}

function Get-WTEventIndicatorCategory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$Event
    )

    $combinedText = @(
        $Event.WerEventName
        $Event.Message
        $Event.MessageShort
    ) -join "`n"

    if (Test-WTTextContainsAny -Text $combinedText -Patterns @(
        'WindowsWcpOtherFailure3',
        'WindowsWcp',
        'CBS\.log',
        'CbsPersist',
        'componentstore',
        'Component Based Servicing',
        'WinSxS',
        'scavenge\.cpp'
    )) {
        return 'Servicing'
    }

    if (Test-WTTextContainsAny -Text $combinedText -Patterns @(
        'StoreAgentSearchUpdatePackagesFailure',
        'StoreAgentScanForUpdatesFailure',
        'StoreAgentInstallFailure',
        'AppxDeploymentFailureBlue',
        'Microsoft Store',
        'StoreAgent'
    )) {
        return 'Store'
    }

    if (Test-WTTextContainsAny -Text $combinedText -Patterns @(
        'crashpad_log',
        'MicrosoftEdgeUpdate\.exe',
        'msedge_installer\.log',
        'webview',
        'WebView'
    )) {
        return 'Edge'
    }

    return $null
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
    $kernelLastBootUpTime = $null
    $uptimeDays = $null
    $kernelUptimeDays = $null
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
        $kernelLastBootUpTime = $lastBoot
        $kernelUptimeDays = $uptimeDays
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
        KernelLastBootUpTime  = $kernelLastBootUpTime
        UptimeDays            = $uptimeDays
        KernelUptimeDays      = $kernelUptimeDays
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
            KernelLastBootUpTime      = $null
            UptimeDays                = $null
            KernelUptimeDays          = $null
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
        KernelLastBootUpTime      = $SystemInfo.KernelLastBootUpTime
        UptimeDays                = $SystemInfo.UptimeDays
        KernelUptimeDays          = $SystemInfo.KernelUptimeDays
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
            KernelLastBootUpTime     = $null
            KernelUptimeDays         = $null
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
        KernelLastBootUpTime    = ConvertTo-WTDisplayValue -Value $PerformanceInfo.KernelLastBootUpTime
        KernelUptimeDays        = $PerformanceInfo.KernelUptimeDays
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
        KernelLastBootUpTime   = if ($os) { ConvertTo-WTDateTimeString -Value $os.LastBootUpTime } else { $null }
        KernelUptimeDays       = $uptimeDays
        ProcessorName          = if ($processorName) { $processorName } else { $null }
        NumberOfLogicalProcessors = $logicalProcessors
        NumberOfProcessors     = $physicalProcessors
        CpuLoadPercent         = $cpuLoadPercent
        TopProcessesByCPU      = $topProcessesByCpu
        TopProcessesByMemory   = $topProcessesByMemory
    }
}

function Get-WTPowerBootInfo {
    [CmdletBinding()]
    param(
        [ValidateRange(1, 90)]
        [int]$Days = 7,

        [AllowNull()]
        [string]$Mode = 'Standard'
    )

    $windowEnd = Get-Date
    $windowStart = $windowEnd.AddDays(-1 * [math]::Abs($Days))
    $limit = Get-WTEventCollectionLimit -Mode $Mode
    $os = $null
    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
    }
    catch {
        $os = $null
    }

    $kernelLastBootUpTime = $null
    $kernelUptimeDays = $null
    if ($os) {
        $kernelLastBootUpTime = ConvertTo-WTDateTimeString -Value $os.LastBootUpTime
        $kernelUptimeDays = ConvertTo-WTUptimeDays -LastBootUpTime $os.LastBootUpTime
    }

    $fastStartupEnabled = $null
    try {
        $hiberboot = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power' -Name HiberbootEnabled -ErrorAction Stop
        if ($null -ne $hiberboot.HiberbootEnabled) {
            if ([int]$hiberboot.HiberbootEnabled -eq 1) {
                $fastStartupEnabled = $true
            }
            elseif ([int]$hiberboot.HiberbootEnabled -eq 0) {
                $fastStartupEnabled = $false
            }
        }
    }
    catch {
        $fastStartupEnabled = $null
    }

    $systemIds = @(12, 13, 41, 42, 107, 1074, 6005, 6006, 6008)
    $rawEvents = @()
    try {
        $rawEvents = @(Get-WinEvent -FilterHashtable @{ LogName = 'System'; StartTime = $windowStart; Id = $systemIds } -MaxEvents ($limit + 1) -ErrorAction Stop)
    }
    catch {
        $rawEvents = @()
    }

    if ($rawEvents.Count -gt $limit) {
        $rawEvents = @($rawEvents | Select-Object -First $limit)
    }

    $events = @()
    foreach ($event in $rawEvents) {
        $events += ConvertTo-WTEventRecord -EventRecord $event
    }

    $bootEvents = @($events | Where-Object { $_.Id -in @(12, 6005) -and ($_.ProviderName -match 'Kernel-General|EventLog') })
    $shutdownEvents = @($events | Where-Object { $_.Id -in @(13, 6006, 1074) -and ($_.ProviderName -match 'Kernel-General|EventLog|User32') })
    $plannedShutdownEvents = @($events | Where-Object { $_.Id -eq 1074 -and $_.ProviderName -eq 'User32' })
    $unexpectedShutdownEvents = @($events | Where-Object { $_.Id -in @(41, 6008) -and ($_.ProviderName -match 'Kernel-Power|EventLog') })
    $sleepHibernateEvents = @($events | Where-Object { $_.Id -eq 42 -and $_.ProviderName -match 'Kernel-Power' })
    $resumeEvents = @($events | Where-Object { $_.Id -eq 107 -and $_.ProviderName -match 'Kernel-Power' })

    return [pscustomobject]@{
        WindowStart                   = $windowStart
        WindowEnd                     = $windowEnd
        Days                          = $Days
        FastStartupEnabled            = $fastStartupEnabled
        KernelLastBootUpTime          = $kernelLastBootUpTime
        KernelUptimeDays              = $kernelUptimeDays
        Events                        = @($events)
        BootEvents                    = @($bootEvents)
        ShutdownEvents                = @($shutdownEvents)
        PlannedShutdownEvents         = @($plannedShutdownEvents)
        UnexpectedShutdownEvents      = @($unexpectedShutdownEvents)
        SleepHibernateEvents          = @($sleepHibernateEvents)
        ResumeEvents                  = @($resumeEvents)
        RecentBootEventsCount         = @($bootEvents).Count
        RecentShutdownEventsCount     = @($shutdownEvents).Count
        RecentPlannedShutdownEventsCount = @($plannedShutdownEvents).Count
        RecentUnexpectedShutdownEventsCount = @($unexpectedShutdownEvents).Count
        RecentSleepHibernateEventsCount = @($sleepHibernateEvents).Count
        RecentResumeEventsCount       = @($resumeEvents).Count
        LastBootEventTime             = if ($bootEvents.Count -gt 0) { ($bootEvents | Sort-Object -Property TimeCreated -Descending | Select-Object -First 1).TimeCreated } else { $null }
        LastShutdownEventTime         = if ($shutdownEvents.Count -gt 0) { ($shutdownEvents | Sort-Object -Property TimeCreated -Descending | Select-Object -First 1).TimeCreated } else { $null }
        LastPlannedShutdownEventTime  = if ($plannedShutdownEvents.Count -gt 0) { ($plannedShutdownEvents | Sort-Object -Property TimeCreated -Descending | Select-Object -First 1).TimeCreated } else { $null }
        LastUnexpectedShutdownEventTime = if ($unexpectedShutdownEvents.Count -gt 0) { ($unexpectedShutdownEvents | Sort-Object -Property TimeCreated -Descending | Select-Object -First 1).TimeCreated } else { $null }
        LastResumeEventTime           = if ($resumeEvents.Count -gt 0) { ($resumeEvents | Sort-Object -Property TimeCreated -Descending | Select-Object -First 1).TimeCreated } else { $null }
        PowerCycleInterpretation      = $null
        UptimeInterpretation          = $null
    }
}

function ConvertTo-WTNormalizedPowerBootInfo {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [psobject]$PowerBootInfo
    )

    if (-not $PowerBootInfo) {
        return [pscustomobject]@{
            KernelLastBootUpTime              = $null
            KernelUptimeDays                  = $null
            FastStartupEnabled                = $null
            RecentBootEventsCount             = 0
            RecentShutdownEventsCount         = 0
            RecentPlannedShutdownEventsCount  = 0
            RecentUnexpectedShutdownEventsCount = 0
            RecentSleepHibernateEventsCount   = 0
            RecentResumeEventsCount           = 0
            LastBootEventTime                 = $null
            LastShutdownEventTime             = $null
            LastPlannedShutdownEventTime      = $null
            LastUnexpectedShutdownEventTime   = $null
            LastResumeEventTime               = $null
            PowerCycleInterpretation          = 'No recent shutdown/start events found'
            UptimeInterpretation              = 'Kernel uptime is based on LastBootUpTime. Fast Startup status could not be determined.'
            BootEvents                        = @()
            ShutdownEvents                    = @()
            PlannedShutdownEvents             = @()
            UnexpectedShutdownEvents          = @()
            SleepHibernateEvents              = @()
            ResumeEvents                      = @()
            Events                            = @()
        }
    }

    $fastStartupText = ConvertTo-WTEnabledDisabledUnknown -Value $PowerBootInfo.FastStartupEnabled
    $powerCycleInterpretation = 'No recent shutdown/start events found'
    if ($PowerBootInfo.RecentUnexpectedShutdownEventsCount -gt 0) {
        $powerCycleInterpretation = 'Unexpected shutdown detected'
    }
    elseif ($PowerBootInfo.RecentShutdownEventsCount -gt 0 -or $PowerBootInfo.RecentBootEventsCount -gt 0) {
        $powerCycleInterpretation = 'Recent shutdown/start events found'
    }
    if ($fastStartupText -eq 'Enabled') {
        $powerCycleInterpretation = '{0}. Fast Startup may cause kernel uptime to persist across user shutdowns.' -f $powerCycleInterpretation
    }

    $uptimeInterpretation = 'Kernel uptime is based on LastBootUpTime. Fast Startup status could not be determined.'
    if ($fastStartupText -eq 'Enabled') {
        $uptimeInterpretation = 'Kernel uptime may not reset after user shutdown because Fast Startup is enabled.'
    }
    elseif ($fastStartupText -eq 'Disabled') {
        $uptimeInterpretation = 'Kernel uptime usually represents time since last full boot.'
    }

    return [pscustomobject]@{
        KernelLastBootUpTime              = ConvertTo-WTDisplayValue -Value $PowerBootInfo.KernelLastBootUpTime
        KernelUptimeDays                  = $PowerBootInfo.KernelUptimeDays
        FastStartupEnabled                = $PowerBootInfo.FastStartupEnabled
        RecentBootEventsCount             = [int]$PowerBootInfo.RecentBootEventsCount
        RecentShutdownEventsCount         = [int]$PowerBootInfo.RecentShutdownEventsCount
        RecentPlannedShutdownEventsCount  = [int]$PowerBootInfo.RecentPlannedShutdownEventsCount
        RecentUnexpectedShutdownEventsCount = [int]$PowerBootInfo.RecentUnexpectedShutdownEventsCount
        RecentSleepHibernateEventsCount   = [int]$PowerBootInfo.RecentSleepHibernateEventsCount
        RecentResumeEventsCount           = [int]$PowerBootInfo.RecentResumeEventsCount
        LastBootEventTime                 = $PowerBootInfo.LastBootEventTime
        LastShutdownEventTime             = $PowerBootInfo.LastShutdownEventTime
        LastPlannedShutdownEventTime      = $PowerBootInfo.LastPlannedShutdownEventTime
        LastUnexpectedShutdownEventTime   = $PowerBootInfo.LastUnexpectedShutdownEventTime
        LastResumeEventTime               = $PowerBootInfo.LastResumeEventTime
        PowerCycleInterpretation          = $powerCycleInterpretation
        UptimeInterpretation              = $uptimeInterpretation
        BootEvents                        = @($PowerBootInfo.BootEvents)
        ShutdownEvents                    = @($PowerBootInfo.ShutdownEvents)
        PlannedShutdownEvents             = @($PowerBootInfo.PlannedShutdownEvents)
        UnexpectedShutdownEvents          = @($PowerBootInfo.UnexpectedShutdownEvents)
        SleepHibernateEvents              = @($PowerBootInfo.SleepHibernateEvents)
        ResumeEvents                      = @($PowerBootInfo.ResumeEvents)
        Events                            = @($PowerBootInfo.Events)
    }
}

function Get-WTUpdateInfo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$Report,

        [ValidateRange(1, 90)]
        [int]$Days = 7,

        [AllowNull()]
        [string]$Mode = 'Standard'
    )

    $windowEnd = Get-Date
    $windowStart = $windowEnd.AddDays(-1 * [math]::Abs($Days))
    $updateLimit = Get-WTUpdateCollectionLimit -Mode $Mode
    $setupLimit = Get-WTSetupCollectionLimit -Mode $Mode
    $logsUnavailable = @()

    $hotfixLimit = switch ($Mode) {
        'Quick' { 10 }
        'Full' { 50 }
        default { 20 }
    }

    $hotfixes = @()
    try {
        if (Get-Command -Name Get-HotFix -ErrorAction SilentlyContinue) {
            $hotfixes = @(Get-HotFix -ErrorAction Stop | ForEach-Object {
                [pscustomobject]@{
                    HotFixID    = $_.HotFixID
                    Description = $_.Description
                    InstalledOn = ConvertTo-WTDateTimeString -Value $_.InstalledOn
                    InstalledBy = $_.InstalledBy
                    InstalledOnSort = ConvertTo-WTDateSortValue -Value $_.InstalledOn
                }
            } | Sort-Object -Property @{ Expression = { $_.InstalledOnSort }; Descending = $true }, @{ Expression = { $_.HotFixID }; Ascending = $true })

            if ($hotfixes.Count -gt $hotfixLimit) {
                [void](Add-WTExecutionWarning -Report $Report -Scope 'Updates' -Message ('Hotfix list truncated to the most recent {0} entries for mode {1}.' -f $hotfixLimit, $Mode))
                $hotfixes = @($hotfixes | Select-Object -First $hotfixLimit)
            }
        }
        else {
            [void](Add-WTExecutionWarning -Report $Report -Scope 'Updates' -Message 'Get-HotFix is not available on this system.')
        }
    }
    catch {
        [void](Add-WTExecutionWarning -Report $Report -Scope 'Updates' -Message ('Get-HotFix failed: {0}' -f $_.Exception.Message))
        $hotfixes = @()
    }

    $serviceNames = @('wuauserv', 'bits', 'cryptsvc', 'trustedinstaller', 'usosvc', 'WaaSMedicSvc')
    $services = @()
    foreach ($serviceName in $serviceNames) {
        $services += Get-WTServiceSnapshot -Name $serviceName
    }

    $pendingRenameInfo = Get-WTPendingFileRenameOperationsInfo -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name 'PendingFileRenameOperations'
    $pendingRebootDetails = [ordered]@{
        CbsRebootPending            = Get-WTRegistryBooleanState -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending' -Name 'RebootPending' -Mode Presence
        WuRebootRequired            = Get-WTRegistryBooleanState -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired' -Name 'RebootRequired' -Mode Presence
        PendingFileRenameOperations = $pendingRenameInfo.Present
        PendingFileRenameOperationsCount = $pendingRenameInfo.Count
        PendingFileRenameOperationsSample = @($pendingRenameInfo.Sample)
        UpdateExeVolatile           = Get-WTRegistryBooleanState -Path 'HKLM:\SOFTWARE\Microsoft\Updates' -Name 'UpdateExeVolatile' -Mode NonZero
    }

    $updatePendingReboot = ($pendingRebootDetails.CbsRebootPending -eq $true -or $pendingRebootDetails.WuRebootRequired -eq $true -or $pendingRebootDetails.UpdateExeVolatile -eq $true)
    $genericPendingFileRename = ($pendingRebootDetails.PendingFileRenameOperations -eq $true)
    $anyRestartSignal = ($updatePendingReboot -or $genericPendingFileRename)
    $pendingRebootDetails['UpdatePendingReboot'] = $updatePendingReboot
    $pendingRebootDetails['GenericPendingFileRename'] = $genericPendingFileRename
    $pendingRebootDetails['AnyRestartSignal'] = $anyRestartSignal
    $pendingRebootDetails['AnyPendingReboot'] = $updatePendingReboot
    $pendingRebootDetails['PendingRebootSource'] = Get-WTPendingRebootSource -PendingDetails ([pscustomobject]$pendingRebootDetails)
    $pendingRebootDetails = [pscustomobject]$pendingRebootDetails

    $cbsLog = Get-WTFileMetadataInfo -Path 'C:\Windows\Logs\CBS\CBS.log'
    $dismLog = Get-WTFileMetadataInfo -Path 'C:\Windows\Logs\DISM\dism.log'

    $wuLogName = 'Microsoft-Windows-WindowsUpdateClient/Operational'
    $wuIds = @(19, 20, 25, 31, 34, 41, 43, 44, 100, 101, 102, 103)
    $wuRawEvents = @()
    $wuQueryFailures = 0
    foreach ($query in @(
        @{
            LogName   = $wuLogName
            StartTime = $windowStart
            Id        = $wuIds
        },
        @{
            LogName   = $wuLogName
            StartTime = $windowStart
            Level     = 2
        },
        @{
            LogName   = $wuLogName
            StartTime = $windowStart
            Level     = 3
        }
    )) {
        try {
            $wuRawEvents += @(Get-WinEvent -FilterHashtable $query -MaxEvents ($updateLimit + 1) -ErrorAction Stop)
        }
        catch {
            $reason = $_.Exception.Message
            if ($reason -match '(?i)(no matching events|no se encontraron eventos|no events were found|no events match|no hay eventos)') {
                continue
            }
            $wuQueryFailures++
            [void](Add-WTExecutionWarning -Report $Report -Scope 'Updates' -Message ('Event collection failed for {0}: {1}' -f $wuLogName, $reason))
        }
    }

    if ($wuRawEvents.Count -eq 0 -and $wuQueryFailures -gt 0) {
        $logsUnavailable += [pscustomobject]@{
            LogName = $wuLogName
            Reason  = 'Unable to read Windows Update operational events.'
        }
    }

    $wuEvents = @()
    $seenKeys = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($event in @($wuRawEvents | Sort-Object -Property TimeCreated -Descending)) {
        $record = ConvertTo-WTEventRecord -EventRecord $event
        $key = Get-WTEventKey -Event $record
        if ($key -and $seenKeys.Add($key)) {
            $wuEvents += $record
        }
        if ($wuEvents.Count -ge $updateLimit) {
            break
        }
    }
    if ($wuEvents.Count -ge $updateLimit) {
        [void](Add-WTExecutionWarning -Report $Report -Scope 'Updates' -Message ('Event collection limit reached for {0}. Results may be incomplete.' -f $wuLogName))
    }

    $setupLogName = 'Setup'
    $setupRawEvents = @()
    try {
        $setupRawEvents = @(Get-WinEvent -FilterHashtable @{ LogName = $setupLogName; StartTime = $windowStart } -MaxEvents ($setupLimit + 1) -ErrorAction Stop)
    }
    catch {
        $reason = $_.Exception.Message
        if ($reason -match '(?i)(no matching events|no se encontraron eventos|no events were found|no events match|no hay eventos)') {
            $setupRawEvents = @()
        }
        else {
            $logsUnavailable += [pscustomobject]@{
                LogName = $setupLogName
                Reason  = $reason
            }
            [void](Add-WTExecutionWarning -Report $Report -Scope 'Updates' -Message ('Event collection failed for {0}: {1}' -f $setupLogName, $reason))
        }
        $setupRawEvents = @()
    }

    $setupEvents = @()
    $setupSeenKeys = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($event in @($setupRawEvents | Sort-Object -Property TimeCreated -Descending)) {
        $record = ConvertTo-WTEventRecord -EventRecord $event
        $key = Get-WTEventKey -Event $record
        if ($key -and $setupSeenKeys.Add($key)) {
            $setupEvents += $record
        }
        if ($setupEvents.Count -ge $setupLimit) {
            break
        }
    }
    if ($setupEvents.Count -ge $setupLimit) {
        [void](Add-WTExecutionWarning -Report $Report -Scope 'Updates' -Message ('Event collection limit reached for {0}. Results may be incomplete.' -f $setupLogName))
    }

    $servicingWerEvents = @()
    $storeWerEvents = @()
    $edgeUpdateWerEvents = @()
    $eventSource = @($Report.Normalized.Events.NonCriticalWerEvents)
    if ($eventSource.Count -eq 0 -and $Report.Raw.Events -and $Report.Raw.Events.ApplicationEvents) {
        $eventSource = @($Report.Raw.Events.ApplicationEvents | Where-Object { $_.ProviderName -eq 'Windows Error Reporting' -and $_.Id -eq 1001 })
    }

    foreach ($evt in $eventSource) {
        if (-not $evt) {
            continue
        }
        $category = Get-WTEventIndicatorCategory -Event $evt
        switch ($category) {
            'Servicing' { $servicingWerEvents += $evt }
            'Store'     { $storeWerEvents += $evt }
            'Edge'      { $edgeUpdateWerEvents += $evt }
        }
    }

    return [pscustomobject]@{
        WindowStart        = $windowStart
        WindowEnd          = $windowEnd
        Days               = $Days
        Mode               = $Mode
        Hotfixes           = @($hotfixes)
        Services           = @($services)
        PendingReboot      = $pendingRebootDetails.UpdatePendingReboot
        PendingRebootDetails = $pendingRebootDetails
        CbsLog             = $cbsLog
        DismLog            = $dismLog
        WindowsUpdateEvents = @($wuEvents)
        SetupEvents        = @($setupEvents)
        ServicingWerEvents = @($servicingWerEvents)
        StoreWerEvents     = @($storeWerEvents)
        EdgeUpdateWerEvents = @($edgeUpdateWerEvents)
        LogsUnavailable    = @($logsUnavailable)
    }
}

function ConvertTo-WTNormalizedUpdateInfo {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [psobject]$UpdateInfo
    )

    if (-not $UpdateInfo) {
        return [pscustomobject]@{
            WindowStart              = $null
            WindowEnd                = $null
            Days                     = $null
            HotfixCount              = 0
            RecentHotfixes           = @()
            LastHotfixInstalledOn    = $null
            Services                 = @()
            ServicesNotRunning       = @()
            ServicesUnknown          = @()
            PendingReboot            = $null
            GenericRestartSignal     = $null
            PendingFileRenameOperationsPresent = $null
            PendingFileRenameOperationsCount = $null
            PendingFileRenameOperationsSample = @()
            PendingRebootSource      = 'Unknown'
            PendingRebootDetails     = [pscustomobject]@{
                CbsRebootPending            = $null
                WuRebootRequired            = $null
                PendingFileRenameOperations = $null
                PendingFileRenameOperationsCount = $null
                PendingFileRenameOperationsSample = @()
                UpdateExeVolatile           = $null
                UpdatePendingReboot         = $null
                GenericPendingFileRename     = $null
                AnyRestartSignal            = $null
                AnyPendingReboot            = $null
                PendingRebootSource         = 'Unknown'
            }
            CbsLog                   = $null
            DismLog                  = $null
            WindowsUpdateEvents      = @()
            WindowsUpdateErrorEvents = @()
            WindowsUpdateWarningEvents = @()
            WindowsUpdateSuccessEvents = @()
            WindowsUpdateFailureCount = 0
            WindowsUpdateWarningCount = 0
            SetupEvents              = @()
            ServicingWerEvents       = @()
            StoreWerEvents           = @()
            EdgeUpdateWerEvents      = @()
            ServicingHealthSummary   = 'No servicing indicators detected'
            UpdateHealthSummary      = 'Windows Update log unavailable'
            StoreUpdateSummary       = 'No Store update indicators detected'
            EdgeUpdateSummary        = 'No Edge/WebView update indicators detected'
            RecentWindowsUpdateEvents = @()
            RecentUpdateIndicators    = @()
            LogsUnavailable          = @()
        }
    }

    $hotfixes = @($UpdateInfo.Hotfixes | Where-Object { $_ })
    $services = @($UpdateInfo.Services | Where-Object { $_ })
    $recentHotfixes = @($hotfixes | Sort-Object -Property @{ Expression = { ConvertTo-WTDateSortValue -Value $_.InstalledOn }; Descending = $true }, @{ Expression = { $_.HotFixID }; Ascending = $true })
    $lastHotfixInstalledOn = $null
    if ($recentHotfixes.Count -gt 0) {
        $lastHotfixInstalledOn = $recentHotfixes[0].InstalledOn
    }

    $servicesNotRunning = @(
        $services | Where-Object {
            $_.Exists -and $_.Name -in @('wuauserv', 'bits', 'cryptsvc') -and $_.Status -notin @('Running', 'Unknown')
        }
    )
    $servicesUnknown = @(
        $services | Where-Object {
            -not $_.Exists -or $_.Status -eq 'Unknown' -or $_.StartType -eq 'Unknown'
        }
    )

    $pending = $UpdateInfo.PendingRebootDetails
    $pendingState = $null
    $genericRestartSignal = $null
    $pendingFileRenamePresent = $null
    $pendingFileRenameCount = $null
    $pendingFileRenameSample = @()
    $pendingSource = 'Unknown'
    if ($pending) {
        $pendingState = $pending.UpdatePendingReboot
        $genericRestartSignal = $pending.AnyRestartSignal
        $pendingFileRenamePresent = $pending.PendingFileRenameOperations
        $pendingFileRenameCount = $pending.PendingFileRenameOperationsCount
        $pendingFileRenameSample = @($pending.PendingFileRenameOperationsSample)
        $pendingSource = $pending.PendingRebootSource
    }

    $windowsUpdateEvents = @($UpdateInfo.WindowsUpdateEvents | Where-Object { $_ })
    $windowsUpdateErrorEvents = @()
    $windowsUpdateWarningEvents = @()
    $windowsUpdateSuccessEvents = @()

    foreach ($evt in $windowsUpdateEvents) {
        $combinedText = @($evt.Message, $evt.MessageShort) -join "`n"
        if ($evt.Id -eq 19) {
            $windowsUpdateSuccessEvents += $evt
        }
        if ($evt.Level -eq 2 -or $combinedText -match '(?i)\berror\b|\bfailed\b|fall[oó]|0x[0-9a-f]{3,}') {
            $windowsUpdateErrorEvents += $evt
            continue
        }
        if ($evt.Level -eq 3) {
            $windowsUpdateWarningEvents += $evt
        }
    }

    $windowsUpdateFailureCount = @($windowsUpdateErrorEvents).Count
    $windowsUpdateWarningCount = @($windowsUpdateWarningEvents).Count

    $recentWindowsUpdateEvents = @(
        $windowsUpdateEvents |
            Sort-Object -Property @{
                Expression = {
                    if ($_.Level -eq 2 -or ($_.Message -match '(?i)\berror\b|\bfailed\b|fall[oó]|0x[0-9a-f]{3,}')) { 0 }
                    elseif ($_.Level -eq 3) { 1 }
                    elseif ($_.Id -eq 19) { 2 }
                    else { 3 }
                }
            }, @{ Expression = { $_.TimeCreated }; Descending = $true } |
            Select-Object -First 10
    )

    $servicingWerEvents = @($UpdateInfo.ServicingWerEvents | Where-Object { $_ })
    $storeWerEvents = @($UpdateInfo.StoreWerEvents | Where-Object { $_ })
    $edgeUpdateWerEvents = @($UpdateInfo.EdgeUpdateWerEvents | Where-Object { $_ })

    $recentUpdateIndicators = @()
    foreach ($indicatorType in @('Servicing', 'Store', 'Edge')) {
        $sourceEvents = switch ($indicatorType) {
            'Servicing' { $servicingWerEvents }
            'Store'     { $storeWerEvents }
            'Edge'      { $edgeUpdateWerEvents }
        }

        foreach ($evt in @($sourceEvents | Sort-Object -Property TimeCreated -Descending)) {
            if (-not $evt) {
                continue
            }
            $recentUpdateIndicators += [pscustomobject]@{
                TimeCreated  = $evt.TimeCreated
                Type         = $indicatorType
                WerEventName = $evt.WerEventName
                MessageShort = $evt.MessageShort
            }
            if ($recentUpdateIndicators.Count -ge 15) {
                break
            }
        }
        if ($recentUpdateIndicators.Count -ge 15) {
            break
        }
    }

    $servicingHealthSummary = 'No servicing indicators detected'
    if ($servicingWerEvents.Count -gt 0 -and $UpdateInfo.CbsLog -and $UpdateInfo.CbsLog.IsRecentlyModified) {
        $servicingHealthSummary = 'Potential servicing/component store issues detected'
    }
    elseif ($servicingWerEvents.Count -gt 0) {
        $servicingHealthSummary = 'Servicing WER indicators detected'
    }
    elseif ($UpdateInfo.CbsLog -and $UpdateInfo.CbsLog.IsRecentlyModified) {
        $servicingHealthSummary = 'CBS log recently modified'
    }

    $updateHealthSummary = 'No Windows Update errors detected'
    $wuLogsUnavailable = @($UpdateInfo.LogsUnavailable | Where-Object { $_.LogName -eq 'Microsoft-Windows-WindowsUpdateClient/Operational' })
    if ($wuLogsUnavailable.Count -gt 0 -and $windowsUpdateEvents.Count -eq 0) {
        $updateHealthSummary = 'Windows Update log unavailable'
    }
    elseif ($windowsUpdateFailureCount -ge 1) {
        $updateHealthSummary = 'Windows Update errors detected'
    }
    elseif ($windowsUpdateWarningCount -ge 1) {
        $updateHealthSummary = 'Windows Update warnings detected'
    }

    $storeUpdateSummary = 'No Store update indicators detected'
    if ($storeWerEvents.Count -gt 0) {
        $storeUpdateSummary = 'Store update WER indicators detected'
    }

    $edgeUpdateSummary = 'No Edge/WebView update indicators detected'
    if ($edgeUpdateWerEvents.Count -gt 0) {
        $edgeUpdateSummary = 'Edge/WebView update WER indicators detected'
    }

    return [pscustomobject]@{
        WindowStart              = $UpdateInfo.WindowStart
        WindowEnd                = $UpdateInfo.WindowEnd
        Days                     = $UpdateInfo.Days
        HotfixCount              = @($recentHotfixes).Count
        RecentHotfixes           = @($recentHotfixes)
        LastHotfixInstalledOn    = $lastHotfixInstalledOn
        Services                 = @($services)
        ServicesNotRunning       = @($servicesNotRunning)
        ServicesUnknown          = @($servicesUnknown)
        PendingReboot            = $pendingState
        GenericRestartSignal      = $genericRestartSignal
        PendingFileRenameOperationsPresent = $pendingFileRenamePresent
        PendingFileRenameOperationsCount = $pendingFileRenameCount
        PendingFileRenameOperationsSample = @($pendingFileRenameSample)
        PendingRebootSource       = $pendingSource
        PendingRebootDetails     = $pending
        CbsLog                   = $UpdateInfo.CbsLog
        DismLog                  = $UpdateInfo.DismLog
        WindowsUpdateEvents      = @($windowsUpdateEvents)
        WindowsUpdateErrorEvents = @($windowsUpdateErrorEvents)
        WindowsUpdateWarningEvents = @($windowsUpdateWarningEvents)
        WindowsUpdateSuccessEvents = @($windowsUpdateSuccessEvents)
        WindowsUpdateFailureCount = $windowsUpdateFailureCount
        WindowsUpdateWarningCount = $windowsUpdateWarningCount
        SetupEvents              = @($UpdateInfo.SetupEvents | Where-Object { $_ })
        ServicingWerEvents       = @($servicingWerEvents)
        StoreWerEvents           = @($storeWerEvents)
        EdgeUpdateWerEvents      = @($edgeUpdateWerEvents)
        ServicingHealthSummary   = $servicingHealthSummary
        UpdateHealthSummary      = $updateHealthSummary
        StoreUpdateSummary       = $storeUpdateSummary
        EdgeUpdateSummary        = $edgeUpdateSummary
        RecentWindowsUpdateEvents = @($recentWindowsUpdateEvents)
        RecentUpdateIndicators    = @($recentUpdateIndicators)
        LogsUnavailable          = @($UpdateInfo.LogsUnavailable)
    }
}

function Export-WTUpdateCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$Report,

        [AllowNull()]
        [psobject]$UpdateInfo
    )

    $rawDirectory = Join-Path -Path $Report.Metadata.ReportDirectory -ChildPath 'raw'
    try {
        [void][System.IO.Directory]::CreateDirectory($rawDirectory)
    }
    catch {
        [void](Add-WTExecutionWarning -Report $Report -Scope 'Updates' -Message ('Unable to create raw update output directory: {0}' -f $_.Exception.Message))
        return $null
    }

    $path = ConvertTo-WTAbsolutePath -Path (Join-Path -Path $rawDirectory -ChildPath 'events-windowsupdate.csv')
    $header = 'TimeCreated,LogName,ProviderName,Id,LevelDisplayName,MessageShort,MachineName'

    try {
        $lines = New-Object System.Collections.Generic.List[string]
        $lines.Add($header) | Out-Null
        $rows = @()
        if ($UpdateInfo -and $UpdateInfo.WindowsUpdateEvents) {
            $rows = @($UpdateInfo.WindowsUpdateEvents)
        }
        foreach ($row in $rows) {
            if (-not $row) {
                continue
            }
            $timeText = ConvertTo-WTDisplayValue -Value (ConvertTo-WTDateTimeString -Value $row.TimeCreated) -Fallback ''
            $cells = @(
                $timeText,
                $row.LogName,
                $row.ProviderName,
                $row.Id,
                $row.LevelDisplayName,
                $row.MessageShort,
                $row.MachineName
            )
            $escaped = foreach ($cell in $cells) {
                if ($null -eq $cell) { '' } else { ('"{0}"' -f (('' + $cell) -replace '"', '""')) }
            }
            $lines.Add(($escaped -join ',')) | Out-Null
        }

        [System.IO.File]::WriteAllLines($path, $lines.ToArray(), [System.Text.Encoding]::UTF8)
    }
    catch {
        [void](Add-WTExecutionWarning -Report $Report -Scope 'Updates' -Message ('Unable to write Windows Update CSV export: {0}' -f $_.Exception.Message))
    }

    return [pscustomobject]@{
        RawDirectory = $rawDirectory
        CsvPath      = $path
    }
}

function Invoke-WTUpdateRules {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$Report
    )

    $updates = $Report.Normalized.Updates
    if (-not $updates) {
        return
    }

    $wuErrors = @($updates.WindowsUpdateErrorEvents)
    $wuWarnings = @($updates.WindowsUpdateWarningEvents)
    $wuSuccess = @($updates.WindowsUpdateSuccessEvents)
    $pending = $updates.PendingRebootDetails
    $servicesNotRunning = @($updates.ServicesNotRunning)
    $servicingWerEvents = @($updates.ServicingWerEvents)
    $storeWerEvents = @($updates.StoreWerEvents)
    $edgeWerEvents = @($updates.EdgeUpdateWerEvents)

    if ($wuErrors.Count -ge 1) {
        $latestWuError = @($wuErrors | Sort-Object -Property TimeCreated -Descending | Select-Object -First 1)
        $ids = @($wuErrors | Group-Object -Property Id | Sort-Object -Property @{ Expression = { $_.Count }; Descending = $true }, @{ Expression = { $_.Name }; Ascending = $true } | Select-Object -First 5 | ForEach-Object { '{0}({1})' -f $_.Name, $_.Count })
        $severity = if ($wuErrors.Count -ge 3) { 'High' } else { 'Medium' }
        [void](Add-WTFinding -Report $Report -Id 'WT-UPD-WU-ERRORS' -Category 'Updates' -Severity $severity -Title 'Windows Update errors detected' -Description 'Windows Update logged one or more error events during the analysis window.' -Evidence @("Count=$($wuErrors.Count)", ('LastEvent={0}' -f (ConvertTo-WTDateTimeString -Value $latestWuError[0].TimeCreated)), ('Ids={0}' -f ($ids -join ', ')), ('MessageShort={0}' -f $latestWuError[0].MessageShort)) -Recommendation 'Review Windows Update history, network/proxy access, disk space, pending reboot and servicing health before attempting repairs.' -Source 'Invoke-WTUpdateRules' -Status 'Warning')
    }

    if ($pending -and $pending.UpdatePendingReboot -eq $true) {
        [void](Add-WTFinding -Report $Report -Id 'WT-UPD-PENDING-REBOOT' -Category 'Updates' -Severity 'Low' -Title 'Pending reboot detected' -Description 'Windows indicates that a reboot may be pending due to updates, CBS, or update installer state.' -Evidence @(
            ('CbsRebootPending={0}' -f (ConvertTo-WTYesNoUnknown -Value $pending.CbsRebootPending)),
            ('WuRebootRequired={0}' -f (ConvertTo-WTYesNoUnknown -Value $pending.WuRebootRequired)),
            ('UpdateExeVolatile={0}' -f (ConvertTo-WTYesNoUnknown -Value $pending.UpdateExeVolatile)),
            ('UpdatePendingReboot={0}' -f (ConvertTo-WTYesNoUnknown -Value $pending.UpdatePendingReboot)),
            ('PendingRebootSource={0}' -f (ConvertTo-WTDisplayValue -Value $pending.PendingRebootSource -Fallback 'None'))
        ) -Recommendation 'Schedule a controlled restart before further troubleshooting, especially before running repair actions.' -Source 'Invoke-WTUpdateRules' -Status 'Warning')
    }

    if ($pending -and $pending.UpdatePendingReboot -ne $true -and $pending.GenericPendingFileRename -eq $true) {
        $sampleText = 'None'
        if ($pending.PendingFileRenameOperationsSample -and @($pending.PendingFileRenameOperationsSample).Count -gt 0) {
            $sampleText = @($pending.PendingFileRenameOperationsSample | Select-Object -First 5) -join '; '
        }
        [void](Add-WTFinding -Report $Report -Id 'WT-UPD-PENDING-FILE-RENAME' -Category 'Updates' -Severity 'Info' -Title 'Pending file rename operations detected' -Description 'Windows has pending file rename operations registered. This may be caused by installers, drivers, antivirus, agents or third-party software and does not necessarily mean Windows Update requires a reboot.' -Evidence @(
            ('PendingFileRenameOperations={0}' -f (ConvertTo-WTYesNoUnknown -Value $pending.PendingFileRenameOperations)),
            ('PendingFileRenameOperationsCount={0}' -f (Get-WTDisplayValue -Value $pending.PendingFileRenameOperationsCount -Fallback 'Unknown')),
            ('Sample={0}' -f $sampleText),
            'UpdatePendingReboot=False',
            ('CbsRebootPending={0}' -f (ConvertTo-WTYesNoUnknown -Value $pending.CbsRebootPending)),
            ('WuRebootRequired={0}' -f (ConvertTo-WTYesNoUnknown -Value $pending.WuRebootRequired))
        ) -Recommendation 'Correlate with recent software installations, driver updates or security agents. Do not treat this alone as Windows Update requiring a reboot.' -Source 'Invoke-WTUpdateRules' -Status 'Info')
    }

    if ($servicesNotRunning.Count -gt 0) {
        $serviceEvidence = @()
        foreach ($svc in $servicesNotRunning) {
            $serviceEvidence += ('{0}={1} ({2})' -f $svc.Name, $svc.Status, $svc.StartType)
        }
        [void](Add-WTFinding -Report $Report -Id 'WT-UPD-SERVICES-NOT-RUNNING' -Category 'Updates' -Severity 'Medium' -Title 'Windows Update related services are not running' -Description 'One or more Windows Update related services are not running.' -Evidence $serviceEvidence -Recommendation 'Review service startup type and recent service control events. Do not restart services automatically from this tool.' -Source 'Invoke-WTUpdateRules' -Status 'Warning')
    }

    $servicingWerNames = @($servicingWerEvents | Select-Object -ExpandProperty WerEventName | Where-Object { $_ } | Group-Object | Sort-Object -Property @{ Expression = { $_.Count }; Descending = $true }, @{ Expression = { $_.Name }; Ascending = $true } | Select-Object -First 5 | ForEach-Object { '{0}({1})' -f $_.Name, $_.Count })
    $storeWerNames = @($storeWerEvents | Select-Object -ExpandProperty WerEventName | Where-Object { $_ } | Group-Object | Sort-Object -Property @{ Expression = { $_.Count }; Descending = $true }, @{ Expression = { $_.Name }; Ascending = $true } | Select-Object -First 5 | ForEach-Object { '{0}({1})' -f $_.Name, $_.Count })
    $edgeWerNames = @($edgeWerEvents | Select-Object -ExpandProperty WerEventName | Where-Object { $_ } | Group-Object | Sort-Object -Property @{ Expression = { $_.Count }; Descending = $true }, @{ Expression = { $_.Name }; Ascending = $true } | Select-Object -First 5 | ForEach-Object { '{0}({1})' -f $_.Name, $_.Count })

    if ($updates.CbsLog -and $updates.CbsLog.IsRecentlyModified -and $servicingWerEvents.Count -gt 0) {
        $latestServicing = @($servicingWerEvents | Sort-Object -Property TimeCreated -Descending | Select-Object -First 1)
        [void](Add-WTFinding -Report $Report -Id 'WT-UPD-SERVICING-INDICATORS' -Category 'Updates' -Severity 'Medium' -Title 'Potential Windows servicing or component store issue detected' -Description 'Windows logged servicing-related WER events and CBS.log was recently modified.' -Evidence @(
            ('CBSLastWriteTime={0}' -f (ConvertTo-WTDateTimeString -Value $updates.CbsLog.LastWriteTime)),
            ('ServicingWerCount={0}' -f $servicingWerEvents.Count),
            ('WerEventNames={0}' -f ($servicingWerNames -join ', ')),
            ('ExampleMessageShort={0}' -f $latestServicing[0].MessageShort)
        ) -Recommendation 'Review CBS.log and Windows Update history. If symptoms persist, consider manual DISM/SFC repair in a controlled remediation step, not from this read-only triage.' -Source 'Invoke-WTUpdateRules' -Status 'Warning')
    }

    foreach ($evt in @($servicingWerEvents)) {
        if ($evt.WerEventName -eq 'WindowsWcpOtherFailure3') {
            $latest = @($servicingWerEvents | Where-Object { $_.WerEventName -eq 'WindowsWcpOtherFailure3' } | Sort-Object -Property TimeCreated -Descending | Select-Object -First 1)
            [void](Add-WTFinding -Report $Report -Id 'WT-UPD-WCP-FAILURE' -Category 'Updates' -Severity 'Medium' -Title 'Windows component servicing WER failure detected' -Description 'Windows Error Reporting logged WindowsWcpOtherFailure3 events, which may indicate component servicing or CBS issues.' -Evidence @(
                ('Count={0}' -f @($servicingWerEvents | Where-Object { $_.WerEventName -eq 'WindowsWcpOtherFailure3' }).Count),
                ('LastEvent={0}' -f (ConvertTo-WTDateTimeString -Value $latest[0].TimeCreated)),
                ('MessageShort={0}' -f $latest[0].MessageShort),
                ('WerEventNames={0}' -f ($servicingWerNames -join ', '))
            ) -Recommendation 'Review CBS.log, DISM log, update history and component store health. Treat this as a servicing signal, not an application crash.' -Source 'Invoke-WTUpdateRules' -Status 'Warning')
            break
        }
    }

    if ($storeWerEvents.Count -ge 5) {
        $latestStore = @($storeWerEvents | Sort-Object -Property TimeCreated -Descending | Select-Object -First 1)
        [void](Add-WTFinding -Report $Report -Id 'WT-UPD-STORE-WER' -Category 'Updates' -Severity 'Low' -Title 'Microsoft Store update related WER events detected' -Description 'Windows logged multiple StoreAgent or AppX deployment WER events.' -Evidence @(
            ('Count={0}' -f $storeWerEvents.Count),
            ('WerEventNames={0}' -f ($storeWerNames -join ', ')),
            ('LastEvent={0}' -f (ConvertTo-WTDateTimeString -Value $latestStore[0].TimeCreated))
        ) -Recommendation 'Review only if users report Store app, AppX, winget, WebView, or packaged app update problems.' -Source 'Invoke-WTUpdateRules' -Status 'Warning')
    }

    if ($edgeWerEvents.Count -ge 3) {
        $latestEdge = @($edgeWerEvents | Sort-Object -Property TimeCreated -Descending | Select-Object -First 1)
        [void](Add-WTFinding -Report $Report -Id 'WT-UPD-EDGEUPDATE-WER' -Category 'Updates' -Severity 'Low' -Title 'Edge/WebView update related WER events detected' -Description 'Windows logged Edge Update or WebView update related WER events.' -Evidence @(
            ('Count={0}' -f $edgeWerEvents.Count),
            ('WerEventNames={0}' -f ($edgeWerNames -join ', ')),
            ('LastEvent={0}' -f (ConvertTo-WTDateTimeString -Value $latestEdge[0].TimeCreated))
        ) -Recommendation 'Review only if users report browser, WebView, Teams, Outlook, embedded web component or app update symptoms.' -Source 'Invoke-WTUpdateRules' -Status 'Warning')
    }

    $wuLogsUnavailable = @($updates.LogsUnavailable | Where-Object { $_.LogName -eq 'Microsoft-Windows-WindowsUpdateClient/Operational' })
    $hasNoIssues = ($wuErrors.Count -eq 0 -and $wuWarnings.Count -eq 0 -and $pending -and $pending.UpdatePendingReboot -eq $false -and $pending.GenericPendingFileRename -eq $false -and $servicesNotRunning.Count -eq 0 -and $servicingWerEvents.Count -eq 0 -and $storeWerEvents.Count -eq 0 -and $edgeWerEvents.Count -eq 0 -and $wuLogsUnavailable.Count -eq 0)
    if ($hasNoIssues) {
        [void](Add-WTFinding -Report $Report -Id 'WT-UPD-NO-ISSUES' -Category 'Updates' -Severity 'Info' -Title 'No Windows Update or servicing issues detected' -Description 'No relevant Windows Update, CBS, servicing or Store update indicators were detected in the selected window.' -Evidence @("Days=$($updates.Days)", "HotfixCount=$($updates.HotfixCount)", 'WindowsUpdateFailureCount=0') -Recommendation 'No action required based on this scope.' -Source 'Invoke-WTUpdateRules' -Status 'Pass')
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

    $powerBoot = $Report.Normalized.PowerBoot
    $kernelUptimeDays = $null
    if ($powerBoot -and $powerBoot.KernelUptimeDays -ne $null) {
        $kernelUptimeDays = $powerBoot.KernelUptimeDays
    }
    elseif ($system.KernelUptimeDays -ne $null) {
        $kernelUptimeDays = $system.KernelUptimeDays
    }
    elseif ($system.UptimeDays -ne $null) {
        $kernelUptimeDays = $system.UptimeDays
    }

    $fastStartupEnabled = $null
    if ($powerBoot) {
        $fastStartupEnabled = $powerBoot.FastStartupEnabled
    }

    if ($kernelUptimeDays -ne $null) {
        if ($kernelUptimeDays -ge 30) {
            $uptimeSeverity = 'Medium'
            $uptimeStatus = 'Warning'
            if ($fastStartupEnabled -eq $true) {
                $uptimeSeverity = 'Low'
                $uptimeStatus = 'Info'
            }
            Add-WTFinding -Report $Report -Id 'WT-SYS-UPTIME-VERY-LONG' -Category 'System' -Severity $uptimeSeverity -Title 'Kernel uptime is very long' -Description 'The kernel has been active for 30 days or more. This may reflect a full boot delay or preserved kernel state from Fast Startup.' -Evidence @("KernelUptimeDays=$kernelUptimeDays", ('FastStartupEnabled={0}' -f (ConvertTo-WTEnabledDisabledUnknown -Value $fastStartupEnabled))) -Recommendation 'Interpret this as kernel uptime, not necessarily physical power-on time. Correlate with restart and shutdown history.' -Source 'Invoke-WTSystemRules' -Status $uptimeStatus
        }
        elseif ($kernelUptimeDays -ge 14) {
            $uptimeSeverity = 'Low'
            $uptimeStatus = 'Warning'
            if ($fastStartupEnabled -eq $true) {
                $uptimeSeverity = 'Info'
                $uptimeStatus = 'Info'
            }
            Add-WTFinding -Report $Report -Id 'WT-SYS-UPTIME-LONG' -Category 'System' -Severity $uptimeSeverity -Title 'Kernel uptime is long' -Description 'The kernel has been active for 14 days or more. Fast Startup can preserve kernel state across normal user shutdowns.' -Evidence @("KernelUptimeDays=$kernelUptimeDays", ('FastStartupEnabled={0}' -f (ConvertTo-WTEnabledDisabledUnknown -Value $fastStartupEnabled))) -Recommendation 'Use kernel uptime together with shutdown and restart events for interpretation.' -Source 'Invoke-WTSystemRules' -Status $uptimeStatus
        }
    }
}

function Invoke-WTPowerBootRules {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$Report
    )

    $powerBoot = $Report.Normalized.PowerBoot
    if (-not $powerBoot) {
        return
    }

    if ($powerBoot.FastStartupEnabled -eq $true) {
        Add-WTFinding -Report $Report -Id 'WT-PWR-FAST-STARTUP-INFO' -Category 'PowerBoot' -Severity 'Info' -Title 'Fast Startup is enabled' -Description 'Windows Fast Startup is enabled. Kernel uptime may not reset after normal user shutdowns.' -Evidence @("FastStartupEnabled=true", ('KernelUptimeDays={0}' -f (ConvertTo-WTDisplayValue -Value $powerBoot.KernelUptimeDays)), ('RecentShutdownEventsCount={0}' -f $powerBoot.RecentShutdownEventsCount)) -Recommendation 'Do not interpret kernel uptime as physical power-on time. Use restart events and power state history for context.' -Source 'Invoke-WTPowerBootRules' -Status 'Info'
    }

    if ($powerBoot.RecentShutdownEventsCount -gt 0 -or $powerBoot.RecentPlannedShutdownEventsCount -gt 0) {
        Add-WTFinding -Report $Report -Id 'WT-PWR-RECENT-SHUTDOWNS' -Category 'PowerBoot' -Severity 'Info' -Title 'Recent shutdown or restart events detected' -Description 'Windows logged normal shutdown or restart events during the analysis window.' -Evidence @("Count=$($powerBoot.RecentShutdownEventsCount)", ('LastShutdownEventTime={0}' -f (ConvertTo-WTDateTimeString -Value $powerBoot.LastShutdownEventTime)), ('LastPlannedShutdownEventTime={0}' -f (ConvertTo-WTDateTimeString -Value $powerBoot.LastPlannedShutdownEventTime))) -Recommendation 'Correlate with user reports and remote management history.' -Source 'Invoke-WTPowerBootRules' -Status 'Info'
    }

    if ($powerBoot.RecentUnexpectedShutdownEventsCount -gt 0) {
        Add-WTFinding -Report $Report -Id 'WT-PWR-UNEXPECTED-SHUTDOWN' -Category 'PowerBoot' -Severity 'High' -Title 'Unexpected shutdown or power loss detected' -Description 'Windows logged unexpected shutdown or Kernel-Power events.' -Evidence @("Count=$($powerBoot.RecentUnexpectedShutdownEventsCount)", ('LastUnexpectedShutdownEventTime={0}' -f (ConvertTo-WTDateTimeString -Value $powerBoot.LastUnexpectedShutdownEventTime))) -Recommendation 'Investigate power loss, forced shutdown, freeze, BSOD, storage or hardware instability.' -Source 'Invoke-WTPowerBootRules' -Status 'Warning'
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

function Get-WTEventCollectionLimit {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [string]$Mode
    )

    switch ($Mode) {
        'Quick' { return 300 }
        'Full' { return 3000 }
        default { return 1000 }
    }
}

function ConvertTo-WTEventMessageShort {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [string]$Message
    )

    if ([string]::IsNullOrWhiteSpace($Message)) {
        return 'No message available'
    }

    $shortMessage = ($Message -replace '[\r\n]+', ' ').Trim()
    if ($shortMessage.Length -gt 500) {
        return $shortMessage.Substring(0, 500)
    }

    return $shortMessage
}

function Get-WTArrayCountSafe {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return 0
    }

    try {
        return @($Value).Count
    }
    catch {
        return 0
    }
}

function Get-WTEventKey {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [psobject]$Event
    )

    if (-not $Event) {
        return $null
    }

    $timeText = $null
    try {
        if ($Event.TimeCreated) {
            $timeText = $Event.TimeCreated.ToString('o')
        }
    }
    catch {
        $timeText = $null
    }

    if (-not $timeText) {
        $timeText = 'Unknown'
    }

    return ('{0}|{1}|{2}|{3}|{4}' -f $timeText, $Event.LogName, $Event.ProviderName, $Event.Id, $Event.MessageShort)
}

function Normalize-WTProcessName {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object]$Value,

        [switch]$RequireExecutable
    )

    if ($null -eq $Value) {
        return $null
    }

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    $text = ($text -split '[\r\n]+' | Select-Object -First 1)
    if ($text -match ',') {
        $commaParts = $text -split ',', 2
        if ($commaParts.Count -gt 0 -and -not [string]::IsNullOrWhiteSpace($commaParts[0])) {
            $text = $commaParts[0]
        }
    }
    $text = $text.Trim().Trim('"').TrimEnd(',').Trim()
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    if ($text -match '^[A-Za-z]:\\|^\\\\|[\\/]' ) {
        try {
            $text = [System.IO.Path]::GetFileName($text)
        }
        catch {
            $text = $null
        }
        if ([string]::IsNullOrWhiteSpace($text)) {
            return $null
        }
        $text = $text.Trim().Trim('"').TrimEnd(',').Trim()
    }

    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    if ($text -match '^[A-Fa-f0-9]{8}$') {
        return $null
    }

    if ($RequireExecutable.IsPresent -and $text -notmatch '\.exe$') {
        return $null
    }

    if (-not $RequireExecutable.IsPresent -and $text -notmatch '\.[A-Za-z0-9]{2,4}$' -and $text -notmatch '\.exe$') {
        return $null
    }

    return $text
}

function Normalize-WTProcessPath {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return $null
    }

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    $text = ($text -split '[\r\n]+' | Select-Object -First 1)
    if ($text -match ',') {
        $commaParts = $text -split ',', 2
        if ($commaParts.Count -gt 0 -and -not [string]::IsNullOrWhiteSpace($commaParts[0])) {
            $text = $commaParts[0]
        }
    }
    $cutMarkers = @(
        'Ruta de módulo con errores:',
        'Id\. de informe:',
        'Nombre completo del paquete con errores:',
        'Faulting module path:',
        'Report Id:',
        'Faulting package full name:',
        'Faulting application path:',
        'Ruta de aplicación con errores:'
    )
    foreach ($marker in $cutMarkers) {
        if ($text -match $marker) {
            $text = [regex]::Split($text, $marker)[0]
        }
    }

    $text = $text.Trim().Trim('"').TrimEnd(',').Trim()
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    if ($text -match '^[A-Za-z]:\\|^\\\\') {
        return $text
    }

    return $null
}

function Get-WTWerEventName {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [string]$Message
    )

    if ([string]::IsNullOrWhiteSpace($Message)) {
        return $null
    }

    foreach ($pattern in @('Nombre de evento:\s*([^\r\n]+)', 'Event Name:\s*([^\r\n]+)')) {
        if ($Message -match $pattern) {
            return $Matches[1].Trim()
        }
    }

    return $null
}

function Get-WTApplicationErrorProcessInfo {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [string]$Message,

        [AllowNull()]
        [string]$MessageShort
    )

    $texts = @()
    if (-not [string]::IsNullOrWhiteSpace($Message)) {
        $texts += $Message
    }
    if (-not [string]::IsNullOrWhiteSpace($MessageShort) -and $MessageShort -ne $Message) {
        $texts += $MessageShort
    }

    if ($texts.Count -eq 0) {
        return [pscustomobject]@{
            ProcessName   = $null
            ProcessPath   = $null
            SourcePattern = $null
        }
    }

    $processName = $null
    $processPath = $null
    $sourcePattern = $null

    foreach ($text in $texts) {
        $lines = @($text -split "`r?`n")
        foreach ($line in $lines) {
            if (-not $processName -and $line -match '^\s*Nombre de aplicación con errores:') {
                $value = $line.Substring($line.IndexOf(':') + 1).Trim()
                if ($value -match ',') {
                    $value = ($value -split ',', 2)[0]
                }
                $candidateName = Normalize-WTProcessName -Value $value -RequireExecutable
                if ($candidateName) {
                    $processName = $candidateName
                    $sourcePattern = 'SpanishFaultingApplicationName'
                }
            }

            if (-not $processPath -and $line -match '^\s*Ruta de aplicación con errores:') {
                $value = $line.Substring($line.IndexOf(':') + 1).Trim()
                $candidatePath = Normalize-WTProcessPath -Value $value
                if ($candidatePath) {
                    $processPath = $candidatePath
                    if (-not $sourcePattern) {
                        $sourcePattern = 'SpanishFaultingApplicationPath'
                    }
                }
            }

            if (-not $processPath -and $line -match '^\s*Ruta de módulo con errores:') {
                $value = $line.Substring($line.IndexOf(':') + 1).Trim()
                $candidatePath = Normalize-WTProcessPath -Value $value
                if ($candidatePath) {
                    $processPath = $candidatePath
                    if (-not $sourcePattern) {
                        $sourcePattern = 'SpanishFaultingModulePath'
                    }
                }
            }

            if (-not $processName -and $line -match '^\s*Faulting application name:') {
                $value = $line.Substring($line.IndexOf(':') + 1).Trim()
                if ($value -match ',') {
                    $value = ($value -split ',', 2)[0]
                }
                $candidateName = Normalize-WTProcessName -Value $value -RequireExecutable
                if ($candidateName) {
                    $processName = $candidateName
                    if (-not $sourcePattern) {
                        $sourcePattern = 'EnglishFaultingApplicationName'
                    }
                }
            }

            if (-not $processPath -and $line -match '^\s*Faulting application path:') {
                $value = $line.Substring($line.IndexOf(':') + 1).Trim()
                $candidatePath = Normalize-WTProcessPath -Value $value
                if ($candidatePath) {
                    $processPath = $candidatePath
                    if (-not $sourcePattern) {
                        $sourcePattern = 'EnglishFaultingApplicationPath'
                    }
                }
            }

            if (-not $processPath -and $line -match '^\s*Faulting module path:') {
                $value = $line.Substring($line.IndexOf(':') + 1).Trim()
                $candidatePath = Normalize-WTProcessPath -Value $value
                if ($candidatePath) {
                    $processPath = $candidatePath
                    if (-not $sourcePattern) {
                        $sourcePattern = 'EnglishFaultingModulePath'
                    }
                }
            }

            if ($processName -and $processPath) {
                break
            }
        }

        if ($processName -and $processPath) {
            break
        }
    }

    if (-not $processName -and $processPath) {
        try {
            $processName = Normalize-WTProcessName -Value ([System.IO.Path]::GetFileName($processPath)) -RequireExecutable
        }
        catch {
            $processName = $null
        }
    }

    return [pscustomobject]@{
        ProcessName   = $processName
        ProcessPath   = $processPath
        SourcePattern = $sourcePattern
    }
}

function Get-WTEventProcessName {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [string]$Message,

        [AllowNull()]
        [string]$MessageShort,

        [AllowNull()]
        [string]$ProviderName,

        [AllowNull()]
        [object]$Id,

        [AllowNull()]
        [string]$WerEventName
    )

    $processName = $null
    $processPath = $null
    $sourcePattern = $null

    if ($ProviderName -eq 'Application Error' -and $Id -eq 1000) {
        $appInfo = Get-WTApplicationErrorProcessInfo -Message $Message -MessageShort $MessageShort
        if ($appInfo) {
            $processName = $appInfo.ProcessName
            $processPath = $appInfo.ProcessPath
            $sourcePattern = $appInfo.SourcePattern
        }
    }

    if (-not $processName -and -not $processPath) {
        $texts = @()
        if (-not [string]::IsNullOrWhiteSpace($Message)) {
            $texts += $Message
        }
        if (-not [string]::IsNullOrWhiteSpace($MessageShort) -and $MessageShort -ne $Message) {
            $texts += $MessageShort
        }

        if ($texts.Count -gt 0) {
            $crashSignal = $false
            if ($WerEventName -and $WerEventName -match '^(APPCRASH|BEX|CLR20r3|MoAppCrash)$') {
                $crashSignal = $true
            }
            elseif (($ProviderName -eq 'Windows Error Reporting' -and $Id -eq 1001) -and (
                (-not [string]::IsNullOrWhiteSpace($Message) -and $Message -match '(?im)^\s*(Nombre de evento|Event Name):\s*(APPCRASH|BEX|CLR20r3|MoAppCrash)\b') -or
                (-not [string]::IsNullOrWhiteSpace($MessageShort) -and $MessageShort -match '(?im)^\s*(Nombre de evento|Event Name):\s*(APPCRASH|BEX|CLR20r3|MoAppCrash)\b')
            )) {
                $crashSignal = $true
            }

            if ($crashSignal) {
                foreach ($text in $texts) {
                    if ($text -match '(?im)^\s*P1:\s*(?<candidate>[^\r\n]+)') {
                        $candidate = Normalize-WTProcessName -Value $Matches.candidate -RequireExecutable
                        if ($candidate) {
                            $processName = $candidate
                            $sourcePattern = 'WerP1CrashExecutable'
                            break
                        }
                    }
                }
            }
        }
    }

    return [pscustomobject]@{
        ProcessName   = $processName
        ProcessPath   = $processPath
        SourcePattern = $sourcePattern
    }
}

function Invoke-WTEventParserSelfTest {
    [CmdletBinding()]
    param()

    $tests = @(
        [pscustomobject]@{
            Name = 'Application Error 1000 Spanish'
            Message = @'
Nombre de aplicación con errores: mxdhcp.exe, versión: 1.1.0.3, marca de tiempo: 0x67b30e7f
Nombre del módulo con errores: mxdhcp.exe, versión: 1.1.0.3, marca de tiempo: 0x67b30e7f
Código de excepción: 0xc0000005
Desplazamiento con errores: 0x0000000000003110
Id. de proceso con errores: 0x162C
Tiempo de inicio de aplicación con errores: 0x1DCD41AA9CEC325
Ruta de aplicación con errores: C:\Program Files\NComputing\vSpace Server Software\mxdhcp.exe
Ruta de módulo con errores: C:\Program Files\NComputing\vSpace Server Software\mxdhcp.exe
Id. de informe: 7eef0426-7d19-4b99-b2dc-6887c9f23c11
Nombre completo del paquete con errores:
Id. de aplicación relacionado con el paquete con errores:
'@
            MessageShort = $null
            ProviderName = 'Application Error'
            Id = 1000
            WerEventName = $null
            ExpectedProcessName = 'mxdhcp.exe'
            ExpectedProcessPath = 'C:\Program Files\NComputing\vSpace Server Software\mxdhcp.exe'
            ExpectedSourcePattern = 'SpanishFaultingApplicationName'
        },
        [pscustomobject]@{
            Name = 'WER APPCRASH Spanish'
            Message = @'
Nombre de evento: APPCRASH
Firma del problema:
P1: mxdhcp.exe
P2: 1.1.0.3
'@
            MessageShort = $null
            ProviderName = 'Windows Error Reporting'
            Id = 1001
            WerEventName = 'APPCRASH'
            ExpectedProcessName = 'mxdhcp.exe'
            ExpectedProcessPath = $null
            ExpectedSourcePattern = 'WerP1CrashExecutable'
        },
        [pscustomobject]@{
            Name = 'WER NonCritical GDIObjectLeak'
            Message = @'
Nombre de evento: GDIObjectLeak
Firma del problema:
P1: MeshAgent.exe
P2: 0.0.0.0
'@
            MessageShort = $null
            ProviderName = 'Windows Error Reporting'
            Id = 1001
            WerEventName = 'GDIObjectLeak'
            ExpectedProcessName = $null
            ExpectedProcessPath = $null
            ExpectedSourcePattern = $null
        },
        [pscustomobject]@{
            Name = 'WER NonCritical AppxDeploymentFailureBlue'
            Message = @'
Nombre de evento: AppxDeploymentFailureBlue
Firma del problema:
P1: 80073CFB
P2: Windows.Desktop
'@
            MessageShort = $null
            ProviderName = 'Windows Error Reporting'
            Id = 1001
            WerEventName = 'AppxDeploymentFailureBlue'
            ExpectedProcessName = $null
            ExpectedProcessPath = $null
            ExpectedSourcePattern = $null
        }
    )

    $results = @()
    $failed = $false

    foreach ($test in $tests) {
        $actual = Get-WTEventProcessName -Message $test.Message -MessageShort $test.MessageShort -ProviderName $test.ProviderName -Id $test.Id -WerEventName $test.WerEventName
        $pass = (
            $actual.ProcessName -eq $test.ExpectedProcessName -and
            $actual.ProcessPath -eq $test.ExpectedProcessPath -and
            $actual.SourcePattern -eq $test.ExpectedSourcePattern
        )

        if ($pass) {
            Write-Host ('PASS: {0}' -f $test.Name) -ForegroundColor Green
        }
        else {
            $failed = $true
            Write-Host ('FAIL: {0}' -f $test.Name) -ForegroundColor Red
            Write-Host ('  Expected ProcessName={0}' -f (ConvertTo-WTDisplayValue -Value $test.ExpectedProcessName)) -ForegroundColor Red
            Write-Host ('  Actual   ProcessName={0}' -f (ConvertTo-WTDisplayValue -Value $actual.ProcessName)) -ForegroundColor Red
            Write-Host ('  Expected ProcessPath={0}' -f (ConvertTo-WTDisplayValue -Value $test.ExpectedProcessPath)) -ForegroundColor Red
            Write-Host ('  Actual   ProcessPath={0}' -f (ConvertTo-WTDisplayValue -Value $actual.ProcessPath)) -ForegroundColor Red
            Write-Host ('  Expected SourcePattern={0}' -f (ConvertTo-WTDisplayValue -Value $test.ExpectedSourcePattern)) -ForegroundColor Red
            Write-Host ('  Actual   SourcePattern={0}' -f (ConvertTo-WTDisplayValue -Value $actual.SourcePattern)) -ForegroundColor Red
        }

        $results += [pscustomobject]@{
            Name = $test.Name
            Passed = $pass
            ExpectedProcessName = $test.ExpectedProcessName
            ActualProcessName = $actual.ProcessName
            ExpectedProcessPath = $test.ExpectedProcessPath
            ActualProcessPath = $actual.ProcessPath
            ExpectedSourcePattern = $test.ExpectedSourcePattern
            ActualSourcePattern = $actual.SourcePattern
        }
    }

    if ($failed) {
        Write-Host 'FAIL' -ForegroundColor Red
        return [pscustomobject]@{
            Passed = $false
            Results = $results
        }
    }

    Write-Host 'PASS' -ForegroundColor Green
    Write-Host 'Self-test summary:' -ForegroundColor Green
    foreach ($result in $results) {
        Write-Host ('- {0}: PASS' -f $result.Name) -ForegroundColor Green
    }

    return [pscustomobject]@{
        Passed = $true
        Results = $results
    }
}

function ConvertTo-WTEventRecord {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$EventRecord
    )

    $message = $null
    try {
        $message = $EventRecord.Message
    }
    catch {
        $message = $null
    }

    $timeCreated = $null
    try {
        $timeCreated = $EventRecord.TimeCreated
    }
    catch {
        $timeCreated = $null
    }

    $messageShort = ConvertTo-WTEventMessageShort -Message $message
    $werEventName = Get-WTWerEventName -Message $message
    $processInfo = Get-WTEventProcessName -Message $message -MessageShort $messageShort -ProviderName $EventRecord.ProviderName -Id $EventRecord.Id -WerEventName $werEventName

    return [pscustomobject]@{
        TimeCreated     = $timeCreated
        LogName         = ConvertTo-WTDisplayValue -Value $EventRecord.LogName
        ProviderName    = ConvertTo-WTDisplayValue -Value $EventRecord.ProviderName
        Id              = if ($null -ne $EventRecord.Id) { [int]$EventRecord.Id } else { $null }
        LevelDisplayName = ConvertTo-WTDisplayValue -Value $EventRecord.LevelDisplayName
        Level           = if ($null -ne $EventRecord.Level) { [int]$EventRecord.Level } else { $null }
        Message         = if ([string]::IsNullOrWhiteSpace($message)) { 'No message available' } else { $message }
        MessageShort    = $messageShort
        MachineName     = ConvertTo-WTDisplayValue -Value $EventRecord.MachineName
        AffectedProcess = $processInfo.ProcessName
        AffectedPath    = $processInfo.ProcessPath
        ProcessSourcePattern = $processInfo.SourcePattern
        WerEventName    = $werEventName
    }
}

function Get-WTEventInfo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$Report,

        [ValidateRange(1, 90)]
        [int]$Days = 7,

        [AllowNull()]
        [string]$Mode = 'Standard'
    )

    $windowEnd = Get-Date
    $windowStart = $windowEnd.AddDays(-1 * [math]::Abs($Days))
    $limit = Get-WTEventCollectionLimit -Mode $Mode

    $systemIds = @(41, 6008, 1074, 6005, 6006, 6009, 12, 13, 1001, 7000, 7001, 7009, 7011, 7022, 7023, 7024, 7031, 7034, 7045)
    $applicationIds = @(1000, 1001)
    $systemEvents = @()
    $applicationEvents = @()
    $logsUnavailable = @()
    $eventCountsByLog = @()

    $logDefinitions = @(
        [pscustomobject]@{ LogName = 'System'; Ids = $systemIds; Target = 'SystemEvents' },
        [pscustomobject]@{ LogName = 'Application'; Ids = $applicationIds; Target = 'ApplicationEvents' }
    )

    foreach ($logDefinition in $logDefinitions) {
        $rawEvents = @()
        try {
            $query = @{
                LogName   = $logDefinition.LogName
                StartTime = $windowStart
                Id        = $logDefinition.Ids
            }
            $rawEvents = @(Get-WinEvent -FilterHashtable $query -MaxEvents ($limit + 1) -ErrorAction Stop)
        }
        catch {
            $reason = $_.Exception.Message
            if ($reason -notmatch '(?i)(no matching events|no se encontraron eventos|no events were found|no events match|no hay eventos)') {
                $logsUnavailable += [pscustomobject]@{
                    LogName = $logDefinition.LogName
                    Reason  = $reason
                }
                [void](Add-WTExecutionWarning -Report $Report -Scope 'Events' -Message ('Event collection failed for {0}: {1}' -f $logDefinition.LogName, $reason))
            }
            continue
        }

        if ($rawEvents.Count -gt $limit) {
            [void](Add-WTExecutionWarning -Report $Report -Scope 'Events' -Message ('Event collection limit reached for {0}. Results may be incomplete.' -f $logDefinition.LogName))
            $rawEvents = @($rawEvents | Select-Object -First $limit)
        }

        $normalizedEvents = @()
        foreach ($event in $rawEvents) {
            $normalizedEvents += ConvertTo-WTEventRecord -EventRecord $event
        }

        if ($logDefinition.Target -eq 'SystemEvents') {
            $systemEvents = $normalizedEvents
        }
        else {
            $applicationEvents = $normalizedEvents
        }

        $eventCountsByLog += [pscustomobject]@{
            LogName = $logDefinition.LogName
            Count   = @($normalizedEvents).Count
        }
    }

    $allEvents = @($systemEvents + $applicationEvents | Sort-Object -Property TimeCreated -Descending)

    $eventCountsById = @(
        $allEvents |
            Group-Object -Property Id |
            Sort-Object -Property @{ Expression = { $_.Count }; Descending = $true }, @{ Expression = { $_.Name } } |
            ForEach-Object {
                [pscustomobject]@{
                    Id    = $_.Name
                    Count = $_.Count
                }
            }
    )

    $eventCountsByProvider = @(
        $allEvents |
            Group-Object -Property ProviderName |
            Sort-Object -Property @{ Expression = { $_.Count }; Descending = $true }, @{ Expression = { $_.Name } } |
            ForEach-Object {
                [pscustomobject]@{
                    ProviderName = $_.Name
                    Count        = $_.Count
                }
            }
    )

    return [pscustomobject]@{
        WindowStart          = $windowStart
        WindowEnd            = $windowEnd
        Days                 = $Days
        LogsQueried          = @('System', 'Application')
        LogsUnavailable      = @($logsUnavailable)
        SystemEvents         = @($systemEvents)
        ApplicationEvents    = @($applicationEvents)
        AllEvents            = @($allEvents)
        EventCountsByLog     = @($eventCountsByLog)
        EventCountsById      = @($eventCountsById)
        EventCountsByProvider = @($eventCountsByProvider)
    }
}

function ConvertTo-WTNormalizedEventInfo {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [psobject]$EventInfo
    )

    if (-not $EventInfo) {
        return [pscustomobject]@{
            WindowStart              = $null
            WindowEnd                = $null
            Days                     = $null
            TotalEvents              = 0
            SystemEventCount         = 0
            ApplicationEventCount    = 0
            UnexpectedShutdownEvents = @()
            KernelPowerEvents        = @()
            BugCheckEvents           = @()
            NormalShutdownEvents     = @()
            ServiceFailureEvents     = @()
            ServiceInstallEvents     = @()
            ApplicationCrashEvents   = @()
            WindowsErrorReportingEvents = @()
            NonCriticalWerEvents     = @()
            ApplicationCrashSummaryByProcess = @()
            EventsById               = @()
            EventsByProvider         = @()
            EventCountsByLog         = @()
            EventCountsById          = @()
            EventCountsByProvider    = @()
            RecentCriticalEvents     = @()
            RecentWarningEvents      = @()
            RecentImportantEvents    = @()
            LogsQueried              = @()
            LogsUnavailable          = @()
            AllEvents                = @()
            SystemEvents             = @()
            ApplicationEvents        = @()
        }
    }

    $systemEvents = @($EventInfo.SystemEvents | Where-Object { $_ })
    $applicationEvents = @($EventInfo.ApplicationEvents | Where-Object { $_ })
    $allEvents = @($EventInfo.AllEvents | Where-Object { $_ } | Sort-Object -Property TimeCreated -Descending)

    $unexpectedShutdownEvents = @($systemEvents | Where-Object { $_.Id -in @(41, 6008) })
    $kernelPowerEvents = @($systemEvents | Where-Object { $_.Id -eq 41 })
    $bugCheckEvents = @($systemEvents | Where-Object {
        $_.Id -eq 1001 -and (
            $_.ProviderName -match 'BugCheck|System Error Reporting|SystemErrorReporting' -or
            $_.MessageShort -match 'bugcheck'
        )
    })
    $normalShutdownEvents = @($systemEvents | Where-Object { $_.Id -in @(1074, 6006, 13) })
    $serviceFailureIds = @(7000, 7001, 7009, 7011, 7022, 7023, 7024, 7031, 7034)
    $serviceFailureEvents = @($systemEvents | Where-Object { $_.ProviderName -eq 'Service Control Manager' -and $_.Id -in $serviceFailureIds })
    $serviceInstallEvents = @($systemEvents | Where-Object { $_.ProviderName -eq 'Service Control Manager' -and $_.Id -eq 7045 })
        $applicationErrorEvents = @($applicationEvents | Where-Object { $_.ProviderName -eq 'Application Error' -and $_.Id -eq 1000 })
        $windowsErrorReportingEvents = @($applicationEvents | Where-Object { $_.ProviderName -eq 'Windows Error Reporting' -and $_.Id -eq 1001 })
        $nonCriticalWerEvents = @($windowsErrorReportingEvents | Where-Object {
            $name = $_.WerEventName
            if ([string]::IsNullOrWhiteSpace($name)) { $name = $_.MessageShort }
            -not ($name -match 'APPCRASH|BEX|CLR20r3|MoAppCrash|BugCheck')
        })
        $applicationCrashEvents = @(
            $applicationErrorEvents +
            ($windowsErrorReportingEvents | Where-Object {
                $name = $_.WerEventName
                if ([string]::IsNullOrWhiteSpace($name)) { $name = $_.MessageShort }
                $name -match 'APPCRASH|BEX|CLR20r3|MoAppCrash'
            })
        )
        $applicationCrashEvents = @($applicationCrashEvents | Sort-Object -Property TimeCreated -Descending)
        $applicationCrashSummaryByProcess = @(
            $applicationCrashEvents |
            ForEach-Object {
                $summaryProcess = $_.AffectedProcess
                if ([string]::IsNullOrWhiteSpace($summaryProcess) -and $_.AffectedPath) {
                    try {
                        $summaryProcess = Normalize-WTProcessName -Value ([System.IO.Path]::GetFileName($_.AffectedPath))
                    }
                    catch {
                        $summaryProcess = $null
                    }
                }
                if ([string]::IsNullOrWhiteSpace($summaryProcess)) {
                    $fallbackProcessInfo = Get-WTEventProcessName -Message $_.Message -MessageShort $_.MessageShort -ProviderName $_.ProviderName -Id $_.Id -WerEventName $_.WerEventName
                    if ($fallbackProcessInfo -and -not [string]::IsNullOrWhiteSpace($fallbackProcessInfo.ProcessName)) {
                        $summaryProcess = $fallbackProcessInfo.ProcessName
                    }
                }
                if ([string]::IsNullOrWhiteSpace($summaryProcess)) {
                    $summaryProcess = 'Unknown'
                }

                [pscustomobject]@{
                    SummaryProcess     = $summaryProcess
                    TimeCreated        = $_.TimeCreated
                    MessageShort       = $_.MessageShort
                    OriginalEvent      = $_
                }
            } |
            Group-Object -Property SummaryProcess |
            Sort-Object -Property @{ Expression = { $_.Count }; Descending = $true }, @{ Expression = { $_.Name } } |
            ForEach-Object {
                $groupEvents = @($_.Group | Sort-Object -Property TimeCreated)
                [pscustomobject]@{
                    ProcessName        = if ([string]::IsNullOrWhiteSpace($_.Name)) { 'Unknown' } else { $_.Name }
                    Count              = $_.Count
                    LastEvent          = $groupEvents[-1].TimeCreated
                    FirstEvent         = $groupEvents[0].TimeCreated
                    ExampleMessageShort = $groupEvents[0].MessageShort
                }
            }
    )

    $eventsById = @(
        $allEvents |
            Group-Object -Property Id |
            Sort-Object -Property @{ Expression = { $_.Count }; Descending = $true }, @{ Expression = { $_.Name } } |
            ForEach-Object {
                [pscustomobject]@{
                    Id    = $_.Name
                    Count = $_.Count
                }
            }
    )

    $eventsByProvider = @(
        $allEvents |
            Group-Object -Property ProviderName |
            Sort-Object -Property @{ Expression = { $_.Count }; Descending = $true }, @{ Expression = { $_.Name } } |
            ForEach-Object {
                [pscustomobject]@{
                    ProviderName = $_.Name
                    Count        = $_.Count
                }
            }
    )

    $recentCriticalEvents = @($allEvents | Where-Object { $_.LevelDisplayName -in @('Critical', 'Error') } | Select-Object -First 20)
    $recentWarningEvents = @($allEvents | Where-Object { $_.LevelDisplayName -eq 'Warning' } | Select-Object -First 20)

    $applicationErrorCrashEventsForImportant = @($applicationCrashEvents | Where-Object { $_.ProviderName -eq 'Application Error' } | Sort-Object -Property TimeCreated -Descending)
    $werAppCrashEventsForImportant = @($applicationCrashEvents | Where-Object { $_.ProviderName -eq 'Windows Error Reporting' -and ($_.WerEventName -eq 'APPCRASH' -or $_.Message -match '(?im)^\s*(Nombre de evento|Event Name):\s*APPCRASH\b') } | Sort-Object -Property TimeCreated -Descending)

    $recentImportantEvents = @()
    $recentImportantEvents += @($unexpectedShutdownEvents | Sort-Object -Property TimeCreated -Descending | Select-Object -First 20)
    $recentImportantEvents += @($bugCheckEvents | Sort-Object -Property TimeCreated -Descending | Select-Object -First 20)
    $recentImportantEvents += @($serviceFailureEvents | Sort-Object -Property TimeCreated -Descending | Select-Object -First 20)
    $recentImportantEvents += @($applicationErrorCrashEventsForImportant | Select-Object -First 20)
    $recentImportantEvents += @($werAppCrashEventsForImportant | Select-Object -First 20)
    $recentImportantEvents += @($serviceInstallEvents | Sort-Object -Property TimeCreated -Descending | Select-Object -First 20)

    if (@($recentImportantEvents).Count -lt 20) {
        $remainingSlots = 20 - @($recentImportantEvents).Count
        $nonCriticalLimit = [math]::Min(2, $remainingSlots)
        if ($nonCriticalLimit -gt 0) {
            $recentImportantEvents += @($nonCriticalWerEvents | Sort-Object -Property TimeCreated -Descending | Select-Object -First $nonCriticalLimit)
        }
    }

    $dedupe = New-Object 'System.Collections.Generic.HashSet[string]'
    $uniqueImportantEvents = @()
    foreach ($evt in @($recentImportantEvents | Where-Object { $_ })) {
        $key = Get-WTEventKey -Event $evt
        if ($key -and $dedupe.Add($key)) {
            $uniqueImportantEvents += $evt
        }
    }
    $recentImportantEvents = @($uniqueImportantEvents | Select-Object -First 20)

    return [pscustomobject]@{
        WindowStart                 = $EventInfo.WindowStart
        WindowEnd                   = $EventInfo.WindowEnd
        Days                        = $EventInfo.Days
        TotalEvents                 = @($allEvents).Count
        SystemEventCount            = @($systemEvents).Count
        ApplicationEventCount       = @($applicationEvents).Count
        UnexpectedShutdownEvents    = @($unexpectedShutdownEvents)
        KernelPowerEvents           = @($kernelPowerEvents)
        BugCheckEvents              = @($bugCheckEvents)
        NormalShutdownEvents        = @($normalShutdownEvents)
        ServiceFailureEvents        = @($serviceFailureEvents)
        ServiceInstallEvents        = @($serviceInstallEvents)
        ApplicationCrashEvents      = @($applicationCrashEvents)
        WindowsErrorReportingEvents  = @($windowsErrorReportingEvents)
        NonCriticalWerEvents        = @($nonCriticalWerEvents)
        ApplicationCrashSummaryByProcess = @($applicationCrashSummaryByProcess)
        EventsById                  = @($eventsById)
        EventsByProvider            = @($eventsByProvider)
        EventCountsByLog            = @($EventInfo.EventCountsByLog)
        EventCountsById             = @($EventInfo.EventCountsById)
        EventCountsByProvider       = @($EventInfo.EventCountsByProvider)
        RecentCriticalEvents        = @($recentCriticalEvents)
        RecentWarningEvents         = @($recentWarningEvents)
        RecentImportantEvents       = @($recentImportantEvents)
        LogsQueried                 = @($EventInfo.LogsQueried)
        LogsUnavailable             = @($EventInfo.LogsUnavailable)
        AllEvents                   = @($allEvents)
        SystemEvents                = @($systemEvents)
        ApplicationEvents           = @($applicationEvents)
    }
}

function Export-WTEventCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$Report,

        [AllowNull()]
        [psobject]$EventInfo
    )

    $rawDirectory = Join-Path -Path $Report.Metadata.ReportDirectory -ChildPath 'raw'
    try {
        [void][System.IO.Directory]::CreateDirectory($rawDirectory)
    }
    catch {
        [void](Add-WTExecutionWarning -Report $Report -Scope 'Events' -Message ('Unable to create raw event output directory: {0}' -f $_.Exception.Message))
        return $null
    }

    $Report.Metadata.RawDirectory = $rawDirectory
    $Report.Metadata.EventCsvSystemPath = ConvertTo-WTAbsolutePath -Path (Join-Path -Path $rawDirectory -ChildPath 'events-system.csv')
    $Report.Metadata.EventCsvApplicationPath = ConvertTo-WTAbsolutePath -Path (Join-Path -Path $rawDirectory -ChildPath 'events-application.csv')
    $Report.Metadata.EventCsvAllPath = ConvertTo-WTAbsolutePath -Path (Join-Path -Path $rawDirectory -ChildPath 'events-all.csv')

    $header = 'TimeCreated,LogName,ProviderName,Id,LevelDisplayName,MessageShort,MachineName,AffectedProcess,AffectedPath,WerEventName'

    $writeCsv = {
        param(
            [string]$Path,
            [object[]]$Rows
        )

        $lines = New-Object System.Collections.Generic.List[string]
        $lines.Add($header) | Out-Null
        foreach ($row in @($Rows)) {
            if (-not $row) {
                continue
            }
            $timeText = ConvertTo-WTDisplayValue -Value (ConvertTo-WTDateTimeString -Value $row.TimeCreated) -Fallback ''
            $cells = @(
                $timeText,
                $row.LogName,
                $row.ProviderName,
                $row.Id,
                $row.LevelDisplayName,
                $row.MessageShort,
                $row.MachineName,
                $row.AffectedProcess,
                $row.AffectedPath,
                $row.WerEventName
            )
            $escaped = foreach ($cell in $cells) {
                if ($null -eq $cell) { '' }
                else {
                    ('"{0}"' -f (('' + $cell) -replace '"', '""'))
                }
            }
            $lines.Add(($escaped -join ',')) | Out-Null
        }

        [System.IO.File]::WriteAllLines($Path, $lines.ToArray(), [System.Text.Encoding]::UTF8)
    }

    try {
        $systemRows = @($EventInfo.SystemEvents)
        $applicationRows = @($EventInfo.ApplicationEvents)
        $allRows = @($EventInfo.AllEvents)
        & $writeCsv -Path $Report.Metadata.EventCsvSystemPath -Rows $systemRows
        & $writeCsv -Path $Report.Metadata.EventCsvApplicationPath -Rows $applicationRows
        & $writeCsv -Path $Report.Metadata.EventCsvAllPath -Rows $allRows
    }
    catch {
        [void](Add-WTExecutionWarning -Report $Report -Scope 'Events' -Message ('Unable to write event CSV exports: {0}' -f $_.Exception.Message))
    }

    return [pscustomobject]@{
        RawDirectory             = $rawDirectory
        SystemCsvPath            = $Report.Metadata.EventCsvSystemPath
        ApplicationCsvPath        = $Report.Metadata.EventCsvApplicationPath
        AllCsvPath               = $Report.Metadata.EventCsvAllPath
    }
}

function Invoke-WTEventRules {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$Report
    )

    $events = $Report.Normalized.Events
    if (-not $events) {
        return
    }

    if (@($events.UnexpectedShutdownEvents).Count -gt 0 -and -not $Report.Normalized.PowerBoot) {
        $ids = @($events.UnexpectedShutdownEvents | Select-Object -ExpandProperty Id -Unique | Sort-Object)
        $latest = @($events.UnexpectedShutdownEvents | Sort-Object -Property TimeCreated -Descending | Select-Object -First 1)
        Add-WTFinding -Report $Report -Id 'WT-EVT-UNEXPECTED-SHUTDOWN' -Category 'Events' -Severity 'High' -Title 'Unexpected shutdown or power loss detected' -Description 'Windows logged one or more unexpected shutdown or Kernel-Power events within the analysis window.' -Evidence @("Count=$(@($events.UnexpectedShutdownEvents).Count)", ('LastEvent={0}' -f (ConvertTo-WTDateTimeString -Value $latest[0].TimeCreated)), ('Ids={0}' -f ($ids -join ', '))) -Recommendation 'Review whether the device experienced power loss, forced shutdown, freeze, crash, or hardware instability. Correlate with BugCheck and Disk events.' -Source 'Invoke-WTEventRules' -Status 'Warning'
    }

    if (@($events.BugCheckEvents).Count -gt 0) {
        $latestBugcheck = @($events.BugCheckEvents | Sort-Object -Property TimeCreated -Descending | Select-Object -First 1)
        $providerNames = @($events.BugCheckEvents | Select-Object -ExpandProperty ProviderName -Unique | Sort-Object)
        Add-WTFinding -Report $Report -Id 'WT-EVT-BUGCHECK-DETECTED' -Category 'Events' -Severity 'Critical' -Title 'Possible BSOD or bugcheck detected' -Description 'Windows logged a bugcheck or system error reporting event within the analysis window.' -Evidence @("Count=$(@($events.BugCheckEvents).Count)", ('LastEvent={0}' -f (ConvertTo-WTDateTimeString -Value $latestBugcheck[0].TimeCreated)), ('Providers={0}' -f ($providerNames -join ', ')), ('MessageShort={0}' -f $latestBugcheck[0].MessageShort) ) -Recommendation 'Review minidump files, driver updates, recent hardware/software changes, and correlate with Kernel-Power events.' -Source 'Invoke-WTEventRules' -Status 'Fail'
    }

    if (@($events.ServiceFailureEvents).Count -ge 5) {
        $ids = @($events.ServiceFailureEvents | Group-Object -Property Id | Sort-Object -Property @{ Expression = { $_.Count }; Descending = $true }, @{ Expression = { $_.Name } } | Select-Object -First 5 | ForEach-Object { '{0}({1})' -f $_.Name, $_.Count })
        $latestServiceFailure = @($events.ServiceFailureEvents | Sort-Object -Property TimeCreated -Descending | Select-Object -First 1)
        $providers = @($events.ServiceFailureEvents | Select-Object -ExpandProperty ProviderName -Unique | Sort-Object)
        Add-WTFinding -Report $Report -Id 'WT-EVT-SERVICE-FAILURES' -Category 'Events' -Severity 'Medium' -Title 'Repeated service failures detected' -Description 'Multiple Service Control Manager failure events were found.' -Evidence @("Count=$(@($events.ServiceFailureEvents).Count)", ('Ids={0}' -f ($ids -join ', ')), ('Providers={0}' -f ($providers -join ', ')), ('LastEvent={0}' -f (ConvertTo-WTDateTimeString -Value $latestServiceFailure[0].TimeCreated))) -Recommendation 'Review affected services, startup type, dependencies and recent software changes.' -Source 'Invoke-WTEventRules' -Status 'Warning'
    }

    if (@($events.ServiceInstallEvents).Count -gt 0) {
        $latestServiceInstall = @($events.ServiceInstallEvents | Sort-Object -Property TimeCreated -Descending | Select-Object -First 1)
        $messages = @($events.ServiceInstallEvents | Select-Object -First 3 | ForEach-Object { $_.MessageShort })
        Add-WTFinding -Report $Report -Id 'WT-EVT-SERVICE-INSTALL' -Category 'Events' -Severity 'Info' -Title 'Service installation events detected' -Description 'Windows logged service installation events within the analysis window.' -Evidence @("Count=$(@($events.ServiceInstallEvents).Count)", ('LastEvent={0}' -f (ConvertTo-WTDateTimeString -Value $latestServiceInstall[0].TimeCreated)), ('RecentMessages={0}' -f ($messages -join ' | '))) -Recommendation 'Validate whether recently installed services are expected, especially if they occurred near crashes or unexpected shutdowns.' -Source 'Invoke-WTEventRules' -Status 'Info'
    }

    $crashCount = Get-WTArrayCountSafe -Value $events.ApplicationCrashEvents
    $nonCriticalWerCount = Get-WTArrayCountSafe -Value $events.NonCriticalWerEvents

    if ($crashCount -gt 0) {
        $latestCrash = @($events.ApplicationCrashEvents | Sort-Object -Property TimeCreated -Descending | Select-Object -First 1)
        $procNames = @($events.ApplicationCrashSummaryByProcess | Select-Object -ExpandProperty ProcessName | Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and $_ -ne 'Unknown' } | Select-Object -First 5)
        if ($procNames.Count -eq 0) {
            $procNames = @('Unknown')
        }
        $topProcess = if ($events.ApplicationCrashSummaryByProcess -and @($events.ApplicationCrashSummaryByProcess).Count -gt 0 -and -not [string]::IsNullOrWhiteSpace($events.ApplicationCrashSummaryByProcess[0].ProcessName) -and $events.ApplicationCrashSummaryByProcess[0].ProcessName -ne 'Unknown') { $events.ApplicationCrashSummaryByProcess[0].ProcessName } else { 'Unknown' }
        $status = if ($crashCount -ge 3) { 'Warning' } else { 'Info' }
        $severity = if ($crashCount -ge 3) { 'Medium' } else { 'Low' }
        Add-WTFinding -Report $Report -Id 'WT-EVT-APP-CRASHES' -Category 'Events' -Severity $severity -Title 'Repeated application crashes detected' -Description 'Multiple real application crash events were found.' -Evidence @("Count=$crashCount", ('Processes={0}' -f ($procNames -join ', ')), ('TopProcess={0}' -f $topProcess), ('LastEvent={0}' -f (ConvertTo-WTDateTimeString -Value $latestCrash[0].TimeCreated))) -Recommendation 'Review affected application, crash frequency, recent updates, add-ins, dependencies and user impact.' -Source 'Invoke-WTEventRules' -Status $status
    }

    if ($nonCriticalWerCount -ge 5) {
        $werNames = @(
            $events.NonCriticalWerEvents |
                ForEach-Object { if ($_.WerEventName) { $_.WerEventName } else { $_.MessageShort } } |
                Where-Object { $_ } |
                Group-Object |
                Sort-Object -Property @{ Expression = { $_.Count }; Descending = $true }, @{ Expression = { $_.Name }; Ascending = $true } |
                Select-Object -First 5 |
                ForEach-Object { '{0}({1})' -f $_.Name, $_.Count }
        )
        $latestWer = @($events.NonCriticalWerEvents | Sort-Object -Property TimeCreated -Descending | Select-Object -First 1)
        Add-WTFinding -Report $Report -Id 'WT-EVT-WER-NONCRITICAL' -Category 'Events' -Severity 'Info' -Title 'Non-critical Windows Error Reporting events detected' -Description 'Windows logged multiple non-critical WER reports that do not appear to be application crashes.' -Evidence @("Count=$nonCriticalWerCount", ('Names={0}' -f ($werNames -join ', ')), ('LastEvent={0}' -f (ConvertTo-WTDateTimeString -Value $latestWer[0].TimeCreated))) -Recommendation 'Review only if users report symptoms related to the affected application.' -Source 'Invoke-WTEventRules' -Status 'Info'
    }

    $hasEventData = @($events.AllEvents).Count -gt 0 -or @($events.LogsUnavailable).Count -lt @($events.LogsQueried).Count
    if ($hasEventData -and @($events.UnexpectedShutdownEvents).Count -eq 0 -and @($events.BugCheckEvents).Count -eq 0 -and @($events.ServiceFailureEvents).Count -eq 0 -and @($events.ServiceInstallEvents).Count -eq 0 -and $crashCount -eq 0) {
        Add-WTFinding -Report $Report -Id 'WT-EVT-NO-CRITICAL-EVENTS' -Category 'Events' -Severity 'Info' -Title 'No critical System/Application events detected' -Description 'No relevant critical System or Application events were found in the selected analysis window.' -Evidence @("Days=$($events.Days)", "TotalEvents=$($events.TotalEvents)") -Recommendation 'No action required based on this event scope.' -Source 'Invoke-WTEventRules' -Status 'Pass'
    }
}

function Invoke-WTEventCorrelationRules {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$Report
    )

    $events = $Report.Normalized.Events
    if (-not $events) {
        return
    }

    $unexpectedEvents = @($events.UnexpectedShutdownEvents | Sort-Object -Property TimeCreated)
    $bugcheckEvents = @($events.BugCheckEvents | Sort-Object -Property TimeCreated)
    $serviceInstallEvents = @($events.ServiceInstallEvents | Sort-Object -Property TimeCreated)

    $bestBugcheckMatch = $null
    foreach ($shutdownEvent in $unexpectedEvents) {
        foreach ($bugcheckEvent in $bugcheckEvents) {
            if ($null -eq $shutdownEvent.TimeCreated -or $null -eq $bugcheckEvent.TimeCreated) {
                continue
            }
            $delta = [math]::Abs(($shutdownEvent.TimeCreated - $bugcheckEvent.TimeCreated).TotalMinutes)
            if ($delta -le 30 -and (-not $bestBugcheckMatch -or $delta -lt $bestBugcheckMatch.DeltaMinutes)) {
                $bestBugcheckMatch = [pscustomobject]@{
                    ShutdownEvent = $shutdownEvent
                    BugCheckEvent = $bugcheckEvent
                    DeltaMinutes  = [math]::Round($delta, 2)
                }
            }
        }
    }

    if ($bestBugcheckMatch) {
        Add-WTFinding -Report $Report -Id 'WT-CORR-BSOD-REBOOT' -Category 'Correlation' -Severity 'Critical' -Title 'BugCheck correlated with unexpected reboot' -Description 'A bugcheck event was found near an unexpected shutdown or Kernel-Power event.' -Evidence @(('ShutdownTime={0}' -f (ConvertTo-WTDateTimeString -Value $bestBugcheckMatch.ShutdownEvent.TimeCreated)), ('BugCheckTime={0}' -f (ConvertTo-WTDateTimeString -Value $bestBugcheckMatch.BugCheckEvent.TimeCreated)), ('DeltaMinutes={0}' -f $bestBugcheckMatch.DeltaMinutes), ('Ids={0}/{1}' -f $bestBugcheckMatch.ShutdownEvent.Id, $bestBugcheckMatch.BugCheckEvent.Id)) -Recommendation 'Prioritize BSOD troubleshooting. Review minidump, drivers, firmware, storage and recent changes.' -Source 'Invoke-WTEventCorrelationRules' -Status 'Fail'
    }

    $bestServiceMatch = $null
    foreach ($shutdownEvent in $unexpectedEvents) {
        foreach ($serviceEvent in $serviceInstallEvents) {
            if ($null -eq $shutdownEvent.TimeCreated -or $null -eq $serviceEvent.TimeCreated) {
                continue
            }
            $delta = ($shutdownEvent.TimeCreated - $serviceEvent.TimeCreated).TotalHours
            if ($delta -ge 0 -and $delta -le 24) {
                $deltaMinutes = [math]::Round(($shutdownEvent.TimeCreated - $serviceEvent.TimeCreated).TotalMinutes, 2)
                if (-not $bestServiceMatch -or $deltaMinutes -lt $bestServiceMatch.DeltaMinutes) {
                    $bestServiceMatch = [pscustomobject]@{
                        ShutdownEvent = $shutdownEvent
                        ServiceEvent = $serviceEvent
                        DeltaMinutes = $deltaMinutes
                    }
                }
            }
        }
    }

    if ($bestServiceMatch) {
        Add-WTFinding -Report $Report -Id 'WT-CORR-SERVICE-INSTALL-REBOOT' -Category 'Correlation' -Severity 'Medium' -Title 'Service installation occurred before unexpected reboot' -Description 'A service installation event was detected shortly before an unexpected reboot or power event.' -Evidence @(('ServiceInstallTime={0}' -f (ConvertTo-WTDateTimeString -Value $bestServiceMatch.ServiceEvent.TimeCreated)), ('ShutdownTime={0}' -f (ConvertTo-WTDateTimeString -Value $bestServiceMatch.ShutdownEvent.TimeCreated)), ('DeltaMinutes={0}' -f $bestServiceMatch.DeltaMinutes), ('MessageShort={0}' -f $bestServiceMatch.ServiceEvent.MessageShort)) -Recommendation 'Validate whether the installed service is expected and whether it may relate to instability.' -Source 'Invoke-WTEventCorrelationRules' -Status 'Warning'
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
            [void](Add-WTExecutionWarning -Report $report -Scope 'OutputPath' -Message ('Fallback output path used because the requested path could not be created. Requested: {0}. Fallback: {1}. Reason: {2}' -f $resolvedOutput.RequestedPath, $resolvedOutput.FallbackPath, $resolvedOutput.FallbackReason))
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

        $powerBootRaw = Invoke-WTSafeCollector -Report $report -Name 'PowerBootInfo' -ScriptBlock {
            Get-WTPowerBootInfo -Days $Days -Mode $mode
        }
        if ($powerBootRaw) {
            $report.Raw.PowerBoot = $powerBootRaw
            Invoke-WTSafeStep -Report $report -Name 'NormalizePowerBoot' -ScriptBlock {
                $report.Normalized.PowerBoot = ConvertTo-WTNormalizedPowerBootInfo -PowerBootInfo $report.Raw.PowerBoot
            } | Out-Null
        }

        $eventRaw = Invoke-WTSafeCollector -Report $report -Name 'EventInfo' -ScriptBlock {
            Get-WTEventInfo -Report $report -Days $Days -Mode $mode
        }
        if ($eventRaw) {
            $report.Raw.Events = $eventRaw
        }
        Invoke-WTSafeStep -Report $report -Name 'NormalizeEvents' -ScriptBlock {
            $report.Normalized.Events = ConvertTo-WTNormalizedEventInfo -EventInfo $report.Raw.Events
        } | Out-Null
        Invoke-WTSafeStep -Report $report -Name 'ExportEventCsv' -ScriptBlock {
            Export-WTEventCsv -Report $report -EventInfo $report.Raw.Events | Out-Null
        } | Out-Null

        $eventProcessParsingWarningNeeded = $false
        $eventProcessParsingWarningMessage = $null
        foreach ($evt in @($report.Normalized.Events.ApplicationCrashEvents)) {
            if (-not $evt) {
                continue
            }
            $hasProcessPattern = $false
            if ($evt.ProviderName -eq 'Application Error' -and $evt.Id -eq 1000 -and -not [string]::IsNullOrWhiteSpace($evt.Message)) {
                if ($evt.Message -match 'Nombre de aplicación con errores:') {
                    $hasProcessPattern = $true
                    $eventProcessParsingWarningMessage = 'Application Error event contains a Spanish faulting application name but process extraction failed.'
                }
                elseif ($evt.Message -match 'Faulting application name:') {
                    $hasProcessPattern = $true
                    $eventProcessParsingWarningMessage = 'Application Error event contains an English faulting application name but process extraction failed.'
                }
            }
            elseif ($evt.ProviderName -eq 'Windows Error Reporting' -and $evt.Id -eq 1001) {
                if (($evt.WerEventName -eq 'APPCRASH') -or ($evt.Message -match 'Nombre de evento:\s*APPCRASH|Event Name:\s*APPCRASH')) {
                    $hasProcessPattern = $true
                    $eventProcessParsingWarningMessage = 'Windows Error Reporting APPCRASH event contains a process name pattern but process extraction failed.'
                }
            }

            if ($hasProcessPattern -and [string]::IsNullOrWhiteSpace($evt.AffectedProcess)) {
                $eventProcessParsingWarningNeeded = $true
                break
            }
        }

        if ($eventProcessParsingWarningNeeded) {
            if ([string]::IsNullOrWhiteSpace($eventProcessParsingWarningMessage)) {
                $eventProcessParsingWarningMessage = 'Application crash event contains a process name pattern but process extraction failed.'
            }
            [void](Add-WTExecutionWarning -Report $report -Scope 'EventProcessParsing' -Message $eventProcessParsingWarningMessage)
        }

        $updateRaw = Invoke-WTSafeCollector -Report $report -Name 'UpdateInfo' -ScriptBlock {
            Get-WTUpdateInfo -Report $report -Days $Days -Mode $mode
        }
        if ($updateRaw) {
            $report.Raw.Updates = $updateRaw
        }
        Invoke-WTSafeStep -Report $report -Name 'NormalizeUpdates' -ScriptBlock {
            $report.Normalized.Updates = ConvertTo-WTNormalizedUpdateInfo -UpdateInfo $report.Raw.Updates
        } | Out-Null
        Invoke-WTSafeStep -Report $report -Name 'ExportUpdateCsv' -ScriptBlock {
            Export-WTUpdateCsv -Report $report -UpdateInfo $report.Raw.Updates | Out-Null
        } | Out-Null

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

        [void](Add-WTFinding -Report $report `
            -Id 'WT-RO-001' `
            -Category 'Framework' `
            -Severity 'Info' `
            -Title 'Tool executed in read-only mode' `
            -Description 'WinTriage is running in a non-remediating, read-only diagnostic mode.' `
            -Evidence @('IsReadOnly=true') `
            -Recommendation 'No action required.' `
            -Source 'Invoke-WinTriage' `
            -Status 'Info' `
            -RequiresAdmin $false)

        [void](Add-WTFinding -Report $report `
            -Id 'WT-MODE-001' `
            -Category 'Framework' `
            -Severity 'Info' `
            -Title ('Execution mode: {0}' -f $mode) `
            -Description 'The selected operating mode controls the breadth of future collectors.' `
            -Evidence @("Mode=$mode", "Days=$Days") `
            -Recommendation 'No action required.' `
            -Source 'Invoke-WinTriage' `
            -Status 'Info' `
            -RequiresAdmin $false)

        Invoke-WTSafeStep -Report $report -Name 'UpdateRules' -ScriptBlock {
            Invoke-WTUpdateRules -Report $report
        } | Out-Null
        Invoke-WTSafeStep -Report $report -Name 'SystemRules' -ScriptBlock {
            Invoke-WTSystemRules -Report $report
        } | Out-Null
        Invoke-WTSafeStep -Report $report -Name 'DiskRules' -ScriptBlock {
            Invoke-WTDiskRules -Report $report
        } | Out-Null
        Invoke-WTSafeStep -Report $report -Name 'PerformanceRules' -ScriptBlock {
            Invoke-WTPerformanceRules -Report $report
        } | Out-Null
        Invoke-WTSafeStep -Report $report -Name 'PowerBootRules' -ScriptBlock {
            Invoke-WTPowerBootRules -Report $report
        } | Out-Null
        Invoke-WTSafeStep -Report $report -Name 'EventRules' -ScriptBlock {
            Invoke-WTEventRules -Report $report
        } | Out-Null
        Invoke-WTSafeStep -Report $report -Name 'EventCorrelationRules' -ScriptBlock {
            Invoke-WTEventCorrelationRules -Report $report
        } | Out-Null

        if (-not $isAdmin) {
            [void](Add-WTSkippedCheck -Report $report -Check 'AdminExample' -Reason 'Skipped example because the script is not running elevated.')
            [void](Add-WTFinding -Report $report `
                -Id 'WT-SKIP-001' `
                -Category 'Privilege' `
                -Severity 'Skipped' `
                -Title 'Example admin-only diagnostic skipped' `
                -Description 'This is a placeholder for a future admin-only check.' `
                -Evidence @('Not running as administrator') `
                -Recommendation 'Rerun elevated only if a future admin-only collector requires it.' `
                -Source 'Invoke-WinTriage' `
                -Status 'Skipped' `
                -RequiresAdmin $true)
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

        $technicalErrors = @(
            $report.Execution.Errors | Where-Object {
                $_.Impact -in @('Partial', 'Fatal') -and $_.AffectsReportIntegrity -eq $true
            }
        )

        if ($report.Summary.Critical -gt 0 -or $report.Summary.High -gt 0) {
            $exitCode = 1
        }
        elseif ($technicalErrors.Count -gt 0) {
            $exitCode = 2
        }
        else {
            $exitCode = 0
        }

        $report.Metadata.ExitCode = $exitCode
        $report.Metadata.FinishedAt = (Get-Date).ToString('o')

        if (-not $script:WTIsJsonOnly) {
            try {
                Export-WTMarkdownReport -Report $report -Path $markdownPath | Out-Null
                $report.Metadata.MarkdownGenerated = $true
            }
            catch {
                [void](Add-WTExecutionWarning -Report $report -Scope 'MarkdownExport' -Message ('Unable to write Markdown report. {0}' -f $_.Exception.Message))
                $report.Metadata.MarkdownGenerated = $false
            }
        }

        if ($OpenReport.IsPresent -and -not $script:WTIsJsonOnly -and $report.Metadata.MarkdownGenerated) {
            try {
                Invoke-Item -LiteralPath $markdownPath
            }
            catch {
                [void](Add-WTExecutionWarning -Report $report -Scope 'OpenReport' -Message ('Unable to open Markdown report. {0}' -f $_.Exception.Message))
            }
        }

        $report.Metadata.JsonGenerated = $true
        Export-WTJsonReport -Report $report -Path $jsonPath | Out-Null
        Write-WTConsoleSummary -Report $report -NoColor:$script:WTNoColor

        return $exitCode
    }
    catch {
        if ($report) {
            [void](Add-WTExecutionError -Report $report -Scope 'Fatal' -Message $_.Exception.Message -Impact 'Fatal' -AffectsReportIntegrity $true)
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
                    [void](Add-WTExecutionWarning -Report $report -Scope 'MarkdownExport' -Message ('Unable to write Markdown report after fatal error. {0}' -f $_.Exception.Message))
                }
            }
            try {
                $report.Metadata.JsonGenerated = $true
                Export-WTJsonReport -Report $report -Path $jsonPath | Out-Null
            }
            catch {
                $report.Metadata.JsonGenerated = $false
                [void](Add-WTExecutionWarning -Report $report -Scope 'JsonExport' -Message ('Unable to write JSON report after fatal error. {0}' -f $_.Exception.Message))
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

if ($SelfTestEventParser.IsPresent) {
    $selfTestResult = Invoke-WTEventParserSelfTest
    if ($UseExitCode.IsPresent) {
        if ($selfTestResult.Passed) {
            exit 0
        }
        exit 1
    }
}
else {
    $script:WTExitCode = Invoke-WinTriage
    if ($UseExitCode.IsPresent) {
        exit $script:WTExitCode
    }
}
