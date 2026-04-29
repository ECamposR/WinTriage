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
    [switch]$OpenReport
)

# WinTriage is read-only by design.
# It collects diagnostic data and generates reports without modifying system configuration.

$script:WTVersion = '0.1.0'
$script:WTIsJsonOnly = $JsonOnly.IsPresent
$script:WTNoColor = $NoColor.IsPresent

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

    $timestamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
    $root = Join-Path -Path $BasePath -ChildPath $Hostname
    $reportDir = Join-Path -Path $root -ChildPath $timestamp

    [void][System.IO.Directory]::CreateDirectory($reportDir)
    return (ConvertTo-WTAbsolutePath -Path $reportDir)
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
            ReportDirectory   = $OutputDirectory
            JsonPath          = $null
            MarkdownPath      = $null
            JsonGenerated     = $false
            MarkdownGenerated = $false
            ExitCode          = $null
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
    Write-Host $lines[7] -ForegroundColor DarkGray
    Write-Host $lines[8] -ForegroundColor Green
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

function Get-WTBasicSystemInfo {
    [CmdletBinding()]
    param()

    $os = $null
    $computer = $null
    $bios = $null
    try { $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop } catch { }
    try { $computer = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop } catch { }
    try { $bios = Get-CimInstance -ClassName Win32_BIOS -ErrorAction Stop } catch { }

    return [pscustomobject]@{
        OS              = if ($os) { [pscustomobject]@{ Caption = $os.Caption; Version = $os.Version; BuildNumber = $os.BuildNumber; LastBootUpTime = $os.LastBootUpTime } } else { $null }
        ComputerSystem  = if ($computer) { [pscustomobject]@{ Manufacturer = $computer.Manufacturer; Model = $computer.Model; PartOfDomain = $computer.PartOfDomain; Domain = $computer.Domain } } else { $null }
        BIOS            = if ($bios) { [pscustomobject]@{ SerialNumber = $bios.SerialNumber } } else { $null }
    }
}

function Get-WTMinimalDiskInfo {
    [CmdletBinding()]
    param()

    $drives = @()
    try {
        $drives = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction Stop
    }
    catch {
        $drives = @()
    }

    $result = @()
    foreach ($drive in $drives) {
        $sizeBytes = [double]$drive.Size
        $freeBytes = [double]$drive.FreeSpace
        $sizeKnown = $sizeBytes -gt 0
        $freePercent = $null
        $status = 'OK'

        if ($sizeKnown) {
            $freePercent = [math]::Round(($freeBytes / $sizeBytes) * 100, 2)
        }
        else {
            $status = 'UnknownSize'
        }

        $result += [pscustomobject]@{
            DriveLetter = $drive.DeviceID
            FileSystem  = $drive.FileSystem
            SizeKnown   = $sizeKnown
            SizeGB      = if ($sizeKnown) { [math]::Round($sizeBytes / 1GB, 2) } else { $null }
            FreeGB      = if ($freeBytes -gt 0) { [math]::Round($freeBytes / 1GB, 2) } else { 0 }
            FreePercent = $freePercent
            Status      = $status
        }
    }

    return $result
}

function Get-WTMinimalPerformanceInfo {
    [CmdletBinding()]
    param()

    $os = $null
    $cpu = $null
    try { $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop } catch { }
    try { $cpu = Get-CimInstance -ClassName Win32_Processor -ErrorAction Stop } catch { }

    $totalRamGB = $null
    $freeRamGB = $null
    if ($os) {
        $totalRamGB = [math]::Round(([double]$os.TotalVisibleMemorySize * 1KB) / 1GB, 2)
        $freeRamGB = [math]::Round(([double]$os.FreePhysicalMemory * 1KB) / 1GB, 2)
    }

    return [pscustomobject]@{
        TotalRamGB = $totalRamGB
        FreeRamGB  = $freeRamGB
        CpuCount   = if ($cpu) { @($cpu).Count } else { $null }
        UptimeDays = if ($os -and $os.LastBootUpTime) { [math]::Round(((Get-Date) - $os.LastBootUpTime).TotalDays, 2) } else { $null }
    }
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

        $basicSystem = Invoke-WTSafeCollector -Report $report -Name 'BasicSystemInfo' -ScriptBlock {
            Get-WTBasicSystemInfo
        }

        if ($basicSystem) {
            $system = $basicSystem
            $report.Raw.System = $system
            $report.Context.Manufacturer = if ($system.ComputerSystem) { $system.ComputerSystem.Manufacturer } else { $null }
            $report.Context.Model = if ($system.ComputerSystem) { $system.ComputerSystem.Model } else { $null }
            $report.Context.SerialNumber = if ($system.BIOS) { $system.BIOS.SerialNumber } else { $null }
            $report.Context.IsDomainJoined = if ($system.ComputerSystem) { [bool]$system.ComputerSystem.PartOfDomain } else { $null }
            $report.Context.DomainName = if ($system.ComputerSystem -and $system.ComputerSystem.PartOfDomain) { $system.ComputerSystem.Domain } else { $null }
            $report.Normalized.System = $system
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

        $diskInfo = Get-WTMinimalDiskInfo
        $report.Raw.Disk = $diskInfo
        $report.Normalized.Disk = $diskInfo

        $perfInfo = Get-WTMinimalPerformanceInfo
        $report.Raw.Performance = $perfInfo
        $report.Normalized.Performance = $perfInfo

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

        Update-WTSummary -Report $report | Out-Null

        if ($report.Summary.Critical -gt 0 -or $report.Summary.High -gt 0) {
            $exitCode = 1
        }
        elseif (@($report.Execution.Errors).Count -gt 0) {
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

        $report.Metadata.JsonGenerated = $true
        Export-WTJsonReport -Report $report -Path $jsonPath | Out-Null
        Write-WTConsoleSummary -Report $report -NoColor:$script:WTNoColor

        return $exitCode
    }
    catch {
        try {
            Write-Host 'WinTriage failed during initialization. No report was generated.' -ForegroundColor Red
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
