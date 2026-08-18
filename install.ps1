param(
    [switch]$Remove,

    [string]$Recipients,

    [string]$FailureRecipient = 'cole.winstead@stantec.com',

    [switch]$SendTest
)

$ErrorActionPreference = 'Stop'
$taskName = 'MDOT Standards Daily Monitor'
$dataDirectory = Join-Path $env:LOCALAPPDATA 'MDOTStandardsMonitor'
$configFile = Join-Path $dataDirectory 'config.json'
$monitorScript = Join-Path $PSScriptRoot 'monitor.py'
$requirementsFile = Join-Path $PSScriptRoot 'requirements.txt'

if ($Remove) {
    $existing = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
    if ($null -ne $existing) {
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
        Write-Host "Removed scheduled task '$taskName'. Local history and configuration were retained in $dataDirectory."
    }
    else {
        Write-Host "Scheduled task '$taskName' is not installed."
    }
    return
}

if (-not (Test-Path -LiteralPath $monitorScript -PathType Leaf)) {
    throw "Monitor program was not found: $monitorScript"
}

$pythonLauncher = (Get-Command py.exe -ErrorAction SilentlyContinue).Source
if (-not $pythonLauncher) {
    $pythonLauncher = (Get-Command python.exe -ErrorAction SilentlyContinue).Source
}

$dependencyCheckArguments = if ([IO.Path]::GetFileName($pythonLauncher) -ieq 'py.exe') {
    @('-3.14', '-c', 'import pymupdf')
}
else {
    @('-c', 'import pymupdf')
}
& $pythonLauncher @dependencyCheckArguments 2>$null
if ($LASTEXITCODE -ne 0) {
    if (-not (Test-Path -LiteralPath $requirementsFile -PathType Leaf)) {
        throw "Dependency list was not found: $requirementsFile"
    }
    Write-Host 'Installing PDF and Excel comparison support...'
    $installDependencyArguments = if ([IO.Path]::GetFileName($pythonLauncher) -ieq 'py.exe') {
        @('-3.14', '-m', 'pip', 'install', '--user', '-r', $requirementsFile)
    }
    else {
        @('-m', 'pip', 'install', '--user', '-r', $requirementsFile)
    }
    & $pythonLauncher @installDependencyArguments
    if ($LASTEXITCODE -ne 0) {
        throw 'Required Python packages could not be installed.'
    }
}
if (-not $pythonLauncher) {
    throw 'Python was not found. Install Python 3.11 or newer and rerun this installer.'
}

$chromeCandidates = @(
    (Join-Path $env:ProgramFiles 'Google\Chrome\Application\chrome.exe'),
    (Join-Path ${env:ProgramFiles(x86)} 'Google\Chrome\Application\chrome.exe'),
    (Join-Path $env:LOCALAPPDATA 'Google\Chrome\Application\chrome.exe')
)
if (-not ($chromeCandidates | Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Leaf) })) {
    throw 'Google Chrome was not found.'
}

$outlookRegistration = Get-ItemProperty -LiteralPath 'Registry::HKEY_CLASSES_ROOT\Outlook.Application\CLSID' -ErrorAction SilentlyContinue
if (-not $outlookRegistration) {
    throw 'Classic Outlook automation is not registered on this computer.'
}

$nonInteractiveInstall = $PSBoundParameters.ContainsKey('Recipients') -or [bool]$env:MDOT_MONITOR_RECIPIENTS
$rawRecipients = if ($nonInteractiveInstall) {
    if ($PSBoundParameters.ContainsKey('Recipients')) { $Recipients } else { $env:MDOT_MONITOR_RECIPIENTS }
}
else {
    Read-Host 'Enter team email addresses or distribution lists, separated by semicolons'
}
$recipientAddresses = @($rawRecipients -split '[;,]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
if ($recipientAddresses.Count -eq 0) {
    throw 'At least one recipient is required.'
}
foreach ($recipient in $recipientAddresses) {
    if ($recipient -notmatch '^[^\s;@]+@[^\s;@]+\.[^\s;@]+$') {
        throw "Recipient does not look like an email address: $recipient"
    }
}
if ($FailureRecipient -notmatch '^[^\s;@]+@[^\s;@]+\.[^\s;@]+$') {
    throw "Failure recipient does not look like an email address: $FailureRecipient"
}

New-Item -ItemType Directory -Path $dataDirectory -Force | Out-Null
@{
    page_url = 'https://mdot.ms.gov/portal/engineering_standards_guides_manuals'
    recipients = $recipientAddresses
    failure_recipient = $FailureRecipient
    extra_documents = @()
    dynamic_documents = @(
        @{
            title = 'RWD Workflow Training ORD'
            source_url = 'https://mdot.ms.gov/documents/Roadway%20Design/Manuals/CADD/RWD%20Cadd%20Manual.pdf'
            match_host = 'pwdocs.mdot.state.ms.us'
            match_path_contains = '/Resources/Services/ProjectWise/Download.ashx/View'
            section = 'Dynamically monitored documents'
        }
    )
} | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath $configFile -Encoding UTF8

$pythonArguments = if ([IO.Path]::GetFileName($pythonLauncher) -ieq 'py.exe') {
    "-3.14 `"$monitorScript`" check"
}
else {
    "`"$monitorScript`" check"
}
$action = New-ScheduledTaskAction -Execute $pythonLauncher -Argument $pythonArguments -WorkingDirectory $PSScriptRoot
$trigger = New-ScheduledTaskTrigger -Weekly -WeeksInterval 1 -DaysOfWeek Monday, Tuesday, Wednesday, Thursday, Friday -At '12:00 PM'
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -WakeToRun -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Hours 6)
$principal = New-ScheduledTaskPrincipal -UserId ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name) -LogonType Interactive -RunLevel Limited
$task = New-ScheduledTask -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Description 'Checks MDOT engineering standards and emails the configured team when content changes.'
Register-ScheduledTask -TaskName $taskName -InputObject $task -Force | Out-Null

Write-Host 'Creating the initial baseline. This downloads and hashes each linked MDOT document and may take several minutes.'
$initializeArguments = if ([IO.Path]::GetFileName($pythonLauncher) -ieq 'py.exe') { @('-3.14', $monitorScript, 'initialize') } else { @($monitorScript, 'initialize') }
& $pythonLauncher @initializeArguments
if ($LASTEXITCODE -ne 0) {
    Write-Warning "The task was installed, but baseline creation failed. Review $dataDirectory\monitor.log. The scheduled run will retry."
}

$shouldSendTest = $SendTest
if (-not $nonInteractiveInstall -and -not $SendTest) {
    $sendTestAnswer = Read-Host 'Send a test email now? [y/N]'
    $shouldSendTest = $sendTestAnswer -match '^(y|yes)$'
}
if ($shouldSendTest) {
    $testArguments = if ([IO.Path]::GetFileName($pythonLauncher) -ieq 'py.exe') { @('-3.14', $monitorScript, 'send-test') } else { @($monitorScript, 'send-test') }
    & $pythonLauncher @testArguments
}

$installed = Get-ScheduledTask -TaskName $taskName
Write-Host "Installed '$($installed.TaskName)' for 12:00 PM on weekdays."
Write-Host "Configuration and logs: $dataDirectory"
