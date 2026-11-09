$ErrorActionPreference = 'Stop'
$installer = Join-Path (Split-Path $PSScriptRoot -Parent) 'install.ps1'
$global:TestPythonExecutable = (Get-Command python.exe -ErrorAction Stop).Source
$global:TestBackgroundPython = Join-Path (Split-Path $global:TestPythonExecutable -Parent) 'pythonw.exe'

function Get-Command { [pscustomobject]@{ Source = 'Test-Python' } }
function Test-Python {
    & $global:TestPythonExecutable @args
    if ($LASTEXITCODE -ne 0) { throw 'Python executable resolution failed.' }
}
function Test-Path {
    param($LiteralPath, $PathType)
    if ($LiteralPath -eq $global:TestBackgroundPython) { return $true }
    Microsoft.PowerShell.Management\Test-Path -LiteralPath $LiteralPath -PathType $PathType
}
function New-ScheduledTaskAction {
    param($Execute, $Argument, $WorkingDirectory)
    [pscustomobject]@{ Execute = $Execute; Arguments = $Argument; WorkingDirectory = $WorkingDirectory }
}
function Get-ScheduledTask {
    param($TaskName, $ErrorAction)
    [pscustomobject]@{ TaskName = $TaskName; TaskPath = '\' }
}
function Set-ScheduledTask {
    param($TaskName, $TaskPath, $Action)
    $global:TestUpdatedTask = [pscustomobject]@{ TaskName = $TaskName; TaskPath = $TaskPath; Action = $Action }
}
function New-Item { throw 'BackgroundOnly must not modify local data.' }
function Set-Content { throw 'BackgroundOnly must not overwrite configuration.' }
function Register-ScheduledTask { throw 'BackgroundOnly must not replace the task.' }

& $installer -BackgroundOnly
$root = Split-Path $PSScriptRoot -Parent
if ($global:TestUpdatedTask.TaskName -ne 'MDOT Standards Daily Monitor' -or
    $global:TestUpdatedTask.TaskPath -ne '\' -or
    $global:TestUpdatedTask.Action.Execute -ne $global:TestBackgroundPython -or
    $global:TestUpdatedTask.Action.Arguments -ne "`"$(Join-Path $root 'monitor.py')`" check" -or
    $global:TestUpdatedTask.Action.WorkingDirectory -ne $root) {
    throw 'Incorrect background task action.'
}
Write-Host 'Background-only task update passed (real Python, Task Scheduler mocked).'
