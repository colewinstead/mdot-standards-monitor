param(
    [Parameter(Mandatory = $true)]
    [string]$Recipients,

    [Parameter(Mandatory = $true)]
    [string]$Subject,

    [Parameter(Mandatory = $true)]
    [string]$HtmlBodyFile
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $HtmlBodyFile -PathType Leaf)) {
    throw "HTML body file was not found: $HtmlBodyFile"
}

$recipientList = @($Recipients -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
if ($recipientList.Count -eq 0) {
    throw 'At least one recipient is required.'
}

$outlook = $null
$mail = $null
try {
    $outlook = New-Object -ComObject Outlook.Application
    $mail = $outlook.CreateItem(0)
    $mail.To = $recipientList -join ';'
    $mail.Subject = $Subject
    $mail.HTMLBody = Get-Content -LiteralPath $HtmlBodyFile -Raw -Encoding UTF8
    $mail.Send()
}
finally {
    if ($null -ne $mail) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($mail)
    }
    if ($null -ne $outlook) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($outlook)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
