param(
    [Parameter(Mandatory = $true)]
    [string]$Recipients,

    [Parameter(Mandatory = $true)]
    [string]$Subject,

    [Parameter(Mandatory = $true)]
    [string]$HtmlBodyFile,

    [string]$InlineImagesJsonFile
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
$inlineAttachments = @()
try {
    $outlook = New-Object -ComObject Outlook.Application
    $mail = $outlook.CreateItem(0)
    $mail.To = $recipientList -join ';'
    $mail.Subject = $Subject
    $mail.HTMLBody = Get-Content -LiteralPath $HtmlBodyFile -Raw -Encoding UTF8
    if ($InlineImagesJsonFile) {
        if (-not (Test-Path -LiteralPath $InlineImagesJsonFile -PathType Leaf)) {
            throw "Inline-image manifest was not found: $InlineImagesJsonFile"
        }
        $inlineImages = @(Get-Content -LiteralPath $InlineImagesJsonFile -Raw -Encoding UTF8 | ConvertFrom-Json)
        foreach ($image in $inlineImages) {
            if (-not (Test-Path -LiteralPath $image.path -PathType Leaf)) {
                throw "Inline image was not found: $($image.path)"
            }
            $attachment = $mail.Attachments.Add($image.path)
            $attachment.PropertyAccessor.SetProperty(
                'http://schemas.microsoft.com/mapi/proptag/0x3712001F',
                [string]$image.content_id
            )
            $attachment.PropertyAccessor.SetProperty(
                'http://schemas.microsoft.com/mapi/proptag/0x7FFE000B',
                $true
            )
            $inlineAttachments += $attachment
        }
    }
    $mail.Send()
}
finally {
    foreach ($attachment in $inlineAttachments) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($attachment)
    }
    if ($null -ne $mail) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($mail)
    }
    if ($null -ne $outlook) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($outlook)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
