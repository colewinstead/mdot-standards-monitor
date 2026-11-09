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
$stage = 'starting Outlook'
try {
    $outlook = New-Object -ComObject Outlook.Application
    $stage = 'creating the message'
    $mail = $outlook.CreateItem(0)
    $mail.To = $recipientList -join ';'
    $mail.Subject = $Subject
    $mail.HTMLBody = Get-Content -LiteralPath $HtmlBodyFile -Raw -Encoding UTF8
    $stage = 'resolving recipients'
    if (-not $mail.Recipients.ResolveAll()) {
        $unresolved = @(
            for ($index = 1; $index -le $mail.Recipients.Count; $index++) {
                $recipient = $mail.Recipients.Item($index)
                try {
                    if (-not $recipient.Resolved) { $recipient.Name }
                }
                finally {
                    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($recipient)
                }
            }
        )
        throw "Outlook could not resolve recipient(s): $($unresolved -join ', ')"
    }
    if ($InlineImagesJsonFile) {
        if (-not (Test-Path -LiteralPath $InlineImagesJsonFile -PathType Leaf)) {
            throw "Inline-image manifest was not found: $InlineImagesJsonFile"
        }
        # Windows PowerShell 5.1 preserves a top-level JSON array as one
        # pipeline object. Wrapping ConvertFrom-Json in @() therefore creates
        # a nested array, and property access concatenates every image path.
        $inlineImages = Get-Content -LiteralPath $InlineImagesJsonFile -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($image in $inlineImages) {
            if (-not (Test-Path -LiteralPath $image.path -PathType Leaf)) {
                throw "Inline image was not found: $($image.path)"
            }
            $stage = "attaching inline image $($image.path)"
            $attachment = $mail.Attachments.Add([string]$image.path)
            $stage = "setting the content ID for $($image.path)"
            $attachment.PropertyAccessor.SetProperty(
                'http://schemas.microsoft.com/mapi/proptag/0x3712001F',
                [string]$image.content_id
            )
            $stage = "hiding inline attachment $($image.path)"
            $attachment.PropertyAccessor.SetProperty(
                'http://schemas.microsoft.com/mapi/proptag/0x7FFE000B',
                $true
            )
            $inlineAttachments += $attachment
        }
        # Outlook can reject Send() with E_INVALIDARG when attachment MAPI
        # properties have not yet been committed to the MailItem.
        $stage = 'saving the message with inline attachments'
        $mail.Save()
    }
    $stage = 'sending the message'
    $mail.Send()
}
catch {
    throw "Outlook failed while $stage. $($_.Exception.Message)"
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
