
function Get-OutlookFolder {
    param (
        $Namespace,
        [string]$FolderPath
    )

    $parts = $FolderPath -split '\\'
    $folder = $Namespace.GetDefaultFolder(6)  # 6 = olFolderInbox

    for ($i = 1; $i -lt $parts.Count; $i++) {
        $folder = $folder.Folders.Item($parts[$i])
        if (-not $folder) {
            throw "Folder path not found: $FolderPath"
        }
    }

    return $folder
}

function Get-OutlookHeaders {
    [CmdletBinding()]
    param (
        [string]$FolderName
    )

    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $folder = Get-OutlookFolder -Namespace $namespace -FolderPath $FolderName

    $results = @()

    foreach ($mail in $folder.Items) {
        if ($mail -is [__ComObject] -and $mail.MessageClass -like "IPM.Note*") {
            $headers = $mail.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E")

            $obj = [PSCustomObject]@{
                Subject         = $mail.Subject
                Sender          = $mail.SenderName
                SentOn          = $mail.SentOn
                MessageID       = if ($headers -match "Message-ID:\s*(.+)") { $matches[1] } else { $null }
                ReturnPath      = if ($headers -match "Return-Path:\s*(.+)") { $matches[1] } else { $null }
                SPF             = if ($headers -match "spf=(pass|fail|softfail|neutral)") { $matches[1] } else { $null }
                DKIM            = if ($headers -match "dkim=(pass|fail|neutral)") { $matches[1] } else { $null }
                DMARC           = if ($headers -match "dmarc=(pass|fail|none|quarantine|reject)") { $matches[1] } else { $null }
                XOriginatingIP  = if ($headers -match "X-Originating-IP:\s*\[?([^\]]+)\]?") { $matches[1] } else { $null }
                SenderIP        = if ($headers -match "X-Originating-IP:\s*\[?([^\]]+)\]?") {
                                    $matches[1]
                                 } elseif ($headers -match "Received:\s*from\s+\S+\s+\((\d{1,3}(?:\.\d{1,3}){3})\)") {
                                    $matches[1]
                                 } else {
                                    $null
                                 }
                Action          = if ($headers -match "spf=(pass|fail|softfail|neutral)") {
                                    switch ($matches[1]) {
                                        'pass'      { 'Allow' }
                                        'fail'      { 'Reject' }
                                        'softfail'  { 'Quarantine' }
                                        'neutral'   { 'Review' }
                                        default     { 'Unknown' }
                                    }
                                 } else {
                                    'Unknown'
                                 }
                PermittedSender = if ($headers -match "spf=pass") { $true } else { $false }
            }

            $results += $obj
        }
    }

    return $results
}

function Get-EnvFile {
    param (
        [string]$envFilePath
    )

    if (-Not (Test-Path $envFilePath)) {
        # Create the .env file with example values
        @"
OUTLOOK_FOLDER=Inbox\Forensic Review
OUTPUT_PATH=.\ParsedHeaders.xlsx
"@ | Out-File -FilePath $envFilePath -Encoding utf8

        Write-Output "The .env file was not found and has been created with example values. Please update it with your actual credentials."
        exit
    }
    # Load the .env file
    Get-Content $envFilePath | ForEach-Object {
        if ($_ -match "^(.*?)=(.*)$") {
            [System.Environment]::SetEnvironmentVariable($matches[1], $matches[2].Trim())
        }
    }
}