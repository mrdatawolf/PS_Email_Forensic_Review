param (
    [switch]$AsExcel = $false
)
$folderName = "Inbox\Forensic Review"
$outputPath = ".\ParsedHeaders.xlsx"
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


if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}
Import-Module ImportExcel

try {
    $data = Get-OutlookHeaders -FolderName $folderName
} catch {
    Write-Error "Failed to retrieve headers from folder '$folderName'. Error: $_"
    exit 1
}
if ($data) {
    if ($AsExcel) {
        $data | Export-Excel -Path $outputPath -AutoSize -WorksheetName "Headers"
        Write-Host "Exported header data to $outputPath"
    } else {
        $htmlPath = [System.IO.Path]::ChangeExtension([System.IO.Path]::GetTempFileName(), "html")
        $html = "<html><head><title>Email Header Report</title><style>body{font-family:Arial;}table{border-collapse:collapse;width:100%;}th,td{border:1px solid #ccc;padding:8px;}th{background:#f2f2f2;}</style></head><body><h2>Email Header Report</h2><table><tr>"

        $headers = $data[0].psobject.Properties.Name
        foreach ($h in $headers) { $html += "<th>$h</th>" }
        $html += "</tr>"
        foreach ($row in $data) {
            $html += "<tr>"
            foreach ($h in $headers) { $html += "<td>$($row.$h)</td>" }
            $html += "</tr>"
        }
        $html += "</table></body></html>"
        Set-Content -Path $htmlPath -Value $html -Encoding UTF8
        Start-Process $htmlPath
    }
} else {
    Write-Warning "No data extracted."
}
