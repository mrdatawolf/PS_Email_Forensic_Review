param (
    [switch]$AsExcel = $false
)
Import-Module -Name "$PSScriptRoot\GetOutlookHeaderData.psm1" -Force

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}
Import-Module ImportExcel

$envFilePath = "$PSScriptRoot\.env"
Get-EnvFile -envFilePath $envFilePath
$folderName = $env:OUTLOOK_FOLDER
$outputPath = $env:OUTPUT_PATH

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
                $htmlPath = "$PSScriptRoot\\EmailHeaderReport.html"
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
