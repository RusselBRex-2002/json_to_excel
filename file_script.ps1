$jsonFilePath = "C:\Users\shabu.a\Desktop\JSONtoExcel\file.json"
$excelFilePath = "C:\Users\shabu.a\Desktop\JSONtoExcel\file.xlsx"

$jsonData = Get-Content -Path $jsonFilePath | ConvertFrom-Json

$dataArray = @()
foreach ($item in $jsonData) {
    $dataArray += [pscustomobject]$item
}

$dataArray | Export-Excel -Path $excelFilePath -WorksheetName "Sheet1" -AutoSize