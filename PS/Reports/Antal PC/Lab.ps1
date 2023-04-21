Param(
    [string]$ivantiSheet = "Grunddata",
    [string]$computerFile = "C:\\temp\\pytest.xlsx",
    [string]$computerSheet = "Sheet2",
    [string]$totalSheet = "Total"
  )

Import-Module ImportExcel

Write-Host "Start!"


Write-Host "Open $computerFile"
$sheet = Import-Excel -Path $computerFile  -WorksheetName $computerSheet

#$sheet.Cells[2,3].Formula = "=[@B]*Sheet1!A2"
$t = 1
foreach($s in $sheet)
{
    $t++
    #=[@B]*Sheet1!A2
    Write-Host $s
    $s.'C' = "=B$t*Sheet1!A2"
    
 #   $sheet.Cells[$s.'C'].Formula = "=([@B]*Sheet1!A2)"
}
$sheet | Export-Excel -Path $computerFile -WorksheetName $computerSheet -Show

