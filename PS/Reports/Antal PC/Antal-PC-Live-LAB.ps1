﻿#https://superuser.com/questions/1669028/how-to-read-excel-using-powershell
#https://github.com/dfinke/ImportExcel
#https://adamtheautomator.com/powershell-excel/

Param(
    [string]$ivantiFile = "",
    [string]$ivantiSheet = "Grunddata",
    [string]$computerFile = "",
    [string]$computerSheet = "Innevarnde Månad",
    [string]$totalSheet = "Total"
  )

Import-Module ImportExcel

Write-Host "Start!"
# <-------- Fix variables for date names
$lastmonth = (Get-Date).AddMonths(-2) | Select-Object Year, Month
$lastmonthYear = "$($lastmonth.Year)-$($lastmonth.Month)"
$thismonth =(Get-Date) | Select-Object Year, Month
#$thismonthYear = "$($thismonth.Year)-$($thismonth.Month)"

# Price
$capioPrice = 380
$extPrice = 100

# <-------- Back up $computerFile
Write-Host "Back up file $computerFile"
Copy-Item -Path $computerFile -Destination "NSV antal PC $lastmonthYear.xlsx"

# <-------- Get sheets from $computerFile
Write-Host "Open $computerFile"
$sheetComputer = Import-Excel -Path $computerFile  -WorksheetName $computerSheet
$sheetTotal = Import-Excel -Path $computerFile  -WorksheetName $totalSheet

# <-------- Get sheets from $ivantiFile
Write-Host "Open $ivantiFile"
$sheetIvanti = Import-Excel -Path $ivantiFile  -WorksheetName $ivantiSheet 

# <-------- Backup $computerSheet and set name of sheet to last month 4234
Write-host "Backup $computerSheet and set name of sheet to last month"
$sheetComputer | Export-Excel $computerFile -WorksheetName "$($lastmonth.Year)-$($lastmonth.Month)" -TableName "Antal_ PC_$($lastmonth.Year)$($lastmonth.Month)" -AutoSize  -MoveBefore $computerSheet

# <-------- Filter $ivantiFile
Write-host "Filter $ivantiFile"
$filterIvanti = $sheetIvanti | Select-Object Department, "Computer - Hostname" | Where-Object "Computer - Hostname" -NotLike ""

# <-------- Group Ivanti File
Write-Host "Group Ivanti File"
$a = $filterIvanti | Group-Object  -Property Department -NoElement | Select-Object name, count | Sort-Object name -Descending

Write-Host $a.Count

foreach($kst in $sheetComputer)
{
    Write-Host $kst
    if($kst.'Typ av dator' -eq "Capio dator")
    {
        foreach($tmpKst in $a)
        {       
            
            if($tmpKst.name -eq $kst.kst)
            {
                

                $kst.'Antal datorer' = $tmpKst.Count
                $sheetComputer.Cells[$kst.'Månadskostn IT tjänst konto 6541'].Formula = "=[@[Antal datorer]]*Tabell2[@[2022 kost/PC/mån]]"
                #$kst.'PC-kost per mån' = "=[@[Antal datorer]]*Tabell2[@[2022 kost/PC/mån]]"
                #$kst.'PC-kost per år' = '=[@[Antal datorer]]*Tabell2[@[2022 kost/PC/mån]]*12'
                #$kst.'Månadskostn IT tjänst konto 6541' = '=[@[Antal datorer]]*Tabell2[@[2022 kost/PC/mån]]'

                <#
                $kst.'Antal datorer' = $tmpKst.Count
                $kst.'PC-kost per mån' = $tmpKst.Count * $capioPrice
                $kst.'PC-kost per år' = ($tmpKst.Count * $capioPrice)*12
                $kst.'Månadskostn IT tjänst konto 6541' = $tmpKst.Count * $capioPrice
                #>
                #Write-Host $kst.Kst $kst.'Antal datorer' $tmpKst.name $tmpKst.count
            }
        }
    }
    if($kst.'Typ av dator' -eq "Extern dator")
    {
        $kst.'PC-kost per mån' = $kst.'Antal datorer' * $extPrice
        $kst.'PC-kost per år' = ($kst.'Antal datorer' * $extPrice)*12
        $kst.'Månadskostn IT tjänst konto 6541' = $kst.'Antal datorer' * $extPrice
    }
}
# <-------- Save sheet, Innevarande månad sheet done
Write-Host "Saving Excel file..."
$sheetComputer | Export-Excel $computerFile -WorksheetName $computerSheet


# <-------- start total
# <-------- open new sheet
Write-Host "Start Total"
$sheetComputer = Import-Excel -Path $computerFile  -WorksheetName $computerSheet

# <-------- Count different computer Type
$extComp = 0
$capioComp = 0

foreach($kst in $sheetComputer)
{
    if($kst.'Typ av dator' -eq "Extern dator")
    {   
        $extComp += $kst.'Antal datorer'
        #Write-Host $kst.'Typ av dator' $kst.Kst $kst.'Antal datorer'   
    }
    if($kst.'Typ av dator' -eq "Capio dator")
    {   
        $capioComp += $kst.'Antal datorer'
        #Write-Host $kst.'Typ av dator' $kst.Kst $kst.'Antal datorer'   
    }
}

# <-------- Write "Total" sheet
foreach($r in $sheetTotal)
{
    if($r.'Typ av dator' -eq "Capio dator")
    {
        $r.'Antal PC 2022' = $capioComp
        #$r.'Tot månadskost 2022' =  $r.'Antal PC 2022' * $r.'2022 kost/PC/mån'
        #$r.'Tot årskost 2022' = $r.'Tot månadskost 2022' * 12
        
        $r.'Tot månadskost 2022' =  "=C2*D2"
        $r.'Tot årskost 2022' = "=E2*12"
    }
    if($r.'Typ av dator' -eq "Extern dator")
    {
        $r.'Antal PC 2022' = $extComp
        #$r.'Tot månadskost 2022' =  $r.'Antal PC 2022' * $r.'2022 kost/PC/mån'
        #$r.'Tot årskost 2022' = $r.'Tot månadskost 2022' * 12
        
        $r.'Tot månadskost 2022' =  "=C3*D3"
        $r.'Tot årskost 2022' = "=E3*12"
    }

}

# <-------- Save file and show
Write-Host "Done, saving file $computerFile"
$sheetTotal | Export-Excel $computerFile -WorksheetName "Total" -Show
#$sheetComputer = Import-Excel -Path $computerFile  -WorksheetName $computerSheet
#$sheetComputer | Export-Excel $computerFile -AutoSize -IncludePivotTable -NoTotalsInPivot -PivotRows "Region (Capio)" -PivotData @{"Antal datorer"='sum'} -MoveToEnd

Write-Host "Antal Extern dator" $extComp
Write-Host "Antal Capio dator" $capioComp
#Save new data in $computerSheet
#$sheetComputer | Export-Excel $computerFile -WorksheetName $computerSheet -Show 