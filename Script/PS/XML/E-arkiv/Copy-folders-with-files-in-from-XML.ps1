
#kopierar alla mappar med attachment i
[CmdletBinding()]
Param(
  [string]$NewPath = '',
  [string]$PDFPath = '',
  [string]$XMLPath = ''
)


$xmlfiles = Get-ChildItem -Path $XMLPath -Include '*.XML' -Recurs
$antal_filer = $xmlfiles | measure
Write-host ""
write-host "Antal filer: " $antal_filer.Count -ForegroundColor Cyan
Write-host ""
Write-host ""
$c = 0


ForEach($xmlfile in $xmlfiles)
{
    $c++
    write-host "Hanterar...$c av " $antal_filer.count  -foregroundcolor DarkYellow
    Write-Host $xmlfile
    [XML]$XMLData = Get-Content $xmlfile

    $PatCodenumber = $XMLData.JournalData.PatientData.Patient.Codenumber
    $PDFFolder = "$PDFPath\$PatCodenumber"

    Copy-Item -Path $PDFFolder -Destination $NewPath -Container -Recurse
    
   
    }