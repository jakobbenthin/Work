#Copy PDF from XML TAG.

[CmdletBinding()]
Param(
  [string]$SveaPath = '',
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

    if(-not ($PDFFolder | Test-Path))
    {
        
        Write-Host "Skapar PDF-mapp: $PDFFolder" -ForegroundColor Yellow
        New-item $PDFFolder -ItemType Directory
        Write-Host "Mapp skapad" -ForegroundColor Green
    }


    foreach($Attachment in $XMLData.JournalData.PatientData.Document.DocAttachment.AttachmentName)
    {
        if($Attachment -ne $null)
        {
            $AttachmentFile = Split-Path $Attachment -Leaf
            
            $SveaFile = "$SveaPath\$AttachmentFile"
            if(-not ($SveaFile | Test-Path))
            {
                Write-Host "Hittar ej fil... $SveaFile" -ForegroundColor Red
            }
            else 
            {
                Write-Host "Kopierar Bilaga $AttachmentFile" -ForegroundColor Yellow
                Copy-Item -Path $SveaFile -Destination $PDFFolder -Recurse
                Write-Host "Kopiering klar $AttachmentFile" -ForegroundColor Green
            }
        }
    }
    foreach($LABAttachment in $XMLData.JournalData.PatientData.Lab.LabResult.LabResultGroup.LabResultItem.ResultImage)
    {
        if($LABAttachment -ne $null)
        {
            $LABAttachmentFile = (Split-Path $LABAttachment -leaf)
            $SveaFile = "$SveaPath\$LabAttachmentFile"
            if(-not ($SveaFile | Test-Path))
            {
                Write-Host "Hittar ej fil... $SveaFile" -ForegroundColor Red
            }
            else 
            {
                Write-Host "Kopierar LAB $LabAttachmentFile" -ForegroundColor Yellow
                Copy-Item -Path $SveaFile -Destination $PDFFolder -Recurse
                Write-Host "Kopiering klar $LabAttachmentFile" -ForegroundColor Green
            }
        }
    }

}