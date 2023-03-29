#Count PDF-TAG in XML filer. if filer exists.

[CmdletBinding()]
Param(
    [string]$PDFPath = '',
  [string]$XMLPath = ''
)


$xmlfiles = Get-ChildItem -Path $XMLPath -Include '*.XML' -Recurs
$antal_filer = $xmlfiles | measure
 
$c = 0
Function LogWrite
{
    Param([string]$logstring)
    Add-Content $Logfile -Value $logstring
    #Write-Host $logstring
}


$Logfile = "$PDFPath\PDF_inventering.log"
if(-not ($Logfile | Test-Path))
{
    new-item $Logfile
}
LogWrite "Antal filer $antal_filer ;;;; $XMLPath"
$date = Get-Date
LogWrite $date
LogWrite "Räknar PDF i XML $XMLPath $PDFPath"

    $totalpdfinxml = 0
    $totalpdfinfolder = 0
    $totalpdfinfoldermissing = 0

ForEach($xmlfile in $xmlfiles)
{
    $c++
    
    [XML]$XMLData = Get-Content $xmlfile -Encoding UTF8

   
    foreach($Attachment in $XMLData.JournalData.PatientData.Document.DocAttachment.AttachmentName)
    {
        if($Attachment -ne $null)
        {          
            $totalpdfinxml++
            if(-not ("$PDFPath$Attachment" | Test-Path))
            {
                LogWrite "PDF Fil saknas $PDFPath$Attachment;;; XML: $xmlfile"
                $totalpdfinfoldermissing++
            }
            else 
            {
                $totalkpdfinfolder++
            }
        } 
    }

    foreach($LABAttachment in $XMLData.JournalData.PatientData.Lab.LabResult.LabResultGroup.LabResultItem.ResultImage)
    {
        if($LABAttachment -ne $null)
        {
          $totalpdfinxml++  
            
            if(-not ("$PDFPath$LABAttachment" | Test-Path))
            {
                LogWrite "PDF Fil saknas $Attachment;;; XML: $xmlfile"
                $totalpdfinfoldermissing++
            }
            else 
            {
                $totalpdfinfolder++
            }
        }
    }
    Clear-Host
    write-host "Räknar...$c av " $antal_filer.count  -foregroundcolor Yellow
    Write-Host "Antal PDF i XML: $totalpdfinxml"
    Write-Host "Antal PDF i Mapp: $totalpdfinfolder"
    Write-Host "Antal PDF saknas: $totalpdfinfoldermissing" -ForegroundColor Red
}
LogWrite "Antal PDF i XML: $totalpdfinxml"
LogWrite "Antal PDF i Mapp: $totalpdfinfolder"
LogWrite "Antal PDF saknas: $totalpdfinfoldermissing"