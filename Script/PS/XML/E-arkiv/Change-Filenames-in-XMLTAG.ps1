#ändrar PDF och XML filnamn efter fel i Malmö
#http://sorastog.blogspot.com/2014/05/update-xml-file-using-powershell.html

[CmdletBinding()]
Param(
  [string]$PDFPath = 'C:\temp\hst\pdf',
  [string]$XMLPath = 'C:\temp\hst\xml'
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
    [XML]$XMLData = Get-Content $xmlfile -Encoding UTF8




    $PatCodenumber = $XMLData.JournalData.PatientData.Patient.Codenumber
    $PDFFolder = "$PDFPath\$PatCodenumber"

    foreach($Attachment in $XMLData.JournalData.PatientData.Document.DocAttachment.AttachmentName)
    {
        if($Attachment -ne $null)
        {
            $SveaFile = "$PDFPath$Attachment"

            if(-not ($SveaFile | Test-Path))
            {

                $node = $XMLData.JournalData.PatientData.Document.DocAttachment | Where {$_.AttachmentName -eq $Attachment}
                
                $AttachmentFile = Split-Path $Attachment -Leaf


                if($AttachmentFile -like "*BREV*")
                {

                    $NewAttachmentFileName = $AttachmentFile.Substring(5, $AttachmentFile.Length -5)

                    
                    $node.AttachmentName = "\$PatCodenumber\" + $NewAttachmentFileName
                    
                    #$XMLData.Save($xmlfile)
                                        
                    $attach = $node.AttachmentName
                    #Write-Host $attach
                    
                    $temp_sveafile = $SveaFile.Replace("1500" , "1520")

                    Write-Host $temp_sveafile
                    Write-host $SveaFile
                    
                    Rename-Item -Path $temp_sveafile -NewName $SveaFile
                    
                }
            }
        }

       }  
}