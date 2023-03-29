#ändrar Filnamn
#http://sorastog.blogspot.com/2014/05/update-xml-file-using-powershell.html

[CmdletBinding()]
Param(
  [string]$PDFPath = 'C:\temp\hst - kopia - kopia\PDF',
  [string]$removeString = '1520_',
  [string]$replaceString = ''
)


$pdfFiles = Get-ChildItem -Path $PDFPath -Include '*.PDF' -Recurs
$antal_filer = $pdfFiles | Measure-Object
Write-host ""
write-host "Antal filer: " $antal_filer.Count -ForegroundColor Cyan
Write-host ""
Write-host ""
$c = 0


ForEach($pdfFile in $pdfFiles)
{
    $c++
    write-host "Hanterar...$c av " $antal_filer.count  -foregroundcolor DarkYellow
       
    $tmpFile = Split-Path $pdfFile -Leaf
    
    if($tmpFile -like "*REMISS*")
    {

        #$newFile = $tmpFile.Substring(5, $AttachmentFile.Length -5)
        $newFile = $tmpFile.Replace($removeString, $replaceString)
        
        Rename-Item -Path $pdfFile -NewName $newFile
        Write-Host $pdfFile -ForegroundColor Yellow
        Write-Host $newFile -ForegroundColor Green
        
    }
     
}