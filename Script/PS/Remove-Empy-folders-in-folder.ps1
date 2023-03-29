[CmdletBinding()]
Param(
  [string]$folderPath = 'C:\temp\MMAPDF'
)

function Remove-Emptyfolder($path)
{
    foreach($subfolder in Get-ChildItem -Force -LiteralPath  $path -Directory)
    {
        #Call Function recursively
        Remove-Emptyfolder -path $subfolder.FullName
    }
    $subItems = Get-ChildItem -Force -LiteralPath $path

    if($null -eq $subItems)
    {
        Write-Host "Remove empty folder... $path" -ForegroundColor Red
        Remove-Item -Force -Recurse -LiteralPath $path
    }
    else 
    {
        Write-Host "$path has content.." -ForegroundColor Green
    }
}

Remove-Emptyfolder -path $folderPath