#Gets files from a file and moves them to new folder.-

[CmdletBinding()]
param (
    [Parameter()]
    [string]$valideringsfil = '',
    [string]$xmlurl = '',
    [string]$copyurl = ''
)

$Files = Get-Content $valideringsfil
$Source = $xmlurl
$Destination = $copyurl



foreach ($SourceFile in $Files) {
    $pos = $SourceFile.IndexOf(",")

    Move-Item $Source$SourceFile -Destination $Destination
    Write-Host $SourceFile.Substring(0, $pos) " kopierad" -ForegroundColor DarkMagenta
}