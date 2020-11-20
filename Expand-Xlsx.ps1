param ([string] $Path, [string] $DestinationPath)

if ([string]::IsNullOrEmpty($Path) -or
    !(Test-Path $Path -Include '*.xlsx' -PathType Leaf) -or
    [string]::IsNullOrEmpty($DestinationPath) -or
    (Test-Path $DestinationPath))
{
    exit 1
}

New-Item $DestinationPath -ItemType 'Directory' | Out-Null
[string] $xlsxFilePath = $Path
[string] $xlsxFileBaseName =
    $xlsxFilePath -creplace '^(?:.*\\)?(.*?)\.xlsx$', '$1'
[string] $zipFilePath = $DestinationPath + '\' + $xlsxFileBaseName + '.zip'
Copy-Item $xlsxFilePath $zipFilePath
[string] $xlsxDirPath = $zipFilePath -creplace '^(.*)\.zip$', '$1'
Expand-Archive $zipFilePath $xlsxDirPath
Remove-Item $zipFilePath
exit 0
