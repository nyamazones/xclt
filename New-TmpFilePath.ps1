[string] $tmpFilePath = ''

while ($true)
{
    [string] $sysTmpDirPath = [System.IO.Path]::GetTempPath()
    [string] $guid = New-Guid
    $tmpFilePath = $sysTmpDirPath + $guid

    if (!(Test-Path $tmpFilePath))
    {
        break
    }
}

$tmpFilePath
exit 0
