param ([string] $Path, [string] $SheetName)

if ([string]::IsNullOrEmpty($Path) -or
    [string]::IsNullOrEmpty($SheetName))
{
    exit 1
}

[System.Text.RegularExpressions.Match[]] $matches = (
    Invoke-Expression ('powershell -File ' +
        $PSScriptRoot + '\Get-Worksheet.ps1 ' + $Path + ' ' + $SheetName) |
    Select-String '(?s)<c.*?>.*?</c>' -CaseSensitive -AllMatches
).Matches

if ($LastExitCode -ne 0)
{
    exit 1
}

if ($matches.Length -eq 0)
{
    exit 1
}

foreach ($match in $matches)
{
    $match.Value
}

exit 0
