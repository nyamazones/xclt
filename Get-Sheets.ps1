using namespace System.Text.RegularExpressions

param ([string] $Path)

if ([string]::IsNullOrEmpty($Path))
{
    exit 1
}

[Match[]] $matches = (
    Invoke-Expression ($PSScriptRoot + '\Get-Workbook ' + $Path) |
    Select-String '(?s)<sheet[^s].*?/>' -CaseSensitive -AllMatches
).Matches

if ($LastExitCode -ne 0)
{
    exit 1
}

[string] $tmpSheets = ''

foreach ($match in $matches)
{
    $tmpSheets +=
        $match.Value -creplace
            '^.*?name="(.*?)".*?r:id="(.*?)".*?$', "`$1: `$2`r`n"
}

$tmpSheets = $tmpSheets.TrimEnd()

[Match[]] $matches = (
    Invoke-Expression ($PSScriptRoot + '\Get-WorkbookRels ' + $Path) |
    Select-String '(?s)<Relationship[^s].*?/>' -CaseSensitive -AllMatches
).Matches

if ($LastExitCode -ne 0)
{
    exit 1
}

[string] $rels = ''

foreach ($match in $matches)
{
    $rels +=
        $match.Value -creplace
            '^.*?Id="(.*?)".*?Target="(.*?)".*?$', "`$1: `$2`r`n"
}

$rels = $rels.TrimEnd()

[string] $sheets = ''

foreach ($tmpSheet in $tmpSheets -split "`r`n")
{
    [string] $relId = $tmpSheet -creplace '^.*?: (.*?)$', '$1'
    [string] $sheetFilePath =
        $rels -creplace "(?s)^.*?\r\n${relId}: (.*?)(?:\r\n.*?)?$", '$1'
    $sheetFilePath = $Path + '\xl\' + $sheetFilePath -creplace '/', '\'
    $sheets += $tmpSheet -creplace '^(.*?): .*?$', "`$1: $sheetFilePath`r`n"
}

$sheets.TrimEnd()
exit 0
