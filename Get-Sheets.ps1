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

[string] $tmp_sheets = ''

foreach ($match in $matches)
{
    $tmp_sheets +=
        $match.Value -creplace
            '^.*?name="(.*?)".*?r:id="(.*?)".*?$', "`$1: `$2`r`n"
}

$tmp_sheets = $tmp_sheets.TrimEnd()

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

foreach ($tmp_sheet in $tmp_sheets -split "`r`n")
{
    [string] $relId = $tmp_sheet -creplace '^.*?: (.*?)$', '$1'
    [string] $sheetFilePath =
        $rels -creplace "(?s)^.*?\r\n${relId}: (.*?)(?:\r\n.*?)?$", '$1'
    $sheetFilePath = $Path + '\xl\' + $sheetFilePath -creplace '/', '\'
    $sheets += $tmp_sheet -creplace '^(.*?): .*?$', "`$1: $sheetFilePath`r`n"
}

$sheets.TrimEnd()
exit 0
