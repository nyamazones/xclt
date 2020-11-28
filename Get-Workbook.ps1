param ([string] $Path)

if ([string]::IsNullOrEmpty($Path) -or
    !(Test-Path $Path -PathType Container))
{
    exit 1
}

[string] $workbookFilePath = $Path + '\xl\workbook.xml'

if (!(Test-Path $workbookFilePath -PathType Leaf))
{
    exit 1
}

[System.Text.RegularExpressions.Match[]] $matches = (
    Get-Content $workbookFilePath -Raw -Encoding UTF8 |
    Select-String '(?s)<workbook.*?>.*?</workbook>' -CaseSensitive -AllMatches
).Matches

if ($matches.Length -ne 1)
{
    exit 1
}

$matches[0].Value
exit 0
