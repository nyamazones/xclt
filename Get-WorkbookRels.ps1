param ([string] $Path)

if ([string]::IsNullOrEmpty($Path) -or
    !(Test-Path $Path -PathType Container))
{
    exit 1
}

[string] $workbookRelsFilePath = $Path + '\xl\_rels\workbook.xml.rels'

if (!(Test-Path $workbookRelsFilePath -PathType Leaf))
{
    exit 1
}

[System.Text.RegularExpressions.Match[]] $matches = (
    Get-Content $workbookRelsFilePath -Raw -Encoding UTF8 |
    Select-String '(?s)<Relationships.*?>.*?</Relationships>' `
        -CaseSensitive -AllMatches
).Matches

if ($matches.Length -ne 1)
{
    exit 1
}

$matches[0].Value
exit 0
