param ([string] $Path, [string] $SheetName)

if ([string]::IsNullOrEmpty($Path) -or
    [string]::IsNullOrEmpty($SheetName))
{
    exit 1
}

[string] $sheets = Invoke-Expression ($PSScriptRoot + '\Get-Sheets ' + $Path)

if ($LastExitCode -ne 0)
{
    exit 1
}

foreach ($sheet in $sheets -split "`r`n")
{
    [string] $tmpSheetName = $sheet -creplace '^(.*?): .*?$', '$1'

    if (!($tmpSheetName -cmatch "^$SheetName$"))
    {
        continue
    }

    [string] $sheetFilePath = $sheet -creplace '^.*?: (.*?)$', '$1'

    [System.Text.RegularExpressions.Match[]] $matches = (
        Get-Content $sheetFilePath -Raw -Encoding UTF8 |
        Select-String '(?s)<worksheet.*?>.*?</worksheet>' `
            -CaseSensitive -AllMatches
    ).Matches

    if ($matches.Length -ne 1)
    {
        exit 1
    }

    $matches[0].Value
    exit 0
}

exit 1
