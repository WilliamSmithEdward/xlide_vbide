<#
.SYNOPSIS
    Writes a minimal valid workbook for the harness to open.

.DESCRIPTION
    The harness needs Excel to start with a document. Excel that is started with no document does
    not publish itself for automation for about ten seconds, which is most of the cost of a check.

    The file is generated rather than committed as a binary so its contents are reviewable, and it
    is the smallest thing Excel will accept: a content-type map, the package relationships, a
    workbook with one sheet, and an empty sheet.
#>
[CmdletBinding()]
param(
    [string] $Path
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (-not $Path) {
    $here = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
    $Path = Join-Path $here 'fixtures\scratch.xlsx'
}

Add-Type -AssemblyName System.IO.Compression.FileSystem

$parts = [ordered] @{
    '[Content_Types].xml' = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>
'@

    '_rels/.rels' = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
'@

    'xl/workbook.xml' = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>
'@

    'xl/_rels/workbook.xml.rels' = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>
'@

    'xl/worksheets/sheet1.xml' = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>
'@
}

New-Item -ItemType Directory -Force -Path (Split-Path $Path) | Out-Null
if (Test-Path $Path) { Remove-Item $Path -Force }

$archive = [System.IO.Compression.ZipFile]::Open($Path, 'Create')
try {
    foreach ($name in $parts.Keys) {
        $entry = $archive.CreateEntry($name)
        $stream = $entry.Open()
        try {
            $bytes = [System.Text.Encoding]::UTF8.GetBytes($parts[$name])
            $stream.Write($bytes, 0, $bytes.Length)
        }
        finally {
            $stream.Dispose()
        }
    }
}
finally {
    $archive.Dispose()
}

Write-Host ("Wrote {0} ({1:N0} bytes)." -f $Path, (Get-Item $Path).Length)
