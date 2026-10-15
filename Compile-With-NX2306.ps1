# Windows PowerShell 5.1. Compile only; never executes the NX journal.
param([Parameter(Mandatory=$true)][string]$NxManagedDirectory)
$ErrorActionPreference = 'Stop'
$sourceFile = Join-Path (Split-Path $PSScriptRoot -Parent) 'NX2306_Drawing_Checker.vb'
$references = @('System.dll','System.Core.dll','System.Drawing.dll','System.Windows.Forms.dll','Microsoft.VisualBasic.dll')
foreach ($name in @('NXOpen.dll','NXOpen.UF.dll')) {
    $dll = Join-Path $NxManagedDirectory $name
    if (-not (Test-Path -LiteralPath $dll)) { throw "Missing NX assembly: $dll" }
    $references += $dll
}
$utilities = Join-Path $NxManagedDirectory 'NXOpen.Utilities.dll'
if (Test-Path -LiteralPath $utilities) { $references += $utilities }
Add-Type -TypeDefinition ([IO.File]::ReadAllText($sourceFile)) -Language VisualBasic -ReferencedAssemblies $references
Write-Output 'PASS: journal compiled against the supplied NX assemblies.'
Write-Output 'NX extraction and Excel runtime tests must still be run inside NX 2306.'
