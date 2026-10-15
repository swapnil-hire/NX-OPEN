# Run with Windows PowerShell 5.1 (powershell.exe), not PowerShell 7.
# Compiles the actual pure comparison source with only the unused NX handle type
# changed to System.Object; no NX installation is needed for these engine tests.
$ErrorActionPreference = 'Stop'
$sourceFile = Join-Path (Split-Path $PSScriptRoot -Parent) 'NX2306_Drawing_Checker.vb'
$source = [IO.File]::ReadAllText($sourceFile)
$start = $source.IndexOf('Public Class Item')
$finish = $source.IndexOf("' -------------------- NX extraction")
$testStart = $source.IndexOf('Public Class EngineTests')
if ($start -lt 0 -or $finish -lt $start -or $testStart -lt 0) { throw 'Cannot locate the source sections.' }
$core = $source.Substring($start, $finish - $start).Replace('Public Native As NXObject', 'Public Native As Object')
$imports = @'
Option Strict On
Option Explicit On
Imports System
Imports System.Text
Imports System.Collections.Generic
Imports System.Globalization
Imports System.Text.RegularExpressions
Imports System.Security.Cryptography
'@
$testCode = $imports + "`r`n" + $core + "`r`n" + $source.Substring($testStart)
Add-Type -TypeDefinition $testCode -Language VisualBasic -ReferencedAssemblies @('System.dll','System.Core.dll','Microsoft.VisualBasic.dll')
[EngineTests]::Run()
Write-Output 'PASS: compiled VB comparison-engine regression tests.'
Write-Output 'This does not validate NX extraction, ownership, highlighting, Excel COM or source preservation in NX.'
