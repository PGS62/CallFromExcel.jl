# Installer for the C#/Excel-DNA CallFromExcel add-in (dotnet-addin), replacing the old
# installer\Install.ps1 (pure-VBA JuliaExcel.xlam world) at the repo root. That script stays
# as-is, serving users who haven't upgraded yet.
#
# Non-interactive by design: no MsgBox dialogs (unlike the old installer) - this is meant to
# be callable unattended from a build script or Julia's own installme(), as well as directly
# by a user. All status is written to the output stream; failures throw a terminating error
# and this exits non-zero.
#
# Targets Windows PowerShell 5.1 (ships on every Windows install) deliberately - no PS7-only
# syntax (ternary operator, null-coalescing, etc.) so nothing extra needs to be installed.

[CmdletBinding()]
param(
    # Left blank and resolved below (relative to the standard repo layout) rather than
    # defaulted here using $PSScriptRoot directly - it isn't reliably populated yet while a
    # param block's own default-value expressions are evaluated.
    [string]$XllSourcePath = "",
    [string]$DllSourcePath = "",
    [string]$TlbSourcePath = "",
    # %APPDATA%\Microsoft\AddIns is Excel's own per-user add-ins folder - the "best practice"
    # no-admin-needed location (it's what Excel's own Add-ins > Browse dialog opens to by
    # default), rather than a self-invented folder like the old installer's
    # C:\Users\Public\CallFromExcel\. A dedicated subfolder (rather than loose files directly in
    # AddIns\) keeps our 2-3 files tidy without affecting how Excel loads the add-in - that's
    # driven entirely by the full path recorded in the registry (see Install-ExcelAddin
    # below), not by which folder the file happens to sit in.
    #
    # Deliberately per-user, unlike the old installer's shared-for-all-users choice: that
    # choice existed to avoid "Excel Link Hell" (a workbook's VBA project stores an absolute
    # path to a referenced .xlam, which breaks when shared between users whose copies live at
    # different paths) - but this add-in's primary, recommended VBA pattern is late-bound
    # CreateObject("CallFromExcel.ComServer") (ProgID-based, no path stored anywhere) and
    # XLL worksheet functions (resolved by function name against loaded add-ins, not by path
    # either) - see dotnet-addin/README.md. Only an optional early-bound Tools > References to
    # the .tlb would store a path, and that's documented as an opt-in, not the recommended
    # pattern, so the per-user Link Hell risk doesn't apply here the way it did for the xlam.
    [string]$DestFolder = (Join-Path $env:APPDATA "Microsoft\AddIns\CallFromExcel"),
    [string]$OldXlamName = "JuliaExcel.xlam"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# dist\ is the committed copy of the Release build (see ..\Update-Dist.ps1).
if ($XllSourcePath -eq "") { $XllSourcePath = Join-Path $PSScriptRoot "..\dist\CallFromExcel.AddIn64-packed.xll" }
if ($DllSourcePath -eq "") { $DllSourcePath = Join-Path $PSScriptRoot "..\dist\CallFromExcel.AddIn.dll" }
if ($TlbSourcePath -eq "") { $TlbSourcePath = Join-Path $PSScriptRoot "..\dist\CallFromExcel.AddIn.tlb" }

$XllDestFileName = "CallFromExcel.AddIn64-packed.xll"
$DllDestFileName = "CallFromExcel.AddIn.dll"
$TlbDestFileName = "CallFromExcel.AddIn.tlb"

function Get-ExcelOptionsRegPath {
    # Matches whichever Office version is actually installed by reading back the version
    # Excel itself reports - see Get-OfficeVersion below. Avoids hardcoding a version number
    # that will eventually be wrong.
    param([string]$OfficeVersion)
    return "HKCU:\Software\Microsoft\Office\$OfficeVersion\Excel\Options"
}

function Test-RegValueExists {
    param([string]$Path, [string]$Name)
    if (-not (Test-Path -Path $Path)) { return $false }
    return $null -ne (Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue)
}

function Get-RegValue {
    param([string]$Path, [string]$Name, [string]$Default = "")
    if (-not (Test-RegValueExists -Path $Path -Name $Name)) { return $Default }
    return (Get-ItemProperty -Path $Path -Name $Name).$Name
}

function Set-RegValue {
    param([string]$Path, [string]$Name, [string]$Value)
    if (-not (Test-Path -Path $Path)) { New-Item -Path $Path -Force | Out-Null }
    Set-ItemProperty -Path $Path -Name $Name -Value $Value -Type String
}

function Remove-RegValue {
    param([string]$Path, [string]$Name)
    Remove-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
}

# -----------------------------------------------------------------------------------------------
# Function  : Remove-ExcelAddinFromRegistry / Install-ExcelAddin
# Purpose   : Ported unchanged (in spirit) from the old installer\Install.ps1 at the repo root -
#             same OPEN/OPENn registry mechanism Excel itself uses to auto-load an add-in, same
#             "matched by filename anywhere in an existing value" dedup logic. Kept as a close
#             port rather than a shared module so this script stays self-contained and the old
#             installer is untouched for users who haven't upgraded yet.
# -----------------------------------------------------------------------------------------------
function Remove-ExcelAddinFromRegistry {
    param([string]$RegPath, [string]$AddinNameFragment)

    $i = 0
    $numAddins = 0
    while ($true) {
        $i++
        $valueName = "OPEN" + $(if ($i -gt 1) { [string]($i - 1) } else { "" })
        if (Test-RegValueExists -Path $RegPath -Name $valueName) { $numAddins++ } else { break }
    }

    $allValues = @()
    $found = $false
    for ($j = 0; $j -lt $numAddins; $j++) {
        $valueName = "OPEN" + $(if ($j -gt 0) { [string]$j } else { "" })
        $value = Get-RegValue -Path $RegPath -Name $valueName -Default ""
        $allValues += [PSCustomObject]@{ Name = $valueName; Value = $value }
        if ($value.ToLower().Contains($AddinNameFragment.ToLower())) { $found = $true }
    }
    if (-not $found) { return $null }

    $removedValue = ($allValues | Where-Object { $_.Value.ToLower().Contains($AddinNameFragment.ToLower()) } | Select-Object -First 1).Value

    foreach ($entry in $allValues) { Remove-RegValue -Path $RegPath -Name $entry.Name }

    $k = 0
    foreach ($entry in $allValues) {
        if (-not $entry.Value.ToLower().Contains($AddinNameFragment.ToLower())) {
            $k++
            $valueName = "OPEN" + $(if ($k -gt 1) { [string]($k - 1) } else { "" })
            Set-RegValue -Path $RegPath -Name $valueName -Value $entry.Value
        }
    }
    return $removedValue
}

function Install-ExcelAddin {
    param([string]$RegPath, [string]$AddinFullName)

    $i = 0
    $numAddins = 0
    while ($true) {
        $i++
        $valueName = "OPEN" + $(if ($i -gt 1) { [string]($i - 1) } else { "" })
        if (Test-RegValueExists -Path $RegPath -Name $valueName) {
            $numAddins++
            $existing = Get-RegValue -Path $RegPath -Name $valueName -Default ""
            if ($existing.ToLower().Contains($AddinFullName.ToLower())) { return }
        } else {
            break
        }
    }
    $valueName = "OPEN" + $(if ($numAddins -gt 0) { [string]$numAddins } else { "" })
    Set-RegValue -Path $RegPath -Name $valueName -Value "/R `"$AddinFullName`""
}

# -----------------------------------------------------------------------------------------------
# Function  : Get-OfficeVersion
# Purpose   : Finds Excel's version by launching it via COM, matching the old installer's own
#             approach (reading the registry directly is unreliable across machines that have
#             had multiple Office versions installed over time).
# -----------------------------------------------------------------------------------------------
function Get-OfficeVersion {
    $excelApp = New-Object -ComObject Excel.Application
    try {
        $excelApp.Visible = $false
        return $excelApp.Version
    } finally {
        $excelApp.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelApp) | Out-Null
    }
}

# -----------------------------------------------------------------------------------------------
# Function  : Register-ComServer
# Purpose   : Registers CallFromExcel.AddIn's two ComVisible classes (ComServer, JuliaDictionary)
#             for COM activation, writing directly to HKCU:\Software\Classes\ rather than
#             shelling out to RegAsm - RegAsm's /tlb and (effectively) /codebase options always
#             write to HKEY_CLASSES_ROOT\/HKLM, needing admin rights; redirecting that to HKCU
#             is the documented workaround (generate a /regfile .reg script and remap
#             HKEY_CLASSES_ROOT -> HKEY_CURRENT_USER\Software\Classes), but RegAsm's reflection
#             step also needs ExcelDna.Integration.dll resolvable, which isn't shipped as a
#             loose file (it's embedded only inside the packed .xll - see JuliaFunctions.cs's
#             own notes on this). Writing the registry structure directly avoids both problems;
#             it's simple, fixed boilerplate for a ComVisible/ClassInterface(None) class with no
#             early-bound (vtable) registration needed, since this add-in's recommended VBA
#             pattern is exclusively late-bound (CreateObject/As Object), which only needs
#             CLSID+ProgID+InprocServer32 - no Interface IID registration at all (confirmed by
#             comparing against RegAsm's own /regfile output, which likewise omits one).
#
#             The CLSIDs below must match the [Guid(...)] attributes on ComServer.cs's
#             ComServer and WireValues.cs's JuliaDictionary classes exactly - update both places
#             together if either ever changes.
# -----------------------------------------------------------------------------------------------
function Register-ComServer {
    param([string]$DllPath)

    $classes = @(
        @{ ProgId = "CallFromExcel.ComServer"; TypeName = "CallFromExcel.AddIn.ComServer"; Clsid = "{0581209a-2052-43b8-8875-83bfe419e8c5}" },
        @{ ProgId = "CallFromExcel.JuliaDictionary"; TypeName = "CallFromExcel.AddIn.JuliaDictionary"; Clsid = "{51faa46c-bed1-43fc-9571-91e7c7bd7385}" }
    )

    foreach ($class in $classes) {
        $progIdPath = "HKCU:\Software\Classes\$($class.ProgId)"
        Set-RegValue -Path $progIdPath -Name "(Default)" -Value $class.ProgId
        Set-RegValue -Path "$progIdPath\CLSID" -Name "(Default)" -Value $class.Clsid

        $clsidPath = "HKCU:\Software\Classes\CLSID\$($class.Clsid)"
        Set-RegValue -Path $clsidPath -Name "(Default)" -Value $class.ProgId
        Set-RegValue -Path "$clsidPath\ProgId" -Name "(Default)" -Value $class.ProgId

        $inprocPath = "$clsidPath\InprocServer32"
        Set-RegValue -Path $inprocPath -Name "(Default)" -Value "mscoree.dll"
        Set-RegValue -Path $inprocPath -Name "ThreadingModel" -Value "Both"
        Set-RegValue -Path $inprocPath -Name "Class" -Value $class.TypeName
        Set-RegValue -Path $inprocPath -Name "Assembly" -Value "CallFromExcel.AddIn, Version=0.0.0.0, Culture=neutral, PublicKeyToken=null"
        Set-RegValue -Path $inprocPath -Name "RuntimeVersion" -Value "v4.0.30319"
        Set-RegValue -Path $inprocPath -Name "CodeBase" -Value ([uri]$DllPath).AbsoluteUri
    }
}

# Returns the PE "Machine" field of EXCEL.EXE (0x014C = 32-bit x86, 0x8664 = x64, 0xAA64 =
# ARM64), or $null if EXCEL.EXE can't be located. Reads the executable itself because Office's
# own bitness registry values aren't stored consistently across versions/install types.
function Get-ExcelMachineType {
    $appPathKey = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe"
    $excelPath = Get-RegValue -Path $appPathKey -Name "(Default)"
    if ($excelPath -eq "" -or -not (Test-Path -Path $excelPath -PathType Leaf)) { return $null }
    $stream = [System.IO.File]::OpenRead($excelPath)
    try {
        $reader = New-Object System.IO.BinaryReader($stream)
        $stream.Position = 0x3C
        $stream.Position = $reader.ReadInt32() + 4
        return $reader.ReadUInt16()
    } finally {
        $stream.Dispose()
    }
}

function Copy-RequiredFile {
    param([string]$SourcePath, [string]$DestPath)
    if (-not (Test-Path -Path $SourcePath -PathType Leaf)) {
        throw "Cannot find file: $SourcePath - it should be committed in dotnet-addin\dist\ (refresh it with dotnet build -c Release in dotnet-addin\CallFromExcel.AddIn)."
    }
    if (Test-Path -Path $DestPath -PathType Leaf) {
        try { Set-ItemProperty -Path $DestPath -Name IsReadOnly -Value $false -ErrorAction Stop } catch {}
    }
    Copy-Item -Path $SourcePath -Destination $DestPath -Force
}

#***************************************************************************************
# Effective start of this script.
#***************************************************************************************

if (Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue) {
    throw "Excel is currently running. Close every Excel window (check Task Manager for a background EXCEL.EXE process too) and run this script again."
}

$machineType = Get-ExcelMachineType
if ($machineType -eq 0x014C) {
    throw "This is 32-bit Excel. CallFromExcel requires 64-bit Excel (the default for Microsoft 365) - nothing has been installed."
}
if ($null -eq $machineType) {
    Write-Output "Could not locate EXCEL.EXE to check whether Excel is 64-bit - continuing (CallFromExcel requires 64-bit Excel)."
}

Write-Output "Detecting Office version..."
$officeVersion = Get-OfficeVersion
$regPath = Get-ExcelOptionsRegPath -OfficeVersion $officeVersion
Write-Output "Found Office version $officeVersion"

# 1) Uninstall the old pure-VBA JuliaExcel.xlam, if present - a user upgrading from the old
# world. Finds the file via whatever path is actually recorded in the registry (not a
# hardcoded assumed path), so this works even if it was installed somewhere other than the old
# installer's own default location.
$oldXlamValue = Remove-ExcelAddinFromRegistry -RegPath $regPath -AddinNameFragment $OldXlamName
if ($null -ne $oldXlamValue) {
    Write-Output "Found old $OldXlamName registered as an add-in - removing."
    $oldXlamPath = $oldXlamValue -replace '^/R\s+"?', '' -replace '"$', ''
    if (Test-Path -Path $oldXlamPath -PathType Leaf) {
        try {
            Set-ItemProperty -Path $oldXlamPath -Name IsReadOnly -Value $false -ErrorAction Stop
        } catch {}
        try {
            Remove-Item -Path $oldXlamPath -Force -ErrorAction Stop
            Write-Output "Deleted old add-in file: $oldXlamPath"
        } catch {
            Write-Output "Could not delete old add-in file (continuing): $oldXlamPath - $($_.Exception.Message)"
        }
    }
} else {
    Write-Output "Old $OldXlamName is not currently installed - nothing to uninstall."
}

# 1b) Uninstall the pre-release C# add-in that was named JuliaExcel (dev_csharp_port branch of
# JuliaExcel.jl) before the rename to CallFromExcel: its add-in entry, files and COM classes.
if ($null -ne (Remove-ExcelAddinFromRegistry -RegPath $regPath -AddinNameFragment "JuliaExcel.AddIn64-packed.xll")) {
    Write-Output "Removed the pre-rename JuliaExcel.AddIn64-packed.xll add-in entry."
}
$preRenameFolder = Join-Path $env:APPDATA "Microsoft\AddIns\JuliaExcel"
if (Test-Path -Path $preRenameFolder) {
    try {
        Remove-Item -Path $preRenameFolder -Recurse -Force -ErrorAction Stop
        Write-Output "Deleted pre-rename add-in folder: $preRenameFolder"
    } catch {
        Write-Output "Could not delete pre-rename add-in folder (continuing): $preRenameFolder - $($_.Exception.Message)"
    }
}
foreach ($key in @("JuliaExcel.AddIn.ComServer", "JuliaExcel.AddIn.JuliaDictionary",
                   "CLSID\{9D7B5C1A-2E4F-4A8B-9C3D-7E1A5B6F8C2D}", "CLSID\{7E4C9F2A-1B8D-4A6E-8C3F-9D2B5A7E1F6C}")) {
    $path = "HKCU:\Software\Classes\$key"
    if (Test-Path -Path $path) {
        Remove-Item -Path $path -Recurse -Force
        Write-Output "Removed pre-rename COM registration: $key"
    }
}

# 2) Copy the new add-in's files to a per-user location that needs no admin rights.
Write-Output "Installing to $DestFolder"
if (-not (Test-Path -Path $DestFolder)) { New-Item -ItemType Directory -Path $DestFolder -Force | Out-Null }

$xllDestPath = Join-Path $DestFolder $XllDestFileName
$dllDestPath = Join-Path $DestFolder $DllDestFileName
$tlbDestPath = Join-Path $DestFolder $TlbDestFileName

Copy-RequiredFile -SourcePath $XllSourcePath -DestPath $xllDestPath
Copy-RequiredFile -SourcePath $DllSourcePath -DestPath $dllDestPath
# The .tlb is optional - only needed for early-bound Tools > References (see README); don't
# fail the whole install if it's missing (e.g. not yet built with RegAsm /tlb).
if (Test-Path -Path $TlbSourcePath -PathType Leaf) {
    Copy-RequiredFile -SourcePath $TlbSourcePath -DestPath $tlbDestPath
    Write-Output "Copied $tlbDestPath (Tools > References > Browse to this for early-bound IntelliSense)."
} else {
    Write-Output "No .tlb found at $TlbSourcePath - skipping (early-bound Tools > References won't be available; late-bound CreateObject still works)."
}

Remove-ExcelAddinFromRegistry -RegPath $regPath -AddinNameFragment $XllDestFileName | Out-Null
Install-ExcelAddin -RegPath $regPath -AddinFullName $xllDestPath
Write-Output "Registered $xllDestPath to auto-load with Excel."

Register-ComServer -DllPath $dllDestPath
Write-Output "Registered CallFromExcel.ComServer for COM (CreateObject(""CallFromExcel.ComServer""))."

Write-Output ""
Write-Output "CallFromExcel is installed. Its functions (JuliaLaunch, JuliaCall, JuliaEval, ...) will be available the next time Excel starts."
exit 0
