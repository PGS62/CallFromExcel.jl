# Copies the Release build of the add-in into dotnet-addin\dist\ (committed to git, so that
# Pkg.add users - who never build the C# themselves - get a ready-built add-in for
# installer\Install.ps1 to install), generates the add-in's type library there (for VBA's
# Tools > References, so early-bound code gets IntelliSense), and records a hash of the
# add-in's source files in dist\source-hash.txt. test/runtests.jl recomputes that hash and
# fails if the source has changed since dist\ was last refreshed.
#
# Runs automatically after every Release build (see CallFromExcel.AddIn.csproj), which passes
# the path of ExcelDna.Integration.dll.
#
# Hash definition - must match source_hash in test/runtests.jl exactly: SHA-256 over every
# file under CallFromExcel.AddIn\ except its bin\ and obj\ folders, in ordinal order of their
# relative paths (forward slashes), each contributing <path> LF <bytes> LF, with all CR bytes
# removed so that git's line-ending conversion can't change the result.

param([Parameter(Mandatory = $true)][string]$ExcelDnaIntegrationPath)

$ErrorActionPreference = "Stop"

$projectDir = Join-Path $PSScriptRoot "CallFromExcel.AddIn"
$buildDir = Join-Path $projectDir "bin\Release\net48"
$distDir = Join-Path $PSScriptRoot "dist"

if (-not (Test-Path $distDir)) { New-Item -ItemType Directory -Path $distDir | Out-Null }
Copy-Item (Join-Path $buildDir "publish\CallFromExcel.AddIn64-packed.xll") $distDir -Force
Copy-Item (Join-Path $buildDir "CallFromExcel.AddIn.dll") $distDir -Force

# The type library, made as TlbExp.exe would (TlbExp itself isn't part of the .NET SDK).
# The Ribbon class is COM-visible and derives from an Excel-DNA class, so the exporter also
# needs a type library for ExcelDna.Integration. That one is written to a temporary folder
# and not shipped: VBA only reads it if asked about the Ribbon class, which no one needs.
# The library is renamed from the assembly-derived CallFromExcel_AddIn to CallFromExcel, so VBA
# code writes CallFromExcel.ComServer.
Add-Type -TypeDefinition @"
using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

public static class TypeLibMaker
{
    [ComImport, Guid("00020406-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface ICreateTypeLib
    {
        void CreateTypeInfo(); void SetName([MarshalAs(UnmanagedType.LPWStr)] string name);
        void SetVersion(); void SetGuid(); void SetDocString();
        void SetHelpFileName(); void SetHelpContext(); void SetLcid(); void SetLibFlags();
        void SaveAllChanges();
    }

    class Sink : ITypeLibExporterNotifySink
    {
        public string RefDir;
        public void ReportEvent(ExporterEventKind kind, int code, string message)
        {
            if (kind != ExporterEventKind.NOTIF_TYPECONVERTED) Console.WriteLine(message);
        }
        public object ResolveRef(Assembly asm)
        {
            return Export(asm, Path.Combine(RefDir, asm.GetName().Name + ".tlb"), this, null);
        }
    }

    static object Export(Assembly asm, string tlbPath, Sink sink, string libName)
    {
        object lib = new TypeLibConverter().ConvertAssemblyToTypeLib(asm, tlbPath, TypeLibExporterFlags.None, sink);
        if (libName != null) ((ICreateTypeLib)lib).SetName(libName);
        ((ICreateTypeLib)lib).SaveAllChanges();
        return lib;
    }

    public static void Make(string dllPath, string tlbPath, string excelDnaIntegrationPath, string refDir)
    {
        AppDomain.CurrentDomain.AssemblyResolve += (s, e) =>
            new AssemblyName(e.Name).Name == "ExcelDna.Integration" ? Assembly.LoadFrom(excelDnaIntegrationPath) : null;
        Export(Assembly.LoadFrom(dllPath), tlbPath, new Sink { RefDir = refDir }, "CallFromExcel");
    }
}
"@
$refDir = Join-Path ([IO.Path]::GetTempPath()) ("CallFromExcel-tlb-" + [Guid]::NewGuid())
New-Item -ItemType Directory -Path $refDir | Out-Null
try {
    [TypeLibMaker]::Make((Join-Path $buildDir "CallFromExcel.AddIn.dll"), (Join-Path $distDir "CallFromExcel.AddIn.tlb"),
        $ExcelDnaIntegrationPath, $refDir)
} finally {
    Remove-Item $refDir -Recurse -Force
}

$projectRoot = (Resolve-Path $projectDir).Path.TrimEnd('\') + '\'
$files = Get-ChildItem $projectDir -Recurse -File | Where-Object {
    $rel = $_.FullName.Substring($projectRoot.Length)
    -not ($rel.StartsWith("bin\") -or $rel.StartsWith("obj\"))
}
$relPaths = [string[]]($files | ForEach-Object { $_.FullName.Substring($projectRoot.Length).Replace('\', '/') })
[Array]::Sort($relPaths, [StringComparer]::Ordinal)

# Latin-1 maps each byte to exactly one char and back, so this strips CR bytes byte-exactly.
$latin1 = [Text.Encoding]::GetEncoding(28591)
$stream = New-Object System.IO.MemoryStream
foreach ($rel in $relPaths) {
    $pathBytes = [Text.Encoding]::UTF8.GetBytes($rel)
    $stream.Write($pathBytes, 0, $pathBytes.Length)
    $stream.WriteByte(10)
    $text = $latin1.GetString([IO.File]::ReadAllBytes((Join-Path $projectRoot $rel))).Replace("`r", "")
    $content = $latin1.GetBytes($text)
    $stream.Write($content, 0, $content.Length)
    $stream.WriteByte(10)
}
$sha = [Security.Cryptography.SHA256]::Create()
$hash = ($sha.ComputeHash($stream.ToArray()) | ForEach-Object { $_.ToString("x2") }) -join ""

[IO.File]::WriteAllText((Join-Path $distDir "source-hash.txt"), $hash + "`n")
Write-Output "Updated $distDir (source hash $hash)"
