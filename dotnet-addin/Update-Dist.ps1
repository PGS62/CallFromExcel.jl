# Copies the Release build of the add-in into dotnet-addin\dist\ (committed to git, so that
# Pkg.add users - who never build the C# themselves - get a ready-built add-in for
# installer\Install.ps1 to install), generates the add-in's type library there (for VBA's
# Tools > References, so early-bound code gets IntelliSense), and records a hash of the
# add-in's source files in dist\source-hash.txt. test/runtests.jl recomputes that hash and
# fails if the source has changed since dist\ was last refreshed.
#
# Runs automatically after every Release build (see CallFromExcel.AddIn.csproj).
#
# Hash definition - must match source_hash in test/runtests.jl exactly: SHA-256 over every
# file under CallFromExcel.AddIn\ except its bin\ and obj\ folders, in ordinal order of their
# relative paths (forward slashes), each contributing <path> LF <bytes> LF, with all CR bytes
# removed so that git's line-ending conversion can't change the result.

$ErrorActionPreference = "Stop"

$projectDir = Join-Path $PSScriptRoot "CallFromExcel.AddIn"
$buildDir = Join-Path $projectDir "bin\Release\net48"
$distDir = Join-Path $PSScriptRoot "dist"

if (-not (Test-Path $distDir)) { New-Item -ItemType Directory -Path $distDir | Out-Null }
Copy-Item (Join-Path $buildDir "publish\CallFromExcel.AddIn64-packed.xll") $distDir -Force
Copy-Item (Join-Path $buildDir "CallFromExcel.AddIn.dll") $distDir -Force

# The type library (for VBA's Tools > References), made as TlbExp.exe would - TlbExp itself
# isn't part of the .NET SDK - but from a stand-in assembly built here in memory, holding only
# copies of ICallFromExcelComServer and ComServer (same names, GUIDs, attributes, method order,
# parameter names, optional flags and defaults; ComServer's methods are empty stubs). Exporting
# the add-in's own DLL would also export its Ribbon class, which derives from Excel-DNA classes
# implementing Office interfaces. The exporter finds those through Office's registered type
# libraries, so the file would differ between machines with and without Office (as on CI).
# The library is named CallFromExcel (not the assembly-derived CallFromExcel_AddIn), so VBA
# code writes CallFromExcel.ComServer.
Add-Type -TypeDefinition @"
using System;
using System.Linq;
using System.Reflection;
using System.Reflection.Emit;
using System.Runtime.InteropServices;

public static class TypeLibMaker
{
    [ComImport, Guid("00020406-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface ICreateTypeLib
    {
        // In vtable order; only SetName and SaveAllChanges are called.
        void CreateTypeInfo(); void SetName([MarshalAs(UnmanagedType.LPWStr)] string name);
        void SetVersion(); void SetGuid(); void SetDocString();
        void SetHelpFileName(); void SetHelpContext(); void SetLcid(); void SetLibFlags();
        void SaveAllChanges();
    }

    class Sink : ITypeLibExporterNotifySink
    {
        public void ReportEvent(ExporterEventKind kind, int code, string message)
        {
            if (kind != ExporterEventKind.NOTIF_TYPECONVERTED) Console.WriteLine(message);
        }
        public object ResolveRef(Assembly asm)
        {
            throw new InvalidOperationException("The type library would depend on " + asm.FullName);
        }
    }

    // The run-time equivalent of a type read from the reflection-only copy of the add-in.
    // Only framework types (object, string, bool, int, ...) can appear in the COM interface.
    static Type Runtime(Type t, Type srcIface, Type iface)
    {
        if (t == srcIface) return iface;
        if (t.Assembly.GetName().Name != "mscorlib") throw new NotSupportedException("Unsupported type " + t.FullName);
        return Type.GetType(t.AssemblyQualifiedName, true);
    }

    static CustomAttributeBuilder Copy(CustomAttributeData d, Type srcIface, Type iface)
    {
        Type[] types = d.Constructor.GetParameters().Select(p => Runtime(p.ParameterType, srcIface, iface)).ToArray();
        object[] args = d.ConstructorArguments.Select((a, i) =>
            a.Value is Type ? Runtime((Type)a.Value, srcIface, iface) :
            types[i].IsEnum ? Enum.ToObject(types[i], a.Value) : a.Value).ToArray();
        if (d.NamedArguments.Count > 0) throw new NotSupportedException("Named attribute arguments: " + d);
        return new CustomAttributeBuilder(Runtime(d.Constructor.DeclaringType, srcIface, iface).GetConstructor(types), args);
    }

    public static void Make(string dllPath, string tlbPath, string ifaceName, string className, string libName)
    {
        AppDomain.CurrentDomain.ReflectionOnlyAssemblyResolve += (s, e) => Assembly.ReflectionOnlyLoad(e.Name);
        Assembly src = Assembly.ReflectionOnlyLoadFrom(dllPath);
        Type srcIface = src.GetType(ifaceName, true), srcClass = src.GetType(className, true);

        var name = new AssemblyName(src.GetName().Name) { Version = src.GetName().Version };
        AssemblyBuilder ab = AppDomain.CurrentDomain.DefineDynamicAssembly(name, AssemblyBuilderAccess.Run);
        foreach (var d in CustomAttributeData.GetCustomAttributes(src).Where(d => d.Constructor.DeclaringType.Name == "GuidAttribute"))
            ab.SetCustomAttribute(Copy(d, null, null));
        ModuleBuilder mb = ab.DefineDynamicModule(name.Name);

        TypeBuilder ib = mb.DefineType(ifaceName, TypeAttributes.Public | TypeAttributes.Interface | TypeAttributes.Abstract);
        foreach (var d in CustomAttributeData.GetCustomAttributes(srcIface)) ib.SetCustomAttribute(Copy(d, null, null));
        var srcMethods = srcIface.GetMethods().OrderBy(m => m.MetadataToken).ToArray();
        var ifaceMethods = srcMethods.Select(m => Define(ib, m, MethodAttributes.Abstract, null, null)).ToArray();
        Type iface = ib.CreateType();

        TypeBuilder cb = mb.DefineType(className, TypeAttributes.Public | TypeAttributes.Class, typeof(object), new[] { iface });
        foreach (var d in CustomAttributeData.GetCustomAttributes(srcClass)) cb.SetCustomAttribute(Copy(d, srcIface, iface));
        cb.DefineDefaultConstructor(MethodAttributes.Public);
        for (int i = 0; i < srcMethods.Length; i++)
        {
            MethodBuilder impl = Define(cb, srcMethods[i], MethodAttributes.Final, srcIface, iface);
            ILGenerator il = impl.GetILGenerator();
            il.Emit(OpCodes.Newobj, typeof(NotSupportedException).GetConstructor(Type.EmptyTypes));
            il.Emit(OpCodes.Throw);
            cb.DefineMethodOverride(impl, ifaceMethods[i]);
        }
        cb.CreateType();

        object lib = new TypeLibConverter().ConvertAssemblyToTypeLib(ab, tlbPath, TypeLibExporterFlags.None, new Sink());
        ((ICreateTypeLib)lib).SetName(libName);
        ((ICreateTypeLib)lib).SaveAllChanges();
    }

    static MethodBuilder Define(TypeBuilder tb, MethodInfo m, MethodAttributes extra, Type srcIface, Type iface)
    {
        ParameterInfo[] ps = m.GetParameters();
        MethodBuilder b = tb.DefineMethod(m.Name,
            MethodAttributes.Public | MethodAttributes.Virtual | MethodAttributes.HideBySig | MethodAttributes.NewSlot | extra,
            Runtime(m.ReturnType, null, null), ps.Select(p => Runtime(p.ParameterType, null, null)).ToArray());
        for (int i = 0; i < ps.Length; i++)
        {
            ParameterBuilder pb = b.DefineParameter(i + 1, ps[i].Attributes, ps[i].Name);
            if ((ps[i].Attributes & ParameterAttributes.HasDefault) != 0) pb.SetConstant(ps[i].RawDefaultValue);
            foreach (var d in CustomAttributeData.GetCustomAttributes(ps[i])) pb.SetCustomAttribute(Copy(d, null, null));
        }
        return b;
    }
}
"@ -ReferencedAssemblies System.Core
[TypeLibMaker]::Make((Join-Path $buildDir "CallFromExcel.AddIn.dll"), (Join-Path $distDir "CallFromExcel.AddIn.tlb"),
    "CallFromExcel.AddIn.ICallFromExcelComServer", "CallFromExcel.AddIn.ComServer", "CallFromExcel")

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
