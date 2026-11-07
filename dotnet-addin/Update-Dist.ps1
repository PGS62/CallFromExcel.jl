# Copies the Release build of the add-in into dotnet-addin\dist\ (committed to git, so that
# Pkg.add users - who never build the C# themselves - get a ready-built add-in for
# installer\Install.ps1 to install), and records a hash of the add-in's source files in
# dist\source-hash.txt. test/runtests.jl recomputes that hash and fails if the source has
# changed since dist\ was last refreshed.
#
# Runs automatically after every Release build (see JuliaExcel.AddIn.csproj).
#
# Hash definition - must match source_hash in test/runtests.jl exactly: SHA-256 over every
# file under JuliaExcel.AddIn\ except its bin\ and obj\ folders, in ordinal order of their
# relative paths (forward slashes), each contributing <path> LF <bytes> LF, with all CR bytes
# removed so that git's line-ending conversion can't change the result.

$ErrorActionPreference = "Stop"

$projectDir = Join-Path $PSScriptRoot "JuliaExcel.AddIn"
$buildDir = Join-Path $projectDir "bin\Release\net48"
$distDir = Join-Path $PSScriptRoot "dist"

if (-not (Test-Path $distDir)) { New-Item -ItemType Directory -Path $distDir | Out-Null }
Copy-Item (Join-Path $buildDir "publish\JuliaExcel.AddIn64-packed.xll") $distDir -Force
Copy-Item (Join-Path $buildDir "JuliaExcel.AddIn.dll") $distDir -Force

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
