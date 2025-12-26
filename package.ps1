$major = 1
$timestamp = Get-Date -Format "yyMMddHH"
$version = "$major.$timestamp"
$zipName = "Checkmeg_$version.zip"
$packageDir = "package"

# Update Version.h
$versionHeader = "#pragma once`n#define CHECKMEG_VERSION `"$version`""
Set-Content -Path "src/Version.h" -Value $versionHeader

Write-Host "Building..."
Stop-Process -Name Checkmeg -Force -ErrorAction SilentlyContinue
llvm-rc src/resource.rc /FO resource.res
if ($LASTEXITCODE -ne 0) {
    Write-Error "Resource compilation failed."
    exit 1
}

clang++ -std=c++20 -O2 -municode -DUNICODE -D_UNICODE -DWIN32_LEAN_AND_MEAN -DNOMINMAX -I. .\src\main.cpp .\src\SupabaseAuth.cpp resource.res -o .\Checkmeg.exe -luser32 -lgdi32 -lshell32 -lole32 -lcomdlg32 -lcomctl32 -ladvapi32 -lgdiplus -lwinhttp

if ($LASTEXITCODE -ne 0) {
    Write-Error "Build failed."
    exit 1
}

if (Test-Path $packageDir) {
    Remove-Item $packageDir -Recurse -Force
}
New-Item -ItemType Directory -Path $packageDir | Out-Null

Copy-Item "Checkmeg.exe" -Destination $packageDir

$zipPath = Join-Path $packageDir $zipName
Compress-Archive -Path "$packageDir\Checkmeg.exe" -DestinationPath $zipPath

Write-Host "Package created: $zipPath"
