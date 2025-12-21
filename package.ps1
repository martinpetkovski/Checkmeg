$major = 1
$timestamp = Get-Date -Format "yyMMddHH"
$version = "$major.$timestamp"
$zipName = "Checkmeg_$version.zip"
$packageDir = "package"

Write-Host "Building..."
Stop-Process -Name Checkmeg -Force -ErrorAction SilentlyContinue
clang++ -std=c++20 -O2 -municode -DUNICODE -D_UNICODE -DWIN32_LEAN_AND_MEAN -DNOMINMAX -I. .\src\main.cpp -o .\Checkmeg.exe -luser32 -lgdi32 -lshell32 -lole32 -lcomdlg32 -lcomctl32 -ladvapi32 -lgdiplus

if ($LASTEXITCODE -ne 0) {
    Write-Error "Build failed."
    exit 1
}

if (Test-Path $packageDir) {
    Remove-Item $packageDir -Recurse -Force
}
New-Item -ItemType Directory -Path $packageDir | Out-Null

Copy-Item "Checkmeg.exe" -Destination $packageDir
Copy-Item "checkmeg.png" -Destination $packageDir
Copy-Item "README.md" -Destination $packageDir -ErrorAction SilentlyContinue

$zipPath = Join-Path $packageDir $zipName
Compress-Archive -Path "$packageDir\*" -DestinationPath $zipPath

Write-Host "Package created: $zipPath"
