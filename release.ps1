$ErrorActionPreference = 'Stop'
& .\package.ps1
if ($LASTEXITCODE -ne 0) {
    throw "Packaging failed."
}
$packageDir = "package"
$zipFile = Get-ChildItem -Path $packageDir -Filter "Checkmeg_*.zip" |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

if (-not $zipFile) {
    Write-Error "Could not find package zip file."
    exit 1
}
$version = $zipFile.Name -replace "Checkmeg_", "" -replace ".zip", ""

Write-Host "Creating release for version $version..."
if (-not (Get-Command gh -ErrorAction SilentlyContinue)) {
    throw "GitHub CLI 'gh' is not installed or not on PATH. Install it from https://cli.github.com/ and run 'gh auth login'."
}

gh release create "v$version" $zipFile.FullName --title "Checkmeg v$version" --generate-notes

if ($LASTEXITCODE -eq 0) {
    Write-Host "Release v$version created successfully."
} else {
    Write-Error "Failed to create release."
}
