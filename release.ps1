# Run the package script
.\package.ps1

# Find the generated zip file
$packageDir = "package"
$zipFile = Get-ChildItem -Path $packageDir -Filter "Checkmeg_*.zip" | Select-Object -First 1

if (-not $zipFile) {
    Write-Error "Could not find package zip file."
    exit 1
}

# Extract version from filename Checkmeg_1.25122101.zip
# $zipFile.Name is Checkmeg_1.25122101.zip
# We want 1.25122101
$version = $zipFile.Name -replace "Checkmeg_", "" -replace ".zip", ""

Write-Host "Creating release for version $version..."

# Create GitHub release
# gh release create <tag> <files> --title <title> --generate-notes
gh release create "v$version" $zipFile.FullName --title "Checkmeg v$version" --generate-notes

if ($LASTEXITCODE -eq 0) {
    Write-Host "Release v$version created successfully."
} else {
    Write-Error "Failed to create release."
}
