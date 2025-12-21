$ErrorActionPreference = 'Stop'

# In PowerShell 7+, native commands can be promoted to errors when they write to stderr.
# We intentionally probe for existence (e.g., "release not found"), so keep those probes non-terminating.
if (Test-Path variable:PSNativeCommandUseErrorActionPreference) {
    $PSNativeCommandUseErrorActionPreference = $false
}

function Assert-Command([string] $Name, [string] $InstallHint) {
    if (-not (Get-Command $Name -ErrorAction SilentlyContinue)) {
        throw $InstallHint
    }
}

function Get-DefaultGitRemote() {
    $remotes = @(git remote) | Where-Object { $_ -and $_.Trim().Length -gt 0 }
    if ($remotes -contains 'origin') { return 'origin' }
    if ($remotes.Count -gt 0) { return $remotes[0] }
    throw "No git remotes found. Add a remote (e.g. 'origin') or push manually."
}

function Test-LocalTagExists([string] $Tag) {
    git show-ref --tags --verify --quiet "refs/tags/$Tag" | Out-Null
    return ($LASTEXITCODE -eq 0)
}

function Test-RemoteTagExists([string] $Remote, [string] $Tag) {
    $out = git ls-remote --tags $Remote "refs/tags/$Tag" 2>$null
    return (-not [string]::IsNullOrWhiteSpace($out))
}

function Test-GhReleaseExists([string] $Tag) {
    $old = $ErrorActionPreference
    try {
        $ErrorActionPreference = 'Continue'
        gh release view $Tag --json tagName 1>$null 2>$null
        return ($LASTEXITCODE -eq 0)
    } finally {
        $ErrorActionPreference = $old
    }
}

& .\package.ps1
if ($LASTEXITCODE -ne 0) { throw "Packaging failed." }
$packageDir = "package"
$zipFile = Get-ChildItem -Path $packageDir -Filter "Checkmeg_*.zip" |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

if (-not $zipFile) {
    Write-Error "Could not find package zip file."
    exit 1
}
$version = $zipFile.Name -replace "Checkmeg_", "" -replace ".zip", ""

$tag = "v$version"

Write-Host "Creating release for $tag..."

Assert-Command gh "GitHub CLI 'gh' is not installed or not on PATH. Install it from https://cli.github.com/ and run 'gh auth login'."
Assert-Command git "Git is not installed or not on PATH. Install Git for Windows from https://git-scm.com/download/win."

git rev-parse --is-inside-work-tree 1>$null 2>$null
if ($LASTEXITCODE -ne 0) {
    throw "This script must be run from inside a git repository."
}

$remote = Get-DefaultGitRemote

# If rerunning a release: delete existing GitHub release + git tags, then recreate.
if (Test-GhReleaseExists $tag) {
    Write-Host "Existing GitHub release '$tag' found. Deleting..."
    # Prefer cleanup-tag if supported; fall back if not.
    try {
        gh release delete $tag -y --cleanup-tag
    } catch {
        gh release delete $tag -y
    }
}

if (Test-RemoteTagExists $remote $tag) {
    Write-Host "Existing remote tag '$tag' found on '$remote'. Deleting..."
    git push --delete $remote $tag
    if ($LASTEXITCODE -ne 0) { throw "Failed to delete remote tag '$tag' from '$remote'." }
}

if (Test-LocalTagExists $tag) {
    Write-Host "Existing local tag '$tag' found. Deleting..."
    git tag -d $tag
    if ($LASTEXITCODE -ne 0) { throw "Failed to delete local tag '$tag'." }
}

Write-Host "Tagging current commit as '$tag' and pushing to '$remote'..."
git tag $tag
if ($LASTEXITCODE -ne 0) { throw "Failed to create local tag '$tag'." }

git push $remote $tag
if ($LASTEXITCODE -ne 0) { throw "Failed to push tag '$tag' to '$remote'." }

gh release create $tag $zipFile.FullName --title "Checkmeg $tag" --generate-notes
if ($LASTEXITCODE -ne 0) {
    throw "Failed to create GitHub release '$tag'."
}

Write-Host "Release $tag created successfully." -ForegroundColor Green
