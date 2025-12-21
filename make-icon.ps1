<#
.SYNOPSIS
	Convert a PNG logo to a multi-size Windows .ico file.

.DESCRIPTION
	Generates an ICO that contains PNG images at multiple sizes (supported by Windows).
	This script uses System.Drawing (works on Windows; best in Windows PowerShell 5.1).

.EXAMPLE
	./make-icon.ps1

.EXAMPLE
	./make-icon.ps1 -InputPng ./checkmeg.png -OutputIco ./checkmeg.ico -Sizes 16,32,48,64,128,256 -Force
#>

[CmdletBinding()]
param(
	[Parameter()] [string] $InputPng = "./checkmeg.png",
	[Parameter()] [string] $OutputIco = "./checkmeg.ico",
	[Parameter()] [ValidateNotNullOrEmpty()] [int[]] $Sizes = @(16, 32, 48, 64, 128, 256),
	[Parameter()] [switch] $Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Resolve-FullPath([string] $Path) {
	return [System.IO.Path]::GetFullPath((Resolve-Path -LiteralPath $Path).Path)
}

function Get-PngBytesForSize {
	param(
		[Parameter(Mandatory)] [string] $PngPath,
		[Parameter(Mandatory)] [int] $Size
	)

	Add-Type -AssemblyName System.Drawing | Out-Null

	$src = $null
	$dst = $null
	$g = $null
	$ms = $null

	try {
		$src = [System.Drawing.Image]::FromFile($PngPath)

		$dst = New-Object System.Drawing.Bitmap $Size, $Size, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
		$dst.SetResolution($src.HorizontalResolution, $src.VerticalResolution)

		$g = [System.Drawing.Graphics]::FromImage($dst)
		$g.CompositingMode = [System.Drawing.Drawing2D.CompositingMode]::SourceCopy
		$g.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
		$g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
		$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
		$g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
		$g.Clear([System.Drawing.Color]::Transparent)

		$rect = New-Object System.Drawing.Rectangle 0, 0, $Size, $Size
		$g.DrawImage($src, $rect)

		$ms = New-Object System.IO.MemoryStream
		$dst.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
		return $ms.ToArray()
	}
	finally {
		if ($g) { $g.Dispose() }
		if ($dst) { $dst.Dispose() }
		if ($src) { $src.Dispose() }
		if ($ms) { $ms.Dispose() }
	}
}

function Write-IcoFile {
	param(
		[Parameter(Mandatory)] [string] $OutPath,
		[Parameter(Mandatory)] [hashtable[]] $Images
	)

	$fs = $null
	$bw = $null

	try {
		$fs = [System.IO.File]::Open($OutPath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
		$bw = New-Object System.IO.BinaryWriter $fs

		# ICONDIR
		$bw.Write([UInt16]0) # reserved
		$bw.Write([UInt16]1) # type: 1 = icon
		$bw.Write([UInt16]$Images.Count)

		$offset = 6 + (16 * $Images.Count)

		foreach ($img in $Images) {
			$size = [int]$img.Size
			$data = [byte[]]$img.Data

			$w = if ($size -ge 256) { [byte]0 } else { [byte]$size }
			$h = if ($size -ge 256) { [byte]0 } else { [byte]$size }

			$bw.Write($w)                    # bWidth
			$bw.Write($h)                    # bHeight
			$bw.Write([byte]0)               # bColorCount
			$bw.Write([byte]0)               # bReserved
			$bw.Write([UInt16]1)             # wPlanes
			$bw.Write([UInt16]32)            # wBitCount
			$bw.Write([UInt32]$data.Length)  # dwBytesInRes
			$bw.Write([UInt32]$offset)       # dwImageOffset

			$offset += $data.Length
		}

		foreach ($img in $Images) {
			$bw.Write([byte[]]$img.Data)
		}
	}
	finally {
		if ($bw) { $bw.Dispose() }
		if ($fs) { $fs.Dispose() }
	}
}

if (-not (Test-Path -LiteralPath $InputPng)) {
	throw "Input PNG not found: $InputPng"
}

$Sizes = $Sizes | Where-Object { $_ -gt 0 } | Sort-Object -Unique
if ($Sizes.Count -eq 0) {
	throw "No valid sizes provided."
}

if ((Test-Path -LiteralPath $OutputIco) -and -not $Force) {
	throw "Output already exists: $OutputIco (use -Force to overwrite)"
}

$inputFull = Resolve-FullPath $InputPng
$outFull = [System.IO.Path]::GetFullPath($OutputIco)

Write-Host "Converting PNG → ICO" -ForegroundColor Cyan
Write-Host "  Input : $inputFull"
Write-Host "  Output: $outFull"
Write-Host "  Sizes : $($Sizes -join ', ')"

$images = @()
foreach ($s in $Sizes) {
	$bytes = Get-PngBytesForSize -PngPath $inputFull -Size $s
	$images += @{ Size = $s; Data = $bytes }
}

Write-IcoFile -OutPath $outFull -Images $images

Write-Host "Done." -ForegroundColor Green