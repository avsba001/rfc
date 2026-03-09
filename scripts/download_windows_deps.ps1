$ErrorActionPreference = 'Stop'

Write-Host 'Installing LibreOffice (x86)...'
choco install libreoffice-fresh --x86 -y --no-progress

Write-Host 'Installing ImageMagick (x86)...'
choco install imagemagick --x86 -y --params "'/NoLegacy /NoPath'" --no-progress

$repoRoot = Resolve-Path "$PSScriptRoot/.."
$thirdParty = Join-Path $repoRoot 'third_party'
$loDst = Join-Path $thirdParty 'LibreOffice'
$imDst = Join-Path $thirdParty 'ImageMagick'

if (Test-Path $loDst) { Remove-Item $loDst -Recurse -Force }
if (Test-Path $imDst) { Remove-Item $imDst -Recurse -Force }

New-Item -ItemType Directory -Force -Path $loDst | Out-Null
New-Item -ItemType Directory -Force -Path $imDst | Out-Null

$loCandidates = @(
  'C:\Program Files (x86)\LibreOffice',
  'C:\Program Files\LibreOffice'
)
$loSrc = $loCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $loSrc) { throw 'LibreOffice install path not found.' }

$imCandidates = @(
  'C:\Program Files (x86)\ImageMagick*',
  'C:\Program Files\ImageMagick*'
)
$imSrc = Get-ChildItem -Directory $imCandidates -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $imSrc) { throw 'ImageMagick install path not found.' }

Write-Host "Copying LibreOffice from $loSrc"
Copy-Item -Path "$loSrc\*" -Destination $loDst -Recurse -Force

Write-Host "Copying ImageMagick from $($imSrc.FullName)"
Copy-Item -Path "$($imSrc.FullName)\*" -Destination $imDst -Recurse -Force
