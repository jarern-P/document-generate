$ErrorActionPreference = "Stop"

Write-Host "========================================"
Write-Host " Document Generator - Build App"
Write-Host "========================================"

# ----------------------------------------
# Project paths
# ----------------------------------------
$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$TargetDir = Join-Path $ProjectRoot "target"
$DistDir = Join-Path $ProjectRoot "dist"

$JarName = "document-generator.jar"
$JarPath = Join-Path $TargetDir $JarName

$AppName = "DocumentGenerator"
$AppVersion = "1.0.0"

# ----------------------------------------
# Check Java
# ----------------------------------------
Write-Host ""
Write-Host "[1/4] Checking Java 21..."

$javaExe = Get-Command java -ErrorAction SilentlyContinue

if (-not $javaExe) {
    Write-Error "Java was not found in PATH."
}

Write-Host "Java found: $($javaExe.Source)"

# Get Java version without PowerShell treating stderr as an error
$javaVersion = & cmd.exe /c "`"$($javaExe.Source)`" -version 2>&1"

$javaVersion | ForEach-Object {
    Write-Host $_
}
# ----------------------------------------
# Check jpackage
# ----------------------------------------
Write-Host ""
Write-Host "[2/4] Checking jpackage..."

$jpackage = Get-Command jpackage -ErrorAction SilentlyContinue

if (-not $jpackage) {
    Write-Error "jpackage was not found. Please install JDK 21 and make sure JDK\bin is in PATH."
}

Write-Host "jpackage found: $($jpackage.Source)"

# ----------------------------------------
# Build JAR
# ----------------------------------------
Write-Host ""
Write-Host "[3/4] Building JAR..."

Push-Location $ProjectRoot

try {
    & ".\mvnw.cmd" clean package

    if ($LASTEXITCODE -ne 0) {
        throw "Maven build failed."
    }
}
finally {
    Pop-Location
}

if (-not (Test-Path $JarPath)) {
    Write-Error "JAR was not found: $JarPath"
}

Write-Host "JAR created:"
Write-Host $JarPath

# ----------------------------------------
# Prepare dist
# ----------------------------------------
Write-Host ""
Write-Host "[4/4] Creating Windows application..."

if (Test-Path $DistDir) {
    Remove-Item $DistDir -Recurse -Force
}

New-Item -ItemType Directory -Path $DistDir -Force | Out-Null

# ----------------------------------------
# Run jpackage
# ----------------------------------------
# ----------------------------------------
# Run jpackage
# ----------------------------------------
& jpackage `
    --type app-image `
    --name $AppName `
    --app-version $AppVersion `
    --input $TargetDir `
    --main-jar $JarName `
    --main-class "docgen.Main" `
    --dest $DistDir

if ($LASTEXITCODE -ne 0) {
    throw "jpackage failed."
}

Write-Host ""
Write-Host "========================================"
Write-Host " BUILD SUCCESS"
Write-Host "========================================"
Write-Host ""

Get-ChildItem $DistDir

Write-Host ""
Write-Host "Installer created in:"
Write-Host $DistDir