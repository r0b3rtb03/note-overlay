Param(
    [switch]$Clean
)

$spec = "note_assistant.spec"
$dist = "."
$work = "build_pyinstaller_temp"

if ($Clean) {
    Write-Output "Cleaning previous build artifacts..."
    Remove-Item -Recurse -Force "$dist\Service Host*" -ErrorAction SilentlyContinue
    Remove-Item -Recurse -Force $work -ErrorAction SilentlyContinue
}

Write-Output "Building EXE from $spec..."
python -m PyInstaller --distpath $dist --workpath $work --noconfirm $spec

if ($LASTEXITCODE -eq 0) {
    Write-Output "Build succeeded: $dist"
} else {
    Write-Error "Build failed with exit code $LASTEXITCODE"
    exit $LASTEXITCODE
}
