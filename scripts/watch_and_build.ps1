# Watch workspace for changes to .py, .spec, .ico and trigger build_exe.ps1
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$workspace = Resolve-Path "$scriptDir\.."
$watchPath = $workspace.Path

Write-Output "Watching: $watchPath"

$fsw = New-Object System.IO.FileSystemWatcher $watchPath, "*.*"
$fsw.IncludeSubdirectories = $true
$fsw.EnableRaisingEvents = $true

# Debounce timer
$timer = New-Object System.Timers.Timer 500
$timer.AutoReset = $false
$timer.add_Elapsed({
    try {
        Write-Output "Change settled; running build..."
        & "$scriptDir\build_exe.ps1"
    } catch {
        Write-Error $_
    }
})

$action = {
    $path = $Event.SourceEventArgs.FullPath
    if ($path -match '\.py$' -or $path -match '\.spec$' -or $path -match '\.ico$') {
        # Restart debounce timer
        $timer.Stop()
        $timer.Start()
    }
}

Register-ObjectEvent $fsw Changed -Action $action | Out-Null
Register-ObjectEvent $fsw Created -Action $action | Out-Null
Register-ObjectEvent $fsw Renamed -Action $action | Out-Null
Register-ObjectEvent $fsw Deleted -Action $action | Out-Null

Write-Output "Press Enter to stop watching."
[Console]::ReadLine() | Out-Null

# Cleanup
Unregister-Event -SourceIdentifier * -ErrorAction SilentlyContinue
$fsw.EnableRaisingEvents = $false
$fsw.Dispose()
$timer.Stop()
$timer.Dispose()
Write-Output "Watcher stopped."