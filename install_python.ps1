# install_python.ps1
# Script to download and install Python silently for the current user, and install dependencies.

$ErrorActionPreference = "Stop"

$pythonDir = "$env:USERPROFILE\AppData\Local\Programs\Python\Python310"
$pythonExe = "$pythonDir\python.exe"

# 1. Download Python 3.10.11 Installer if not already installed
if (-not (Test-Path $pythonExe)) {
    Write-Host "Python not found at $pythonExe. Downloading installer..."
    $url = "https://www.python.org/ftp/python/3.10.11/python-3.10.11-amd64.exe"
    $installerPath = "$env:TEMP\python-3.10.11.exe"
    
    # Check if installer already exists in TEMP to save download time
    if (-not (Test-Path $installerPath)) {
        Write-Host "Downloading from $url to $installerPath..."
        Invoke-WebRequest -Uri $url -OutFile $installerPath
    }
    
    Write-Host "Running Python installer silently..."
    # /quiet: silent installation
    # InstallAllUsers=0: user-only install (does not require admin rights)
    # PrependPath=1: add to user environment variable PATH
    # TargetDir: specify target directory
    $arguments = "/quiet InstallAllUsers=0 PrependPath=1 TargetDir=$pythonDir"
    $process = Start-Process -FilePath $installerPath -ArgumentList $arguments -Wait -PassThru
    
    if ($process.ExitCode -ne 0) {
        Write-Error "Python installation failed with exit code $($process.ExitCode)"
    }
    Write-Host "Python installation completed successfully."
} else {
    Write-Host "Python is already installed at $pythonExe."
}

# 2. Add Python to current process path to make it available immediately in this run
$pathsToInject = @(
    $pythonDir,
    "$pythonDir\Scripts"
)

foreach ($path in $pathsToInject) {
    if ($env:PATH -notlike "*$path*") {
        $env:PATH = "$path;$env:PATH"
        Write-Host "Added $path to current process PATH."
    }
}

# 3. Verify Python works in current session
Write-Host "Verifying Python version..."
& python --version

# 4. Install required libraries
Write-Host "Upgrading pip..."
& python -m pip install --upgrade pip --quiet

Write-Host "Installing pandas, numpy, matplotlib, seaborn, python-docx, python-pptx, and scikit-learn..."
& pip install pandas numpy matplotlib seaborn python-docx python-pptx scikit-learn --quiet

Write-Host "All libraries installed successfully."
