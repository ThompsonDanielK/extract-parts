# Extract-Parts.ps1
# Windows PowerShell 5.1 compatible.
# Produced and iteratively refined with GitHub Copilot.
# Features:
# - Up to $MaxParallel RAR archive sets are extracted concurrently using Start-Job.
# - Flatten extracted files into a user-selected destination folder.
# - Convert extracted and previously-existing .nsp/.xci files to .nsz/.xcz via NSZ.
# - Ask before installing missing dependencies (WinRAR via winget; NSZ via its official PyPI package in a virtual environment).
# - Ask the user to select an existing legal prod.keys/keys.txt file if none is found.
# - Show numbered, separate waiting-queue rows in the same PowerShell progress display.
# - Remove source RAR volume sets only when extraction, conversion (when applicable), and destination copy succeed.
# - Optionally delete standalone .nsp/.xci originals only after a verified output has been copied successfully.
# - Timestamped diagnostic log, transcript, and per-archive WinRAR logs for troubleshooting.
# - Work files are staged in a hidden folder inside the output folder, so finished files are moved into place
#   by rename (no second full copy) and large extractions don't fill the system drive's %TEMP%.
#
# Usage examples:
#   .\Extract-Parts.ps1
#   .\Extract-Parts.ps1 -SourceFolder "C:\Games\Roms" -OutputFolder "C:\Games\Output" -DeleteArchives -DeleteStandalone
#   .\Extract-Parts.ps1 -NonInteractive -SourceFolder "D:\Roms" -DryRun
#   .\Extract-Parts.ps1 -NonInteractive -SourceFolder "D:\Roms" -DeleteArchives
#   .\Extract-Parts.ps1 -SourceFolder "D:\Roms" -MaxParallel 2
#
# Behavior:
#   - If command-line flags are provided, they take precedence.
#   - Otherwise, the script asks interactively to confirm deletions.
#   - -NonInteractive never shows dialogs or prompts: -SourceFolder is required, -OutputFolder defaults to
#     <source>\Extracted, every yes/no prompt is answered "no", and duplicate filenames are skipped.
#     Add -Force to auto-approve dependency installs (never deletions). Exit code is 1 if anything failed.
#   - -DryRun only reports what would be processed; nothing is extracted, converted, installed, or deleted.
#   - -NoDelete processes everything but never deletes source files.
#   - -MaxParallel sets the number of concurrent WinRAR jobs (default 6; lower it for spinning hard drives).

# --------------------------- Configuration ---------------------------
param(
    [string]$SourceFolder = "",
    [string]$OutputFolder,
    [switch]$DeleteArchives,
    [switch]$DeleteStandalone,
    [switch]$DryRun,
    [switch]$NoDelete,
    [switch]$NonInteractive,
    [switch]$Force,
    [ValidateRange(1, 32)][int]$MaxParallel = 6
)

$QueuePreviewLimit = 12
$DeleteArchivesOnSuccess = $null
$DefaultOutputFolder = if ([string]::IsNullOrWhiteSpace($SourceFolder)) { Join-Path $env:USERPROFILE 'Extracted' } else { Join-Path $SourceFolder 'Extracted' }
$NszInstallFolder = Join-Path $env:LOCALAPPDATA "ExtractParts\tools"
$NszExecutableName = "nsz-cli-windows-x64.exe"
$NszPythonVenvFolder = Join-Path $NszInstallFolder "venv"

# ------------------------------ Setup -------------------------------
$ErrorActionPreference = 'Stop'
# Note: $script:OutputFolder is the -OutputFolder parameter itself (script-scope params share the scope), so it must not be reset here.
$script:RarExe = $null
$script:NszExe = $null
$script:KeysFile = $null
$script:CompressionEnabled = $false
$script:DeleteStandaloneSources = $false
$script:DeleteArchivesConfirmed = $false
$script:DryRun = [bool]$DryRun
$script:NoDelete = [bool]$NoDelete
$script:NonInteractive = [bool]$NonInteractive
$script:Force = [bool]$Force

# Deletion preferences are resolved after the scan, in the "Deletion preferences" section.
$script:StopProcessing = $false
$script:SuccessActions = @('Copy', 'Overwrite', 'KeepBoth')
$script:StagingRoot = $null
$script:LastActiveScanSummary = ''
$script:OverwriteAll = $false
$script:SkipAll = $false
$script:KeepBothAll = $false
$script:PreviousQueueProgressIds = @()
$script:LastProgressScan = [datetime]::MinValue
$script:CurrentWork = 'Preparing'
$script:ProgressStage = 'Archives'
$script:Active = @()
$script:NextArchiveIndex = 0
$script:CompletedArchives = 0
$script:CopiedCount = 0
$script:SkippedCount = 0
$script:FailedCopyCount = 0
$script:ErrorCount = 0
$script:WarningCount = 0
$script:DeletedArchiveCount = 0
$script:ConversionSuccessCount = 0
$script:ConversionErrorCount = 0
$script:StandaloneProcessedCount = 0
$script:StandaloneQueue = @()
$script:StandaloneNextIndex = 0
$script:TotalArchives = 0
$script:TotalStandalone = 0
$script:Archives = @()
$script:RunLogFolder = Join-Path $env:LOCALAPPDATA 'ExtractParts\Logs'
$script:RunId = Get-Date -Format 'yyyyMMdd_HHmmss'
try {
    New-Item -ItemType Directory -Path $script:RunLogFolder -Force -ErrorAction Stop | Out-Null
} catch {
    $script:RunLogFolder = Join-Path $env:TEMP 'ExtractParts_Logs'
    New-Item -ItemType Directory -Path $script:RunLogFolder -Force -ErrorAction SilentlyContinue | Out-Null
}
$script:DiagnosticsLogPath = Join-Path $script:RunLogFolder ("ExtractParts_{0}.log" -f $script:RunId)
$script:TranscriptPath = Join-Path $script:RunLogFolder ("ExtractParts_{0}_transcript.txt" -f $script:RunId)
$script:TranscriptStarted = $false
$script:WorkerLogCounter = 0
try {
    Start-Transcript -Path $script:TranscriptPath -Force | Out-Null
    $script:TranscriptStarted = $true
} catch {
    # Diagnostics logging remains available if transcripts are unsupported in this host.
}

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('DEBUG','INFO','WARN','ERROR')][string]$Level = 'INFO'
    )
    $line = '[{0}][{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'), $Level, $Message
    try { Add-Content -LiteralPath $script:DiagnosticsLogPath -Value $line -Encoding UTF8 -ErrorAction Stop } catch {}
}

function Stop-RunLogging {
    if ($script:TranscriptStarted) {
        try { Stop-Transcript | Out-Null } catch {}
        $script:TranscriptStarted = $false
    }
}

Write-Log -Message "Run started. PowerShell=$($PSVersionTable.PSVersion); Source=$SourceFolder; Configured parallelism=$MaxParallel" -Level INFO
Write-Log -Message "Diagnostic log: $script:DiagnosticsLogPath; transcript: $script:TranscriptPath" -Level INFO

function Test-DeleteAllowed {
    if ($script:DryRun -or $script:NoDelete) {
        return $false
    }
    return $true
}

function Test-DeleteCandidate {
    param(
        [Parameter(Mandatory)][string]$SourcePath,
        [string[]]$RequiredOutputs = @()
    )

    if ([string]::IsNullOrWhiteSpace($SourcePath)) { return $false }
    if (-not (Test-Path -LiteralPath $SourcePath -PathType Leaf)) { return $false }

    $sourceItem = Get-Item -LiteralPath $SourcePath -ErrorAction SilentlyContinue
    if ($null -eq $sourceItem -or $sourceItem.Length -le 0) { return $false }

    foreach ($outputPath in $RequiredOutputs) {
        if ([string]::IsNullOrWhiteSpace($outputPath)) { continue }
        if (-not (Test-Path -LiteralPath $outputPath -PathType Leaf)) { return $false }
        $outputItem = Get-Item -LiteralPath $outputPath -ErrorAction SilentlyContinue
        if ($null -eq $outputItem -or $outputItem.Length -le 0) { return $false }
    }

    return $true
}

function Read-YesNo {
    param(
        [Parameter(Mandatory)][string]$Prompt,
        [bool]$DefaultNo = $true,
        # Non-destructive prompts (e.g. dependency installs) may be auto-approved by -NonInteractive -Force.
        [switch]$AllowForce
    )
    if ($script:NonInteractive) {
        $autoAnswer = [bool]($script:Force -and $AllowForce)
        Write-Host "Non-interactive mode: answering '$(if ($autoAnswer) { 'yes' } else { 'no' })' to: $Prompt" -ForegroundColor Yellow
        Write-Log -Message "Non-interactive prompt auto-answered ($autoAnswer): '$Prompt'" -Level WARN
        return $autoAnswer
    }

    $suffix = if ($DefaultNo) { '[y/N]' } else { '[Y/n]' }
    while ($true) {
        $answer = [string](Read-Host "$Prompt $suffix")
        $answer = $answer.Trim()
        if ([string]::IsNullOrWhiteSpace($answer)) { return (-not $DefaultNo) }
        if ($answer -match '^(?i:y|yes)$') { return $true }
        if ($answer -match '^(?i:n|no)$') { return $false }
        Write-Host 'Please answer y or n.' -ForegroundColor Yellow
    }
}

function Wait-ForExit {
    if (-not $script:NonInteractive) { [void](Read-Host 'Press Enter to exit') }
}

function Resolve-RarExecutable {
    $candidates = @()
    if (-not [string]::IsNullOrWhiteSpace($env:ProgramFiles)) { $candidates += (Join-Path $env:ProgramFiles 'WinRAR\Rar.exe') }
    $programFilesX86 = [Environment]::GetEnvironmentVariable('ProgramFiles(x86)')
    if (-not [string]::IsNullOrWhiteSpace($programFilesX86)) { $candidates += (Join-Path $programFilesX86 'WinRAR\Rar.exe') }
    if (-not [string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) { $candidates += (Join-Path $env:LOCALAPPDATA 'Programs\WinRAR\Rar.exe') }
    foreach ($candidate in $candidates) {
        if (-not [string]::IsNullOrWhiteSpace($candidate) -and (Test-Path -LiteralPath $candidate -PathType Leaf)) {
            return $candidate
        }
    }
    $command = Get-Command 'Rar.exe' -ErrorAction SilentlyContinue
    if ($command -and $command.Source) { return $command.Source }
    return $null
}

function Install-WinRAR {
    Write-Host "WinRAR's command-line executable is required to extract the RAR archives." -ForegroundColor Yellow
    Write-Host "The script can try to install WinRAR using the Windows package manager (winget)." -ForegroundColor Yellow
    if (-not (Read-YesNo -Prompt 'Install WinRAR now using winget?' -AllowForce)) { Write-Log -Message 'User declined WinRAR installation.' -Level WARN; return $null }

    $winget = Get-Command 'winget.exe' -ErrorAction SilentlyContinue
    if (-not $winget) { $winget = Get-Command 'winget' -ErrorAction SilentlyContinue }
    if (-not $winget -or -not $winget.Source) {
        Write-Host "winget is unavailable. Install Microsoft's App Installer (which provides winget), or install WinRAR manually from https://www.win-rar.com/, then run this script again." -ForegroundColor Red
        return $null
    }

    try {
        Write-Host 'Installing WinRAR with winget...' -ForegroundColor Cyan
        Write-Log -Message 'Starting winget install for RARLab.WinRAR.' -Level INFO
        $wingetArguments = @('install', '--exact', '--id', 'RARLab.WinRAR', '--accept-package-agreements', '--accept-source-agreements')
        & $winget.Source @wingetArguments
        $wingetExitCode = $LASTEXITCODE
        Write-Log -Message "WinRAR winget install exit code: $wingetExitCode" -Level INFO
        if ($wingetExitCode -ne 0) {
            Write-Host "winget returned exit code $wingetExitCode. WinRAR may not have been installed." -ForegroundColor Red
        }
    } catch {
        Write-Host "WinRAR installation failed: $($_.Exception.Message)" -ForegroundColor Red
        Write-Log -Message "WinRAR installation exception: $($_ | Out-String)" -Level ERROR
    }

    $rar = Resolve-RarExecutable
    if (-not $rar) {
        Write-Host "WinRAR's Rar.exe was not detected after the installation attempt. Archive extraction will be skipped." -ForegroundColor Red
        Write-Log -Message 'WinRAR executable was not detected after installation attempt.' -Level ERROR
    } else { Write-Log -Message "WinRAR detected after install attempt: '$rar'." -Level INFO }
    return $rar
}

function Test-NszExecutable {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $false }
    $oldPreference = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    $outputLines = @()
    $code = 999
    try {
        $outputLines = @(& $Path --help 2>&1 | ForEach-Object { [string]$_ })
        $code = $LASTEXITCODE
    } catch {
        $outputLines += $_.Exception.Message
    } finally {
        $ErrorActionPreference = $oldPreference
    }
    if ($code -eq 0) {
        Write-Log -Message "NSZ executable self-test succeeded: '$Path'" -Level INFO
        return $true
    }
    $detail = ($outputLines | Select-Object -Last 12) -join ' | '
    Write-Log -Message "NSZ executable self-test FAILED: path='$Path'; exit=$code; output=$detail" -Level ERROR
    return $false
}

function Resolve-NszExecutable {
    $candidates = @(
        (Join-Path $NszPythonVenvFolder 'Scripts\nsz.exe'),
        (Join-Path $NszInstallFolder $NszExecutableName),
        (Join-Path $NszInstallFolder 'nsz.exe')
    )
    foreach ($candidate in $candidates) {
        if ((Test-Path -LiteralPath $candidate -PathType Leaf) -and (Test-NszExecutable -Path $candidate)) {
            return $candidate
        }
    }

    foreach ($name in @('nsz.exe', 'nsz')) {
        $command = Get-Command $name -ErrorAction SilentlyContinue
        if ($command -and $command.Source -and (Test-NszExecutable -Path $command.Source)) { return $command.Source }
    }
    return $null
}

function Resolve-PythonExecutable {
    $candidates = @()
    foreach ($name in @('python.exe', 'python3.exe', 'python', 'python3')) {
        $command = Get-Command $name -ErrorAction SilentlyContinue
        if ($command -and $command.Source) { $candidates += $command.Source }
    }
    foreach ($candidate in @(
        (Join-Path $env:LOCALAPPDATA 'Programs\Python\Python312\python.exe'),
        (Join-Path $env:LOCALAPPDATA 'Programs\Python\Python311\python.exe'),
        (Join-Path $env:LOCALAPPDATA 'Programs\Python\Python310\python.exe'),
        (Join-Path $env:ProgramFiles 'Python312\python.exe'),
        (Join-Path $env:ProgramFiles 'Python311\python.exe'),
        (Join-Path $env:ProgramFiles 'Python310\python.exe')
    )) { if ($candidate -and (Test-Path -LiteralPath $candidate -PathType Leaf)) { $candidates += $candidate } }

    $seen = @{}
    foreach ($candidate in $candidates) {
        $key = ([string]$candidate).ToLowerInvariant()
        if ($seen.ContainsKey($key)) { continue }
        $seen[$key] = $true
        try {
            $oldPreference = $ErrorActionPreference
            $ErrorActionPreference = 'Continue'
            $versionOutput = @(& $candidate --version 2>&1 | ForEach-Object { [string]$_ })
            $code = $LASTEXITCODE
            $ErrorActionPreference = $oldPreference
            $versionText = $versionOutput -join ' '
            if ($code -eq 0 -and $versionText -match 'Python\s+(?<major>\d+)\.(?<minor>\d+)') {
                $major = [int]$Matches['major']; $minor = [int]$Matches['minor']
                if ($major -gt 3 -or ($major -eq 3 -and $minor -ge 8)) {
                    Write-Log -Message "Found usable Python: '$candidate' ($versionText)" -Level INFO
                    return $candidate
                }
                Write-Log -Message "Python is too old for this setup: '$candidate' ($versionText); requiring Python 3.8+" -Level WARN
            } else {
                Write-Log -Message "Python candidate failed version test: '$candidate'; exit=$code; output=$versionText" -Level WARN
            }
        } catch {
            $ErrorActionPreference = $oldPreference
            Write-Log -Message "Python candidate test failed: '$candidate'; error=$($_.Exception.Message)" -Level WARN
        }
    }

    # Use the py launcher if installed, without assuming its default version.
    $launcher = Get-Command 'py.exe' -ErrorAction SilentlyContinue
    if (-not $launcher) { $launcher = Get-Command 'py' -ErrorAction SilentlyContinue }
    if ($launcher -and $launcher.Source) {
        try {
            $oldPreference = $ErrorActionPreference; $ErrorActionPreference = 'Continue'
            $resolved = @(& $launcher.Source -3 -c 'import sys; print(sys.executable); print(sys.version_info[0]); print(sys.version_info[1])' 2>&1)
            $code = $LASTEXITCODE
            $ErrorActionPreference = $oldPreference
            if ($code -eq 0 -and $resolved.Count -ge 3) {
                $pyPath = [string]$resolved[0]
                $major = [int]$resolved[1]; $minor = [int]$resolved[2]
                if ((Test-Path -LiteralPath $pyPath -PathType Leaf) -and ($major -gt 3 -or ($major -eq 3 -and $minor -ge 8))) {
                    Write-Log -Message "Resolved Python through py launcher: '$pyPath' ($major.$minor)" -Level INFO
                    return $pyPath
                }
            }
        } catch {
            $ErrorActionPreference = $oldPreference
            Write-Log -Message "Python launcher detection failed: $($_.Exception.Message)" -Level WARN
        }
    }
    return $null
}

function Install-Nsz {
    Write-Host "NSZ is needed to convert .nsp to .nsz and .xci to .xcz." -ForegroundColor Yellow
    Write-Host "The standalone NSZ executable appears to be unreliable on some Windows systems. This fallback installs the official NSZ Python package in an isolated virtual environment." -ForegroundColor Yellow
    Write-Host "Installation folder: $NszPythonVenvFolder" -ForegroundColor Yellow
    if (-not (Read-YesNo -Prompt 'Install NSZ and its required Python packages now?' -AllowForce)) {
        Write-Log -Message 'User declined NSZ installation.' -Level WARN
        return $null
    }

    $pythonPath = Resolve-PythonExecutable
    if (-not $pythonPath) {
        Write-Host 'Python 3.8+ was not found. The script can install Python 3.12 using winget first.' -ForegroundColor Yellow
        if (-not (Read-YesNo -Prompt 'Install Python 3.12 using winget?' -AllowForce)) {
            Write-Log -Message 'Python installation was declined; NSZ will be unavailable.' -Level WARN
            return $null
        }
        $winget = Get-Command 'winget.exe' -ErrorAction SilentlyContinue
        if (-not $winget) { $winget = Get-Command 'winget' -ErrorAction SilentlyContinue }
        if (-not $winget -or -not $winget.Source) {
            Write-Host 'winget is unavailable. Install Python 3.12 manually from https://www.python.org/downloads/windows/ and rerun this script.' -ForegroundColor Red
            Write-Log -Message 'Python install failed: winget not available.' -Level ERROR
            return $null
        }
        try {
            Write-Host 'Installing Python 3.12 with winget...' -ForegroundColor Cyan
            Write-Log -Message 'Starting winget install for Python.Python.3.12.' -Level INFO
            & $winget.Source install --exact --id Python.Python.3.12 --accept-package-agreements --accept-source-agreements
            $pythonInstallCode = $LASTEXITCODE
            Write-Log -Message "winget Python install exit code: $pythonInstallCode" -Level INFO
        } catch {
            Write-Log -Message "winget Python installation exception: $($_.Exception.Message)" -Level ERROR
        }
        # A newly installed interpreter may not yet be present in this session's PATH.
        $pythonPath = Resolve-PythonExecutable
        if (-not $pythonPath) {
            foreach ($guess in @(
                (Join-Path $env:LOCALAPPDATA 'Programs\Python\Python312\python.exe'),
                (Join-Path $env:ProgramFiles 'Python312\python.exe')
            )) {
                if ($guess -and (Test-Path -LiteralPath $guess -PathType Leaf)) { $pythonPath = $guess; break }
            }
        }
        if (-not $pythonPath) {
            Write-Host 'Python was not detected after installation. Restart PowerShell and rerun the script, or install Python manually.' -ForegroundColor Red
            Write-Log -Message 'Python still not found after winget installation.' -Level ERROR
            return $null
        }
    }

    $oldInstallPreference = $ErrorActionPreference
    try {
        $ErrorActionPreference = 'Continue'
        New-Item -ItemType Directory -Path $NszInstallFolder -Force -ErrorAction Stop | Out-Null
        $venvPython = Join-Path $NszPythonVenvFolder 'Scripts\python.exe'
        if (-not (Test-Path -LiteralPath $venvPython -PathType Leaf)) {
            Write-Host 'Creating an isolated Python environment for NSZ...' -ForegroundColor Cyan
            Write-Log -Message "Creating virtual environment with '$pythonPath' at '$NszPythonVenvFolder'." -Level INFO
            & $pythonPath -m venv $NszPythonVenvFolder 2>&1 | Tee-Object -FilePath (Join-Path $script:RunLogFolder ("PythonVenv_{0}.log" -f $script:RunId)) -Append
            if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $venvPython -PathType Leaf)) {
                throw "Python failed to create the virtual environment (exit code $LASTEXITCODE). See the diagnostic log folder: $script:RunLogFolder"
            }
        }
        Write-Host 'Installing/updating the official NSZ package and dependencies (pycryptodome, zstandard, enlighten)...' -ForegroundColor Cyan
        Write-Log -Message "Updating pip in venv '$NszPythonVenvFolder'." -Level INFO
        & $venvPython -m pip install --upgrade pip 2>&1 | Tee-Object -FilePath (Join-Path $script:RunLogFolder ("NszInstall_{0}.log" -f $script:RunId)) -Append
        if ($LASTEXITCODE -ne 0) { throw "pip upgrade failed with exit code $LASTEXITCODE." }

        Write-Log -Message 'Installing/updating official PyPI package: nsz.' -Level INFO
        & $venvPython -m pip install --upgrade nsz 2>&1 | Tee-Object -FilePath (Join-Path $script:RunLogFolder ("NszInstall_{0}.log" -f $script:RunId)) -Append
        if ($LASTEXITCODE -ne 0) { throw "Installing the nsz package failed with exit code $LASTEXITCODE." }

        $venvNsz = [string](Join-Path $NszPythonVenvFolder 'Scripts\nsz.exe')
        if ([string]::IsNullOrWhiteSpace($venvNsz)) { throw 'NSZ installation produced an empty executable path.' }
        if (-not (Test-Path -LiteralPath $venvNsz -PathType Leaf)) { throw "The pip installation completed but nsz.exe was not found at '$venvNsz'." }
        if (-not (Test-NszExecutable -Path $venvNsz)) { throw 'The installed NSZ Python package failed its --help startup test.' }
        Write-Host "NSZ and its dependencies are installed: $venvNsz" -ForegroundColor Green
        Write-Log -Message "NSZ Python package installed successfully at '$venvNsz'." -Level INFO
        return $venvNsz.Trim()
    } catch {
        Write-Host "NSZ installation failed: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "See installation details under: $script:RunLogFolder" -ForegroundColor Yellow
        Write-Log -Message "NSZ installation failed: $($_ | Out-String)" -Level ERROR
        return $null
    } finally {
        $ErrorActionPreference = $oldInstallPreference
    }
}

function Find-KeysFile {
    $nszPath = $script:NszExe
    if ([string]::IsNullOrWhiteSpace([string]$nszPath)) {
        $nszDirectory = $NszInstallFolder
    } else {
        $nszDirectory = Split-Path -Parent $nszPath
        if ([string]::IsNullOrWhiteSpace([string]$nszDirectory)) {
            $nszDirectory = $NszInstallFolder
        }
    }

    $candidates = @(
        (Join-Path $HOME '.switch\prod.keys'),
        (Join-Path $HOME '.switch\keys.txt'),
        (Join-Path $NszInstallFolder 'prod.keys'),
        (Join-Path $NszInstallFolder 'keys.txt'),
        (Join-Path $nszDirectory 'prod.keys'),
        (Join-Path $nszDirectory 'keys.txt'),
        (Join-Path $PSScriptRoot 'prod.keys'),
        (Join-Path $PSScriptRoot 'keys.txt'),
        (Join-Path (Get-Location).Path 'prod.keys'),
        (Join-Path (Get-Location).Path 'keys.txt')
    )
    foreach ($candidate in $candidates) {
        if (-not [string]::IsNullOrWhiteSpace([string]$candidate) -and (Test-Path -LiteralPath $candidate -PathType Leaf)) {
            return $candidate
        }
    }
    return $null
}

function Select-KeysFile {
    Write-Host ''
    Write-Host 'NSZ requires a compatible keys file that you must obtain legally. The script will not download keys.' -ForegroundColor Yellow
    if ($script:NonInteractive) {
        Write-Host 'Non-interactive mode: no keys file found in the default locations, so compression will be skipped.' -ForegroundColor Yellow
        return $null
    }
    Write-Host 'Choose an existing prod.keys or keys.txt file in the file picker, or cancel to skip compression.' -ForegroundColor Yellow
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = 'Select your existing prod.keys or keys.txt file'
    $dialog.Filter = 'Switch keys files (prod.keys or keys.txt)|prod.keys;keys.txt|All files (*.*)|*.*'
    $dialog.CheckFileExists = $true
    $dialog.Multiselect = $false
    try {
        while ((Show-DialogOnTop -Dialog $dialog) -eq [System.Windows.Forms.DialogResult]::OK) {
            $chosen = $dialog.FileName
            $leaf = [System.IO.Path]::GetFileName($chosen)
            if (($leaf -ieq 'prod.keys' -or $leaf -ieq 'keys.txt') -and (Test-Path -LiteralPath $chosen -PathType Leaf)) {
                return $chosen
            }
            Write-Host 'The selected file must be named prod.keys or keys.txt. Choose again, or cancel to skip compression.' -ForegroundColor Red
        }
        return $null
    } finally {
        $dialog.Dispose()
    }
}

# Console-hosted dialogs often open behind the console window; a hidden TopMost owner brings them to the front.
function Show-DialogOnTop {
    param([Parameter(Mandatory)][System.Windows.Forms.CommonDialog]$Dialog)
    $owner = New-Object System.Windows.Forms.Form
    try {
        $owner.TopMost = $true
        $owner.ShowInTaskbar = $false
        $owner.StartPosition = 'CenterScreen'
        $owner.Size = New-Object System.Drawing.Size(1, 1)
        $owner.Opacity = 0
        $owner.Show()
        $owner.Activate()
        return $Dialog.ShowDialog($owner)
    } finally {
        $owner.Dispose()
    }
}

function Select-Folder {
    param(
        [Parameter(Mandatory)][string]$Description,
        [Parameter(Mandatory)][string]$InitialFolder
    )
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = $Description
    $dialog.SelectedPath = $InitialFolder
    $dialog.ShowNewFolderButton = $true
    try {
        $result = Show-DialogOnTop -Dialog $dialog
        if ($result -eq [System.Windows.Forms.DialogResult]::OK -and -not [string]::IsNullOrWhiteSpace($dialog.SelectedPath)) {
            return $dialog.SelectedPath
        }
        return $null
    } finally {
        $dialog.Dispose()
    }
}

function Get-ArchiveStarts {
    param([Parameter(Mandatory)][string]$Folder)
    @(
        Get-ChildItem -LiteralPath $Folder -File -ErrorAction SilentlyContinue |
            Where-Object {
                ($_.Name -match '(?i)\.part0*1\.rar$') -or
                (($_.Extension -ieq '.rar') -and ($_.Name -notmatch '(?i)\.part\d+\.rar$'))
            } |
            Sort-Object Name
    )
}

function Get-LooseGameFiles {
    param([string[]]$Folders)
    $seen = @{}
    $files = @()
    foreach ($folder in $Folders) {
        if ([string]::IsNullOrWhiteSpace($folder) -or -not (Test-Path -LiteralPath $folder -PathType Container)) { continue }
        foreach ($file in (Get-ChildItem -LiteralPath $folder -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -in @('.nsp', '.xci') })) {
            $key = $file.FullName.ToLowerInvariant()
            if (-not $seen.ContainsKey($key)) {
                $seen[$key] = $true
                $files += $file
            }
        }
    }
    @($files | Sort-Object FullName)
}

function Get-EligibleFiles {
    param([Parameter(Mandatory)][string]$Folder)
    if (-not (Test-Path -LiteralPath $Folder -PathType Container)) { return @() }
    @(
        Get-ChildItem -LiteralPath $Folder -File -Recurse -Force -ErrorAction SilentlyContinue |
            Where-Object {
                $_.Extension -ine '.url' -and
                $_.FullName -notmatch '(?i)(^|[\\/])Internet\s*Shortcuts?([\\/]|$)'
            }
    )
}

function Get-ArchiveFilesToDelete {
    param([Parameter(Mandatory)][System.IO.FileInfo]$Archive)

    $allFiles = @(Get-ChildItem -LiteralPath $Archive.DirectoryName -File -ErrorAction SilentlyContinue)
    $partOne = [regex]::Match($Archive.Name, '^(?<base>.+)\.part0*1\.rar$', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if ($partOne.Success) {
        $base = $partOne.Groups['base'].Value
        $pattern = '^' + [regex]::Escape($base) + '\.part\d+\.rar$'
        return @($allFiles | Where-Object { $_.Name -match $pattern } | Sort-Object Name)
    }

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($Archive.Name)
    $oldVolumePattern = '^' + [regex]::Escape($baseName) + '\.(?:r|s|t|u|v|w|x|y|z)\d{2}$'
    $matching = @($allFiles | Where-Object { $_.Name -ieq $Archive.Name -or $_.Name -match $oldVolumePattern })
    if ($matching.Count -gt 0) { return $matching }
    return @($Archive)
}

function Get-UniqueDestinationPath {
    param([Parameter(Mandatory)][string]$Directory, [Parameter(Mandatory)][string]$FileName)
    $candidate = Join-Path $Directory $FileName
    if (-not (Test-Path -LiteralPath $candidate)) { return $candidate }
    $base = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    $ext = [System.IO.Path]::GetExtension($FileName)
    $n = 2
    do {
        $candidate = Join-Path $Directory ("{0} ({1}){2}" -f $base, $n, $ext)
        $n++
    } while (Test-Path -LiteralPath $candidate)
    return $candidate
}

# Moves a staged (temporary) file into the flat output folder. Sources are always throwaway files in the
# staging area, so a move is used: on the same volume it is an instant rename instead of a full data copy.
# Returns an object with Action ('Copy','Overwrite','KeepBoth','Skipped','Failed','Quit') and the actual Destination.
function Move-FlatFile {
    param([Parameter(Mandatory)][System.IO.FileInfo]$File)
    if ($script:StopProcessing) { return [pscustomobject]@{ Action = 'Quit'; Destination = $null } }

    $destination = Join-Path $script:OutputFolder $File.Name
    $sourceLength = $File.Length
    Write-Log -Message "Destination handling started: source='$($File.FullName)'; target='$destination'; bytes=$sourceLength." -Level DEBUG
    $action = 'Copy'
    if (Test-Path -LiteralPath $destination) {
        if ($script:OverwriteAll) {
            $action = 'Overwrite'
        } elseif ($script:SkipAll -or $script:NonInteractive) {
            $script:SkippedCount++
            Write-Host "Skipping duplicate: $($File.Name)" -ForegroundColor DarkYellow
            Write-Log -Message "Duplicate skipped (SkipAll or non-interactive): '$destination'." -Level WARN
            return [pscustomobject]@{ Action = 'Skipped'; Destination = $destination }
        } elseif ($script:KeepBothAll) {
            $action = 'KeepBoth'
        } else {
            Write-Host "`nDuplicate filename detected:" -ForegroundColor Yellow
            Write-Host "  Existing: $destination"
            Write-Host "  Incoming: $($File.Name) ($('{0:N1}' -f ($sourceLength / 1MB)) MB)"
            $choice = $null
            while ($choice -notin @('O', 'S', 'B', 'A', 'K', 'L', 'Q')) {
                Write-Host 'Choose: [O]verwrite, [S]kip, [B] keep both, [A] overwrite all, [K] skip all, [L] keep both all, [Q] stop processing'
                $choice = ([string](Read-Host 'Choice')).Trim().ToUpperInvariant()
            }
            switch ($choice) {
                'O' { $action = 'Overwrite' }
                'B' { $action = 'KeepBoth' }
                'A' { $script:OverwriteAll = $true; $action = 'Overwrite' }
                'L' { $script:KeepBothAll = $true; $action = 'KeepBoth' }
                'K' { $script:SkipAll = $true }
                'Q' { $script:StopProcessing = $true; Write-Log -Message "User selected Q during duplicate handling for '$($File.Name)'." -Level WARN; return [pscustomobject]@{ Action = 'Quit'; Destination = $null } }
            }
            if ($choice -in @('S', 'K')) {
                $script:SkippedCount++
                Write-Log -Message "Duplicate skipped by user (choice $choice): '$destination'." -Level WARN
                return [pscustomobject]@{ Action = 'Skipped'; Destination = $destination }
            }
        }
        if ($action -eq 'KeepBoth') { $destination = Get-UniqueDestinationPath -Directory $script:OutputFolder -FileName $File.Name }
    }

    try {
        if ($action -eq 'Overwrite') {
            Move-Item -LiteralPath $File.FullName -Destination $destination -Force -ErrorAction Stop
        } else {
            Move-Item -LiteralPath $File.FullName -Destination $destination -ErrorAction Stop
        }
        $written = Get-Item -LiteralPath $destination -ErrorAction Stop
        if ($written.Length -ne $sourceLength) {
            throw "Size mismatch after move: expected $sourceLength bytes, found $($written.Length) bytes."
        }
        $script:CopiedCount++
        Write-Log -Message "Move $action succeeded: '$($File.FullName)' -> '$destination' ($sourceLength bytes)." -Level INFO
        return [pscustomobject]@{ Action = $action; Destination = $destination }
    } catch {
        $script:FailedCopyCount++
        Write-Log -Message "Move failed: source='$($File.FullName)'; destination='$destination'; exception=$($_ | Out-String)" -Level ERROR
        Write-Host "Copy failed for '$($File.Name)': $($_.Exception.Message)" -ForegroundColor Red
        return [pscustomobject]@{ Action = 'Failed'; Destination = $null }
    }
}

# Invoked in a background job so the dashboard can keep refreshing while NSZ works.
$script:NszWorker = {
    param([string]$Executable, [string]$SourceFile, [string]$OutputDirectory, [string]$KeysPath)
    $ErrorActionPreference = 'Continue'
    $lines = @()
    $code = 999
    try {
        $arguments = @('--keys', $KeysPath, '--verify', '--keep', '--minimal-output', '--output', $OutputDirectory, '-C', $SourceFile)
        $lines += "Executable: $Executable"
        $lines += "Source: $SourceFile"
        $lines += "Output directory: $OutputDirectory"
        $lines += "Keys file: $KeysPath"
        $lines += "Arguments: $($arguments -join ' ')"
        $processOutput = @(& $Executable @arguments 2>&1 | ForEach-Object { [string]$_ })
        $code = $LASTEXITCODE
        $lines += $processOutput
        $lines += "NSZ exit code: $code"
    } catch {
        $lines += "Exception invoking NSZ: $($_.Exception.Message)"
        $lines += ($_ | Out-String)
    }
    [pscustomobject]@{ ExitCode = [int]$code; Output = (($lines | Select-Object -Last 30) -join [Environment]::NewLine).Trim(); FullOutput = ($lines -join [Environment]::NewLine) }
}

# Invoked in a background job to extract one archive set. The worker saves its complete output to a unique log.
$script:ExtractionWorker = {
    param([string]$Executable, [string]$ArchivePath, [string]$TempFolder, [string]$WorkerLogPath)
    $ErrorActionPreference = 'Continue'
    $captured = New-Object System.Collections.Generic.List[string]
    $exitCode = 999
    $startedAt = Get-Date
    function Add-WorkerLogLine {
        param([string]$Text)
        $line = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'), $Text
        [void]$captured.Add($line)
        try { [System.IO.File]::AppendAllText($WorkerLogPath, $line + [Environment]::NewLine, [System.Text.Encoding]::UTF8) } catch {}
    }
    try {
        Add-WorkerLogLine "Starting extraction worker. PowerShell=$($PSVersionTable.PSVersion)"
        Add-WorkerLogLine "Executable='$Executable'"
        Add-WorkerLogLine "Archive='$ArchivePath'"
        Add-WorkerLogLine "TempFolder='$TempFolder'"
        if (-not (Test-Path -LiteralPath $Executable -PathType Leaf)) { throw "WinRAR executable not found: $Executable" }
        if (-not (Test-Path -LiteralPath $ArchivePath -PathType Leaf)) { throw "Archive file not found: $ArchivePath" }
        if (-not (Test-Path -LiteralPath $TempFolder -PathType Container)) { throw "Temporary extraction folder not found: $TempFolder" }
        $targetPath = $TempFolder.TrimEnd('\') + '\'
        $commandDescription = "& `"$Executable`" x -y -p- `"$ArchivePath`" `"$targetPath`""
        Add-WorkerLogLine "Command=$commandDescription"
        Add-WorkerLogLine 'WinRAR output begins:'
        # Do not use -idq: capture normal command output for troubleshooting. The job keeps it off the interactive console.
        # -p- stops WinRAR from asking for a password, which would hang forever in a job with no console.
        $toolOutput = @(& $Executable x -y -p- $ArchivePath $targetPath 2>&1 | ForEach-Object { [string]$_ })
        $nativeCode = $LASTEXITCODE
        foreach ($line in $toolOutput) { Add-WorkerLogLine ([string]$line) }
        if ($null -eq $nativeCode) {
            $exitCode = 999
            Add-WorkerLogLine 'ERROR: PowerShell did not expose LASTEXITCODE after invoking WinRAR.'
        } else {
            $exitCode = [int]$nativeCode
        }
        Add-WorkerLogLine "WinRAR exit code=$exitCode"
        Add-WorkerLogLine ("Elapsed seconds={0:N2}" -f ((Get-Date) - $startedAt).TotalSeconds)
    } catch {
        $exitCode = 998
        Add-WorkerLogLine "EXCEPTION invoking extraction: $($_.Exception.Message)"
        Add-WorkerLogLine ($_ | Out-String)
        if ($_.ScriptStackTrace) { Add-WorkerLogLine "StackTrace: $($_.ScriptStackTrace)" }
    }
    $messageLines = @($captured | Select-Object -Last 25)
    [pscustomobject]@{
        ExitCode = [int]$exitCode
        Message = ($messageLines -join ' | ')
        LogPath = $WorkerLogPath
        Started = $startedAt.ToString('s')
        Finished = (Get-Date).ToString('s')
        OutputLineCount = $captured.Count
    }
}

# Temporary work folders live on the output folder's volume so finished files can be moved into place by rename
# instead of being copied again (and so large extractions don't fill a small system drive under %TEMP%).
function New-StagingFolder {
    param([Parameter(Mandatory)][string]$Prefix)
    if (-not $script:StagingRoot) {
        $script:StagingRoot = Join-Path $script:OutputFolder ('.ExtractParts_staging_{0}' -f $script:RunId)
    }
    if (-not (Test-Path -LiteralPath $script:StagingRoot -PathType Container)) {
        $root = New-Item -ItemType Directory -Path $script:StagingRoot -Force
        try { $root.Attributes = $root.Attributes -bor [System.IO.FileAttributes]::Hidden } catch {}
    }
    $folder = Join-Path $script:StagingRoot ('{0}_{1}' -f $Prefix, ([guid]::NewGuid().ToString('N')))
    New-Item -ItemType Directory -Path $folder -Force | Out-Null
    return $folder
}

function Remove-StagingRoot {
    if ($script:StagingRoot -and (Test-Path -LiteralPath $script:StagingRoot)) {
        Remove-Item -LiteralPath $script:StagingRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Get-RunningExtractionCount {
    @($script:Active | Where-Object { $_.Job.State -notin @('Completed', 'Failed', 'Stopped') }).Count
}

# Keeps up to $MaxParallel extraction jobs running. Safe to call while other work (e.g. NSZ) is in progress.
function Start-PendingExtractions {
    if ($script:StopProcessing) { return }
    while ((Get-RunningExtractionCount) -lt $MaxParallel -and $script:NextArchiveIndex -lt $script:Archives.Count) { Start-NextExtraction }
}

function Start-NextExtraction {
    if ($script:NextArchiveIndex -ge $script:Archives.Count) { return }
    $archive = $script:Archives[$script:NextArchiveIndex]
    $queueNumber = $script:NextArchiveIndex + 1
    $script:NextArchiveIndex++
    $tempFolder = $null
    $safeName = [regex]::Replace([System.IO.Path]::GetFileNameWithoutExtension($archive.Name), '[^A-Za-z0-9._-]', '_')
    $workerLogPath = Join-Path $script:RunLogFolder ("RAR_{0:D3}_{1}_{2}.log" -f $queueNumber, $safeName, ([guid]::NewGuid().ToString('N').Substring(0,8)))
    try {
        $tempFolder = New-StagingFolder -Prefix 'extract'
        Set-Content -LiteralPath $workerLogPath -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff')] RAR worker log for '$($archive.Name)'" -Encoding UTF8
        $job = Start-Job -ScriptBlock $script:ExtractionWorker -ArgumentList $script:RarExe, $archive.FullName, $tempFolder, $workerLogPath -ErrorAction Stop
        $script:Active += [pscustomobject]@{ Archive = $archive; TempFolder = $tempFolder; Job = $job; LogPath = $workerLogPath; QueueNumber = $queueNumber }
        Write-Host "Started extraction [$queueNumber/$($script:TotalArchives)]: $($archive.Name) (job $($job.Id))" -ForegroundColor Cyan
        Write-Log -Message "Started parallel extraction queue #$queueNumber/$($script:TotalArchives): archive='$($archive.FullName)'; jobId=$($job.Id); temp='$tempFolder'; workerLog='$workerLogPath'. Running=$(Get-RunningExtractionCount)/$MaxParallel" -Level INFO
    } catch {
        if ($tempFolder -and (Test-Path -LiteralPath $tempFolder)) { Remove-Item -LiteralPath $tempFolder -Recurse -Force -ErrorAction SilentlyContinue }
        $script:ErrorCount++
        $script:CompletedArchives++
        Write-Host "Could not start extraction for $($archive.Name): $($_.Exception.Message)" -ForegroundColor Red
        Write-Log -Message "Could not start extraction queue #$queueNumber for '$($archive.FullName)': $($_ | Out-String)" -Level ERROR
    }
}

function Remove-ExtractionTemp {
    param([string]$Folder, [System.Management.Automation.Job]$Job)
    if ($Job) { Remove-Job -Job $Job -Force -ErrorAction SilentlyContinue | Out-Null }
    if ($Folder -and (Test-Path -LiteralPath $Folder)) {
        Remove-Item -LiteralPath $Folder -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# Returns only the slice of the queue the dashboard displays, rather than copying the whole pending list on every refresh.
function Get-QueueSlice {
    param([object[]]$Queue, [int]$Start, [int]$Count)
    if ($null -eq $Queue -or $Start -ge $Queue.Count -or $Count -le 0) { return @() }
    $end = [Math]::Min($Queue.Count, $Start + $Count) - 1
    return @($Queue[$Start..$end])
}

function Get-ProgressJobSummary {
    $summary = [ordered]@{
        Running = 0
        Completed = 0
        Failed = 0
        Stopped = 0
        Names = @()
    }

    foreach ($entry in $script:Active) {
        $state = [string]$entry.Job.State
        switch ($state) {
            'Running' { $summary.Running++ }
            'Completed' { $summary.Completed++ }
            'Failed' { $summary.Failed++ }
            'Stopped' { $summary.Stopped++ }
            default { $summary.Running++ }
        }
        $summary.Names += "$($entry.Archive.Name) [$state]"
    }

    return $summary
}

function Update-ProgressDashboard {
    if ($script:ProgressStage -eq 'Archives') {
        $pendingCount = [Math]::Max(0, $script:Archives.Count - $script:NextArchiveIndex)
        $pendingItems = Get-QueueSlice -Queue $script:Archives -Start $script:NextArchiveIndex -Count $QueuePreviewLimit
        $jobSummary = Get-ProgressJobSummary
        $percent = if ($script:TotalArchives -gt 0) { [int](100 * $script:CompletedArchives / $script:TotalArchives) } else { 100 }
        $status = "Processed $($script:CompletedArchives)/$($script:TotalArchives); running $($jobSummary.Running)/$MaxParallel; failed $($jobSummary.Failed); stopped $($jobSummary.Stopped)"
        if (-not [string]::IsNullOrWhiteSpace($script:CurrentWork)) { $status += "; current: $($script:CurrentWork)" }
        Write-Progress -Id 0 -Activity 'Extracting RAR archives in parallel' -Status $status -PercentComplete $percent

        # Rescanning the staging folders is disk I/O, so do it at most once per second and cache the result.
        if (((Get-Date) - $script:LastProgressScan).TotalSeconds -ge 1) {
            $script:LastProgressScan = Get-Date
            $activeBytes = [int64]0
            $activeFiles = 0
            foreach ($entry in $script:Active) {
                foreach ($f in (Get-EligibleFiles -Folder $entry.TempFolder)) {
                    $activeFiles++
                    $activeBytes += [int64]$f.Length
                }
            }
            $script:LastActiveScanSummary = '{0} file(s), {1:N1} MB extracted so far' -f $activeFiles, ($activeBytes / 1MB)
        }
        $activeDetails = if ($jobSummary.Names.Count -gt 0) { $jobSummary.Names -join '; ' } else { 'No active archive jobs' }
        if ($script:LastActiveScanSummary) { $activeDetails = "$($script:LastActiveScanSummary) | $activeDetails" }
        Write-Progress -Id 1 -ParentId 0 -Activity 'Active extraction jobs' -Status $activeDetails -PercentComplete 0
    } else {
        $pendingCount = [Math]::Max(0, $script:StandaloneQueue.Count - $script:StandaloneNextIndex)
        $pendingItems = Get-QueueSlice -Queue $script:StandaloneQueue -Start $script:StandaloneNextIndex -Count $QueuePreviewLimit
        $percent = if ($script:TotalStandalone -gt 0) { [int](100 * $script:StandaloneProcessedCount / $script:TotalStandalone) } else { 100 }
        $status = "Processed $($script:StandaloneProcessedCount)/$($script:TotalStandalone) standalone files"
        if (-not [string]::IsNullOrWhiteSpace($script:CurrentWork)) { $status += "; current: $($script:CurrentWork)" }
        Write-Progress -Id 0 -Activity 'Compressing previously extracted NSP/XCI files' -Status $status -PercentComplete $percent
        $conversionStatus = if ([string]::IsNullOrWhiteSpace($script:CurrentWork)) { 'Idle' } else { $script:CurrentWork }
        Write-Progress -Id 1 -ParentId 0 -Activity 'Active conversions' -Status $conversionStatus -PercentComplete 0
    }

    $currentQueueProgressIds = @()
    $visibleQueueCount = $pendingItems.Count
    for ($slot = 0; $slot -lt $visibleQueueCount; $slot++) {
        $queueId = 2 + $slot
        $currentQueueProgressIds += $queueId
        if ($script:ProgressStage -eq 'Archives') {
            $queueOrdinal = $script:NextArchiveIndex + $slot + 1
            $queueActivity = 'WAITING #{0:D2}/{1:D2}' -f $queueOrdinal, $script:TotalArchives
        } else {
            $queueOrdinal = $script:StandaloneNextIndex + $slot + 1
            $queueActivity = 'WAITING #{0:D2}/{1:D2}' -f $queueOrdinal, $script:TotalStandalone
        }

        # Format arguments (not string interpolation) so braces in filenames can't break -f.
        $details = '{0} - {1:N1} MB' -f $pendingItems[$slot].Name, ([double]$pendingItems[$slot].Length / 1MB)

        Write-Progress -Id $queueId -ParentId 0 -Activity $queueActivity -Status $details -PercentComplete 0
    }
    if ($pendingCount -gt $QueuePreviewLimit) {
        $overflowId = 2 + $QueuePreviewLimit
        $currentQueueProgressIds += $overflowId
        $overflowCount = $pendingCount - $QueuePreviewLimit
        Write-Progress -Id $overflowId -ParentId 0 -Activity 'WAITING: additional files' -Status ("... and {0} more queued" -f $overflowCount) -PercentComplete 0
    }

    foreach ($oldId in $script:PreviousQueueProgressIds) {
        if ($currentQueueProgressIds -notcontains $oldId) {
            Write-Progress -Id $oldId -ParentId 0 -Activity 'Queue row' -Completed
        }
    }
    $script:PreviousQueueProgressIds = $currentQueueProgressIds
}

function Clear-ProgressDashboard {
    foreach ($queueId in $script:PreviousQueueProgressIds) {
        Write-Progress -Id $queueId -ParentId 0 -Activity 'Queue row' -Completed
    }
    $script:PreviousQueueProgressIds = @()
    Write-Progress -Id 1 -Activity 'Active extraction jobs' -Completed
    if ($script:ProgressStage -eq 'Archives') {
        Write-Progress -Id 0 -Activity 'Extracting RAR archives in parallel' -Completed
    } else {
        Write-Progress -Id 0 -Activity 'Compressing previously extracted NSP/XCI files' -Completed
    }
}

function Invoke-NSZConversion {
    param([Parameter(Mandatory)][System.IO.FileInfo]$File)

    $expectedExtension = if ($File.Extension -ieq '.nsp') { '.nsz' } elseif ($File.Extension -ieq '.xci') { '.xcz' } else { return $null }
    $workFolder = New-StagingFolder -Prefix 'nsz'

    $script:CurrentWork = "Compressing $($File.Extension.ToUpperInvariant()) to $($expectedExtension.ToUpperInvariant()): $($File.Name)"
    Write-Log -Message "Starting NSZ conversion: executable='$script:NszExe'; source='$($File.FullName)'; targetExtension='$expectedExtension'; keys='$script:KeysFile'; workFolder='$workFolder'." -Level INFO
    $job = $null
    try {
        # --verify is paired with --keep as required by the NSZ CLI documentation.
        $job = Start-Job -ScriptBlock $script:NszWorker -ArgumentList $script:NszExe, $File.FullName, $workFolder, $script:KeysFile
        while ($job.State -in @('NotStarted', 'Running', 'Blocked', 'Disconnected')) {
            # Keep extraction slots busy while this (often long) conversion runs.
            if ($script:ProgressStage -eq 'Archives') { Start-PendingExtractions }
            Update-ProgressDashboard
            Start-Sleep -Milliseconds 400
        }
        $result = Receive-Job -Job $job -ErrorAction SilentlyContinue | Select-Object -Last 1
        if ($null -eq $result) {
            $jobDiagnostics = @("NSZ background job state: $($job.State)")
            $nszChildJobs = @($job.ChildJobs)
            if ($nszChildJobs.Count -gt 0) {
                foreach ($jobError in @($nszChildJobs[0].Error)) { $jobDiagnostics += ('Job error: ' + [string]$jobError) }
                if ($nszChildJobs[0].JobStateInfo.Reason) { $jobDiagnostics += ('Job reason: ' + $nszChildJobs[0].JobStateInfo.Reason.ToString()) }
            } else {
                $jobDiagnostics += 'The NSZ background job has no child-job object.'
            }
            Write-Log -Message "NSZ conversion job returned no result for '$($File.FullName)': $($jobDiagnostics -join ' | ')" -Level ERROR
            throw "The NSZ conversion job returned no result. $($jobDiagnostics -join ' | ')"
        }
        Write-Log -Message "NSZ raw output for '$($File.FullName)': $([string]$result.FullOutput)" -Level DEBUG
        if ([int]$result.ExitCode -ne 0) {
            $detail = [string]$result.Output
            if ([string]::IsNullOrWhiteSpace($detail)) { $detail = 'NSZ did not provide diagnostic output.' }
            throw "NSZ exited with code $($result.ExitCode). $detail"
        }

        $outputs = @(Get-ChildItem -LiteralPath $workFolder -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -ieq $expectedExtension })
        if ($outputs.Count -ne 1) {
            throw "NSZ exited successfully but produced $($outputs.Count) '$expectedExtension' output files; expected exactly one."
        }
        $outputFile = Get-Item -LiteralPath $outputs[0].FullName
        if ($outputFile.Length -le 0) { throw 'The compressed output file is empty.' }
        $script:ConversionSuccessCount++
        Write-Log -Message "NSZ conversion passed: source='$($File.FullName)'; output='$($outputFile.FullName)'; bytes=$($outputFile.Length)." -Level INFO
        return [pscustomobject]@{ Success = $true; OutputFile = $outputFile; WorkFolder = $workFolder }
    } catch {
        $script:ConversionErrorCount++
        Write-Log -Message "NSZ conversion failed for '$($File.FullName)': $($_ | Out-String)" -Level ERROR
        Write-Host "Conversion failed for '$($File.Name)': $($_.Exception.Message)" -ForegroundColor Red
        if ($script:ProgressStage -eq 'Archives') {
            Write-Host 'The source RAR archive will be kept because conversion did not complete successfully.' -ForegroundColor Yellow
        }
        if (Test-Path -LiteralPath $workFolder) { Remove-Item -LiteralPath $workFolder -Recurse -Force -ErrorAction SilentlyContinue }
        return [pscustomobject]@{ Success = $false; OutputFile = $null; WorkFolder = $null; Message = $_.Exception.Message }
    } finally {
        if ($job) { Remove-Job -Job $job -Force -ErrorAction SilentlyContinue | Out-Null }
    }
}

# Returns the Move-FlatFile result object (Action + Destination).
function Process-OneFile {
    param(
        [Parameter(Mandatory)][System.IO.FileInfo]$File,
        [bool]$IsStandalone = $false
    )
    if ($script:StopProcessing) { return [pscustomobject]@{ Action = 'Quit'; Destination = $null } }

    if ($script:CompressionEnabled -and $File.Extension -in @('.nsp', '.xci')) {
        $conversion = Invoke-NSZConversion -File $File
        if (-not $conversion -or -not $conversion.Success) { return [pscustomobject]@{ Action = 'Failed'; Destination = $null } }
        try {
            $moveResult = Move-FlatFile -File $conversion.OutputFile
            if ($IsStandalone -and $script:DeleteStandaloneSources -and $moveResult.Action -in $script:SuccessActions) {
                if (-not (Test-DeleteAllowed)) {
                    Write-Host "Dry-run/no-delete mode enabled; original standalone file kept: $($File.Name)" -ForegroundColor Yellow
                } elseif (-not (Test-DeleteCandidate -SourcePath $File.FullName -RequiredOutputs @($moveResult.Destination))) {
                    Write-Host "Compressed output missing or invalid; original standalone file kept: $($File.Name)" -ForegroundColor Yellow
                    Write-Log -Message "Delete skipped for standalone '$($File.FullName)' because verification failed for output '$($moveResult.Destination)'." -Level WARN
                } else {
                    try {
                        Remove-Item -LiteralPath $File.FullName -Force -ErrorAction Stop
                        Write-Host "Compressed and deleted original: $($File.Name)" -ForegroundColor Green
                        Write-Log -Message "Deleted standalone source '$($File.FullName)' after verified output '$($moveResult.Destination)'." -Level INFO
                    } catch {
                        Write-Host "Compressed output saved, but original could not be deleted ('$($File.Name)'): $($_.Exception.Message)" -ForegroundColor Yellow
                    }
                }
            }
            return $moveResult
        } finally {
            if ($conversion.WorkFolder -and (Test-Path -LiteralPath $conversion.WorkFolder)) {
                Remove-Item -LiteralPath $conversion.WorkFolder -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }

    return (Move-FlatFile -File $File)
}

function Process-FinishedArchive {
    param([Parameter(Mandatory)]$Entry)

    $exitCode = 999
    $detail = ''
    $jobState = [string]$Entry.Job.State
    $jobResult = $null
    $childJob = $null
    $childJobs = @($Entry.Job.ChildJobs)
    if ($childJobs.Count -gt 0) { $childJob = $childJobs[0] }
    $diagnosticLines = @()
    try {
        if ($jobState -eq 'Completed') {
            $received = @(Receive-Job -Job $Entry.Job -ErrorAction Continue)
            $jobResult = @($received | Where-Object { $null -ne $_ -and $_.PSObject.Properties['ExitCode'] }) | Select-Object -Last 1
            if ($jobResult) {
                $exitCode = [int]$jobResult.ExitCode
                $detail = [string]$jobResult.Message
                if ($jobResult.LogPath) { $Entry.LogPath = [string]$jobResult.LogPath }
            } else {
                $diagnosticLines += 'Job reached Completed state but returned no object containing ExitCode.'
                foreach ($item in $received) { $diagnosticLines += ('Received output type={0}; value={1}' -f $item.GetType().FullName, [string]$item) }
            }
        } else {
            $diagnosticLines += "Background job state was '$jobState', not Completed."
            if ($Entry.Job.JobStateInfo -and $Entry.Job.JobStateInfo.Reason) { $diagnosticLines += ('Job reason: ' + $Entry.Job.JobStateInfo.Reason.ToString()) }
        }
    } catch {
        $diagnosticLines += "Receive-Job exception: $($_.Exception.Message)"
        $diagnosticLines += ($_ | Out-String)
    }

    if ($childJob) {
        try {
            if ($childJob.JobStateInfo -and $childJob.JobStateInfo.Reason) { $diagnosticLines += ('Child job reason: ' + $childJob.JobStateInfo.Reason.ToString()) }
            foreach ($jobError in @($childJob.Error)) {
                $diagnosticLines += ('Job error stream: ' + [string]$jobError)
                if ($jobError.ScriptStackTrace) { $diagnosticLines += ('Job error stack: ' + $jobError.ScriptStackTrace) }
            }
            foreach ($jobWarning in @($childJob.Warning)) { $diagnosticLines += ('Job warning stream: ' + [string]$jobWarning) }
            foreach ($jobInfo in @($childJob.Information)) { $diagnosticLines += ('Job information stream: ' + [string]$jobInfo.MessageData) }
        } catch { $diagnosticLines += "Could not inspect child job streams: $($_.Exception.Message)" }
    }
    if ($Entry.LogPath -and (Test-Path -LiteralPath $Entry.LogPath -PathType Leaf)) {
        $diagnosticLines += "Full per-archive WinRAR log: $($Entry.LogPath)"
    }
    if ($diagnosticLines.Count -gt 0) {
        foreach ($line in $diagnosticLines) { Write-Log -Message "Archive='$($Entry.Archive.Name)' diagnostic: $line" -Level ERROR }
    }
    Write-Log -Message "Finished extraction job. Archive='$($Entry.Archive.FullName)'; jobId=$($Entry.Job.Id); jobState=$jobState; childState=$(if($childJob){$childJob.State}else{'none'}); exitCode=$exitCode; workerLog='$($Entry.LogPath)'; resultReturned=$([bool]$jobResult)" -Level INFO

    $archiveSucceeded = ($jobState -eq 'Completed' -and $null -ne $jobResult -and $exitCode -eq 0)
    if (-not $archiveSucceeded) {
        if ($jobResult -and $exitCode -eq 1) {
            $script:WarningCount++
            Write-Host "Warning: $($Entry.Archive.Name) (WinRAR exit code 1). Source archive kept." -ForegroundColor Yellow
        } elseif ($jobResult -and $exitCode -notin @(0,1,998,999)) {
            $script:ErrorCount++
            Write-Host "Extraction failed: $($Entry.Archive.Name) (WinRAR exit code $exitCode). Source archive kept." -ForegroundColor Red
        } elseif ($jobResult -and $exitCode -eq 998) {
            $script:ErrorCount++
            Write-Host "Extraction could not start correctly for: $($Entry.Archive.Name). Source archive kept." -ForegroundColor Red
        } else {
            $script:ErrorCount++
            Write-Host "Extraction worker failed before a reliable WinRAR exit code was returned: $($Entry.Archive.Name) (job state '$jobState'). Source archive kept." -ForegroundColor Red
        }
        if (-not [string]::IsNullOrWhiteSpace($detail)) { Write-Host "  Worker detail: $detail" -ForegroundColor DarkRed }
        foreach ($line in $diagnosticLines) { Write-Host "  Diagnostic: $line" -ForegroundColor DarkYellow }
        if ($Entry.LogPath) { Write-Host "  Worker log: $($Entry.LogPath)" -ForegroundColor Yellow }
        if ($jobResult -and $exitCode -notin @(0, 998, 999)) { Write-Log -Message "WinRAR reported nonzero exit code $exitCode for '$($Entry.Archive.Name)'; detail='$detail'" -Level ERROR }
        Remove-ExtractionTemp -Folder $Entry.TempFolder -Job $Entry.Job
        $script:CompletedArchives++
        return
    }

    if ($script:StopProcessing) {
        Write-Host "Stopped before processing extracted contents: $($Entry.Archive.Name). Source archive kept." -ForegroundColor Yellow
        Write-Log -Message "Completed archive extraction but user stopped before consuming outputs: '$($Entry.Archive.Name)'." -Level WARN
        Remove-ExtractionTemp -Folder $Entry.TempFolder -Job $Entry.Job
        $script:CompletedArchives++
        return
    }

    $eligibleFiles = @(Get-EligibleFiles -Folder $Entry.TempFolder)
    Write-Log -Message "Archive extraction contents scanned: '$($Entry.Archive.Name)'; eligible files=$($eligibleFiles.Count); temp='$($Entry.TempFolder)'." -Level INFO
    if ($eligibleFiles.Count -eq 0) {
        Write-Host "No eligible files found in $($Entry.Archive.Name); source archive kept." -ForegroundColor Yellow
        $script:WarningCount++
        Remove-ExtractionTemp -Folder $Entry.TempFolder -Job $Entry.Job
        $script:CompletedArchives++
        return
    }

    $allFilesHandled = $true
    # The actual written paths (e.g. 'game.nsz' for 'game.nsp', or 'file (2).ext' for keep-both) are what must be verified.
    $requiredOutputPaths = @()
    foreach ($file in $eligibleFiles) {
        if ($script:StopProcessing) { $allFilesHandled = $false; break }
        Write-Log -Message "Handling extracted file '$($file.FullName)' ($($file.Length) bytes). CompressionEnabled=$($script:CompressionEnabled)." -Level INFO
        $result = Process-OneFile -File $file -IsStandalone:$false
        Write-Log -Message "Handled extracted file '$($file.Name)': result='$($result.Action)'; destination='$($result.Destination)'." -Level INFO
        if ($result.Action -in $script:SuccessActions) {
            $requiredOutputPaths += $result.Destination
        } else {
            $allFilesHandled = $false
        }
    }

    if ($DeleteArchivesOnSuccess -and $allFilesHandled -and -not $script:StopProcessing -and (Test-DeleteAllowed)) {
        $volumeFiles = @(Get-ArchiveFilesToDelete -Archive $Entry.Archive)
        $deleteFailure = $false

        Write-Log -Message "Archive content handling succeeded for '$($Entry.Archive.Name)'; attempting to delete source volumes: $($volumeFiles.Name -join ', ')." -Level INFO
        foreach ($volume in $volumeFiles) {
            if (-not (Test-DeleteCandidate -SourcePath $volume.FullName -RequiredOutputs $requiredOutputPaths)) {
                $deleteFailure = $true
                Write-Host "Delete skipped for archive volume '$($volume.Name)' because required output verification failed." -ForegroundColor Yellow
                Write-Log -Message "Archive volume delete skipped because verification failed for '$($volume.FullName)'." -Level WARN
                continue
            }
            try {
                Remove-Item -LiteralPath $volume.FullName -Force -ErrorAction Stop
                $script:DeletedArchiveCount++
                Write-Log -Message "Deleted source volume '$($volume.FullName)'." -Level INFO
            } catch {
                $deleteFailure = $true
                Write-Host "Could not delete source archive file '$($volume.Name)': $($_.Exception.Message)" -ForegroundColor Red
                Write-Log -Message "Source archive deletion failed: '$($volume.FullName)': $($_ | Out-String)" -Level ERROR
            }
        }
        if (-not $deleteFailure) {
            Write-Host "Extraction, conversion/copy succeeded; deleted source archive set: $($Entry.Archive.Name)" -ForegroundColor Green
        } else {
            $script:WarningCount++
            Write-Host "Output succeeded, but one or more source volume files could not be deleted for: $($Entry.Archive.Name)" -ForegroundColor Yellow
        }
    } elseif ($allFilesHandled -and -not $script:StopProcessing) {
        Write-Host "Extracted: $($Entry.Archive.Name) ($($requiredOutputPaths.Count) file(s)); source archive kept." -ForegroundColor Green
        Write-Log -Message "Source archive kept (deletion not enabled): '$($Entry.Archive.Name)'" -Level INFO
    } else {
        Write-Host "Keeping source archive: $($Entry.Archive.Name) (one or more files failed, were skipped, or processing was stopped)." -ForegroundColor Yellow
        Write-Log -Message "Source archive kept because not all eligible files were copied/converted without skips or errors: '$($Entry.Archive.Name)'" -Level WARN
    }

    Remove-ExtractionTemp -Folder $Entry.TempFolder -Job $Entry.Job
    $script:CompletedArchives++
}

# ------------------------ Input and destination ------------------------
try {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
} catch {
    Write-Host 'Unable to load Windows Forms for the folder/file pickers.' -ForegroundColor Red
    Write-Log -Message 'Unable to load System.Windows.Forms. Exiting.' -Level ERROR
    Wait-ForExit
    Stop-RunLogging
    exit 1
}

if ([string]::IsNullOrWhiteSpace($SourceFolder) -or -not (Test-Path -LiteralPath $SourceFolder -PathType Container)) {
    if ($script:NonInteractive) {
        Write-Host "Non-interactive mode requires a valid -SourceFolder. Given: '$SourceFolder'" -ForegroundColor Red
        Write-Log -Message "Non-interactive run without a valid source folder ('$SourceFolder'). Exiting." -Level ERROR
        Stop-RunLogging
        exit 1
    }
    if (-not [string]::IsNullOrWhiteSpace($SourceFolder)) { Write-Host "Source folder not found: $SourceFolder" -ForegroundColor Yellow }
    $sourcePromptInitial = if ([string]::IsNullOrWhiteSpace($SourceFolder)) { $env:USERPROFILE } else { $SourceFolder }
    $sourceChoice = Select-Folder -Description 'Choose the folder that contains your RAR archive sets or loose .nsp/.xci files.' -InitialFolder $sourcePromptInitial
    if ([string]::IsNullOrWhiteSpace($sourceChoice)) {
        Write-Host 'Cancelled. No files were processed.' -ForegroundColor Yellow
        Write-Log -Message 'User cancelled source-folder selection; no files processed.' -Level WARN
        Stop-RunLogging
        exit 0
    }
    $SourceFolder = $sourceChoice
    $DefaultOutputFolder = Join-Path $SourceFolder 'Extracted'
}
$SourceFolder = [System.IO.Path]::GetFullPath($SourceFolder)
Write-Log -Message "Source folder: '$SourceFolder'." -Level INFO

$detectedArchives = @(Get-ArchiveStarts -Folder $SourceFolder)
Write-Log -Message "Detected $($detectedArchives.Count) archive set(s) directly in source folder '$SourceFolder'." -Level INFO

if (-not [string]::IsNullOrWhiteSpace($script:OutputFolder)) {
    Write-Log -Message "Using -OutputFolder from the command line: '$script:OutputFolder'." -Level INFO
} elseif ($script:NonInteractive) {
    $script:OutputFolder = $DefaultOutputFolder
    Write-Host "Non-interactive mode: using default output folder $script:OutputFolder" -ForegroundColor Yellow
} else {
    # Pre-create the suggested folder so the picker can open on it, but remove it again if it is not chosen.
    $createdDefaultOutput = $false
    if (-not (Test-Path -LiteralPath $DefaultOutputFolder)) {
        try { New-Item -ItemType Directory -Path $DefaultOutputFolder -Force -ErrorAction Stop | Out-Null; $createdDefaultOutput = $true } catch {}
    }
    $outputInitial = if (Test-Path -LiteralPath $DefaultOutputFolder -PathType Container) { $DefaultOutputFolder } else { $SourceFolder }
    $outputChoice = Select-Folder -Description 'Choose the destination for extracted and compressed files. Eligible files are flattened into this folder.' -InitialFolder $outputInitial
    $chosePrecreated = (-not [string]::IsNullOrWhiteSpace($outputChoice)) -and ([System.IO.Path]::GetFullPath($outputChoice).TrimEnd('\') -ieq [System.IO.Path]::GetFullPath($DefaultOutputFolder).TrimEnd('\'))
    if ($createdDefaultOutput -and (-not $chosePrecreated -or $script:DryRun)) {
        if (-not (Get-ChildItem -LiteralPath $DefaultOutputFolder -Force -ErrorAction SilentlyContinue | Select-Object -First 1)) {
            Remove-Item -LiteralPath $DefaultOutputFolder -Force -ErrorAction SilentlyContinue
        }
    }
    if ([string]::IsNullOrWhiteSpace($outputChoice)) {
        Write-Host 'Cancelled. No files were processed.' -ForegroundColor Yellow
        Write-Log -Message 'User cancelled destination selection; no files processed.' -Level WARN
        Stop-RunLogging
        exit 0
    }
    $script:OutputFolder = $outputChoice
}
$script:OutputFolder = [System.IO.Path]::GetFullPath($script:OutputFolder)
if (-not $script:DryRun) {
    try {
        New-Item -ItemType Directory -Path $script:OutputFolder -Force -ErrorAction Stop | Out-Null
    } catch {
        Write-Host "Could not create the output folder '$script:OutputFolder': $($_.Exception.Message)" -ForegroundColor Red
        Write-Log -Message "Could not create output folder '$script:OutputFolder': $($_ | Out-String)" -Level ERROR
        Wait-ForExit
        Stop-RunLogging
        exit 1
    }
}

$looseScanFolders = @($SourceFolder, $script:OutputFolder)
if (Read-YesNo -Prompt 'Would you like to scan another folder for previously extracted .nsp/.xci files to compress?') {
    $additionalFolder = Select-Folder -Description 'Choose the additional folder to scan for .nsp/.xci files.' -InitialFolder $script:OutputFolder
    if (-not [string]::IsNullOrWhiteSpace($additionalFolder)) { $looseScanFolders += $additionalFolder }
}
# De-duplicate folder paths without changing the user's selected paths.
$uniqueFolders = @()
$seenFolders = @{}
foreach ($folder in $looseScanFolders) {
    $normalized = [System.IO.Path]::GetFullPath($folder).TrimEnd('\').ToLowerInvariant()
    if (-not $seenFolders.ContainsKey($normalized)) { $seenFolders[$normalized] = $true; $uniqueFolders += $folder }
}
$looseCandidates = @(Get-LooseGameFiles -Folders $uniqueFolders)
Write-Log -Message "Scanned loose-file folders: $($uniqueFolders -join '; '); standalone NSP/XCI candidates=$($looseCandidates.Count)." -Level INFO

# -------------------------- Dependency checks --------------------------
# A dry run only reports what is missing; it never installs anything.
$archiveProcessingEnabled = ($detectedArchives.Count -gt 0)
if ($archiveProcessingEnabled) {
    $script:RarExe = Resolve-RarExecutable
    if (-not $script:RarExe -and -not $script:DryRun) { Write-Log -Message 'WinRAR command-line executable not found; asking to install it.' -Level WARN; $script:RarExe = Install-WinRAR }
    if ($script:RarExe) { Write-Log -Message "Using WinRAR executable '$script:RarExe'." -Level INFO }
    if (-not $script:RarExe) {
        $archiveProcessingEnabled = $false
        Write-Host 'RAR extraction is disabled because WinRAR could not be located or installed. Source RAR archives will not be touched.' -ForegroundColor Yellow
    }
}

$needsCompression = ($archiveProcessingEnabled -or $looseCandidates.Count -gt 0)
if ($needsCompression) {
    $script:NszExe = Resolve-NszExecutable
    if (-not $script:NszExe -and -not $script:DryRun) { Write-Log -Message 'No working NSZ CLI was detected; asking to install NSZ into an isolated Python environment.' -Level WARN; $script:NszExe = Install-Nsz }
    if ($script:NszExe) {
        Write-Log -Message "Using NSZ executable '$script:NszExe'." -Level INFO
        $script:KeysFile = Find-KeysFile
        if (-not $script:KeysFile -and -not $script:DryRun) { $script:KeysFile = Select-KeysFile }
        if ($script:KeysFile) {
            $script:CompressionEnabled = $true
            Write-Log -Message "Compression enabled with NSZ='$script:NszExe' and keys='$script:KeysFile'." -Level INFO
        } else {
            Write-Host 'NSP/XCI compression is disabled because no compatible keys file was selected. Archives can still be extracted, with NSP/XCI files copied uncompressed.' -ForegroundColor Yellow
        }
    } else {
        Write-Host 'NSP/XCI compression is disabled because NSZ is unavailable or installation was declined. Archive contents, if extracted, will be copied without compression.' -ForegroundColor Yellow
    }
}

if ($archiveProcessingEnabled) { $script:Archives = $detectedArchives } else { $script:Archives = @() }
$script:TotalArchives = $script:Archives.Count
if ($script:CompressionEnabled) {
    $script:StandaloneQueue = $looseCandidates
    $script:TotalStandalone = $script:StandaloneQueue.Count
}

# ------------------------- Deletion preferences -------------------------
# Command-line flags win; otherwise ask. Dry-run/no-delete always wins over both.
if ($script:CompressionEnabled -and $script:TotalStandalone -gt 0) {
    Write-Host "Found $($script:TotalStandalone) existing .nsp/.xci file(s) to convert." -ForegroundColor Cyan
    if (-not (Test-DeleteAllowed)) {
        $script:DeleteStandaloneSources = $false
    } elseif ($DeleteStandalone) {
        $script:DeleteStandaloneSources = $true
        Write-Host 'CLI override enabled: original standalone NSP/XCI files will be deleted only after verification succeeds.' -ForegroundColor Yellow
    } else {
        $script:DeleteStandaloneSources = Read-YesNo -Prompt 'After successful verification, delete the original standalone .nsp/.xci files to save space?'
    }
}

if ($script:Archives.Count -gt 0) {
    if (-not (Test-DeleteAllowed)) {
        $DeleteArchivesOnSuccess = $false
    } elseif ($DeleteArchives) {
        $DeleteArchivesOnSuccess = $true
        Write-Host 'CLI override enabled: original RAR archives will be removed after successful verification.' -ForegroundColor Yellow
    } else {
        $DeleteArchivesOnSuccess = Read-YesNo -Prompt 'After successful verification, delete the original RAR archive sets? This only happens after every eligible file in a set is handled.'
    }
} else {
    $DeleteArchivesOnSuccess = $false
}
$DeleteArchivesOnSuccess = [bool]$DeleteArchivesOnSuccess
$script:DeleteArchivesConfirmed = $DeleteArchivesOnSuccess

# ------------------------------ Dry run -------------------------------
if ($script:DryRun) {
    Write-Host ''
    Write-Host '===================== Dry run summary =====================' -ForegroundColor Cyan
    Write-Host "Source folder:            $SourceFolder"
    Write-Host "Destination:              $script:OutputFolder"
    Write-Host "RAR archive sets:         $($script:Archives.Count) of $($detectedArchives.Count) detected"
    foreach ($archive in ($script:Archives | Select-Object -First 20)) { Write-Host "  - $($archive.Name)" }
    if ($script:Archives.Count -gt 20) { Write-Host "  ... and $($script:Archives.Count - 20) more" }
    Write-Host "Loose NSP/XCI files:      $($script:StandaloneQueue.Count) of $($looseCandidates.Count) found"
    foreach ($loose in ($script:StandaloneQueue | Select-Object -First 20)) { Write-Host "  - $($loose.FullName)" }
    if ($script:StandaloneQueue.Count -gt 20) { Write-Host "  ... and $($script:StandaloneQueue.Count - 20) more" }
    Write-Host "WinRAR:                   $(if ($script:RarExe) { $script:RarExe } elseif ($detectedArchives.Count -gt 0) { 'NOT FOUND (would offer to install)' } else { 'not needed' })"
    Write-Host "NSZ:                      $(if ($script:NszExe) { $script:NszExe } elseif ($needsCompression) { 'NOT FOUND (would offer to install)' } else { 'not needed' })"
    Write-Host "Keys file:                $(if ($script:KeysFile) { $script:KeysFile } elseif ($script:NszExe) { 'NOT FOUND (would ask you to select one)' } else { 'n/a' })"
    Write-Host "Compression:              $($script:CompressionEnabled)"
    Write-Host "Parallel extraction jobs: $MaxParallel"
    Write-Host 'Nothing was extracted, converted, copied, installed, or deleted.' -ForegroundColor Yellow
    Write-Host '==========================================================' -ForegroundColor Cyan
    Write-Log -Message "Dry run finished: archives=$($script:Archives.Count); loose=$($script:StandaloneQueue.Count); compression=$($script:CompressionEnabled)." -Level INFO
    Wait-ForExit
    Stop-RunLogging
    exit 0
}

if ($script:Archives.Count -eq 0 -and $script:StandaloneQueue.Count -eq 0) {
    if ($detectedArchives.Count -eq 0 -and $looseCandidates.Count -eq 0) {
        Write-Host "No .rar archives or standalone .nsp/.xci files were found in the selected scan folders." -ForegroundColor Yellow
    } elseif (-not $script:CompressionEnabled -and $looseCandidates.Count -gt 0) {
        Write-Host 'No standalone files were converted. Install/enable NSZ and select a valid keys file, then run the script again.' -ForegroundColor Yellow
    } else {
        Write-Host 'No tasks are enabled. Files and source archives were left untouched.' -ForegroundColor Yellow
    }
    Write-Host "Diagnostic log: $script:DiagnosticsLogPath"
    Write-Host "Transcript:     $script:TranscriptPath"
    Wait-ForExit
    Stop-RunLogging
    exit 0
}

Write-Host ''
Write-Host "Source folder:            $SourceFolder"
Write-Host "Destination:              $script:OutputFolder"
Write-Host "RAR archive sets:         $($script:Archives.Count)"
$parallelStatus = if ($script:Archives.Count -gt 0) { "ON (up to $MaxParallel concurrent WinRAR jobs)" } else { 'Not needed' }
$compressionStatus = if ($script:CompressionEnabled) { 'ON (NSP -> NSZ, XCI -> XCZ)' } else { 'OFF' }
Write-Host "Parallel extraction:      $parallelStatus"
Write-Host "NSP/XCI compression:      $compressionStatus"
Write-Host "Loose files queued:       $($script:StandaloneQueue.Count)"
Write-Host "Log folder:               $script:RunLogFolder"
if ($script:CompressionEnabled) { Write-Host "NSZ executable:           $script:NszExe"; Write-Host "Keys file:                $script:KeysFile" }
if ($script:Archives.Count -gt 0) {
    if ($DeleteArchivesOnSuccess) {
        Write-Host 'RAR archives: deleted only after every eligible extracted file has been converted/copied and verified.' -ForegroundColor Yellow
    } else {
        Write-Host 'RAR archives: kept.' -ForegroundColor Yellow
    }
}
if ($script:StandaloneQueue.Count -gt 0) {
    if ($script:DeleteStandaloneSources) {
        Write-Host 'Loose NSP/XCI originals: deleted only after the compressed result is verified and in place.' -ForegroundColor Yellow
    } else {
        Write-Host 'Loose NSP/XCI originals: kept.' -ForegroundColor Yellow
    }
}
if ($script:NoDelete) { Write-Host 'No-delete mode active: no source files will be deleted.' -ForegroundColor Yellow }
Write-Host ''

try {
    # ---------------------- Parallel archive extraction ----------------------
    $script:ProgressStage = 'Archives'
    $script:CurrentWork = 'Starting parallel extraction queue'
    Start-PendingExtractions

    while (-not $script:StopProcessing -and ($script:NextArchiveIndex -lt $script:Archives.Count -or $script:Active.Count -gt 0)) {
        # Refill any failed-to-start slots too; otherwise a Start-Job startup error could leave the queue stuck.
        Start-PendingExtractions
        $finished = @($script:Active | Where-Object { $_.Job.State -in @('Completed', 'Failed', 'Stopped') })

        foreach ($entry in $finished) {
            $script:Active = @($script:Active | Where-Object { $_.Job.Id -ne $entry.Job.Id })
            # Refill queue slots immediately: this is the rolling parallel queue, not sequential batches.
            Start-PendingExtractions
            $script:CurrentWork = "Processing extracted files: $($entry.Archive.Name)"
            Process-FinishedArchive -Entry $entry
            if ($script:StopProcessing) { break }
        }
        $script:CurrentWork = ''

        Update-ProgressDashboard
        if ($finished.Count -eq 0) { Start-Sleep -Milliseconds 300 }
    }
    Clear-ProgressDashboard

    # ---------------------- Previously extracted files ----------------------
    if ($script:CompressionEnabled -and $script:StandaloneQueue.Count -gt 0 -and -not $script:StopProcessing) {
        $script:ProgressStage = 'Standalone'
        $script:CurrentWork = 'Starting standalone NSP/XCI conversions'
        while ($script:StandaloneNextIndex -lt $script:StandaloneQueue.Count -and -not $script:StopProcessing) {
            $file = $script:StandaloneQueue[$script:StandaloneNextIndex]
            $script:StandaloneNextIndex++
            $script:CurrentWork = "Converting existing file: $($file.Name)"
            Update-ProgressDashboard
            $result = Process-OneFile -File $file -IsStandalone:$true
            if ($result.Action -eq 'Quit') { break }
            $script:StandaloneProcessedCount++
            Update-ProgressDashboard
        }
        Clear-ProgressDashboard
    }
} finally {
    # Runs on normal completion, on Q, and on Ctrl+C: stop outstanding jobs and remove temporary files.
    foreach ($entry in @($script:Active)) {
        try { if ($entry.Job.State -in @('NotStarted', 'Running')) { Stop-Job -Job $entry.Job -ErrorAction SilentlyContinue } } catch {}
        Remove-ExtractionTemp -Folder $entry.TempFolder -Job $entry.Job
        Write-Log -Message "Stopped unfinished extraction job for '$($entry.Archive.Name)'; source archive kept." -Level WARN
    }
    $script:Active = @()
    Remove-StagingRoot
    Clear-ProgressDashboard
}

$failureTotal = $script:ErrorCount + $script:FailedCopyCount + $script:ConversionErrorCount
Write-Host ''
Write-Host '==================== Run summary ====================' -ForegroundColor Cyan
Write-Host "Source folder:            $SourceFolder"
Write-Host "Output folder:            $script:OutputFolder"
Write-Host "Archives detected:        $($detectedArchives.Count)"
Write-Host "Archives processed:       $($script:CompletedArchives)"
Write-Host "Archives deleted (files): $($script:DeletedArchiveCount)"
Write-Host "Loose files processed:    $($script:StandaloneProcessedCount)"
Write-Host "Moved to output:          $($script:CopiedCount)"
Write-Host "Skipped:                  $($script:SkippedCount)"
Write-Host "Copy failures:            $($script:FailedCopyCount)"
Write-Host "Extraction errors:        $($script:ErrorCount)"
Write-Host "Warnings:                 $($script:WarningCount)"
Write-Host "Compression successes:    $($script:ConversionSuccessCount)"
Write-Host "Compression failures:     $($script:ConversionErrorCount)"
Write-Host "Diagnostic log:           $script:DiagnosticsLogPath"
Write-Host "Transcript log:           $script:TranscriptPath"
if ($script:StopProcessing) { Write-Host 'Processing stopped by user. Unprocessed archives were left untouched.' -ForegroundColor Yellow }
if (-not $script:CompressionEnabled -and ($detectedArchives.Count -gt 0 -or $looseCandidates.Count -gt 0)) {
    Write-Host 'Compression was disabled; existing loose NSP/XCI files were left untouched, and archive contents were copied uncompressed.' -ForegroundColor Yellow
}
Write-Host '====================================================' -ForegroundColor Cyan
Write-Log -Message "Run summary: archivesDetected=$($detectedArchives.Count); archivesProcessed=$($script:CompletedArchives); archivesDeleted=$($script:DeletedArchiveCount); copied=$($script:CopiedCount); skipped=$($script:SkippedCount); copyFailures=$($script:FailedCopyCount); extractionErrors=$($script:ErrorCount); extractionWarnings=$($script:WarningCount); conversionsPassed=$($script:ConversionSuccessCount); conversionErrors=$($script:ConversionErrorCount); standaloneProcessed=$($script:StandaloneProcessedCount)/$($script:TotalStandalone)." -Level INFO
Write-Log -Message "Run finished. Transcript='$script:TranscriptPath'; diagnosticLog='$script:DiagnosticsLogPath'." -Level INFO
Wait-ForExit
Stop-RunLogging
# Nonzero exit code lets scheduled/automated runs detect failures.
exit $(if ($failureTotal -gt 0) { 1 } else { 0 })
