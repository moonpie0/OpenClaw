<#
.SYNOPSIS
  OpenClaw One-Click Installer (Windows) — Skip AI configuration automatically
.DESCRIPTION
  Automatically detects and installs Node.js v22+, Git, then installs and configures OpenClaw.
  After installation, no AI configuration is performed. The process ends directly.
.NOTES
  Usage:
    powershell -ExecutionPolicy Bypass -File install-openclaw.ps1
#>

# ── Self‑repair execution policy: restart with Bypass if needed ──
if ($MyInvocation.MyCommand.Path) {
    try {
        $policy = Get-ExecutionPolicy -Scope Process
        if ($policy -eq "Restricted" -or $policy -eq "AllSigned") {
            Write-Host "  [INFO] Detected execution policy $policy. Restarting with Bypass policy..." -ForegroundColor Blue
            Start-Process -FilePath "powershell.exe" -ArgumentList "-ExecutionPolicy Bypass -File `"$($MyInvocation.MyCommand.Path)`"" -Wait -NoNewWindow
            exit $LASTEXITCODE
        }
    } catch {}
}

# ── Force UTF-8 encoding (prevents garbled text) ──
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding  = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

if ($PSVersionTable.PSVersion.Major -lt 5) {
    Write-Host "  [FAIL] PowerShell 5.0 or later is required. Current version: $($PSVersionTable.PSVersion)" -ForegroundColor Red
    exit 1
}

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# ── Color output functions ──

function Write-Info    { param($Msg) Write-Host "  [INFO] " -ForegroundColor Blue -NoNewline; Write-Host $Msg }
function Write-Ok      { param($Msg) Write-Host "  [OK]   " -ForegroundColor Green -NoNewline; Write-Host $Msg }
function Write-Warn    { param($Msg) Write-Host "  [WARN] " -ForegroundColor Yellow -NoNewline; Write-Host $Msg }
function Write-Err     { param($Msg) Write-Host "  [FAIL] " -ForegroundColor Red -NoNewline; Write-Host $Msg }
function Write-Step    { param($Msg) Write-Host "`n━━━ $Msg ━━━`n" -ForegroundColor Cyan }

# ── Global variables ──

$script:NodeBinDir = $null
$script:NvmManaged = $false
$script:RequiredNodeMajor = 22
$script:Arch = if ($env:PROCESSOR_ARCHITECTURE -eq "ARM64") { "arm64" } else { "x64" }

function Get-LocalAppData {
    if ($env:LOCALAPPDATA) { return $env:LOCALAPPDATA }
    return (Join-Path $HOME "AppData\Local")
}

# ── Utility functions ──

function Refresh-PathEnv {
    $machinePath = [Environment]::GetEnvironmentVariable("PATH", "Machine")
    $userPath = [Environment]::GetEnvironmentVariable("PATH", "User")
    $env:PATH = "$machinePath;$userPath"
}

function Add-ToUserPath {
    param([string]$Dir)
    $currentPath = [Environment]::GetEnvironmentVariable("PATH", "User")
    if ($currentPath -notlike "*$Dir*") {
        [Environment]::SetEnvironmentVariable("PATH", "$Dir;$currentPath", "User")
        $env:PATH = "$Dir;$env:PATH"
        Write-Info "Added $Dir to user PATH"
    }
}

function Ensure-ExecutionPolicy {
    try {
        $currentPolicy = Get-ExecutionPolicy -Scope CurrentUser
        if ($currentPolicy -eq "Restricted" -or $currentPolicy -eq "AllSigned") {
            Write-Info "Current PowerShell execution policy is $currentPolicy. pnpm scripts cannot run."
            Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned -Force
            Write-Ok "Execution policy set to RemoteSigned (current user only)."
        }
    } catch {
        Write-Warn "Could not automatically set execution policy."
        Write-Host "  Please run the following command manually and then reopen the terminal:" -ForegroundColor Yellow
        Write-Host "    Set-ExecutionPolicy -Scope CurrentUser RemoteSigned" -ForegroundColor Cyan
    }
}

function Get-NodeVersion {
    param([string]$NodeExe = "node")
    try {
        $output = & $NodeExe -v 2>$null
        if ($output -match "v(\d+)") {
            $major = [int]$Matches[1]
            if ($major -ge $script:RequiredNodeMajor) {
                return $output.Trim()
            }
        }
    } catch {}
    return $null
}

function Pin-NodePath {
    foreach ($dir in $env:PATH.Split(";")) {
        if (-not $dir) { continue }
        $nodeExe = Join-Path $dir "node.exe"
        if (Test-Path $nodeExe) {
            try {
                $output = & $nodeExe -v 2>$null
                if ($output -match "v(\d+)" -and [int]$Matches[1] -ge $script:RequiredNodeMajor) {
                    $script:NodeBinDir = $dir
                    $rest = ($env:PATH.Split(";") | Where-Object { $_ -ne $dir }) -join ";"
                    $env:PATH = "$dir;$rest"
                    Write-Info "Pinned Node.js v22 path: $dir"
                    return
                }
            } catch {}
        }
    }
}

function Ensure-NodePriority {
    param([string]$NodeV22Dir)

    # Node managed by nvm doesn't need manual PATH adjustment; nvm use already handles it.
    if ($script:NvmManaged) { return }

    if (-not $NodeV22Dir -or -not (Test-Path (Join-Path $NodeV22Dir "node.exe"))) { return }

    $machinePath = [Environment]::GetEnvironmentVariable("PATH", "Machine")
    if (-not $machinePath) { return }
    $machineDirs = $machinePath.Split(";") | Where-Object { $_ }

    $hasConflict = $false
    foreach ($dir in $machineDirs) {
        if ($dir -eq $NodeV22Dir) { continue }
        $nodeExe = Join-Path $dir "node.exe"
        if (Test-Path $nodeExe) {
            try {
                $output = & $nodeExe -v 2>$null
                if ($output -match "v(\d+)" -and [int]$Matches[1] -lt $script:RequiredNodeMajor) {
                    $oldVer = $output.Trim()
                    Write-Warn "Detected low‑version Node.js in system PATH: $dir ($oldVer)"
                    $hasConflict = $true
                }
            } catch {}
        }
    }

    if (-not $hasConflict) { return }

    # Already first? Skip.
    if ($machineDirs[0] -eq $NodeV22Dir) { return }

    Write-Info "Promoting Node.js v22 path to the front of system PATH..."

    $newMachineDirs = @($NodeV22Dir) + ($machineDirs | Where-Object { $_ -ne $NodeV22Dir })
    $newMachinePath = $newMachineDirs -join ";"

    try {
        [Environment]::SetEnvironmentVariable("PATH", $newMachinePath, "Machine")
        Write-Ok "Set Node.js v22 as the system default version."
    } catch {
        Write-Info "Administrator rights required. Requesting elevation..."
        $escaped = $newMachinePath -replace "'", "''"
        try {
            Start-Process -FilePath "powershell.exe" `
                -ArgumentList "-NoProfile -Command `"[Environment]::SetEnvironmentVariable('PATH','$escaped','Machine')`"" `
                -Verb RunAs -Wait -WindowStyle Hidden
            Write-Ok "Set Node.js v22 as the system default version."
        } catch {
            Write-Warn "Could not modify system PATH (user cancelled elevation)."
            Write-Info "Patching openclaw startup script to use the correct Node.js version..."
            Patch-OpenclawShim -NodeV22Dir $NodeV22Dir
            return
        }
    }

    Refresh-PathEnv
    $env:PATH = "$NodeV22Dir;$env:PATH"
}

function Patch-OpenclawShim {
    param([string]$NodeV22Dir)

    $nodeExe = Join-Path $NodeV22Dir "node.exe"
    if (-not (Test-Path $nodeExe)) { return }

    $found = Find-OpenclawBinary
    if (-not $found) { return }

    $shimPath = $found.Path
    if ($shimPath -notlike "*.cmd") { return }

    try {
        $content = Get-Content $shimPath -Raw -Encoding UTF8
        if (-not $content) { return }

        # Already full path? Skip.
        if ($content -like "*$nodeExe*") {
            Write-Ok "openclaw startup script already uses the correct Node.js path."
            return
        }

        # Replace bare `node` calls with full path.
        $patched = $content -replace '(?m)^(@?)node(\.exe)?\s', "`$1`"$nodeExe`" "
        if ($patched -eq $content) {
            $patched = $content -replace '(?m)"node(\.exe)?"\s', "`"$nodeExe`" "
        }

        if ($patched -ne $content) {
            Set-Content -Path $shimPath -Value $patched -Encoding UTF8 -NoNewline
            Write-Ok "Patched openclaw startup script → $nodeExe"
        } else {
            Write-Warn "Could not automatically patch; openclaw.cmd format unexpected."
            Write-Host "  Manual fix:" -ForegroundColor Yellow
            Write-Host "    1. Open System Properties → Advanced → Environment Variables" -ForegroundColor Yellow
            Write-Host "    2. In System Variables PATH, move $NodeV22Dir to the front." -ForegroundColor Yellow
            Write-Host "    3. Or uninstall older Node.js versions and reopen the terminal." -ForegroundColor Yellow
        }
    } catch {
        Write-Warn "Failed to patch openclaw startup script: $_"
    }
}

function Get-NpmCmd {
    if ($script:NodeBinDir) {
        $cmd = Join-Path $script:NodeBinDir "npm.cmd"
        if (Test-Path $cmd) { return $cmd }
    }
    return "npm"
}

function Get-PnpmCmd {
    if ($script:NodeBinDir) {
        $cmd = Join-Path $script:NodeBinDir "pnpm.cmd"
        if (Test-Path $cmd) { return $cmd }
    }
    $defaultPnpmHome = Join-Path (Get-LocalAppData) "pnpm"
    $cmd = Join-Path $defaultPnpmHome "pnpm.cmd"
    if (Test-Path $cmd) { return $cmd }
    try {
        $resolved = (Get-Command pnpm.cmd -ErrorAction Stop).Source
        if (Test-Path $resolved) { return $resolved }
    } catch {}
    return "pnpm.cmd"
}

function Find-OpenclawBinary {
    $searchDirs = @()

    # pnpm bin -g (actual installation location)
    try {
        $pnpmCmd = Get-PnpmCmd
        $pnpmBin = (& $pnpmCmd bin -g 2>$null).Trim()
        if ($pnpmBin -and (Test-Path $pnpmBin)) { $searchDirs += $pnpmBin }
    } catch {}

    # PNPM_HOME
    if ($env:PNPM_HOME -and (Test-Path $env:PNPM_HOME)) { $searchDirs += $env:PNPM_HOME }

    # Default pnpm path + global store subdirectories
    $defaultPnpmHome = Join-Path (Get-LocalAppData) "pnpm"
    if (Test-Path $defaultPnpmHome) {
        $searchDirs += $defaultPnpmHome
        # pnpm global install may be in pnpm\global\<version>\node_modules\.bin
        $pnpmGlobalDir = Join-Path $defaultPnpmHome "global"
        if (Test-Path $pnpmGlobalDir) {
            Get-ChildItem -Path $pnpmGlobalDir -Directory -ErrorAction SilentlyContinue | ForEach-Object {
                $binDir = Join-Path $_.FullName "node_modules\.bin"
                if (Test-Path $binDir) { $searchDirs += $binDir }
            }
        }
    }

    # npm prefix -g (npm global prefix)
    try {
        $npmCmd = Get-NpmCmd
        $npmPrefix = (& $npmCmd prefix -g 2>$null).Trim()
        if ($npmPrefix) {
            if (Test-Path $npmPrefix) { $searchDirs += $npmPrefix }
            $npmBin = Join-Path $npmPrefix "bin"
            if (Test-Path $npmBin) { $searchDirs += $npmBin }
        }
    } catch {}

    # %AppData%\npm (common npm global directory on Windows)
    if ($env:APPDATA) {
        $appDataNpm = Join-Path $env:APPDATA "npm"
        if (Test-Path $appDataNpm) { $searchDirs += $appDataNpm }
    }

    # NodeBinDir
    if ($script:NodeBinDir -and (Test-Path $script:NodeBinDir)) { $searchDirs += $script:NodeBinDir }

    # where.exe lookup
    try {
        $whereResult = & where.exe openclaw 2>$null
        if ($whereResult) {
            $whereResult -split "`r?`n" | ForEach-Object {
                $line = $_.Trim()
                if ($line -and (Test-Path $line)) {
                    $searchDirs += (Split-Path $line -Parent)
                }
            }
        }
    } catch {}

    $searchDirs = $searchDirs | Where-Object { $_ } | Select-Object -Unique
    foreach ($dir in $searchDirs) {
        foreach ($name in @("openclaw.cmd", "openclaw.exe", "openclaw.ps1")) {
            $candidate = Join-Path $dir $name
            if (Test-Path $candidate) {
                return @{ Path = $candidate; Dir = $dir }
            }
        }
    }
    return $null
}

function Get-OpenclawCmd {
    $found = Find-OpenclawBinary
    if ($found) { return $found.Path }
    return "openclaw"
}

function Ensure-PnpmHome {
    $pnpmHome = $env:PNPM_HOME
    if (-not $pnpmHome) {
        $pnpmHome = [Environment]::GetEnvironmentVariable("PNPM_HOME", "User")
    }
    if (-not $pnpmHome) {
        $pnpmHome = Join-Path (Get-LocalAppData) "pnpm"
    }

    $env:PNPM_HOME = $pnpmHome
    if ($env:PATH -notlike "*$pnpmHome*") { $env:PATH = "$pnpmHome;$env:PATH" }

    $savedHome = [Environment]::GetEnvironmentVariable("PNPM_HOME", "User")
    if ($savedHome -ne $pnpmHome) {
        [Environment]::SetEnvironmentVariable("PNPM_HOME", $pnpmHome, "User")
        Write-Info "Persisted PNPM_HOME=$pnpmHome"
    }

    Add-ToUserPath $pnpmHome
}

function Download-File {
    param([string]$Dest, [string[]]$Urls)
    foreach ($url in $Urls) {
        $hostName = ([Uri]$url).Host
        Write-Info "Downloading from $hostName..."
        try {
            Invoke-WebRequest -Uri $url -OutFile $Dest -UseBasicParsing -TimeoutSec 300
            Write-Ok "Download completed"
            return $true
        } catch {
            Write-Warn "Download from $hostName failed. Trying backup source..."
        }
    }
    return $false
}

function Get-LatestNodeVersion {
    param([int]$Major)
    $urls = @(
        "https://npmmirror.com/mirrors/node/latest-v${Major}.x/SHASUMS256.txt",
        "https://nodejs.org/dist/latest-v${Major}.x/SHASUMS256.txt"
    )
    foreach ($url in $urls) {
        try {
            $content = (Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 15).Content
            if ($content -match "node-(v\d+\.\d+\.\d+)") {
                return $Matches[1]
            }
        } catch {}
    }
    return $null
}

# ── Node.js installation ──

function Install-NodeViaNvm {
    $nvmExe = $null
    try {
        $nvmExe = (Get-Command nvm -ErrorAction Stop).Source
    } catch {
        try {
            $nvmOut = & cmd /c "nvm version" 2>$null
            if (-not $nvmOut) { return $false }
        } catch { return $false }
    }

    Write-Info "Detected nvm-windows. Installing Node.js v22 via nvm..."

    # Save original node_mirror setting and restore later
    $nvmHome = if ($env:NVM_HOME) { $env:NVM_HOME } else { Join-Path $env:APPDATA "nvm" }
    $nvmSettings = Join-Path $nvmHome "settings.txt"
    $hadMirror = $false
    $oldMirror = $null
    if (Test-Path $nvmSettings) {
        $settingsContent = Get-Content $nvmSettings -ErrorAction SilentlyContinue
        $mirrorLine = $settingsContent | Where-Object { $_ -match '^node_mirror:\s*(.+)' }
        if ($mirrorLine) {
            $hadMirror = $true
            $oldMirror = ($mirrorLine -replace '^node_mirror:\s*', '').Trim()
        }
    }

    & cmd /c "nvm node_mirror https://npmmirror.com/mirrors/node/" 2>$null | Out-Null

    try {
        try { & cmd /c "nvm install 22" 2>$null | Out-Null } catch {
            Write-Warn "nvm install 22 failed: $_"
            return $false
        }

        # nvm use requires administrator rights (creates symlinks)
        & cmd /c "nvm use 22" 2>$null | Out-Null
        Refresh-PathEnv
        $ver = Get-NodeVersion
        if ($ver) {
            Write-Ok "Node.js $ver installed and switched via nvm."
            $script:NvmManaged = $true
            return $true
        }

        # nvm use may have failed due to insufficient permissions; try elevation
        Write-Info "nvm use requires administrator rights. Requesting elevation..."
        try {
            Start-Process -FilePath "cmd.exe" `
                -ArgumentList "/c nvm use 22" `
                -Verb RunAs -Wait -WindowStyle Hidden
            Refresh-PathEnv
            $ver = Get-NodeVersion
            if ($ver) {
                Write-Ok "Node.js $ver installed and switched via nvm."
                $script:NvmManaged = $true
                return $true
            }
        } catch {
            Write-Warn "nvm use elevation failed (user may have cancelled)."
        }

        Write-Warn "Failed to switch Node.js version via nvm."
        return $false
    } finally {
        # Restore nvm node_mirror setting
        if (Test-Path $nvmSettings) {
            if ($hadMirror) {
                & cmd /c "nvm node_mirror $oldMirror" 2>$null | Out-Null
            } else {
                $lines = Get-Content $nvmSettings -ErrorAction SilentlyContinue
                $lines = $lines | Where-Object { $_ -notmatch '^node_mirror:' }
                Set-Content $nvmSettings -Value $lines -ErrorAction SilentlyContinue
            }
        }
    }
}

function Install-NodeDirect {
    Write-Info "Downloading and installing Node.js v22 directly..."

    $version = Get-LatestNodeVersion -Major 22
    if (-not $version) {
        Write-Err "Could not retrieve Node.js version information. Check network connection."
        return $false
    }
    Write-Info "Latest LTS version: $version"

    $filename = "node-$version-win-$($script:Arch).zip"
    $tmpPath = Join-Path $env:TEMP "openclaw-install"
    $tmpFile = Join-Path $tmpPath $filename
    $extractedName = "node-$version-win-$($script:Arch)"
    $installDir = Join-Path (Get-LocalAppData) "nodejs"

    New-Item -ItemType Directory -Force -Path $tmpPath | Out-Null

    $downloaded = Download-File -Dest $tmpFile -Urls @(
        "https://npmmirror.com/mirrors/node/$version/$filename",
        "https://nodejs.org/dist/$version/$filename"
    )

    if (-not $downloaded) {
        Write-Err "Node.js download failed. Check network connection."
        return $false
    }

    try {
        Write-Info "Extracting and installing..."
        Expand-Archive -Path $tmpFile -DestinationPath $tmpPath -Force
        if (Test-Path $installDir) { Remove-Item $installDir -Recurse -Force }
        Move-Item (Join-Path $tmpPath $extractedName) $installDir

        $env:PATH = "$installDir;$env:PATH"
        Add-ToUserPath $installDir
    } catch {
        Write-Err "Installation failed: $_"
        Remove-Item $tmpPath -Recurse -Force -ErrorAction SilentlyContinue
        return $false
    }

    Remove-Item $tmpPath -Recurse -Force -ErrorAction SilentlyContinue

    $ver = Get-NodeVersion
    if ($ver) {
        Write-Ok "Node.js $ver installed successfully."
        return $true
    }
    Write-Warn "Node.js installation completed but verification failed."
    return $false
}

# ── Git installation ──

function Get-GitVersion {
    try {
        $output = & git --version 2>$null
        return $output.Trim()
    } catch {}

    $gitPaths = @(
        (Join-Path $env:ProgramFiles "Git\cmd"),
        (Join-Path ${env:ProgramFiles(x86)} "Git\cmd")
    )
    foreach ($gp in $gitPaths) {
        $gitExe = Join-Path $gp "git.exe"
        if (Test-Path $gitExe) {
            try {
                $output = & $gitExe --version 2>$null
                if (-not ($env:PATH -like "*$gp*")) { $env:PATH = "$gp;$env:PATH" }
                return $output.Trim()
            } catch {}
        }
    }
    return $null
}

function Get-LatestGitRelease {
    $url = "https://registry.npmmirror.com/-/binary/git-for-windows/"
    try {
        $content = (Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 15).Content
        $regexMatches = [regex]::Matches($content, "v(\d+)\.(\d+)\.(\d+)\.windows\.(\d+)/")
        if ($regexMatches.Count -eq 0) { return $null }

        $best = $regexMatches | Sort-Object {
            [int]$_.Groups[1].Value * 1000000 + [int]$_.Groups[2].Value * 10000 +
            [int]$_.Groups[3].Value * 100 + [int]$_.Groups[4].Value
        } -Descending | Select-Object -First 1

        $version = "$($best.Groups[1].Value).$($best.Groups[2].Value).$($best.Groups[3].Value)"
        $winBuild = $best.Groups[4].Value
        $tag = "v$version.windows.$winBuild"
        $fileVersion = if ($winBuild -eq "1") { $version } else { "$version.$winBuild" }
        return @{ Version = $version; Tag = $tag; FileVersion = $fileVersion }
    } catch {}

    try {
        $ghUrl = "https://api.github.com/repos/git-for-windows/git/releases/latest"
        $content = (Invoke-WebRequest -Uri $ghUrl -UseBasicParsing -TimeoutSec 15).Content
        if ($content -match '\x22tag_name\x22\s*:\s*\x22(v(\d+\.\d+\.\d+)\.windows\.(\d+))\x22') {
            $version = $Matches[2]; $winBuild = $Matches[3]; $tag = $Matches[1]
            $fileVersion = if ($winBuild -eq "1") { $version } else { "$version.$winBuild" }
            return @{ Version = $version; Tag = $tag; FileVersion = $fileVersion }
        }
    } catch {}
    return $null
}

function Install-GitViaWinget {
    try {
        Get-Command winget -ErrorAction Stop | Out-Null
    } catch { return $false }

    Write-Info "Detected winget. Installing Git..."
    try {
        & winget install --id Git.Git -e --source winget --silent --accept-package-agreements --accept-source-agreements 2>$null | Out-Null
    } catch {
        Write-Warn "winget command returned non‑zero exit code. Checking if Git is already available..."
    }

    Refresh-PathEnv
    $ver = Get-GitVersion
    if ($ver) {
        Write-Ok "$ver is now available."
        return $true
    }
    Write-Warn "Git still not detected after winget installation."
    return $false
}

function Install-GitDirect {
    Write-Info "Downloading Git for Windows..."

    $release = Get-LatestGitRelease
    if (-not $release) {
        Write-Err "Could not retrieve Git version information. Check network connection."
        return $false
    }
    Write-Info "Latest version: Git $($release.FileVersion)"

    $archStr = if ($script:Arch -eq "arm64") { "arm64" } else { "64-bit" }
    $filename = "Git-$($release.FileVersion)-$archStr.exe"
    $tmpPath = Join-Path $env:TEMP "openclaw-install"
    $tmpFile = Join-Path $tmpPath $filename

    New-Item -ItemType Directory -Force -Path $tmpPath | Out-Null

    $downloaded = Download-File -Dest $tmpFile -Urls @(
        "https://registry.npmmirror.com/-/binary/git-for-windows/$($release.Tag)/$filename",
        "https://github.com/git-for-windows/git/releases/download/$($release.Tag)/$filename"
    )

    if (-not $downloaded) {
        Write-Err "Git download failed. Check network connection."
        Remove-Item $tmpPath -Recurse -Force -ErrorAction SilentlyContinue
        return $false
    }

    Write-Info "Installing Git silently..."
    try {
        Start-Process -FilePath $tmpFile -ArgumentList "/VERYSILENT","/NORESTART","/NOCANCEL","/SP-","/CLOSEAPPLICATIONS","/RESTARTAPPLICATIONS" -Wait
        Refresh-PathEnv

        $ver = Get-GitVersion
        if ($ver) {
            Write-Ok "$ver installed successfully."
            Remove-Item $tmpPath -Recurse -Force -ErrorAction SilentlyContinue
            return $true
        }
    } catch {
        Write-Err "Git installation failed: $_"
    }

    Remove-Item $tmpPath -Recurse -Force -ErrorAction SilentlyContinue
    return $false
}

# ── Main installation steps ──

function Test-NvmInstalled {
    try { $null = & cmd /c "nvm version" 2>$null; return $true } catch {}
    try { Get-Command nvm -ErrorAction Stop | Out-Null; return $true } catch {}
    return $false
}

function Test-NvmNodeActive {
    param([int]$Major)
    try {
        $list = & cmd /c "nvm list" 2>$null
        if ($list -match "\*\s+$Major\.") { return $true }
    } catch {}
    return $false
}

function Use-NodeV22Dir {
    param([string]$Dir)
    $script:NodeBinDir = $Dir
    $rest = ($env:PATH.Split(";") | Where-Object { $_ -ne $Dir }) -join ";"
    $env:PATH = "$Dir;$rest"
    Add-ToUserPath $Dir
    Ensure-NodePriority -NodeV22Dir $Dir
}

function Step-CheckNode {
    Write-Step "Step 1/7: Prepare Node.js environment"

    $hasNvm = Test-NvmInstalled

    # If nvm is present, try to manage Node version via nvm first.
    if ($hasNvm) {
        Write-Info "Detected nvm-windows..."

        # Check if nvm already has v22+ active
        if (Test-NvmNodeActive -Major 22) {
            $ver = Get-NodeVersion
            if ($ver) {
                Write-Ok "Node.js $ver already active via nvm (>= 22)."
                Pin-NodePath
                $script:NvmManaged = $true
                return $true
            }
        }

        # nvm does not have v22 active; attempt to install and switch
        if (Install-NodeViaNvm) {
            Pin-NodePath
            return $true
        }

        # nvm switching failed; try to use existing v22 (previously installed directly)
        Write-Warn "Failed to switch to Node.js v22 via nvm (usually needs admin rights)."
        Write-Info "Looking for other available Node.js v22..."
    }

    # Check previously directly installed path
    $scriptInstallDir = Join-Path (Get-LocalAppData) "nodejs"
    $scriptNodeExe = Join-Path $scriptInstallDir "node.exe"
    if (Test-Path $scriptNodeExe) {
        $ver = Get-NodeVersion -NodeExe $scriptNodeExe
        if ($ver) {
            Write-Ok "Node.js $ver already installed (>= 22)."
            Use-NodeV22Dir $scriptInstallDir
            return $true
        }
    }

    # Look for qualifying version in PATH
    $ver = Get-NodeVersion
    if ($ver) {
        Write-Ok "Node.js $ver already installed (>= 22)."
        Pin-NodePath
        if ($script:NodeBinDir) { Ensure-NodePriority -NodeV22Dir $script:NodeBinDir }
        return $true
    }

    $existingVer = try { & node -v 2>$null } catch { $null }
    if ($existingVer) {
        Write-Warn "Detected Node.js $existingVer, which is too old. Version 22+ required."
    } else {
        Write-Warn "Node.js not detected."
    }

    Write-Info "Automatically installing Node.js v22..."
    if (Install-NodeDirect) {
        Pin-NodePath
        if ($script:NodeBinDir) { Ensure-NodePriority -NodeV22Dir $script:NodeBinDir }
        return $true
    }

    Write-Err "All installation methods failed. Please check network connection and try again."
    if ($hasNvm) {
        Write-Host ""
        Write-Host "  Suggestion: Open PowerShell as Administrator and run:" -ForegroundColor Yellow
        Write-Host "    nvm install 22" -ForegroundColor Cyan
        Write-Host "    nvm use 22" -ForegroundColor Cyan
        Write-Host "  Then re‑run this installer." -ForegroundColor Yellow
    }
    return $false
}

function Step-CheckGit {
    Write-Step "Step 2/7: Prepare Git environment"

    $ver = Get-GitVersion
    if ($ver) {
        Write-Ok "$ver already installed."
        return $true
    }

    Write-Warn "Git not detected. Automatically installing..."

    if (Install-GitDirect) { return $true }
    if (Install-GitViaWinget) { return $true }

    Write-Err "Automatic Git installation failed. Please install Git manually and retry."
    Write-Host "  Download: https://git-scm.com/downloads"
    return $false
}

function Step-SetMirror {
    Write-Step "Step 3/7: Set npm mirror (China)"

    $env:npm_config_registry = "https://registry.npmmirror.com"
    Write-Ok "npm registry temporarily set to https://registry.npmmirror.com (only for this installation)."
    return $true
}

function Step-InstallPnpm {
    Write-Step "Step 4/7: Install pnpm"

    $pnpmCmd = Get-PnpmCmd
    try {
        $pnpmVer = (& $pnpmCmd -v 2>$null).Trim()
        if ($pnpmVer) {
            Write-Ok "pnpm $pnpmVer already installed. Skipping installation."
            Ensure-PnpmHome
            return $true
        }
    } catch {}

    $npmCmd = Get-NpmCmd
    Write-Info "Installing pnpm..."
    try {
        $savedEAP = $ErrorActionPreference
        $ErrorActionPreference = "SilentlyContinue"
        & $npmCmd install -g pnpm 2>$null | Out-Null
        $npmExit = $LASTEXITCODE
        $ErrorActionPreference = $savedEAP
        if ($npmExit -ne 0) { throw "npm install -g pnpm failed (exit code: $npmExit)" }
        $pnpmCmd = Get-PnpmCmd
        Write-Info "Verifying pnpm installation..."
        $pnpmVer = (& $pnpmCmd -v 2>$null).Trim()
        Write-Ok "pnpm $pnpmVer installed successfully."

        Write-Info "Configuring pnpm global path (pnpm setup)..."
        try { & $pnpmCmd setup 2>$null | Out-Null } catch { Write-Warn "pnpm setup did not complete successfully; continuing anyway." }

        Ensure-PnpmHome
        return $true
    } catch {
        Write-Err "pnpm installation failed: $_"
        return $false
    }
}

function Run-PnpmInstall {
    param([string]$PnpmCmd, [string]$Label = "Install")

    try {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = "cmd.exe"
        $psi.Arguments = "/c `"$PnpmCmd`" add -g openclaw@latest"
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.CreateNoWindow = $true

        # Ensure child process uses the correct Node.js version
        if ($script:NodeBinDir) {
            $childPath = $env:PATH
            $cleanParts = $childPath.Split(";") | Where-Object {
                if (-not $_) { return $false }
                $nodeInDir = Join-Path $_ "node.exe"
                if ((Test-Path $nodeInDir) -and ($_ -ne $script:NodeBinDir)) { return $false }
                return $true
            }
            $psi.EnvironmentVariables["PATH"] = "$($script:NodeBinDir);$($cleanParts -join ';')"
        }

        $proc = [System.Diagnostics.Process]::Start($psi)
        $stdoutTask = $proc.StandardOutput.ReadToEndAsync()
        $stderrTask = $proc.StandardError.ReadToEndAsync()
    } catch {
        Write-Err "Failed to start $Label process: $_"
        return @{ Success = $false; Stderr = ""; Stdout = "" }
    }

    $progress = 0
    $width = 30
    while (-not $proc.HasExited) {
        if ($progress -lt 30) { $progress += 3 }
        elseif ($progress -lt 60) { $progress += 2 }
        elseif ($progress -lt 90) { $progress += 1 }
        if ($progress -gt 90) { $progress = 90 }
        $filled = [math]::Floor($progress * $width / 100)
        $empty = $width - $filled
        $bar = ([string]::new([char]0x2588, $filled)) + ([string]::new([char]0x2591, $empty))
        Write-Host "`r  $Label progress [$bar] $($progress.ToString().PadLeft(3))%" -NoNewline
        Start-Sleep -Seconds 1
    }

    $stdout = $stdoutTask.GetAwaiter().GetResult()
    $stderr = $stderrTask.GetAwaiter().GetResult()

    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8

    $fullBar = [string]::new([char]0x2588, $width)
    if ($proc.ExitCode -eq 0) {
        Write-Host "`r  $Label progress [$fullBar] 100%"
        return @{ Success = $true; Stderr = $stderr; Stdout = $stdout }
    }

    Write-Host "`r  $Label progress [$fullBar] failed"
    return @{ Success = $false; Stderr = $stderr; Stdout = $stdout; ExitCode = $proc.ExitCode }
}

function Step-InstallOpenClaw {
    Write-Step "Step 5/7: Install OpenClaw"

    $gitVer = Get-GitVersion
    if (-not $gitVer) {
        Write-Err "Git is unavailable. OpenClaw dependencies require Git to resolve."
        Write-Host "  Please install Git first: https://git-scm.com/downloads" -ForegroundColor Yellow
        return $false
    }

    Write-Info "Installing OpenClaw, please wait..."

    $pnpmCmd = Get-PnpmCmd
    if (-not (Test-Path $pnpmCmd -ErrorAction SilentlyContinue)) {
        try { Get-Command $pnpmCmd -ErrorAction Stop | Out-Null } catch {
            Write-Err "Cannot find pnpm command."
            return $false
        }
    }

    # Temporary git URL rewrite rules (does not modify user's git config)
    $env:GIT_CONFIG_COUNT = "2"
    $env:GIT_CONFIG_KEY_0 = "url.https://github.com/.insteadOf"
    $env:GIT_CONFIG_VALUE_0 = "git+ssh://git@github.com/"
    $env:GIT_CONFIG_KEY_1 = "url.https://github.com/.insteadOf"
    $env:GIT_CONFIG_VALUE_1 = "ssh://git@github.com/"

    function Try-InstallWithCleanup([string]$PnpmCmd, [ref]$Result) {
        $combinedOutput = "$($Result.Value.Stderr)`n$($Result.Value.Stdout)"
        $isPnpmStoreError = $combinedOutput -match "VIRTUAL_STORE_DIR" -or $combinedOutput -match "broken lockfile" -or $combinedOutput -match "not compatible with current pnpm"
        if ($isPnpmStoreError) {
            Write-Warn "Detected incompatible pnpm global store state. Cleaning up and retrying..."
            $pnpmGlobalDir = Join-Path (Get-LocalAppData) "pnpm\global"
            if (Test-Path $pnpmGlobalDir) {
                Remove-Item $pnpmGlobalDir -Recurse -Force -ErrorAction SilentlyContinue
                Write-Info "Cleaned $pnpmGlobalDir"
            }
            try { & $PnpmCmd store prune 2>$null } catch {}
            $retryResult = Run-PnpmInstall -PnpmCmd $PnpmCmd -Label "Retry install"
            $Result.Value = $retryResult
            return $retryResult.Success
        }
        return $false
    }

    function Clear-GitConfigEnv {
        Remove-Item Env:GIT_CONFIG_COUNT -ErrorAction SilentlyContinue
        for ($i = 0; $i -lt 2; $i++) {
            Remove-Item "Env:GIT_CONFIG_KEY_$i" -ErrorAction SilentlyContinue
            Remove-Item "Env:GIT_CONFIG_VALUE_$i" -ErrorAction SilentlyContinue
        }
    }

    function On-InstallSuccess {
        Clear-GitConfigEnv
        Write-Ok "OpenClaw installation completed."
        Refresh-PathEnv
        Ensure-PnpmHome
        Ensure-ExecutionPolicy
        $found = Find-OpenclawBinary
        if ($found) {
            Add-ToUserPath $found.Dir
            Write-Info "OpenClaw installed to: $($found.Dir)"
        } else {
            try {
                $pnpmBin = (& $pnpmCmd bin -g 2>$null).Trim()
                if ($pnpmBin -and (Test-Path $pnpmBin)) {
                    Add-ToUserPath $pnpmBin
                    Write-Info "pnpm global bin directory: $pnpmBin"
                }
            } catch {}
        }
        return $true
    }

    # ── GitHub mirror helpers ──

    function Set-GitMirror([string]$Mirror) {
        $env:GIT_CONFIG_COUNT = "3"
        $env:GIT_CONFIG_KEY_2 = "url.${Mirror}.insteadOf"
        $env:GIT_CONFIG_VALUE_2 = "https://github.com/"
    }

    function Clear-GitMirror {
        $env:GIT_CONFIG_COUNT = "2"
        Remove-Item Env:GIT_CONFIG_KEY_2 -ErrorAction SilentlyContinue
        Remove-Item Env:GIT_CONFIG_VALUE_2 -ErrorAction SilentlyContinue
    }

    function Install-WithMirrors {
        Write-Warn "You chose to use third‑party GitHub mirrors. Be aware of potential risks."
        $gitHubMirrors = @(
            "https://bgithub.xyz/",
            "https://kkgithub.com/",
            "https://github.ur1.fun/",
            "https://ghproxy.net/https://github.com/",
            "https://gitclone.com/github.com/"
        )

        # Probe mirrors concurrently, filter available ones
        Write-Info "Probing for available mirrors..."
        $available = @()
        $jobs = @()
        foreach ($m in $gitHubMirrors) {
            $testUrl = $m.TrimEnd('/') + "/"
            $jobs += @{ Mirror = $m; Request = $null }
            try {
                $req = [System.Net.HttpWebRequest]::Create($testUrl)
                $req.Method = "HEAD"
                $req.Timeout = 6000
                $req.AllowAutoRedirect = $true
                $jobs[-1].Request = $req
            } catch {}
        }
        foreach ($j in $jobs) {
            if (-not $j.Request) { continue }
            try {
                $resp = $j.Request.GetResponse()
                $resp.Close()
                $available += $j.Mirror
                Write-Ok "Mirror available: $($j.Mirror)"
            } catch {
                Write-Warn "Mirror unavailable: $($j.Mirror)"
            }
        }

        if ($available.Count -eq 0) {
            Write-Err "No mirrors are reachable."
            return $null
        }

        foreach ($mirror in $available) {
            try {
                Set-GitMirror $mirror
                Write-Info "Installing using mirror $mirror..."
                $r = Run-PnpmInstall -PnpmCmd $pnpmCmd -Label "Install"
                Clear-GitMirror
                if ($r.Success) { return $r }
                $rr = $r
                if (Try-InstallWithCleanup $pnpmCmd ([ref]$rr)) { return $rr }
            } catch {
                Clear-GitMirror
            }
        }
        return $null
    }

    function Show-GitHubChoiceMenu {
        Write-Host ""
        Write-Host "  Some dependencies must be downloaded from GitHub, but GitHub is currently unreachable." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "  Choose an option:" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "   1) Use a community GitHub mirror (may contain altered content)" -ForegroundColor White
        Write-Host "   2) Configure a proxy and retry (recommended if you have a proxy)" -ForegroundColor White
        Write-Host "   0) Exit installation" -ForegroundColor White
        Write-Host ""
        return (Read-Host "  Enter choice [0-2]").Trim()
    }

    function Show-ProxyGuide {
        Write-Host ""
        Write-Host "  After enabling your proxy tool, set git proxy with a command like:" -ForegroundColor Yellow
        Write-Host "    git config --global http.https://github.com.proxy http://127.0.0.1:7890" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "  Then re‑run this installer. After installation you can unset it:" -ForegroundColor Yellow
        Write-Host "    git config --global --unset http.https://github.com.proxy" -ForegroundColor Cyan
    }

    # ── Installation ──

    $result = Run-PnpmInstall -PnpmCmd $pnpmCmd -Label "Install"
    if ($result.Success) { return (On-InstallSuccess) }
    if (Try-InstallWithCleanup $pnpmCmd ([ref]$result)) { return (On-InstallSuccess) }

    # Direct installation also failed; offer fallback choices.
    $combinedOutput = "$($result.Stderr)`n$($result.Stdout)"
    $isGitHubError = $combinedOutput -match "github\.com" -and ($combinedOutput -match "git ls-remote" -or $combinedOutput -match "fatal:" -or $combinedOutput -match "Could not resolve" -or $combinedOutput -match "timed out")

    if ($isGitHubError) {
    Write-Warn "Direct installation failed because GitHub is unreachable. Trying mirrors automatically..."
    $mirrorResult = Install-WithMirrors
    if ($mirrorResult -and $mirrorResult.Success) { return (On-InstallSuccess) }
    Write-Err "All mirrors failed."
    Clear-GitConfigEnv
    return $false
}

    # Generic non‑GitHub error
    Clear-GitConfigEnv
    Write-Err "OpenClaw installation failed (exit code: $($result.ExitCode))"
    if ($result.Stderr) {
        Write-Err "Error details:"
        $result.Stderr.Trim().Split("`n") | ForEach-Object { Write-Host "         $_" -ForegroundColor Red }
    }
    if ($result.Stdout) {
        Write-Info "Installation output (last 15 lines):"
        $result.Stdout.Trim().Split("`n") | Select-Object -Last 15 | ForEach-Object { Write-Host "         $_" }
    }
    return $false
}

function Step-Verify {
    Write-Step "Step 6/7: Verify installation"

    Refresh-PathEnv
    Ensure-PnpmHome
    Ensure-ExecutionPolicy

    $found = Find-OpenclawBinary
    if ($found) {
        $binDir = $found.Dir
        if ($env:PATH -notlike "*$binDir*") {
            $env:PATH = "$binDir;$env:PATH"
        }
        Add-ToUserPath $binDir
        Write-Info "OpenClaw installed at: $binDir"

        $ver = $null
        try { $ver = (& $found.Path -v 2>$null).Trim() } catch {}
        if ($ver) {
            Write-Ok "OpenClaw $ver installed successfully!"
            Write-Host "`n  Congratulations! Your lobster is ready!`n" -ForegroundColor Green
            return $true
        }
    }

    # Fallback: pnpm bin -g
    try {
        $pnpmCmd = Get-PnpmCmd
        $pnpmBin = (& $pnpmCmd bin -g 2>$null).Trim()
        if ($pnpmBin -and (Test-Path $pnpmBin)) {
            $env:PATH = "$pnpmBin;$env:PATH"
            Add-ToUserPath $pnpmBin
            Write-Info "Added pnpm global bin directory to PATH: $pnpmBin"

            $openclawCmd = Join-Path $pnpmBin "openclaw.cmd"
            if (Test-Path $openclawCmd) {
                $ver = $null
                try { $ver = (& $openclawCmd -v 2>$null).Trim() } catch {}
                if ($ver) {
                    Write-Ok "OpenClaw $ver installed successfully!"
                    Write-Host "`n   Congratulations! Your lobster is ready!`n" -ForegroundColor Green
                    return $true
                }
            }
        }
    } catch {}

    Write-Err "Installation completed but cannot locate openclaw executable."
    Write-Host ""
    Write-Host "  Troubleshooting steps:" -ForegroundColor Yellow
    Write-Host "    1. Close this terminal, open a new PowerShell window." -ForegroundColor Yellow
    Write-Host "    2. Run `openclaw -v` to check if it works." -ForegroundColor Yellow
    Write-Host "    3. If not, run `pnpm bin -g` to find the pnpm global bin directory." -ForegroundColor Cyan
    Write-Host "    4. Add that directory to your system PATH manually." -ForegroundColor Yellow
    Write-Host ""
    return $false
}

# ========== Modified: skip AI configuration automatically ==========
function Step-Onboard {
    Write-Step "Step 7/7: Configure OpenClaw"
    Write-Info "Skipping AI configuration automatically. To configure manually, run 'openclaw onboard'."
    return $true
}
# ====================================================================

# ── Main function ──

function Main {
    Write-Host ""
    Write-Host "   OpenClaw One‑Click Installer (Auto‑skip configuration)" -ForegroundColor Green
    Write-Host "  ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Blue
    Write-Host ""

    Refresh-PathEnv

    # Check if already installed
    $existingVer = $null
    $found = Find-OpenclawBinary
    if ($found) {
        try { $existingVer = (& $found.Path -v 2>$null).Trim() } catch {}
    }
    if (-not $existingVer) {
        try { $existingVer = (& openclaw -v 2>$null).Trim() } catch {}
    }
    if ($existingVer) {
        if ($found) { Add-ToUserPath $found.Dir }
        Ensure-ExecutionPolicy
    Write-Ok "OpenClaw $existingVer is already installed. No need to reinstall."
    Write-Host "`n Your lobster is ready!`n" -ForegroundColor Green
    return
    }

    if (-not (Step-CheckNode))       { Write-Host "`nPress any key to exit..."; $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown"); return }
    if (-not (Step-CheckGit))        { Write-Host "`nPress any key to exit..."; $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown"); return }
    if (-not (Step-SetMirror))       { Write-Host "`nPress any key to exit..."; $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown"); return }
    if (-not (Step-InstallPnpm))     { Write-Host "`nPress any key to exit..."; $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown"); return }
    if (-not (Step-InstallOpenClaw)) { Write-Host "`nPress any key to exit..."; $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown"); return }
    if (-not (Step-Verify))          { Write-Host "`nPress any key to exit..."; $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown"); return }
    Step-Onboard | Out-Null

    Write-Host ""
    Write-Host "  ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Green
    Write-Host "  Installation complete!" -ForegroundColor Green
    Write-Host "  ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Green
    Write-Host ""
}

Main

# Refresh PATH for the current process after installation
$machinePath = [Environment]::GetEnvironmentVariable("PATH", "Machine")
$userPath = [Environment]::GetEnvironmentVariable("PATH", "User")
$env:PATH = "$userPath;$machinePath"