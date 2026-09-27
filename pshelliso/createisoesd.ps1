#Requires -Version 5.1
#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Slipstream VirtIO drivers into Windows ISOs, convert install.wim to a solid
    install.esd, and build matching standard + unattended ISOs.

.DESCRIPTION
    Uses wimlib apply/capture ("no-mount") instead of DISM /Mount-Wim, which avoids
    the DISM mount corruption that breaks the classic mount-and-service workflow.

    DISM is still used for driver injection, but only against the offline extracted
    directory (dism.exe /Image:<dir>), never against a mounted WIM.

    boot.wim (WinPE) receives only the boot-critical VirtIO drivers (viostor,
    vioscsi, netkvm): other drivers bloat WinPE and can crash Setup. The full driver
    set is injected into every image index of install.wim/install.esd, and the
    unattended ISO autoinstalls the edition selected by $PreferredEditionPatterns
    (e.g. Server Standard with Desktop Experience, or Windows 11 Enterprise) with the
    matching KMS client setup key.

    Image metadata comes from the WIM XML via `wimlib-imagex info --extract-xml`,
    so this script does not depend on the Dism PowerShell module (Get-WindowsImage),
    which is not available when the script is run from PowerShell 7.
matching KMS client setup key.

.PARAMETER TestInQemu
    After each unattended ISO is verified, create a fresh qcow2 and boot the ISO in
    QEMU so the unattended install can be watched end to end. A small raw marker disk
    is attached alongside it and the guest writes its result there, so the run reports
    whether the unattended install actually completed rather than just whether QEMU
    started. This needs a host that can virtualise (WHPX); on any other host the test
    is skipped with a warning, because software emulation is far too slow to install
    Windows.

.EXAMPLE
    .\createisoesd.ps1
    Builds a standard and an unattended ISO for every .iso in $SourceIsoFolder.

.EXAMPLE
    .\createisoesd.ps1 -TestInQemu
    The same, then boots each unattended ISO in QEMU on a new virtio disk.

.NOTES
    Run from an elevated PowerShell 5.1 or 7 session.

    Runtime is dominated by wimlib: every image index is applied to disk, serviced,
    and re-captured. Converting a multi-edition 8 GB Server image to a solid ESD can
    take well over an hour, and solid capture is memory hungry (see --solid-chunk-size).

    -TestInQemu needs QEMU (qemu-system-x86_64.exe and qemu-img.exe) and a host that
    can virtualise, so it is normally used on a different machine from the one that
    just needs to build ISOs. Without either, the ISOs still build and the boot test is
    skipped with a warning.
#>

[CmdletBinding()]
param(
    # After each unattended ISO is verified, create a fresh qcow2 and boot it in QEMU
    # so the unattended install can be watched end to end. QEMU is only used for that
    # test, so a missing install downgrades to a warning.
    [Parameter()][switch]$TestInQemu
)

Set-StrictMode -Off
$ErrorActionPreference = 'Stop'

# PowerShell 7.3+ can promote native stderr into a terminating error. This script
# inspects exit codes explicitly, so keep that behaviour off for 5.1 and 7 alike.
$PSNativeCommandUseErrorActionPreference = $false

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# ====================================================================================
# Configuration
# ====================================================================================
$SourceIsoFolder    = 'C:\ISOs\win'
$OutputIsoFolder    = 'C:\ISO_Output'
$WorkDir            = 'C:\ISO_WorkDir'
$MountDir           = 'C:\ISO_Mount'

$VirtIoBaseDir      = 'C:\VirtIO'
$VirtIoDrivers      = Join-Path $VirtIoBaseDir 'Drivers'
$VirtIoDriversAmd64 = Join-Path $VirtIoBaseDir 'Drivers-amd64'
$VirtIoBootDrivers  = Join-Path $VirtIoBaseDir 'Drivers-boot'
$VirtIoMsiPath      = Join-Path $VirtIoBaseDir 'MSI\virtio-win-gt-x64.msi'
$QemuGaMsiPath      = Join-Path $VirtIoBaseDir 'MSI\qemu-ga-x86_64.msi'

$WimlibDir          = 'C:\Tools\wimlib'
$WimlibPath         = Join-Path $WimlibDir 'wimlib-imagex.exe'

$TemplateXmlPath    = 'C:\vscode\iso-cbase-forge2af\pshelliso\unattend\autounattend.xml'

# ADK Deployment Tools. These are the "version-less" ADK paths; Resolve-Tool also
# probes versioned folders (Windows Kits\10\...) and PATH before giving up.
$AdkDeploymentTools = 'C:\Program Files (x86)\Windows Kits\Assessment and Deployment Kit\Deployment Tools'
$OscdimgPath        = Join-Path $AdkDeploymentTools 'amd64\Oscdimg\oscdimg.exe'
$DismPath           = Join-Path $AdkDeploymentTools 'amd64\DISM\dism.exe'

$StandardIsoLabel   = 'CUST_ISO'
$UnattendIsoLabel   = 'CUST_UNAT'

# --- Answer-file values. These are written to autounattend.xml in clear text.
#     Keep this script off untrusted shares, or replace them with prompted values.
$AdminPassword      = 'P@ssw0rdAdmin!'
$LocalUsername      = 'ITAdmin'
$LocalUserPassword  = 'P@ssw0rdUser!'
$StandardUserAgent  = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'

# --- boot.wim (WinPE) receives ONLY these drivers. Extra drivers in WinPE bloat it
#     and can crash Setup; the full driver set still goes into every install image.
$BootWimDriverNames = @('viostor', 'vioscsi', 'netkvm')

# --- Edition the unattended ISO autoinstalls, tried in order for every source ISO.
#     The first image whose name matches is selected; if nothing matches, the first
#     image is used with a warning. Choose the entry you want:
#       'Standard.*Desktop' -> Windows Server <ver> Standard (Desktop Experience)
#       'Enterprise'        -> Windows 11 / 10 Enterprise
$PreferredEditionPatterns = @(
    'Standard.*Desktop'
    'Enterprise'
    'Standard'
    'Datacenter.*Desktop'
    'Pro'
)

# --- QEMU boot test (-TestInQemu). QEMU is only used by that test, so a missing
#     install is a warning and never a build failure.
$QemuSystemPath     = 'C:\Program Files\qemu\qemu-system-x86_64.exe'
$QemuImgPath        = 'C:\Program Files\qemu\qemu-img.exe'
$QemuVmDir          = 'C:\ISO_Output\qemu'
$QemuDiskSizeGb     = 60
$QemuMemoryMb       = 8192
$QemuCpuCount       = 4
# 'whpx' is the Windows Hypervisor Platform (fast; needs Hyper-V/WHPX enabled).
# Set it to '' to run unaccelerated on a host without WHPX.
$QemuAcceleration   = 'whpx'
# Software emulation can boot a VM on a host with no hypervisor at all, but installing
# Windows that way takes many hours, so -TestInQemu is skipped instead unless this is
# set to $true.
$QemuAllowSoftwareEmulation = $false
# $true keeps the script attached until the QEMU window is closed, so one ISO is
# finished before the next is built. $false detaches the VM and carries on, which also
# means the guest result cannot be read back (the marker needs QEMU to have exited).
$QemuWaitForExit    = $true
# Give up waiting on a forgotten VM after this long; 0 waits forever.
$QemuBootTimeoutMinutes = 120

# Handshake for the marker disk described above. It must fit in 16 bytes and has to
# match the value baked into the guest payload further down, so change it in one place
# only if you really need to.
$GuestMarkerMagic   = 'UNATTEND-TEST-V1'

# ====================================================================================
# Helpers
# ====================================================================================
function Write-Header {
    param([Parameter(Mandatory)][string]$Text)
    Write-Host ''
    Write-Host ('=' * 64) -ForegroundColor Cyan
    Write-Host $Text -ForegroundColor Cyan
    Write-Host ('=' * 64) -ForegroundColor Cyan
}

function Write-Step {
    param([Parameter(Mandatory)][string]$Text)
    Write-Host $Text -ForegroundColor Green
}

function Write-Detail {
    param([Parameter(Mandatory)][string]$Text)
    Write-Host "      -> $Text" -ForegroundColor DarkGray
}

<#
    Runs a native executable with an explicit argument array (never a single
    concatenated command string) and fails loudly on an unexpected exit code.

    Child stdout/stderr is passed through so long-running tools show progress.
    The exit code is left in $script:LastNativeExitCode for callers that need it.
#>
function Invoke-Native {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$FilePath,
        [Parameter()][string[]]$Arguments = @(),
        [Parameter()][int[]]$AllowedExitCodes = @(0),
        [Parameter()][string]$Activity = ''
    )

    if ([string]::IsNullOrWhiteSpace($Activity)) {
        $Activity = Split-Path -Path $FilePath -Leaf
    }

    Write-Verbose ("Executing: {0} {1}" -f $FilePath, ($Arguments -join ' '))

    $global:LASTEXITCODE = 0
    & $FilePath @Arguments
    $exitCode = $LASTEXITCODE
    $script:LastNativeExitCode = $exitCode

    if ($AllowedExitCodes -notcontains $exitCode) {
        throw ("{0} failed with exit code {1}.{2}Command: {3} {4}" -f `
                $Activity, $exitCode, [Environment]::NewLine, $FilePath, ($Arguments -join ' '))
    }
}

<#
    Resolves a tool that may live in a versioned ADK folder, a configured path, or PATH.
    $Candidates may contain literal paths or bare command names.
#>
function Resolve-Tool {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter()][string[]]$Candidates = @(),
        [Parameter()][switch]$Required
    )

    foreach ($candidate in @($Candidates) + @($Name)) {
        if ([string]::IsNullOrWhiteSpace($candidate)) { continue }

        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return (Resolve-Path -LiteralPath $candidate).Path
        }

        $onPath = @(Get-Command -Name $candidate -CommandType Application -ErrorAction SilentlyContinue)
        if ($onPath.Count -gt 0) { return $onPath[0].Source }
    }

    if ($Required) {
        throw "$Name was not found. Checked: $(@($Candidates) -join '; ') and PATH."
    }
    return $null
}

<#
    Reads a response header across PowerShell versions. PS 5.1 exposes Headers as a
    dictionary; PS 7 exposes HttpResponseHeaders, which needs GetValues().
#>
function Get-ResponseHeader {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Response,
        [Parameter(Mandatory)][string]$Name
    )

    try {
        $value = $Response.Headers[$Name]
        if ($value) { return [string]@($value)[0] }
    } catch { }

    try {
        $value = $Response.Headers.GetValues($Name)
        if ($value) { return [string]@($value)[0] }
    } catch { }

    return $null
}

<#
    Cheap structural check that a file really is an ISO9660 image: the primary volume
    descriptor carries the ASCII identifier "CD001" at byte offset 32769. This exists to
    stop a poisoned download (an anti-bot challenge page) from being treated as media.
#>
function Test-IsoImage {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    try {
        if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $false }
        if ((Get-Item -LiteralPath $Path).Length -lt 1MB) { return $false }

        $stream = [System.IO.File]::OpenRead($Path)
        try {
            [void]$stream.Seek(32769, [System.IO.SeekOrigin]::Begin)
            $buffer = New-Object byte[] 5
            if ($stream.Read($buffer, 0, 5) -ne 5) { return $false }
            return ([System.Text.Encoding]::ASCII.GetString($buffer) -eq 'CD001')
        } finally {
            $stream.Dispose()
        }
    } catch {
        Write-Verbose "ISO check failed for '$Path': $($_.Exception.Message)"
        return $false
    }
}

# Returns { Path = ...; Image = <MSFT_DiskImage>; Drive = 'D'; Root = 'D:\' }.
function Mount-IsoImage {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    $diskImage = Mount-DiskImage -ImagePath $Path -PassThru -ErrorAction Stop

    # The volume is not always ready the instant the image mounts.
    $driveLetter = $null
    for ($attempt = 0; $attempt -lt 20 -and -not $driveLetter; $attempt++) {
        Start-Sleep -Milliseconds 250
        $volume = $diskImage | Get-Volume -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($volume -and $volume.DriveLetter) { $driveLetter = [string]$volume.DriveLetter }
    }

    if (-not $driveLetter) {
        try { Dismount-DiskImage -ImagePath $Path -ErrorAction SilentlyContinue | Out-Null } catch { }
        throw "Mounted '$Path' but Windows never assigned it a drive letter."
    }

    return [pscustomobject]@{
        Path  = $Path
        Image = $diskImage
        Drive = $driveLetter
        Root  = '{0}:\' -f $driveLetter
    }
}

function Dismount-IsoImage {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    try {
        Dismount-DiskImage -ImagePath $Path -ErrorAction Stop | Out-Null
    } catch {
        Write-Warning "Could not dismount '$Path': $($_.Exception.Message)"
    }
}

<#
    Clears ReadOnly/System/Hidden on a tree. Files copied off an ISO (and files
    restored by `wimlib apply`) carry these attributes, and they make later
    deletes and overwrites fail with "Access is denied".
#>
function Clear-ReadOnlyAttribute {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) { return }

    $mask = [System.IO.FileAttributes]::ReadOnly -bor
            [System.IO.FileAttributes]::Hidden   -bor
            [System.IO.FileAttributes]::System

    foreach ($item in @(Get-ChildItem -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue)) {
        if ($item.Attributes -band $mask) {
            try {
                $item.Attributes = [System.IO.FileAttributes]($item.Attributes -band (-bnot $mask))
            } catch {
                Write-Verbose "Could not clear attributes on '$($item.FullName)': $($_.Exception.Message)"
            }
        }
    }
}

<#
    Purges a directory by mirroring an empty folder onto it.

    robocopy is used rather than Remove-Item because a tree extracted from a Windows
    image breaks PowerShell 5.1's recursive delete: it aborts the whole recursion at
    the first ordering hiccup ("The directory is not empty") and leaves most of the
    tree behind, and .NET's path APIs stop at MAX_PATH. robocopy deletes strictly
    bottom-up and is long-path aware.
#>
function Invoke-EmptyMirror {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    $empty = Join-Path $env:TEMP ('empty_' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $empty -Force | Out-Null

    try {
        # Exit codes 0-7 are robocopy's success variants; 8+ means a real failure.
        # Output is discarded because a large purge emits one line per file.
        Invoke-Native -FilePath "$env:SystemRoot\System32\robocopy.exe" `
            -Arguments @($empty, $Path, '/MIR', '/R:1', '/W:1', '/NFL', '/NDL', '/NJH', '/NJS', '/NP') `
            -AllowedExitCodes (0..7) -Activity 'robocopy purge' | Out-Null
    } finally {
        Remove-Item -LiteralPath $empty -Recurse -Force -ErrorAction SilentlyContinue
    }
}

<#
    Deletes a directory tree.

    Layered on purpose. A plain recursive delete is fine (and fast) for the ISO
    staging trees, but a tree extracted from a Windows image needs more:

      1. Remove-Item -Recurse -Force does not create the destination and aborts the
         whole walk on the first hiccup, so it is only the cheap first attempt.
      2. An empty-mirror purge handles the bulk of the tree (long paths, deep
         component-store folders) and is retried, because antivirus can briefly lock
         files that were just created in bulk.
      3. Some image files carry DACLs that deny deletion even to Administrators
         (TrustedInstaller-owned), so ownership is taken and ACLs reset as a last
         resort.

    Anything still present after all of that is reported, never ignored. -BestEffort
    downgrades that report to a warning, which is what the staging cleanup uses: a
    directory that cannot be removed is not a reason to abandon an ISO build.
#>
function Remove-DirectoryTree {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter()][switch]$BestEffort
    )

    if (-not (Test-Path -LiteralPath $Path)) { return }

    $resolved = (Resolve-Path -LiteralPath $Path).Path.TrimEnd('\')
    if ($resolved -match '^[A-Za-z]:$') {
        throw "Refusing to remove a drive root: $resolved"
    }

    # Attempt 1: cheap recursive delete.
    Remove-Item -LiteralPath $resolved -Recurse -Force -ErrorAction SilentlyContinue
    if (-not (Test-Path -LiteralPath $resolved)) { return }

    Clear-ReadOnlyAttribute -Path $resolved
    Write-Detail ("Purging '{0}' (large image tree; this can take a while)..." -f $resolved)

    # Attempt 2: empty-mirror purge, retried to ride out transient locks.
    for ($attempt = 1; $attempt -le 2; $attempt++) {
        if ($attempt -gt 1) { Start-Sleep -Seconds 3 }

        try {
            Invoke-EmptyMirror -Path $resolved
        } catch {
            Write-Verbose "Empty-mirror purge reported a problem: $($_.Exception.Message)"
        }

        Remove-Item -LiteralPath $resolved -Recurse -Force -ErrorAction SilentlyContinue
        if (-not (Test-Path -LiteralPath $resolved)) { return }
    }

    # Attempt 3: take ownership and reset ACLs, then purge again.
    try {
        Invoke-Native -FilePath "$env:SystemRoot\System32\takeown.exe" `
            -Arguments @('/F', $resolved, '/R', '/D', 'Y') -Activity 'takeown'
        Invoke-Native -FilePath "$env:SystemRoot\System32\icacls.exe" `
            -Arguments @($resolved, '/reset', '/t', '/c', '/q') -Activity 'icacls'
        Invoke-EmptyMirror -Path $resolved
    } catch {
        Write-Verbose "takeown/icacls recovery reported a problem: $($_.Exception.Message)"
    }

    Remove-Item -LiteralPath $resolved -Recurse -Force -ErrorAction SilentlyContinue
    if (-not (Test-Path -LiteralPath $resolved)) { return }

    if ($BestEffort) {
        Write-Warning "Could not remove '$resolved'; leaving it behind and continuing."
        return
    }

    # Enumerated non-recursively on purpose: if the directory already looks empty yet
    # still cannot be deleted, the volume is the problem, not a file lock.
    $sample = @(Get-ChildItem -LiteralPath $resolved -Force -ErrorAction SilentlyContinue |
        Select-Object -First 5 -ExpandProperty Name)
    $detail = if ($sample.Count -gt 0) { $sample -join '; ' } else { '<nothing visible - the volume is refusing the delete>' }

    Write-Warning 'An empty directory that refuses deletion means the volume metadata is inconsistent; run "chkdsk /f" on it (the system volume needs a reboot).'
    throw "Could not remove '$resolved'. Still present (sample): $detail"
}

<#
    Creates an empty staging directory for a single wimlib apply.

    Every apply gets a fresh directory rather than reusing one fixed path. Deleting an
    extracted image tree is not guaranteed to succeed (see Remove-DirectoryTree), and
    applying on top of a partially purged tree would capture the wrong content, so a
    fresh directory makes the outcome of the previous purge irrelevant.
#>
function New-StagingDirectory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Root,
        [Parameter(Mandatory)][string]$Purpose
    )

    if (-not (Test-Path -LiteralPath $Root)) {
        New-Item -ItemType Directory -Path $Root -Force | Out-Null
    }

    $path = Join-Path $Root ('{0}_{1}' -f $Purpose, [guid]::NewGuid().ToString('N').Substring(0, 8))
    New-Item -ItemType Directory -Path $path -Force | Out-Null

    return $path
}

<#
    Best-effort purge of the staging directories left under the mount root.

    A directory that was not deleted after an apply is retried here and at the start of
    every ISO, so one failure does not cost disk space for the rest of the run. Never
    throws - whatever survives is reported by the caller at the end of the run.
#>
function Clear-StagingRoot {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) { return }

    foreach ($child in @(Get-ChildItem -LiteralPath $Path -Force -ErrorAction SilentlyContinue)) {
        Remove-DirectoryTree -Path $child.FullName -BestEffort
    }
}

<#
    Empties the per-ISO staging tree before the base ISO is copied into it.

    Only leftover files matter here: an undeletable *empty* directory changes nothing
    about the output ISO, but a stale boot.wim/install.esd/$OEM$ would be silently
    included in it. So the purge is best effort and the artifacts this script itself
    generates are then verified to be gone.
#>
function Reset-WorkDirectory {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    if (Test-Path -LiteralPath $Path) {
        Remove-DirectoryTree -Path $Path -BestEffort
    }
    New-Item -ItemType Directory -Path $Path -Force | Out-Null

    $generated = @(
        'autounattend.xml'
        'sources\$OEM$'
        'sources\boot.wim'
        'sources\install.wim'
        'sources\install.esd'
        'sources\new_boot.wim'
        'sources\new_install.esd'
    )

    $stale = @($generated |
        ForEach-Object { Join-Path $Path $_ } |
        Where-Object { Test-Path -LiteralPath $_ })

    if ($stale.Count -gt 0) {
        throw ("Could not empty the staging directory '{0}'; these would end up in the output ISO: {1}" -f $Path, ($stale -join '; '))
    }
}

<#
    Reports whether the volume holding $Path has its NTFS dirty bit set.

    Windows can refuse to delete a directory on a dirty volume even when the directory
    is empty (rd reports "The directory is not empty"), which is what makes a staged
    image tree undeletable. Checking up front turns a confusing mid-run failure into an
    actionable warning; repairing it needs an offline chkdsk.
#>
function Test-VolumeDirty {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    $qualifier = Split-Path -Qualifier $Path
    if ([string]::IsNullOrWhiteSpace($qualifier)) { return $false }

    try {
        $output = & "$env:SystemRoot\System32\fsutil.exe" dirty query $qualifier 2>&1
    } catch {
        return $false
    }

    return [bool]($output -match '(?i)is\s+dirty')
}

<#
    Reads image metadata straight out of the WIM XML. Deliberately avoids
    Get-WindowsImage so the script behaves identically under PowerShell 5.1 and 7.
#>
function Get-WimImages {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$WimPath)

    if (-not (Test-Path -LiteralPath $WimPath -PathType Leaf)) {
        throw "WIM/ESD not found: $WimPath"
    }

    $xmlFile = Join-Path $env:TEMP ("wiminfo_{0}.xml" -f [guid]::NewGuid().ToString('N'))
    try {
        Invoke-Native -FilePath $WimlibPath `
            -Arguments @('info', $WimPath, "--extract-xml=$xmlFile") `
            -Activity 'wimlib-imagex info' | Out-Null

        [xml]$xml = [System.IO.File]::ReadAllText($xmlFile)
    } finally {
        Remove-Item -LiteralPath $xmlFile -Force -ErrorAction SilentlyContinue
    }

    $images = @()
    foreach ($node in @($xml.WIM.IMAGE)) {
        $index = [int]$node.INDEX
        $name = $null
        if ($node.NAME) { $name = [string]$node.NAME }
        elseif ($node.DISPLAYNAME) { $name = [string]$node.DISPLAYNAME }
        if ([string]::IsNullOrWhiteSpace($name)) { $name = "Image $index" }

        $images += [pscustomobject]@{ Index = $index; Name = $name }
    }

    if ($images.Count -eq 0) { throw "No images were found in $WimPath" }
    return $images
}

<#
    Stages only the amd64 drivers for operating systems this script targets. The
    virtio-win ISO uses two different layouts and BOTH must be matched:

        <Driver>\<OS>\amd64\driver.inf   - most drivers
        amd64\<OS>\vioscsi.inf           - the storage drivers (boot critical)

    Three properties of the vendor tree break `dism /Add-Driver`, and all are filtered
    out here rather than worked around at the DISM call:

      * x86/ARM64 INFs are rejected outright;
      * the pre-Windows-10 OS folders (xp, 2k3, 2k8, 2k8R2, 2k12, 2k12R2, w7, w8, w8.1)
        hold boot-critical packages whose signatures modern DISM refuses;
      * `smbus` is rejected in every OS folder, including the current ones, because the
        package calls itself boot-critical while carrying no acceptable signature.

    Any one of those aborts the whole call with "Cannot install non-signed boot-critical
    drivers on amd64 images" (exit code 50) and leaves the image half serviced.
    /ForceUnsigned would silence it, but an unsigned boot-start driver is a real boot
    risk, and the guest-tools MSI installs the excluded drivers properly after setup.

    An optional -DriverNames whitelist restricts the stage to specific drivers; the
    boot.wim stage uses it to keep WinPE to viostor/vioscsi/netkvm. All other filters
    still apply.

    Relative structure is preserved so same-named files from different OS folders
    cannot collide.
#>
function New-Amd64DriverStage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DriverRoot,
        [Parameter(Mandatory)][string]$StageRoot,
        [Parameter()][string[]]$DriverNames = @()
    )

    if (-not (Test-Path -LiteralPath $DriverRoot)) {
        throw "VirtIO driver folder not found: $DriverRoot"
    }

    $allInfs = @(Get-ChildItem -LiteralPath $DriverRoot -Recurse -File -Filter '*.inf' -ErrorAction SilentlyContinue)
    if ($allInfs.Count -eq 0) {
        throw "No .inf files found under $DriverRoot"
    }

    $prefix = $DriverRoot.TrimEnd('\') + '\'

    # Pre-Windows-10 driver folders; see the note above.
    $legacyOsPattern = '(?i)(^|\\)({0})(\\|$)' -f ((
        'xp', '2k3', '2k8', '2k8R2', '2k12', '2k12R2', 'w7', 'w8', 'w8.1' |
            ForEach-Object { [regex]::Escape($_) }
    ) -join '|')

    $unsignableDrivers = @('smbus')

    # Match on the path (any 'amd64' directory segment), not just the leaf folder name.
    $sourceDirs = @(
        $allInfs |
            Where-Object {
                $relative = $_.FullName.Substring($prefix.Length)

                ($relative -match '(?i)(^|\\)amd64(\\|$)') -and
                    ($relative -notmatch $legacyOsPattern) -and
                    # The INF is named after its driver in both layouts.
                    ($unsignableDrivers -notcontains $_.BaseName) -and
                    (-not ($DriverNames.Count -gt 0) -or ($DriverNames -contains $_.BaseName))
            } |
            Select-Object -ExpandProperty DirectoryName -Unique
    )

    if ($sourceDirs.Count -eq 0) {
        $scope = if ($DriverNames.Count -gt 0) { "matching drivers ($($DriverNames -join ', '))" } else { 'drivers' }
        throw "No amd64 $scope were found under $DriverRoot; refusing to inject the unfiltered tree, which contains x86 and legacy INFs that DISM rejects."
    }

    Remove-DirectoryTree -Path $StageRoot

    foreach ($dir in $sourceDirs) {
        if (-not $dir.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "Refusing to stage drivers from outside $DriverRoot : $dir"
        }

        $target = Join-Path $StageRoot $dir.Substring($prefix.Length)
        New-Item -ItemType Directory -Path $target -Force | Out-Null
        Copy-Item -Path (Join-Path $dir '*') -Destination $target -Recurse -Force
    }

    Write-Detail ("Staged {0} amd64 driver folder(s) to {1}" -f $sourceDirs.Count, $StageRoot)

    $skipped = @($unsignableDrivers | Where-Object {
        $allInfs.BaseName -contains $_ -and
            (-not ($DriverNames.Count -gt 0) -or $DriverNames -contains $_)
    })
    if ($skipped.Count -gt 0) {
        Write-Warning ("Not injecting {0}: DISM rejects the package as an unsigned boot-critical driver. The VirtIO guest-tools MSI installs it after setup." -f ($skipped -join ', '))
    }

    return $StageRoot
}

function Get-KmsClientSetupKey {
    param([string]$ImageName)

    switch -Regex ($ImageName) {
        '(?i)Server.*?2025.*?Datacenter' { return 'D764K-2NDRG-47T6Q-P8T8W-CWKP7' }
        '(?i)Server.*?2025.*?Standard'   { return 'TVRH6-WHNXV-HTMHW-8WQKV-DFCQV' }
        '(?i)Server.*?2022.*?Datacenter' { return 'WX4NM-KYWYW-QJJR4-XV3QB-6VM33' }
        '(?i)Server.*?2022.*?Standard'   { return 'VDYBN-27WPP-V4HQT-9VMD4-VMK7H' }
        '(?i)Server.*?2019.*?Datacenter' { return 'WMDGN-G9PQG-XVVXX-R3X43-63DFG' }
        '(?i)Server.*?2019.*?Standard'   { return 'N69G4-B89J2-4G8F4-WWYCC-J464C' }
        '(?i)Server.*?2019.*?Essentials' { return 'WVDHN-86M7X-466P6-VHXV7-YY726' }
        '(?i)Server.*?2016.*?Datacenter' { return 'CB7KF-BWN84-R7R2Y-793K2-8XDDG' }
        '(?i)Server.*?2016.*?Standard'   { return 'WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY' }
        '(?i)Server.*?2016.*?Essentials' { return 'JCKRF-N37P4-C2D82-9YXRT-4M63B' }
        '(?i)Pro.*?Workstation'          { return 'NRG8B-VKK3Q-CXVCJ-9G2XF-HKM4' }
        '(?i)Pro.*?Education'            { return '6TP4R-GNPTD-KYYHQ-7B7DP-J447Y' }
        '(?i)Education'                  { return 'NW6C2-QMPVW-D7KKK-3GKT6-VCFB2' }
        '(?i)Enterprise'                 { return 'NPPR9-FWDCX-D2C8J-H872K-2YT43' }
        '(?i)Pro'                        { return 'W269N-WFGWX-YVC9B-4J6C9-T83GX' }
        '(?i)Home|Core'                  { return 'YTMG3-N6DKC-DKB77-7M9GH-8HVX7' }
        default                          { return $null }
    }
}

<#
    Chooses which edition the unattended ISO will autoinstall.

    Server media ships Standard/Datacenter each as Core and Desktop Experience
    images, and client media ships per-SKU images, so "the first image" is rarely the
    one the operator wants. The preferred patterns are tried in order and the first
    image whose name matches is selected. If nothing matches, the first image is used
    and a warning is printed so the fallback is never silent.
#>
function Select-InstallEdition {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object[]]$Images,
        [Parameter(Mandatory)][string[]]$PreferredPatterns
    )

    foreach ($pattern in $PreferredPatterns) {
        foreach ($image in $Images) {
            if ($image.Name -match $pattern) {
                Write-Detail ("Selected edition '{0}' (source index {1}; matched '{2}')" -f $image.Name, $image.Index, $pattern)
                return $image
            }
        }
    }

    $fallback = $Images[0]
    Write-Warning ("No edition matched the preferred patterns ({0}); falling back to '{1}' (source index {2}). Adjust `$PreferredEditionPatterns in the configuration if that is wrong." -f ($PreferredPatterns -join ', '), $fallback.Name, $fallback.Index)
    return $fallback
}

<#
    Reports how usable this host is for an accelerated VM.

    QEMU's -accel whpx is a hard requirement to QEMU rather than a preference, and WHPX
    itself needs hardware virtualisation exposed to Windows, so the two ways this
    fails need different advice:

      'Available'                 - WHPX is on and the CPU can virtualise
      'VirtualisationUnavailable' - the CPU/firmware cannot virtualise, which is the
                                    usual answer inside a VM with no nested
                                    virtualisation. Nothing can be enabled to fix it.
      'FeatureDisabled'           - the CPU can virtualise but the Windows Hypervisor
                                    Platform feature is switched off, which is fixable.
      'Unknown'                   - the checks could not be completed
#>
function Get-HypervisorStatus {
    [CmdletBinding()]
    param()

    # Win32_Processor exposes the same VT-x/AMD-V flags systeminfo reports as "Hyper-V
    # Requirements". Get-ComputerInfo carries equivalents, but they come back as empty
    # strings on some builds (including this script's own test host), so the CIM class
    # is queried directly. An absent value stays unknown rather than becoming a failure.
    $virtualisationOk = $null
    try {
        $processor = Get-CimInstance -ClassName Win32_Processor -ErrorAction Stop | Select-Object -First 1
        if ($null -ne $processor) {
            $firmware = $processor.VirtualizationFirmwareEnabled
            $monitor  = $processor.VMMonitorModeExtensions

            if ($firmware -eq $false -or $monitor -eq $false) { $virtualisationOk = $false }
            elseif ($firmware -eq $true -and $monitor -eq $true) { $virtualisationOk = $true }
        }
    } catch {
        Write-Verbose "Could not read the processor virtualisation flags: $($_.Exception.Message)"
    }

    # Cast to a string on purpose: the State value must not be tested for truthiness,
    # because a state that happens to be zero would read as "absent".
    $featureState = ''
    try {
        $featureState = [string](Get-WindowsOptionalFeature -Online -FeatureName 'HypervisorPlatform' -ErrorAction Stop).State
    } catch {
        Write-Verbose "Could not read the HypervisorPlatform feature: $($_.Exception.Message)"
    }

    if ($virtualisationOk -eq $false) { return 'VirtualisationUnavailable' }
    if ($featureState -eq 'Enabled') { return 'Available' }
    if (-not [string]::IsNullOrWhiteSpace($featureState)) { return 'FeatureDisabled' }
    return 'Unknown'
}

<#
    Reports whether Windows is itself running under a hypervisor, i.e. this machine is
    a VM. Only used to explain the virtualisation verdict: a guest without nested
    virtualisation cannot host another VM, and no Windows setting will change that.
#>
function Test-RunningInVirtualMachine {
    [CmdletBinding()]
    param()

    try {
        return [bool](Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop).HypervisorPresent
    } catch {
        return $false
    }
}

<#
    (Re)creates the raw marker disk the guest reports into.

    It is deliberately a raw byte file rather than a filesystem: the guest writes its
    verdict at a fixed offset and the host reads the same bytes back, so neither side
    needs a driver, a partition table, or a mount.

    The first sector is a contract shared with the guest payload:

        0..15     handshake magic, space padded
        16..511   the guest's report, NUL padded

    The magic stays in place for the life of the disk so the guest can find it again
    when it reports a second time (GUEST-STARTED, then the verdict), and the host can
    tell an untouched disk from one the guest has written to.
#>
function New-GuestMarkerDisk {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    Remove-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue

    # Windows ignores disks below roughly a megabyte, so leave some room. Only the
    # first sector is ever used.
    $bytes = New-Object byte[] (8MB)
    $magic = [System.Text.Encoding]::ASCII.GetBytes($GuestMarkerMagic.PadRight(16, ' '))
    [Array]::Copy($magic, $bytes, [Math]::Min($magic.Length, 16))
    [System.IO.File]::WriteAllBytes($Path, $bytes)

    return $Path
}

<#
    Reads the marker disk and returns the guest's report, or $null when the disk is not
    there, is not ours (magic missing), or cannot be read yet. An empty string means the
    disk is ours but the guest has not written anything. Never throws: while QEMU runs
    it may hold the file in a way that blocks the read, which just means "no report yet".
#>
function Read-GuestMarker {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $null }

    try {
        $stream = [System.IO.File]::Open($Path, 'Open', 'Read', 'ReadWrite')
    } catch {
        return $null
    }

    try {
        $buffer = New-Object byte[] 512
        $read = $stream.Read($buffer, 0, $buffer.Length)
        if ($read -lt 16) { return $null }

        $magicSeen = [System.Text.Encoding]::ASCII.GetString($buffer, 0, 16).TrimEnd([char]0, [char]32)
        if ($magicSeen -ne $GuestMarkerMagic.TrimEnd()) { return $null }

        return [System.Text.Encoding]::ASCII.GetString($buffer, 16, $read - 16).TrimEnd([char]0, [char]32)
    } finally {
        $stream.Dispose()
    }
}

<#
    Boots a freshly built ISO in QEMU on a new virtio disk.

    The point is to watch an unattended install end to end: partitioning, whether Setup
    can see the VirtIO disk and NIC, and whether the first-logon guest-tools install
    succeeds. Everything here is best effort - the ISO is built and verified already -
    so a missing QEMU or a failed launch only produces a warning. The VM disk is left
    in place so the installed system can be inspected afterwards.

    Results come back through a second, tiny raw disk (see New-GuestMarkerDisk), which
    is why no qcow2 conversion or NTFS mount is needed: Windows has no NBD device, so
    qemu-nbd is not an option here.
#>
function Invoke-QemuBootTest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$IsoPath,
        [Parameter(Mandatory)][string]$Label
    )

    if (-not (Test-Path -LiteralPath $IsoPath -PathType Leaf)) {
        Write-Warning "Not booting $Label in QEMU: '$IsoPath' does not exist."
        return
    }

    # Resolved rather than assumed: qemu-img.exe in particular often arrives on PATH
    # from a package manager without being installed under Program Files.
    $systemExe = Resolve-Tool -Name 'qemu-system-x86_64.exe' -Candidates @($QemuSystemPath)
    $imgExe    = Resolve-Tool -Name 'qemu-img.exe' -Candidates @($QemuImgPath)

    if (-not $systemExe -or -not $imgExe) {
        $systemShown = if ($systemExe) { $systemExe } else { 'not found' }
        $imgShown    = if ($imgExe) { $imgExe } else { 'not found' }
        Write-Warning ("QEMU is not fully available (qemu-system-x86_64.exe: {0}; qemu-img.exe: {1}); skipping the boot test for {2}." -f $systemShown, $imgShown, $Label)
        return
    }

    if (-not (Test-Path -LiteralPath $QemuVmDir)) {
        New-Item -ItemType Directory -Path $QemuVmDir -Force | Out-Null
    }

    # Always start from a blank disk: the test is "does this media install cleanly".
    $diskPath = Join-Path $QemuVmDir ("{0}-test.qcow2" -f $Label)
    Remove-Item -LiteralPath $diskPath -Force -ErrorAction SilentlyContinue

    Write-Step ("[qemu] Creating a {0} GB virtio disk and booting {1}..." -f $QemuDiskSizeGb, (Split-Path $IsoPath -Leaf))
    Invoke-Native -FilePath $imgExe `
        -Arguments @('create', '-f', 'qcow2', $diskPath, ("{0}G" -f $QemuDiskSizeGb)) `
        -Activity 'qemu-img create' | Out-Null

    $markerPath = New-GuestMarkerDisk -Path (Join-Path $QemuVmDir ("{0}-marker.img" -f $Label))

    # Precomputed so no -f operator sits inside the array literal below.
    $vmName    = '{0}-test' -f $Label
    $diskArg   = 'file={0},if=virtio,format=qcow2' -f $diskPath
    $markerArg = 'file={0},if=virtio,format=raw' -f $markerPath

    # The install disk is listed first so it enumerates as Disk 0 in the guest, which
    # is what the answer file targets. The marker disk must come second: it is only
    # 8 MB, so if the order ever flipped the EFI partition would not fit and Setup
    # would fail loudly instead of silently installing onto the wrong disk.
    $qemuArgs = @(
        '-name', $vmName
        '-machine', 'q35'
        '-m', ([string]$QemuMemoryMb)
        '-smp', ([string]$QemuCpuCount)
        '-drive', $diskArg
        '-drive', $markerArg
        '-nic', 'user,model=virtio-net-pci'
        '-cdrom', $IsoPath
        '-boot', 'order=d'
    )
    if (-not [string]::IsNullOrWhiteSpace($QemuAcceleration)) {
        $qemuArgs = @('-accel', $QemuAcceleration) + $qemuArgs
    }

    if (-not $QemuWaitForExit) {
        Start-Process -FilePath $systemExe -ArgumentList $qemuArgs | Out-Null
        Write-Host '      QEMU launched detached; the build continues immediately.' -ForegroundColor DarkGray
        Write-Warning ("-TestInQemu cannot read the guest result in detached mode (`$QemuWaitForExit = `$false): the marker disk is only authoritative once QEMU has exited.")
        Write-Detail ("QEMU disk kept at {0}" -f $diskPath)
        return
    }

    Write-Host '      Close the QEMU window when the install has finished; the build then continues.' -ForegroundColor DarkGray
    try {
        $qemu = Start-Process -FilePath $systemExe -ArgumentList $qemuArgs -PassThru
    } catch {
        Write-Warning ("Could not start QEMU: {0}" -f $_.Exception.Message)
        Write-Warning 'If this host has no WHPX support, set $QemuAcceleration to an empty string in the configuration and retry.'
        return
    }

    # Poll while the VM runs so the result is announced as soon as it lands. The
    # marker only becomes authoritative once QEMU has exited and flushed it.
    $deadline   = if ($QemuBootTimeoutMinutes -gt 0) { (Get-Date).AddMinutes($QemuBootTimeoutMinutes) } else { $null }
    $lastReport = $null
    $timedOut   = $false

    while (-not $qemu.HasExited) {
        Start-Sleep -Seconds 5

        # Announce each new report as it lands: the payload writes GUEST-STARTED first
        # and then the verdict, so the progress is visible while the VM is still up.
        $peek = Read-GuestMarker -Path $markerPath
        if (-not [string]::IsNullOrWhiteSpace($peek) -and $peek -ne $lastReport) {
            Write-Host ("      guest: {0}" -f $peek) -ForegroundColor DarkGray
            $lastReport = $peek
        }

        if ($null -ne $deadline -and (Get-Date) -gt $deadline) {
            $timedOut = $true
            break
        }
    }

    if ($timedOut) {
        Write-Warning ("QEMU is still running after {0} minute(s); leaving it up and moving on without a guest result." -f $QemuBootTimeoutMinutes)
        Write-Detail ("QEMU disk kept at {0}" -f $diskPath)
        return
    }

    $qemu.WaitForExit()
    Write-Detail ("QEMU exited with code {0} after {1} minute(s)" -f $qemu.ExitCode, [int]((Get-Date) - $qemu.StartTime).TotalMinutes)

    $report = Read-GuestMarker -Path $markerPath
    if ([string]::IsNullOrWhiteSpace($report)) {
        Write-Warning ("{0}: the guest never reported back, so Setup did not reach the first-logon script (or the VM was closed before it finished)." -f $Label)
    } elseif ($report -like 'GUEST-OK*') {
        Write-Host ("      OK   unattended install completed: {0}" -f $report) -ForegroundColor Green
    } elseif ($report -like 'GUEST-STARTED*') {
        Write-Warning ("{0}: the first-logon script started but never finished - check \VirtIO\guest-tools-install.log inside the guest." -f $Label)
    } else {
        Write-Warning ("{0}: the guest reported a problem: {1}" -f $Label, $report)
    }

    Write-Detail ("QEMU disk kept at {0}" -f $diskPath)
}

# ====================================================================================
# Tool updates (best effort - hard requirements are checked afterwards)
# ====================================================================================
function Update-Wimlib {
    param([string]$InstallDir)

    Write-Host 'Checking for wimlib-imagex updates...' -ForegroundColor Cyan
    $baseUrl = 'https://wimlib.net/'
    $exePath = Join-Path $InstallDir 'wimlib-imagex.exe'

    try {
        # -UseBasicParsing: without it, Windows PowerShell 5.1 tries to parse HTML
        # with the IE engine, which fails on Server Core.
        $page = (Invoke-WebRequest -Uri $baseUrl -UseBasicParsing -UserAgent $StandardUserAgent -ErrorAction Stop).Content

        $match = [regex]::Match($page, 'href="(downloads/wimlib-(?<ver>\d+(?:\.\d+)*)-windows-x86_64-bin\.zip)"')
        if (-not $match.Success) {
            Write-Warning 'Could not determine the latest wimlib release from wimlib.net; keeping the installed copy.'
            return
        }

        $downloadUrl   = $baseUrl + $match.Groups[1].Value
        $latestVersion = [version]$match.Groups['ver'].Value

        $localVersion = [version]'0.0.0'
        if (Test-Path -LiteralPath $exePath) {
            try {
                $versionOutput = @(& $exePath --version 2>&1) | Select-Object -First 1
                $versionMatch = [regex]::Match([string]$versionOutput, 'wimlib-imagex ([\d\.]+)')
                if ($versionMatch.Success) { $localVersion = [version]$versionMatch.Groups[1].Value }
            } catch {
                Write-Verbose "Could not read the local wimlib version: $($_.Exception.Message)"
            }
        }

        if ($localVersion -ge $latestVersion) {
            Write-Host "wimlib-imagex $localVersion is current." -ForegroundColor DarkGray
            return
        }

        Write-Host "Updating wimlib-imagex from $localVersion to $latestVersion..." -ForegroundColor Yellow
        $tempZip = Join-Path $env:TEMP 'wimlib.zip'
        Invoke-WebRequest -Uri $downloadUrl -OutFile $tempZip -UserAgent $StandardUserAgent -ErrorAction Stop

        New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
        Expand-Archive -LiteralPath $tempZip -DestinationPath $InstallDir -Force
        Remove-Item -LiteralPath $tempZip -Force

        # Some wimlib packages nest the binaries in a version-stamped folder.
        # Flatten it so wimlib-imagex.exe sits next to libwim-*.dll.
        $found = Get-ChildItem -LiteralPath $InstallDir -Recurse -File -Filter 'wimlib-imagex.exe' |
            Select-Object -First 1
        if ($found -and $found.DirectoryName -ne $InstallDir) {
            Copy-Item -Path (Join-Path $found.DirectoryName '*') -Destination $InstallDir -Recurse -Force
        }
    } catch {
        Write-Warning "wimlib update check failed ($($_.Exception.Message)); keeping the installed copy."
    }
}

function Update-VirtIO {
    param([string]$BaseDir)

    Write-Host 'Checking for VirtIO driver updates...' -ForegroundColor Cyan

    $isoUrl       = 'https://fedorapeople.org/groups/virt/virtio-win/direct-downloads/latest-virtio/virtio-win.iso'
    $localIsoPath = Join-Path $BaseDir 'virtio-win.iso'
    $msiDir       = Join-Path $BaseDir 'MSI'

    $downloadIso = $false
    try {
        # This mirror sits behind the Anubis anti-bot proxy, so scripted clients usually get
        # an HTML challenge page instead of file metadata. Treat freshness as best-effort:
        # a missing header just means "keep the local ISO as-is".
        $head = Invoke-WebRequest -Uri $isoUrl -Method Head -UseBasicParsing -ErrorAction Stop

        $remoteModified = $null
        $remoteLength   = $null

        $lastModifiedHeader = Get-ResponseHeader -Response $head -Name 'Last-Modified'
        if ($lastModifiedHeader) {
            try { $remoteModified = [datetime]$lastModifiedHeader } catch { $remoteModified = $null }
        }

        $contentLengthHeader = Get-ResponseHeader -Response $head -Name 'Content-Length'
        if ($contentLengthHeader) {
            try { $remoteLength = [int64]$contentLengthHeader } catch { $remoteLength = $null }
        }

        if (-not $remoteModified -and -not $remoteLength) {
            Write-Warning ('Could not read freshness metadata for the VirtIO ISO (the mirror serves an ' +
                           'anti-bot challenge to scripted clients). Using the local copy as-is.')
        }

        if (-not (Test-Path -LiteralPath $localIsoPath)) {
            $downloadIso = $true
        } elseif ($remoteModified -and (Get-Item -LiteralPath $localIsoPath).LastWriteTimeUtc -lt $remoteModified.ToUniversalTime()) {
            $downloadIso = $true
        } elseif ($remoteLength -and (Get-Item -LiteralPath $localIsoPath).Length -ne $remoteLength) {
            $downloadIso = $true
        }
    } catch {
        Write-Warning "VirtIO update check failed ($($_.Exception.Message))."
        if (-not (Test-Path -LiteralPath $localIsoPath)) { throw "No local VirtIO ISO at $localIsoPath and the update check failed." }
    }

    New-Item -ItemType Directory -Path $BaseDir -Force | Out-Null

    if ($downloadIso) {
        Write-Host 'Downloading the latest VirtIO ISO (this takes a while)...' -ForegroundColor Yellow

        # Download to the side and validate before replacing a known-good ISO: the mirror is
        # fronted by the Anubis anti-bot proxy, which answers scripted clients with HTTP 200
        # and a small HTML challenge page.
        $download = "$localIsoPath.download"
        Remove-Item -LiteralPath $download -Force -ErrorAction SilentlyContinue
        Invoke-WebRequest -Uri $isoUrl -OutFile $download -ErrorAction Stop

        if (-not (Test-IsoImage -Path $download)) {
            $bytes = if (Test-Path -LiteralPath $download) { (Get-Item -LiteralPath $download).Length } else { 0 }
            Remove-Item -LiteralPath $download -Force -ErrorAction SilentlyContinue
            throw ("The VirtIO download did not return an ISO image ($bytes bytes were saved). " +
                   'fedorapeople.org answers automated clients with an anti-bot challenge page. ' +
                   "Download virtio-win.iso in a browser and save it as '$localIsoPath', then re-run.")
        }

        Move-Item -LiteralPath $download -Destination $localIsoPath -Force
    }

    if (-not (Test-IsoImage -Path $localIsoPath)) {
        throw ("'$localIsoPath' is not a valid ISO image. Delete it, download virtio-win.iso " +
               'in a browser, and re-run.')
    }

    # Extraction is gated on a stamp file that is written only after a completely
    # successful extraction, so a half-populated C:\VirtIO is always repaired on the
    # next run instead of being mistaken for "already done".
    $stampFile = Join-Path $BaseDir 'extracted.stamp'
    $isoItem   = Get-Item -LiteralPath $localIsoPath
    $isoStamp  = '{0}|{1}' -f $isoItem.Length, $isoItem.LastWriteTimeUtc.Ticks

    $needsExtract = $downloadIso -or
                    -not (Test-Path -LiteralPath $stampFile) -or
                    -not (Test-Path -LiteralPath $VirtIoMsiPath) -or
                    -not (Test-Path -LiteralPath $QemuGaMsiPath) -or
                    -not (Test-Path -LiteralPath $VirtIoDrivers) -or
                    ((Get-Content -LiteralPath $stampFile -Raw).Trim() -ne $isoStamp)

    if (-not $needsExtract) {
        Write-Host 'VirtIO MSIs and drivers are already extracted and current.' -ForegroundColor DarkGray
        return
    }

    Write-Host 'Extracting VirtIO MSIs and drivers from the ISO...' -ForegroundColor Yellow
    $mount = Mount-IsoImage -Path $localIsoPath
    try {
        $sourceMsi1 = Join-Path $mount.Root 'virtio-win-gt-x64.msi'
        $sourceMsi2 = Join-Path $mount.Root 'guest-agent\qemu-ga-x86_64.msi'

        foreach ($required in @($sourceMsi1, $sourceMsi2)) {
            if (-not (Test-Path -LiteralPath $required)) {
                throw "The VirtIO ISO does not contain the expected file: $required"
            }
        }

        New-Item -ItemType Directory -Path $msiDir -Force | Out-Null
        Copy-Item -LiteralPath $sourceMsi1 -Destination $msiDir -Force
        Copy-Item -LiteralPath $sourceMsi2 -Destination $msiDir -Force

        Remove-DirectoryTree -Path $VirtIoDrivers
        New-Item -ItemType Directory -Path $VirtIoDrivers -Force | Out-Null

        # Copy only the driver folders; the root also holds MSIs/EXEs that DISM
        # must not see. -Exclude does not filter reliably with -Recurse, so filter
        # the items explicitly.
        $driverFolders = @(Get-ChildItem -LiteralPath $mount.Root -Directory -Force |
            Where-Object { $_.Name -ine 'guest-agent' })
        if ($driverFolders.Count -eq 0) { throw "No driver folders found on the VirtIO ISO." }

        foreach ($folder in $driverFolders) {
            # The destination container MUST exist first. Copy-Item onto a non-existent
            # destination with a wildcard source throws
            # "Container cannot be copied onto existing leaf item" instead of creating it.
            $destination = Join-Path $VirtIoDrivers $folder.Name
            New-Item -ItemType Directory -Path $destination -Force | Out-Null
            Copy-Item -Path (Join-Path $folder.FullName '*') -Destination $destination -Recurse -Force
        }

        # Only reached when every copy above succeeded.
        [System.IO.File]::WriteAllText($stampFile, $isoStamp)
    } finally {
        Dismount-IsoImage -Path $localIsoPath
    }
}

# ====================================================================================
# Pre-flight
# ====================================================================================
# These paths are handed to native tools as raw arguments, and oscdimg's -bootdata
# value cannot be quoted reliably from PowerShell 5.1, so keep them space-free.
foreach ($path in @($WorkDir, $MountDir, $VirtIoDrivers, $VirtIoDriversAmd64, $VirtIoBootDrivers, $OutputIsoFolder, $QemuVmDir)) {
    if ($path -match '\s') {
        throw "Path contains a space, which this script does not support: '$path'"
    }
}

# Refresh external tools first (best effort), so a missing wimlib can install itself.
Update-Wimlib -InstallDir $WimlibDir
Update-VirtIO -BaseDir $VirtIoBaseDir

# Resolve against the configured path, the versioned ADK folder, then PATH.
$WimlibPath  = Resolve-Tool -Name 'wimlib-imagex.exe' -Candidates @($WimlibPath) -Required
$OscdimgPath = Resolve-Tool -Name 'oscdimg.exe' -Candidates @(
    $OscdimgPath
    'C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\oscdimg.exe'
) -Required
$DismPath    = Resolve-Tool -Name 'dism.exe' -Candidates @(
    $DismPath
    'C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\DISM\dism.exe'
) -Required

# QEMU is only needed by -TestInQemu, so it is resolved separately from the hard
# requirements above: report up front, but let the ISOs build either way. The boot test
# is also disabled here when this host cannot run it usefully, so the reason is stated
# once before hours of work rather than once per ISO.
#
# Host capability is checked before QEMU itself: if the machine cannot virtualise,
# there is no point sending the operator off to install QEMU for it.
$QemuBootTestEnabled = [bool]$TestInQemu
if ($QemuBootTestEnabled) {
    $hypervisor = Get-HypervisorStatus
    if ($hypervisor -ne 'Available') {
        if ($QemuAllowSoftwareEmulation) {
            Write-Warning 'This host cannot run an accelerated VM, so QEMU will run unaccelerated and the boot test will be very slow. -accel is being dropped for this run.'
            $QemuAcceleration = ''
        } else {
            $reason = switch ($hypervisor) {
                'VirtualisationUnavailable' {
                    if (Test-RunningInVirtualMachine) {
                        'this machine is itself a VM and hardware virtualisation is not exposed to it (no nested virtualisation), which is the usual reason a build host cannot also boot the images it produces'
                    } else {
                        'hardware virtualisation is not available to Windows - check that VT-x/AMD-V is enabled in firmware'
                    }
                }
                'FeatureDisabled'           { 'the "Windows Hypervisor Platform" optional feature is not enabled - run Enable-WindowsOptionalFeature -Online -FeatureName HypervisorPlatform and reboot' }
                default                     { 'this host could not be probed for virtualisation support' }
            }
            Write-Warning ("-TestInQemu is skipped because {0}. Software emulation is far too slow to install Windows, so build the ISOs here and boot them on a host that can virtualise - or set `$QemuAllowSoftwareEmulation = `$true to run the test unaccelerated anyway." -f $reason)
            $QemuBootTestEnabled = $false
        }
    }
}

if ($QemuBootTestEnabled) {
    $resolvedSystem = Resolve-Tool -Name 'qemu-system-x86_64.exe' -Candidates @($QemuSystemPath)
    $resolvedImg    = Resolve-Tool -Name 'qemu-img.exe' -Candidates @($QemuImgPath)
    if ($resolvedSystem) { $QemuSystemPath = $resolvedSystem }
    if ($resolvedImg) { $QemuImgPath = $resolvedImg }

    if (-not (Test-Path -LiteralPath $QemuSystemPath -PathType Leaf) -or
        -not (Test-Path -LiteralPath $QemuImgPath -PathType Leaf)) {
        Write-Warning '-TestInQemu was requested but QEMU is not installed (qemu-system-x86_64.exe / qemu-img.exe not found), so the boot test is skipped. The ISOs will still be built.'
        $QemuBootTestEnabled = $false
    }
}

# Hard requirements - anything missing here must fail before hours of work.
foreach ($requiredFile in @($WimlibPath, $OscdimgPath, $DismPath, $TemplateXmlPath, $VirtIoMsiPath, $QemuGaMsiPath)) {
    if (-not (Test-Path -LiteralPath $requiredFile -PathType Leaf)) {
        throw "Required file not found: $requiredFile"
    }
}
if (-not (Test-Path -LiteralPath $SourceIsoFolder)) {
    throw "Source ISO folder not found: $SourceIsoFolder"
}

# ====================================================================================
# Working directories
# ====================================================================================
foreach ($dir in @($OutputIsoFolder, $WorkDir, $MountDir)) {
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
}

# A dirty volume fights the staging cleanup for the whole run, so say so before hours
# of work rather than after the first delete fails.
if (Test-VolumeDirty -Path $MountDir) {
    Write-Warning ("The volume holding '{0}' has its NTFS dirty bit set. Windows may refuse to delete the staging directories, so cleanup can leave folders behind. Run 'chkdsk /f {1}' to repair it (the system volume needs a reboot)." -f $MountDir, (Split-Path -Qualifier $MountDir))
}

# Retry anything a previous run could not delete.
Clear-StagingRoot -Path $MountDir

$IsoFiles = @(Get-ChildItem -LiteralPath $SourceIsoFolder -Filter '*.iso' -File | Sort-Object Name)
if ($IsoFiles.Count -eq 0) {
    throw "No .iso files found in $SourceIsoFolder"
}

# Driver injection targets, built once and reused for every ISO and every image index.
# boot.wim gets only the boot-critical drivers (see $BootWimDriverNames); the install
# images get the full amd64 set.
$StagedDrivers     = New-Amd64DriverStage -DriverRoot $VirtIoDrivers -StageRoot $VirtIoDriversAmd64
$StagedBootDrivers = New-Amd64DriverStage -DriverRoot $VirtIoDrivers -StageRoot $VirtIoBootDrivers -DriverNames $BootWimDriverNames

# ====================================================================================
# Execution
# ====================================================================================
$failedIsos = @()

foreach ($Iso in $IsoFiles) {
    Write-Header "Processing: $($Iso.Name)"

    $BaseName       = $Iso.BaseName
    $StandardIsoOut = Join-Path $OutputIsoFolder "$BaseName-VirtIO-Standard.iso"
    $UnattendIsoOut = Join-Path $OutputIsoFolder "$BaseName-VirtIO-Unattended.iso"
    $UnattendTarget = Join-Path $WorkDir 'autounattend.xml'
    $OemRoot        = Join-Path $WorkDir 'sources\$OEM$'
    $OemVirtIoDir   = Join-Path $OemRoot '$1\VirtIO'
    $isoBuilt       = $false

    try {
        # ----------------------------------------------------------------------------
        # 1. Copy the base ISO to the working directory
        # ----------------------------------------------------------------------------
        Write-Step '[1/6] Copying base ISO contents...'

        Reset-WorkDirectory -Path $WorkDir
        Clear-StagingRoot -Path $MountDir

        if (-not (Test-IsoImage -Path $Iso.FullName)) {
            throw "Not a valid ISO image: $($Iso.FullName)"
        }

        $mount = Mount-IsoImage -Path $Iso.FullName
        try {
            # robocopy is dramatically faster than Copy-Item for a multi-GB tree.
            # Exit codes 0-7 are success variants; 8+ indicates a real failure.
            Invoke-Native -FilePath "$env:SystemRoot\System32\robocopy.exe" `
                -Arguments @($mount.Root, $WorkDir, '/E', '/MT:16', '/R:1', '/W:1',
                             '/NFL', '/NDL', '/NJH', '/NJS', '/NP') `
                -AllowedExitCodes (0..7) -Activity 'robocopy' | Out-Null
        } finally {
            Dismount-IsoImage -Path $Iso.FullName
        }

        Clear-ReadOnlyAttribute -Path $WorkDir

        $SourceInstallFile = Join-Path $WorkDir 'sources\install.wim'
        if (-not (Test-Path -LiteralPath $SourceInstallFile)) {
            $SourceInstallFile = Join-Path $WorkDir 'sources\install.esd'
        }
        $BootWim     = Join-Path $WorkDir 'sources\boot.wim'
        $EtfsBoot    = Join-Path $WorkDir 'boot\etfsboot.com'
        $EfiSys      = Join-Path $WorkDir 'efi\microsoft\boot\efisys.bin'

        foreach ($requiredFile in @($SourceInstallFile, $BootWim, $EtfsBoot, $EfiSys)) {
            if (-not (Test-Path -LiteralPath $requiredFile -PathType Leaf)) {
                throw "The ISO does not look like bootable Windows media - missing: $requiredFile"
            }
        }

        $InstallImages  = @(Get-WimImages -WimPath $SourceInstallFile)
        $BootImages     = @(Get-WimImages -WimPath $BootWim)
        $InstallEdition = Select-InstallEdition -Images $InstallImages -PreferredPatterns $PreferredEditionPatterns

        Write-Detail ("Install image: {0} ({1} edition(s))" -f (Split-Path $SourceInstallFile -Leaf), $InstallImages.Count)
        Write-Detail ("Unattended ISO will autoinstall: {0}" -f $InstallEdition.Name)

        # ----------------------------------------------------------------------------
        # 2. Slipstream boot.wim via wimlib apply (bypasses DISM mount errors)
        # ----------------------------------------------------------------------------
        Write-Step '[2/6] Slipstreaming VirtIO into boot.wim (wimlib native)...'
        $NewBootWim    = Join-Path $WorkDir 'sources\new_boot.wim'
        $NewSetupIndex = $null

        for ($i = 0; $i -lt $BootImages.Count; $i++) {
            $bootImage = $BootImages[$i]
            $newIndex  = $i + 1
            Write-Detail ("Applying boot index {0} ({1})..." -f $bootImage.Index, $bootImage.Name)

            $stage = New-StagingDirectory -Root $MountDir -Purpose 'boot'
            try {
                Invoke-Native -FilePath $WimlibPath `
                    -Arguments @('apply', $BootWim, "$($bootImage.Index)", $stage) -Activity 'wimlib apply' | Out-Null

                Invoke-Native -FilePath $DismPath `
                    -Arguments @("/Image:$stage", '/Add-Driver', "/Driver:$StagedBootDrivers", '/Recurse') `
                    -Activity 'DISM /Add-Driver (boot)' | Out-Null

                Write-Detail 'Capturing the serviced index back into the WIM...'
                if ($newIndex -eq 1) {
                    Invoke-Native -FilePath $WimlibPath `
                        -Arguments @('capture', $stage, $NewBootWim, $bootImage.Name) -Activity 'wimlib capture' | Out-Null
                } else {
                    Invoke-Native -FilePath $WimlibPath `
                        -Arguments @('append', $stage, $NewBootWim, $bootImage.Name) -Activity 'wimlib append' | Out-Null
                }
            } finally {
                Remove-DirectoryTree -Path $stage -BestEffort
            }

            if ($bootImage.Name -match '(?i)setup') { $NewSetupIndex = $newIndex }
        }

        # The Windows Setup image must be the WIM's bootable image, or the media
        # boots straight into WinPE instead of the installer.
        if (-not $NewSetupIndex) { $NewSetupIndex = if ($BootImages.Count -ge 2) { 2 } else { 1 } }
        Write-Detail ("Marking boot index {0} as the bootable image..." -f $NewSetupIndex)
        Invoke-Native -FilePath $WimlibPath `
            -Arguments @('info', $NewBootWim, "$NewSetupIndex", '--boot') -Activity 'wimlib info --boot' | Out-Null

        Remove-Item -LiteralPath $BootWim -Force
        Rename-Item -LiteralPath $NewBootWim -NewName 'boot.wim'

        # ----------------------------------------------------------------------------
        # 3. Slipstream the install image, compressing straight to a solid ESD
        # ----------------------------------------------------------------------------
        Write-Step '[3/6] Slipstreaming VirtIO into the install image (solid ESD)...'
        $NewInstallEsd = Join-Path $WorkDir 'sources\new_install.esd'

        # wimappend rejects a duplicate image name, so keep names unique.
        $usedNames = @{}
        $installNames = @()
        foreach ($image in $InstallImages) {
            $candidate = $image.Name
            if ($usedNames.ContainsKey($candidate.ToLowerInvariant())) {
                $candidate = "$($image.Name) ($($image.Index))"
            }
            $usedNames[$candidate.ToLowerInvariant()] = $true
            $installNames += $candidate
        }

        for ($i = 0; $i -lt $InstallImages.Count; $i++) {
            $image    = $InstallImages[$i]
            $newIndex = $i + 1
            Write-Detail ("Applying install index {0} ({1})..." -f $image.Index, $image.Name)

            $stage = New-StagingDirectory -Root $MountDir -Purpose 'install'
            try {
                Invoke-Native -FilePath $WimlibPath `
                    -Arguments @('apply', $SourceInstallFile, "$($image.Index)", $stage) -Activity 'wimlib apply' | Out-Null

                Invoke-Native -FilePath $DismPath `
                    -Arguments @("/Image:$stage", '/Add-Driver', "/Driver:$StagedDrivers", '/Recurse') `
                    -Activity 'DISM /Add-Driver' | Out-Null

                Write-Detail 'Repacking to a highly compressed solid ESD...'
                if ($newIndex -eq 1) {
                    Invoke-Native -FilePath $WimlibPath `
                        -Arguments @('capture', $stage, $NewInstallEsd, $installNames[$i], '--solid') `
                        -Activity 'wimlib capture --solid' | Out-Null
                } else {
                    Invoke-Native -FilePath $WimlibPath `
                        -Arguments @('append', $stage, $NewInstallEsd, $installNames[$i], '--solid') `
                        -Activity 'wimlib append --solid' | Out-Null
                }
            } finally {
                Remove-DirectoryTree -Path $stage -BestEffort
            }
        }

        Remove-Item -LiteralPath $SourceInstallFile -Force
        Rename-Item -LiteralPath $NewInstallEsd -NewName 'install.esd'

        # ----------------------------------------------------------------------------
        # 4. Build the standard ISO
        # ----------------------------------------------------------------------------
        Write-Step '[4/6] Creating the standard bootable ISO...'
        # Paths are validated as space-free above, so -bootdata needs no nesting quotes.
        $BootData = "2#p0,e,b$EtfsBoot#pEF,e,b$EfiSys"
        Remove-Item -LiteralPath $StandardIsoOut -Force -ErrorAction SilentlyContinue
        Invoke-Native -FilePath $OscdimgPath `
            -Arguments @("-bootdata:$BootData", '-u2', '-udfver102', '-m', "-l$StandardIsoLabel", $WorkDir, $StandardIsoOut) `
            -Activity 'oscdimg' | Out-Null

        # ----------------------------------------------------------------------------
        # 5. Stage the unattended-answer payload
        # ----------------------------------------------------------------------------
        Write-Step '[5/6] Staging $OEM$ payload and autounattend.xml...'

        New-Item -ItemType Directory -Path $OemVirtIoDir -Force | Out-Null
        Copy-Item -LiteralPath $VirtIoMsiPath -Destination $OemVirtIoDir -Force
        Copy-Item -LiteralPath $QemuGaMsiPath -Destination $OemVirtIoDir -Force

        # The guest-tools installer runs from FirstLogonCommands. It is shipped as a
        # readable .ps1 rather than a base64 -EncodedCommand blob so that a failed
        # image can be diagnosed from the media or from C:\VirtIO after install.
        $GuestToolsScriptName = 'install-guest-tools.ps1'
        $GuestToolsScript = @'
# Deployed to C:\VirtIO by the ISO build; executed from FirstLogonCommands.
# Installs the VirtIO guest tools and records the outcome in
# C:\VirtIO\guest-tools-install.log. Failures also leave a
# C:\VirtIO\guest-tools-install.FAILED marker so a bad image is easy to spot.
# When the media is booted by the build's -TestInQemu VMs, the result is also written
# to a small raw marker disk so the host can verify the run without mounting anything.
$ErrorActionPreference = 'Stop'

$root     = 'C:\VirtIO'
$logFile  = Join-Path $root 'guest-tools-install.log'
$flagFile = Join-Path $root 'guest-tools-install.FAILED'
$msiFiles = @(
    'virtio-win-gt-x64.msi'
    'qemu-ga-x86_64.msi'
)
# 0 = success, 3010 = success + reboot required, 1641 = success + reboot initiated
$successCodes = @(0, 3010, 1641)

function Write-Log {
    param([string]$Message)
    try {
        $stamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        "$stamp  $Message" | Out-File -LiteralPath $logFile -Append -Encoding utf8
    } catch { }
}

function Write-Flag {
    param([string]$Message)
    try { $Message | Out-File -LiteralPath $flagFile -Append -Encoding utf8 } catch { }
}

function Write-TestMarker {
    param([string]$Report)

    # The build attaches a tiny raw disk so the host can tell whether this unattended
    # run finished. It is identified by its magic header, never by disk number (the
    # answer file targets Disk 0, and this disk must never be confused with it).
    # Entirely best effort: nothing here may disturb the guest-tools install.
    try {
        $magic = '__MARKER_MAGIC__'

        foreach ($disk in @(Get-Disk -ErrorAction SilentlyContinue)) {
            if ($disk.IsBoot -or $disk.IsSystem) { continue }

            try {
                if ($disk.IsOffline) { Set-Disk -Number $disk.Number -IsOffline $false }
                if ($disk.IsReadOnly) { Set-Disk -Number $disk.Number -IsReadOnly $false }
            } catch { }

            $path = '\\.\PhysicalDrive' + $disk.Number
            try { $stream = [System.IO.File]::Open($path, 'Open', 'ReadWrite', 'ReadWrite') }
            catch { continue }

            try {
                $probe = New-Object byte[] 16
                [void]$stream.Read($probe, 0, 16)
                if ([System.Text.Encoding]::ASCII.GetString($probe).TrimEnd() -ne $magic) { continue }

                # Keep the magic in place: it is what lets a later report (GUEST-OK /
                # GUEST-FAIL) find this same disk again.
                $block = New-Object byte[] 512
                [Array]::Copy($probe, $block, 16)
                $text = [System.Text.Encoding]::ASCII.GetBytes($Report)
                [Array]::Copy($text, 0, $block, 16, [Math]::Min($text.Length, 512 - 16))
                $stream.Position = 0
                $stream.Write($block, 0, $block.Length)
                $stream.Flush()
            } finally {
                $stream.Dispose()
            }

            return
        }
    } catch { }
}

Write-Log 'Guest tools installation started'
try { Write-TestMarker ('GUEST-STARTED|{0}' -f (Get-Date).ToString('s')) } catch { }
$failed = $false

foreach ($msiName in $msiFiles) {
    $msiPath = Join-Path $root $msiName

    if (-not (Test-Path -LiteralPath $msiPath)) {
        Write-Log  ("MISSING: {0}" -f $msiPath)
        Write-Flag ("MISSING: {0}" -f $msiPath)
        $failed = $true
        continue
    }

    try {
        $startParams = @{
            FilePath     = 'msiexec.exe'
            ArgumentList = @('/i', $msiPath, '/qn', '/norestart')
            Wait         = $true
            PassThru     = $true
        }
        $proc = Start-Process @startParams
        Write-Log ("msiexec /i {0} -> exit {1}" -f $msiName, $proc.ExitCode)

        if ($successCodes -notcontains $proc.ExitCode) {
            Write-Flag ("FAILED: {0} -> exit {1}" -f $msiName, $proc.ExitCode)
            $failed = $true
        }
    } catch {
        Write-Log  ("msiexec /i {0} threw: {1}" -f $msiName, $_.Exception.Message)
        Write-Flag ("EXCEPTION: {0} -> {1}" -f $msiName, $_.Exception.Message)
        $failed = $true
    }
}

if ($failed) {
    Write-Log 'Guest tools installation COMPLETED WITH ERRORS'
} else {
    Write-Log 'Guest tools installation COMPLETED OK'
}

$verdict = if ($failed) { 'GUEST-FAIL' } else { 'GUEST-OK' }
try { Write-TestMarker ('{0}|{1}|virtio-win + qemu-ga' -f $verdict, (Get-Date).ToString('s')) } catch { }
'@

        # The marker handshake must match the magic the host stamps into the marker
        # disk, so it is injected here rather than hard-coded in the payload.
        $GuestToolsScript = $GuestToolsScript.Replace('__MARKER_MAGIC__', $GuestMarkerMagic)

        # Write with an explicit BOM so the bytes do not depend on whether this script
        # runs under Windows PowerShell 5.1 or PowerShell 7.
        [System.IO.File]::WriteAllText(
            (Join-Path $OemVirtIoDir $GuestToolsScriptName),
            $GuestToolsScript,
            [System.Text.UTF8Encoding]::new($true))

        $GuestToolsCommand = 'powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File C:\VirtIO\install-guest-tools.ps1'

        # The image name in the rebuilt ESD can be uniquified ("Edition (2)") when the
        # source media carries duplicates, so substitute the captured name rather than
        # the source name to keep /IMAGE/NAME accurate.
        $selectedPos       = [array]::IndexOf($InstallImages, $InstallEdition)
        $SelectedImageName = $installNames[$selectedPos]

        $DynamicProductKey = Get-KmsClientSetupKey -ImageName $InstallEdition.Name
        $XmlContent = [System.IO.File]::ReadAllText($TemplateXmlPath)

        # .Replace() is literal, so passwords containing $ or \ are safe.
        $XmlContent = $XmlContent.Replace('{{IMAGE_NAME}}', $SelectedImageName)
        $XmlContent = $XmlContent.Replace('{{ENCODED_COMMAND}}', $GuestToolsCommand)
        $XmlContent = $XmlContent.Replace('{{ADMIN_PASSWORD}}', $AdminPassword)
        $XmlContent = $XmlContent.Replace('{{USERNAME}}', $LocalUsername)
        $XmlContent = $XmlContent.Replace('{{USER_PASSWORD}}', $LocalUserPassword)

        if ($DynamicProductKey) {
            $XmlContent = $XmlContent.Replace('{{PRODUCT_KEY}}', $DynamicProductKey)
        } else {
            Write-Warning "No KMS client setup key matched '$($InstallEdition.Name)'; removing <ProductKey> from the answer file."
            $XmlContent = $XmlContent -replace '(?s)<ProductKey>.*?</ProductKey>', ''
            $XmlContent = $XmlContent.Replace('{{PRODUCT_KEY}}', '')
        }

        $leftovers = @([regex]::Matches($XmlContent, '\{\{[A-Z_]+\}\}') | ForEach-Object { $_.Value }) |
            Select-Object -Unique
        if ($leftovers.Count -gt 0) {
            Write-Warning "The answer file still contains unresolved placeholders: $($leftovers -join ', ')"
        }

        [System.IO.File]::WriteAllText($UnattendTarget, $XmlContent, [System.Text.UTF8Encoding]::new($true))

        # ----------------------------------------------------------------------------
        # 6. Build the unattended ISO
        # ----------------------------------------------------------------------------
        Write-Step '[6/6] Creating the unattended bootable ISO...'
        Remove-Item -LiteralPath $UnattendIsoOut -Force -ErrorAction SilentlyContinue
        Invoke-Native -FilePath $OscdimgPath `
            -Arguments @("-bootdata:$BootData", '-u2', '-udfver102', '-m', "-l$UnattendIsoLabel", $WorkDir, $UnattendIsoOut) `
            -Activity 'oscdimg' | Out-Null

        # --- Per-ISO verification -------------------------------------------------
        foreach ($outputIso in @($StandardIsoOut, $UnattendIsoOut)) {
            if (-not (Test-Path -LiteralPath $outputIso -PathType Leaf)) {
                throw "Expected output ISO was not created: $outputIso"
            }
            $sizeMb = [math]::Round((Get-Item -LiteralPath $outputIso).Length / 1MB, 1)
            if ($sizeMb -lt 100) {
                throw "Output ISO looks far too small to be valid: $outputIso ($sizeMb MB)"
            }
            Write-Host ("      OK  {0}  ({1} MB)" -f (Split-Path $outputIso -Leaf), $sizeMb) -ForegroundColor DarkGray
        }

        Write-Host ("Completed: {0}" -f $Iso.Name) -ForegroundColor Green
        $isoBuilt = $true
    } catch {
        Write-Host ("FAILED: {0} - {1}" -f $Iso.Name, $_.Exception.Message) -ForegroundColor Red
        $failedIsos += $Iso.Name
    } finally {
        # Always leave the shared work directories in a reusable state. Cleanup
        # problems must not mask the real failure, so they only warn.
        try {
            Remove-Item -LiteralPath $UnattendTarget -Force -ErrorAction SilentlyContinue
            if (Test-Path -LiteralPath $OemRoot) { Remove-DirectoryTree -Path $OemRoot -BestEffort }
        } catch {
            Write-Warning "Per-ISO cleanup failed: $($_.Exception.Message)"
        }
    }

    # Boot the unattended ISO so the install can be watched. Best effort on purpose:
    # the ISOs are the deliverable and are already on disk, so nothing here may fail
    # the run.
    if ($QemuBootTestEnabled -and $isoBuilt) {
        try {
            Invoke-QemuBootTest -IsoPath $UnattendIsoOut -Label $BaseName
        } catch {
            Write-Warning "QEMU boot test for '$($Iso.Name)' failed: $($_.Exception.Message)"
        }
    }
}

# ====================================================================================
# Cleanup and summary
# ====================================================================================
Write-Host 'Cleaning up working directories...' -ForegroundColor Cyan
Remove-DirectoryTree -Path $WorkDir -BestEffort
New-Item -ItemType Directory -Path $WorkDir -Force | Out-Null
Remove-DirectoryTree -Path $MountDir -BestEffort
New-Item -ItemType Directory -Path $MountDir -Force | Out-Null

# Report anything that survived cleanup rather than silently leaving it behind.
$staleStaging = @(Get-ChildItem -LiteralPath $MountDir -Force -ErrorAction SilentlyContinue)
if ($staleStaging.Count -gt 0) {
    Write-Warning ("{0} staging folder(s) could not be removed from {1}: {2}" -f `
        $staleStaging.Count, $MountDir, (($staleStaging | Select-Object -ExpandProperty Name) -join ', '))
    Write-Warning 'They are retried at the start of the next run. If this keeps happening, run "chkdsk /f" on the volume.'
}

if ($failedIsos.Count -gt 0) {
    Write-Host ''
    Write-Host ("{0} of {1} ISO(s) failed: {2}" -f $failedIsos.Count, $IsoFiles.Count, ($failedIsos -join ', ')) -ForegroundColor Red
    exit 1
}

Write-Host ''
Write-Host "Process complete. ISOs are in $OutputIsoFolder" -ForegroundColor Green
