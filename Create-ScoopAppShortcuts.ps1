<#
.SYNOPSIS
Creates Windows desktop shortcuts for specified Scoop-installed apps with robust exception handling.

.DESCRIPTION
This version uses 'throw' within validation functions, and the top-level 'try/catch'
surrounds all logic, including the initial 'apps' directory check.

NOTE ON ICON POSITIONING:
The script uses WScript.Shell to create the shortcut file (.lnk). This method only creates 
the file and cannot control the icon's position on the desktop screen. Icon arrangement 
is managed solely by the Windows Explorer shell (based on auto-arrange settings and 
internal registry keys), and is not reliably scriptable via standard PowerShell.

# EXECUTION FROM GITHUB (Single Command Line)
# If this script were hosted on GitHub (e.g., in a raw Gist or repo file), 
# you could run it directly using Invoke-RestMethod (irm) and Invoke-Expression (iex):
# iex (irm "https://raw.githubusercontent.com/username/repo/branch/scoop_desktop_shortcuts.ps1")
#>

[CmdletBinding()]
param()

$ScoopPath = "$env:USERPROFILE\scoop"
$DesktopPath = [Environment]::GetFolderPath('Desktop')
$InstalledAppsDir = Join-Path -Path $ScoopPath -ChildPath 'apps'

$ScoopApps = @(
    'vscode',
    'gpu-z',
    'cpu-z' ,
    'hwinfo',
    'vlc',
    '7zip'
    )

function CheckScoopAppsDirectoryExists {
    param(
        [Parameter(Mandatory=$true)][string]$AppsDir
    )
    if (-not (Test-Path $AppsDir -PathType Container)) {
        throw "PrerequisiteError: Scoop 'apps' directory not found at '$AppsDir'. Please check the \$ScoopPath variable and ensure Scoop is installed."
    }
}

function Get-CpuLookupTable {
    # This is a representative, non-exhaustive lookup table for modern CPUs.
    return @{
        # AMD Ryzen 3000 Series
        '3100'   = @{ Class = 'R3' }
        '3600'   = @{ Class = 'R5' }
        '3600X'  = @{ Class = 'R5' }
        '3600XT' = @{ Class = 'R5' }
        '3700X'  = @{ Class = 'R7' }
        '3800X'  = @{ Class = 'R7' }
        '3900X'  = @{ Class = 'R9' }
        # AMD Ryzen 5000 Series
        '5500X3D' = @{ Class = 'R5' }
        '5600'   = @{ Class = 'R5' }
        '5600F'  = @{ Class = 'R5' }
        '5600G'  = @{ Class = 'R5' }
        '5600X'  = @{ Class = 'R5' }
        '5600X3D' = @{ Class = 'R5' }
        '5700G'  = @{ Class = 'R7' }
        '5700X'  = @{ Class = 'R7' }
        '5700X3D' = @{ Class = 'R7' }
        '5800X'  = @{ Class = 'R7' }
        '5800X3D' = @{ Class = 'R7' }
        '5800XT' = @{ Class = 'R7' }
        '5900X'  = @{ Class = 'R9' }
        '5950X'  = @{ Class = 'R9' }
        # AMD Ryzen 7000 / 8000 Series
        '7400'   = @{ Class = 'R5' }
        '7400F'  = @{ Class = 'R5' }
        '7500F'  = @{ Class = 'R5' }
        '7600'   = @{ Class = 'R5' }
        '7600X'  = @{ Class = 'R5' }
        '7600X3D' = @{ Class = 'R5' }
        '7700'   = @{ Class = 'R7' }
        '7700X'  = @{ Class = 'R7' }
        '7800X3D' = @{ Class = 'R7' }
        '7900'   = @{ Class = 'R9' }
        '7900X'  = @{ Class = 'R9' }
        '7900X3D' = @{ Class = 'R9' }
        '7950X'  = @{ Class = 'R9' }
        '7950X3D' = @{ Class = 'R9' }
        '8600G'  = @{ Class = 'R5' }
        '8700G'  = @{ Class = 'R7' }
        # AMD Ryzen 9000 Series
        '9500F'  = @{ Class = 'R5' }
        '9600'   = @{ Class = 'R5' }
        '9600X'  = @{ Class = 'R5' }
        '9700F'  = @{ Class = 'R7' }
        '9700X'  = @{ Class = 'R7' }
        '9800X3D' = @{ Class = 'R7' }
        '9900X'  = @{ Class = 'R9' }
        '9900X3D' = @{ Class = 'R9' }
        '9950X'  = @{ Class = 'R9' }
        '9950X3D' = @{ Class = 'R9' }

        # Intel 10th - 14th Gen
        '10100'  = @{ Class = 'i3' }
        '10100F' = @{ Class = 'i3' }
        '10400F' = @{ Class = 'i5' }
        '10600K' = @{ Class = 'i5' }
        '10700K' = @{ Class = 'i7' }
        '10900K' = @{ Class = 'i9' }
        '11400'  = @{ Class = 'i5' }
        '11600K' = @{ Class = 'i7' }
        '11700K' = @{ Class = 'i7' }
        '11900K' = @{ Class = 'i9' }
        '12100'  = @{ Class = 'i3' }
        '12400F' = @{ Class = 'i5' }
        '12600K' = @{ Class = 'i5' }
        '12700K' = @{ Class = 'i7' }
        '12900K' = @{ Class = 'i9' }
        '13400F' = @{ Class = 'i5' }
        '13500'  = @{ Class = 'i5' }
        '13600K' = @{ Class = 'i5' }
        '13700K' = @{ Class = 'i7' }
        '13900K' = @{ Class = 'i9' }
        '14400F' = @{ Class = 'i5' }
        '14600K' = @{ Class = 'i5' }
        '14700K' = @{ Class = 'i7' }
        '14900K' = @{ Class = 'i9' }
        # Intel Core Ultra Series (Meteor/Raptor Lake)
        '125H'   = @{ Class = 'U5' }
        '135H'   = @{ Class = 'U5' }
        '155H'   = @{ Class = 'U7' }
        '165H'   = @{ Class = 'U7' }
        '185H'   = @{ Class = 'U9' }
        '125U'   = @{ Class = 'U5' }
        '135U'   = @{ Class = 'U5' }
        '155U'   = @{ Class = 'U7' }
        '165U'   = @{ Class = 'U7' }
        # Intel Core Ultra 200 (Arrow Lake) Series
        '245K'   = @{ Class = 'U5' }
        '265K'   = @{ Class = 'U7' }
        '285K'   = @{ Class = 'U9' }
    }
}

function New-ScoopAppShortcut {
    param(
        [Parameter(Mandatory=$true)][string]$AppName,
        [Parameter(Mandatory=$true)][string]$InstalledAppsDir,
        [Parameter(Mandatory=$true)][string]$DesktopPath
    )
    
    try {
        $CurrentVersionDir = Join-Path $InstalledAppsDir "$AppName\current"
        $ManifestPath = Join-Path $CurrentVersionDir "manifest.json"
        Write-Verbose "Looking for manifest at: $ManifestPath"
        if (-not (Test-Path $ManifestPath)) {
            throw "ManifestError: Manifest not found for '$AppName' at '$ManifestPath'."
        }

        $Manifest = Get-Content -Path $ManifestPath | ConvertFrom-Json

        # Prioritize architecture-specific shortcuts if they exist, otherwise fall back to top-level shortcuts.
        # 1 item arrays are convered to systemobjects, force shortcuts to always be treated as arrays
        # See https://www.reddit.com/r/PowerShell/comments/13j9qng/comment/jnkas4y/
        $ShortcutsList = if ($Manifest.architecture.'64bit'.shortcuts) {
            Write-Verbose "Using 64bit architecture shortcuts for '$AppName'."
            $Manifest.architecture.'64bit'.shortcuts # This is an array of arrays, e.g., [["path", "name"], ...]
        } elseif ( $Manifest.shortcuts ) {
            Write-Verbose "Using top-level shortcuts for '$AppName'."
            $Manifest.shortcuts # This is also an array of arrays
        } else {
            throw "ManifestError: Manifest for '$AppName' is missing the required 'shortcuts' key."
        }

        # Determine the correct shortcut source array.
        # If the first element of the list is an array itself (e.g., multiple shortcuts defined), use that first element.
        # Otherwise, the list has been "unwrapped" by ConvertFrom-Json, so the list itself is the source.
        # This robustly handles both [["path", "name"]] and [["path1", "name1"], ["path2", "name2"]] structures.
        $ShortcutSource = if ($ShortcutsList[0] -is [array]) { $ShortcutsList[0] } else { $ShortcutsList }

        # The 'shortcuts' property is an array of arrays, e.g., [["bin\\7zFM.exe", "7-Zip File Manager"]]
        $ExecutableSubPath = $ShortcutSource[0]
        $DesiredShortcutName = $ShortcutSource[1]
        $ExecutableName = [System.IO.Path]::GetFileName($ExecutableSubPath)

        Write-Host "`n--- Processing App: '$AppName' (Shortcut: '$DesiredShortcutName', Executable: '$ExecutableName') ---"
        $AppExecutablePath = Join-Path -Path $CurrentVersionDir -ChildPath $ExecutableSubPath
        Write-Verbose "Resolved executable path: $AppExecutablePath"
        if (-not (Test-Path $AppExecutablePath)) {
            throw "PathResolutionError: Executable '$ExecutableName' not found at expected path: '$AppExecutablePath'."
        }

        $WshShell = New-Object -ComObject WScript.Shell
        $ShortcutFile = Join-Path -Path $DesktopPath -ChildPath "$DesiredShortcutName.lnk"
        $Shortcut = $WshShell.CreateShortcut($ShortcutFile)

        Write-Verbose "Setting shortcut properties for '$ShortcutFile'."
        $Shortcut.TargetPath = $AppExecutablePath
        $Shortcut.WorkingDirectory = $CurrentVersionDir
        # Set the icon to be the first one available in the executable
        $Shortcut.IconLocation = "$AppExecutablePath, 0"

        # The .Save() method guarantees replacement of the existing .lnk file.
        $Shortcut.Save()
        Write-Host "✅ Shortcut created/updated for '$DesiredShortcutName'."
    }
    catch {
        throw "ProcessingError for '$AppName': $($_.Exception.Message)"
    }
}

function Get-ProcessorShortName {
    <#
    .SYNOPSIS
    Gets the short, human-friendly model name of the computer's processor.
    .DESCRIPTION
    This function queries CIM to retrieve the full name of the CPU, then uses a regular
    expression to remove the manufacturer prefix (e.g., "Intel(R)") and any suffixes
    (e.g., "CPU @ 3.60GHz") to return a shortened model name.
    .EXAMPLE
    # Full Name: "AMD Ryzen 7 8700G w/ Radeon 780M Graphics"
    Get-ProcessorShortName
    # Returns: "R7 8700G"
    .OUTPUTS
    [string] The shortened name of the processor. Returns $null on failure.
    #>
    $fullName = (Get-CimInstance -ClassName Win32_Processor).Name.Trim()
    $cpuLookup = Get-CpuLookupTable

    # Iterate through the known CPU models from the lookup table.
    foreach ($modelNumber in $cpuLookup.Keys) {
        # Check if the full CPU name contains the model number.
        # Using "-match" with word boundaries (\b) ensures we match "10700K" but not "10700".
        if ($fullName -match "\b$([regex]::Escape($modelNumber))\b") {
            $class = $cpuLookup[$modelNumber].Class
            Write-Verbose "Found CPU '$modelNumber' in lookup table. Class: '$class' for '$fullName'"
            return "$class $modelNumber"
        }
    }

    # Fallback for CPUs not in the lookup table: use the old regex method for a best-effort guess.
    Write-Verbose "CPU model not found in lookup table. Using regex fallback for '$fullName'."
    return ($fullName -replace '^(?:Intel\(R\)|AMD)\s+', '' -replace '^Core\(TM\)\s+', '' -replace '^Ryzen ', 'R' -replace '\s+(?:CPU|Processor|with Radeon Graphics|w/ Radeon.*|@).*', '').Trim()
}

function Get-GpuShortName {
    <#
    .SYNOPSIS
    Gets the shortened name(s) of the installed graphics card(s).
    .DESCRIPTION
    This function queries CIM to retrieve the name of the video controller(s).
    It removes common manufacturer branding like "AMD", "Radeon", and "NVIDIA GeForce".
    If multiple GPUs are found, their names are returned as a single comma-separated string.
    .EXAMPLE
    # Full Name: "NVIDIA GeForce RTX 4090"
    Get-GpuShortName
    # Returns: "RTX 4090"
    .OUTPUTS
    [string] The shortened name(s) of the GPU(s). Returns $null on failure.
    #>
    # This can return one or more video controllers
    $gpus = Get-CimInstance -ClassName Win32_VideoController
    if ($gpus) {
        Write-Verbose "Found GPU(s): $($gpus.Name -join ', ')"
        # Get the 'Name' property, remove common branding, trim whitespace, and join them.
        $shortNames = $gpus.Name | ForEach-Object { $_.Trim() -replace '^(?:AMD|Radeon|NVIDIA GeForce)\s*', '' -replace '\s+', ' ' }
        return ($shortNames | Where-Object { $_ }) -join ', '
    }
    # If no GPUs are found, this will throw an error that will be caught by the main handler.
    throw "No video controller found."
}

enum MetricGroupType {
    Unassigned
    CPU
    GPU
}

class MsiAfterburnerConfig {
    # Properties
    [string]$FilePath
    [System.Collections.IDictionary]$Data
    [int]$ChangeCount = 0

    # Constructor
    MsiAfterburnerConfig([string]$Path) {
        $this.FilePath = $Path
        $this.Data = @{}
        $this.ChangeCount = 0
        $this._Load()
    }

    # Private method to load and parse the file
    hidden _Load() {
        $rawConfigContent = ''
        if (Test-Path -Path $this.FilePath -PathType Leaf) {
            $rawConfigContent = Get-Content -Path $this.FilePath -Raw
        }

        $currentSection = $null
        $rawConfigContent -split '(\r?\n)' | ForEach-Object {
            $line = $_.Trim()
            if ($line -match '^\[(.+)\]$') {
                $currentSection = $matches[1]
                if (-not ($this.Data.Keys -contains $currentSection)) {
                    $this.Data[$currentSection] = @{}
                }
            } elseif ($line -match '^(.+?)=(.*)') {
                if ($currentSection) {
                    $key = $matches[1].Trim()
                    $value = $matches[2].Trim()
                    $this.Data[$currentSection][$key] = $value
                }
            }
        }

        Write-Verbose "--- Initial Config Object (Loaded from file) ---"
        Write-Verbose ($this.Data | ConvertTo-Json -Depth 5)
    }

    # Public method to set a value
    SetSectionValue([string]$SectionName, [string]$Key, [string]$Value) {
        $section = $this._TryGetSection($SectionName)

        if ($section.Keys -contains $Key) {
            if ( ($section[$Key] -ne $Value)) {
                 Write-Verbose "Changing [$SectionName][$Key]=$Value"
                $this.ChangeCount++
            }
        }
        else{
            Write-Verbose "Creating [$SectionName][$Key]=$Value"
            $this.ChangeCount++
        }

        $section[$Key] = $Value
    }

    # Public method to set a value in a 'Source' section
    SetSourceSectionValue([string]$SourceName, [string]$Key, [string]$Value) {
        $sectionName = "Source $SourceName"
        $this.SetSectionValue($sectionName, $Key, $Value)
    }

    # Public method to clear the value of a key in a 'Source' section
    ClearSourceSectionValue([string]$SourceName, [string]$Key) {
        $sectionName = "Source $SourceName"
        if ($this.Data.Keys -contains $sectionName) {
            $section = $this.Data[$sectionName]
            if ($section.Keys -contains $Key) {
                Write-Verbose "Removing key '$Key' from section [$sectionName]"
                $section.Remove($Key)
                $this.ChangeCount++
            }
        }
    }

    # Private helper to get or create a section
    hidden [System.Collections.IDictionary] _TryGetSection([string]$SectionName) {
        if (-not ($this.Data.Keys -contains $SectionName)) {
            Write-Verbose "Creating [$SectionName]"
            $this.Data[$SectionName] = @{}
        }
        return $this.Data[$SectionName]
    }

    [MetricGroupType] GetMetricGroupType([string]$MetricName) {
        $groupLookup = @{
            'CPU temperature' = [MetricGroupType]::CPU
            'GPU temperature' = [MetricGroupType]::GPU
            'CPU usage'       = [MetricGroupType]::CPU
            'GPU usage'       = [MetricGroupType]::GPU
            'CPU power'       = [MetricGroupType]::CPU
            'Power'           = [MetricGroupType]::GPU
            'RAM usage'       = [MetricGroupType]::CPU
            'Memory usage'    = [MetricGroupType]::GPU
            'Fan speed'       = [MetricGroupType]::GPU
        }

        if (-not ($groupLookup.Keys -contains $MetricName)) {
            return [MetricGroupType]::Unassigned
        }
        
        return $groupLookup[$MetricName]
    }

    # Public method to save the config back to the file
    Save() {
        if ($this.ChangeCount -eq 0) {
            Write-Host "MSI Afterburner configuration is already up to date. No changes made."
            return
        }

        Write-Verbose "--- Final Config Object (Before INI Conversion) ---"
        Write-Verbose ($this.Data | ConvertTo-Json -Depth 5)

        $sb = [System.Text.StringBuilder]::new()
        foreach ($section in $this.Data.Keys) {
            $sb.AppendLine("[$section]") | Out-Null
            $sectionObject = $this.Data[$section]
            if ($sectionObject) {
                foreach ($key in $sectionObject.Keys) {
                    $value = $sectionObject[$key]
                    $sb.AppendLine("$key=$value") | Out-Null
                }
            }
            $sb.AppendLine() | Out-Null
        }

        $sb.ToString().Trim() | Set-Content -Path $this.FilePath -NoNewline
        Write-Host "MSI Afterburner configuration saved."
    }
}

function Configure-MsiAfterburnerMonitoring {
    <#
    .SYNOPSIS
    Configures MSI Afterburner's On-Screen Display monitoring groups if installed via Scoop.
    .DESCRIPTION
    This function checks for a Scoop installation of 'msiafterburner'. If found, it modifies
    the 'MSIAfterburner.cfg' file to group GPU and CPU metrics (Temperature, Utilization,
    Memory Usage, Power, and Fan Speed) under custom names derived from the system's
    hardware. This makes the On-Screen Display cleaner and more organized.
    #>
    param(
        [Parameter(Mandatory=$true)][string]$InstalledAppsDir
    )
    Write-Verbose "Starting MSI Afterburner configuration."

    $appName = 'msiafterburner'
    $appDir = Join-Path -Path $InstalledAppsDir -ChildPath $appName

    if (-not (Test-Path -Path $appDir -PathType Container)) {
        throw "ConfigError: MSI Afterburner directory not found at '$appDir'."
    }

    $profileDir = Join-Path -Path $appDir -ChildPath 'current\Profiles'
    $configFile = Join-Path -Path $profileDir -ChildPath 'MSIAfterburner.cfg'
    
    if (-not (Test-Path -Path $profileDir -PathType Container)) {
        throw "ConfigError: MSI Afterburner 'Profiles' directory not found at '$profileDir'."
    }

    $config = [MsiAfterburnerConfig]::new($configFile)

    $cpuGroupName = Get-ProcessorShortName
    $gpuGroupName = Get-GpuShortName
    Write-Verbose "CPU Group Name set to: '$cpuGroupName'"
    Write-Verbose "GPU Group Name set to: '$gpuGroupName'"
    $groupKey = "Group"

    # Define the definitive string for the 'Sources' key. This is the single source of truth.
    $sourcesString = "+CPU temperature,+GPU temperature,+CPU usage,+GPU usage,+CPU power,+Power,+RAM usage,+Memory usage,+Fan speed,-FB usage,-VID usage,-BUS usage,-Memory usage \process,-Core clock,-Memory clock,-Power percent,-Fan tachometer,-Temp limit,-Power limit,-Voltage limit,-No load limit,-CPU1 temperature,-CPU2 temperature,-CPU3 temperature,-CPU4 temperature,-CPU5 temperature,-CPU6 temperature,-CPU7 temperature,-CPU8 temperature,-CPU9 temperature,-CPU10 temperature,-CPU11 temperature,-CPU12 temperature,-CPU13 temperature,-CPU14 temperature,-CPU15 temperature,-CPU16 temperature,-CPU1 usage,-CPU2 usage,-CPU3 usage,-CPU4 usage,-CPU5 usage,-CPU6 usage,-CPU7 usage,-CPU8 usage,-CPU9 usage,-CPU10 usage,-CPU11 usage,-CPU12 usage,-CPU13 usage,-CPU14 usage,-CPU15 usage,-CPU16 usage,-CPU1 clock,-CPU2 clock,-CPU3 clock,-CPU4 clock,-CPU5 clock,-CPU6 clock,-CPU7 clock,-CPU8 clock,-CPU9 clock,-CPU10 clock,-CPU11 clock,-CPU12 clock,-CPU13 clock,-CPU14 clock,-CPU15 clock,-CPU16 clock,-CPU clock,-CPU1 power,-CPU2 power,-CPU3 power,-CPU4 power,-CPU5 power,-CPU6 power,-CPU7 power,-CPU8 power,-CPU9 power,-CPU10 power,-CPU11 power,-CPU12 power,-CPU13 power,-CPU14 power,-CPU15 power,-CPU16 power,-RAM usage \process,-Commit charge,+Framerate,-Frametime,-Framerate Min,-Framerate Avg,-Framerate Max,-Framerate 1% Low,-Framerate 0.1% Low"
    $config.SetSectionValue("Settings", "Sources", $sourcesString)

    $enabledMetrics = $sourcesString -split ',' | Where-Object { $_.StartsWith('+') } | ForEach-Object { $_.TrimStart('+') }
    foreach ($sourcename in $enabledMetrics) {
        $config.SetSourceSectionValue($sourcename, "ShowInOSD", '1')

        switch ( $config.GetMetricGroupType($sourcename) ) {
            ([MetricGroupType]::CPU) { $config.SetSourceSectionValue($sourcename, $groupKey, $cpuGroupName) }
            ([MetricGroupType]::GPU) { $config.SetSourceSectionValue($sourcename, $groupKey, $gpuGroupName) }
            ([MetricGroupType]::Unassigned) { $config.ClearSourceSectionValue($sourcename, $groupKey) }
        }
    }

    $config.Save()
}

Write-Host "Starting Scoop Desktop Shortcut Creator on PowerShell $($PSVersionTable.PSVersion.ToString())..."

try {
    CheckScoopAppsDirectoryExists -AppsDir $InstalledAppsDir

    foreach ($AppDirectory in Get-ChildItem $InstalledAppsDir -Directory) {
        if ($ScoopApps -contains $AppDirectory.Name) {
            New-ScoopAppShortcut -AppName $AppDirectory.Name -InstalledAppsDir $InstalledAppsDir -DesktopPath $DesktopPath
        }

        if ($AppDirectory.Name -eq 'msiafterburner') {
            Configure-MsiAfterburnerMonitoring -InstalledAppsDir $InstalledAppsDir
        }
    }

    Write-Host "`nFinished processing all apps. Check your desktop for new shortcuts. 🚀"
}
catch {
    Write-Error "❌ Fatal Error during script execution. The process has been halted."
    # Output the full error record, which includes the message, script stack trace, and line number.
    Write-Error $_
}
