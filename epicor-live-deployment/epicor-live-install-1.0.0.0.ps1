# *******************************************
#             -Script Purpose-
# - Used by Intune to install Epicor client
# -     Copy and extract a zip file
# - Create desktop and start menu shortcuts
# *******************************************
# Exit Codes:
# - 0: Success
# - 12: Copy failed
# - 13: Expansion failed
# - 14: Shortcut creation failed
# - 15: Zip file removal failed
# - 16: Script does not have write access
# *******************************************

# ********************************
# *  FUNCTION: Test-WriteAccess  *
# ********************************

# Define the Test-WriteAccess function
function Test-WriteAccess {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Path
    )
    Write-Output "Function: Test-WriteAccess" | Out-File -Append $logFile
    $testPath = Join-Path -Path $Path -ChildPath ("testfile_" + [guid]::NewGuid())
    try {
        New-Item -Path $testPath -ItemType File -ErrorAction Stop | Out-Null
        Remove-Item -Path $testPath -ErrorAction Stop
        return $true
    }
    catch {
        return $false
    }
}

# *********************************
# *  FUNCTION: Test-Shortcut  *
# *********************************

# Define the Test-Shortcut function
function Test-Shortcut {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ShortcutPath
    )
    Write-Output "Function: Test-Shortcut" | Out-File -Append $logFile
    # Check if the shortcut was created
    if (Test-Path -Path $ShortcutPath) {
        # Write-Output "Shortcut was created successfully at $ShortcutPath" | Out-File -Append $logFile
        return $true
    }
    else {
        # Write-Output "Failed to create shortcut at $ShortcutPath" | Out-File -Append $logFile
        return $false
    }
}

# *****************************
# *  DEFINE & INITIALIZE LOG  *
# *****************************
# Get the current timestamp
$startTimestamp = Get-Date
# Define the log directory and log file
$scriptName = $MyInvocation.MyCommand.Name
$logDir = "C:\.EpicorLogs"
$logFolder = Join-Path -Path $logDir -ChildPath "$scriptName"
$logFile = Join-Path -Path $logFolder -ChildPath "$scriptName.log"
# Create the log directory if it doesn't exist
if (!(Test-Path -Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir | Out-Null
}
# Create the log folder if it doesn't exist
if (!(Test-Path -Path $logFolder)) {
    New-Item -ItemType Directory -Path $logFolder | Out-Null
}
Write-Output "Initialize $scriptName log" | Out-File -Append $logFile
# Output the timestamp
Write-Output "$scriptName startTime: $startTimestamp" | Out-File -Append $logFile



# *****************************
# *  DEFINE ZIP FILE DETAILS  *
# *****************************
# Define the source and destination for the zip file
$sourceZip = "161830-LIVE.zip"
$destinationDir = "C:\Epicor\ERPDT"
Write-Output "Succesfully Define the source ($sourceZip) and destination ($destinationDir) for the zip file" | Out-File -Append $logFile

if (Test-Path -Path "C:\Epicor") {
    Write-Output "C:\Epicor directory found" | Out-File -Append $logFile
    Write-Output "Testing write access to $destinationDir" | Out-File -Append $logFile
    if (!(Test-WriteAccess -Path $destinationDir)) {
        Write-Output "The script does not have write access to the destination directory: $destinationDir" | Out-File -Append $logFile
        Write-Output "Attempting to delete existing zip file(s) and/or directory(s)." | Out-File -Append $logFile    
        try {
            # Call Remove-File function with the path to the file to delete
            Remove-Item C:\Epicor\ERPDT\161830-LIVE.zip
        
            # Call Remove-Folder function with the path to the folder to delete
            Remove-Item "C:\Epicor\ERPDT\161830-LIVE"
            Write-Output "Done attempting to delete existing zip file(s) & directory(s)." | Out-File -Append $logFile
    
        }
        catch {
            Write-Output "The script has write access to $destinationDir" | Out-File -Append $logFile
        }
        if (!(Test-WriteAccess -Path $destinationDir)) {
            Write-Output "The script has write access to $destinationDir" | Out-File -Append $logFile
            exit 16 # Script does not have write access to $destinationDir
        }
    }
    else {
        Write-Output "The script has write access to $destinationDir" | Out-File -Append $logFile
    }
}
else {
    Write-Output "C:\Epicor directory doesn't exist" | Out-File -Append $logFile
}



# ***********************************
# *  DEFINE TEMP DIRECTORY DETAILS  *
# ***********************************

# Define a temporary directory for extracting the zip file
$tempDirParent = "C:\.EpicorTemp"
Write-Output "Defined $tempDirParent as the Epicor install temp directory" | Out-File -Append $logFile

# Check if the parent directory exists
if (!(Test-Path -Path $tempDirParent)) {
    # The directory does not exist, so create it
    Write-Output "No temp directory found. Creating.. " | Out-File -Append $logFile
    New-Item -Path $tempDirParent -ItemType Directory | Out-File -Append $logFile
}
Write-Output "Succesfully created/validated parent directory $tempDirParent" | Out-File -Append $logFile

Write-Output "Begin creating temp folders " | Out-File -Append $logFile

$tempDir = Join-Path -Path $tempDirParent -ChildPath "Temp"
# $tempDirZip = Join-Path -Path $tempDirParent -ChildPath "Zip"

# Check if the temp directory exists
if (!(Test-Path -Path $tempDir)) {
    # The directory does not exist, so create it
    Write-Output "The directory does not exist, creating.." | Out-File -Append $logFile
    New-Item -Path $tempDir -ItemType Directory | Out-File -Append $logFile
    Write-Output "Done.." | Out-File -Append $logFile
}

# *****************************
# *      EXPAND ZIP FILE      *
# * Exit 13 = Expansion Error *
# *****************************

Write-Output "Begin expand & move zip file" | Out-File -Append $logFile

# Define the path to the zip file in the temp directory
$sourceZipPath = Join-Path -Path ".\" -ChildPath $sourceZip
# OLD: $tempDestinationZip = Join-Path -Path $tempDirZip -ChildPath $sourceZip
# Write-Output "Successfully defined the path to $sourceZip in $tempDirZip directory" | Out-File -Append $logFile

# Expand the zip file to the temporary directory

try {
    Write-Output "Begin expanding $sourceZip to $tempDir" | Out-File -Append $logFile
    Expand-Archive -Path $sourceZipPath -DestinationPath $tempDir -Force
    Write-Output "Successfully expanded $sourceZip to $tempDir" | Out-File -Append $logFile
}
catch {
    Write-Output "Expand operation failed: $_" | Out-File -Append $logFile
    exit 13 # Expansion failed
}

Write-Output "End expand zip file" | Out-File -Append $logFile

# ************************
# *  CREATE DIRECTORIES  *
# ************************

# Define the directories
$epicorDir = "C:\Epicor"
$erpdtDir = "C:\Epicor\ERPDT"

# Create the directories if they do not exist
if (!(Test-Path -Path $epicorDir)) {
    New-Item -Path $epicorDir -ItemType Directory | Out-Null
    Write-Output "Created directory $epicorDir" | Out-File -Append $logFile
}

if (!(Test-Path -Path $erpdtDir)) {
    New-Item -Path $erpdtDir -ItemType Directory | Out-Null
    Write-Output "Created directory $erpdtDir" | Out-File -Append $logFile
}

# **************************
# *  MOVE EXTRACTED FILES  *
# **************************

try {
    Write-Output "Begin moving the extracted files to $destinationDir" | Out-File -Append $logFile
    Get-ChildItem -Path $tempDir | Move-Item -Destination $destinationDir | Out-File -Append $logFile
    Write-Output "Successfully moved the extracted files to $destinationDir" | Out-File -Append $logFile
}
catch {
    Write-Output "Move operation failed: $_" | Out-File -Append $logFile
    exit 14
}

Write-Output "End expand & move zip file" | Out-File -Append $logFile

# *****************************
# *      CREATE SHORTCUTS     *
# *****************************

Write-Output "Begin Create shortcuts" | Out-File -Append $logFile
try {
    Write-Output "Begin define shortcut details" | Out-File -Append $logFile

    # Define the path to the executable
    $executablePath = "C:\Epicor\ERPDT\161830-LIVE\Client\Epicor.exe"
    Write-Output "executablePath: $executablePath" | Out-File -Append $logFile

    # Define the path to the icon file
    $iconPath = "C:\Epicor\ERPDT\161830-LIVE\Icon\Epicor.ico"
    Write-Output "iconPath: $iconPath" | Out-File -Append $logFile

    # Define the argument for the executable
    $executableArg = "/config=saas5127.sysconfig /KineticHome"
    Write-Output "executableArg: $executableArg" | Out-File -Append $logFile

    # Define the working directory for the shortcut
    $workingDirectory = "C:\Epicor\ERPDT\161830-LIVE\Client"
    Write-Output "workingDirectory: $workingDirectory" | Out-File -Append $logFile

    # Define the paths for the shortcuts
    # Use the Public Desktop and All Users Start Menu
    $desktopShortcut = Join-Path -Path "$env:Public\Desktop" -ChildPath "Epicor Live.lnk"
    $startMenuShortcut = Join-Path -Path "$env:ProgramData\Microsoft\Windows\Start Menu\Programs" -ChildPath "Epicor Live.lnk"
    Write-Output "Defined the paths for the shortcuts" | Out-File -Append $logFile

    # Create a WScript Shell object to create the shortcuts
    $WshShell = New-Object -comObject WScript.Shell
    Write-Output "Created a WScript Shell object to create the shortcuts" | Out-File -Append $logFile
}
catch {
    Write-Output "Shortcut path definition failed: $_" | Out-File -Append $logFile
}
Write-Output "End define shortcut details" | Out-File -Append $logFile

# ****************************
# * CHECK & DELETE SHORTCUTS *
# ****************************

# Check for & delete StartMenu shortcut
if (Test-Path $startMenuShortcut) {
    Write-Output "Start menu shortcut found. Deleting.." | Out-File -Append $logFile
    Remove-Item "$startMenuShortcut" | Out-File -Append $logFile
    Write-Output "Start menu shortcut deleted." | Out-File -Append $logFile
}
else {
    Write-Output "Start menu shortcut not found" | Out-File -Append $logFile
}

# Check for & delete desktop shortcut
if (Test-Path $desktopShortcut) {
    Write-Output "Desktop shortcut found. Deleting.." | Out-File -Append $logFile
    Remove-Item "$desktopShortcut" | Out-File -Append $logFile
    Write-Output "Desktop shortcut deleted." | Out-File -Append $logFile
}
else {
    Write-Output "Desktop shortcut not found" | Out-File -Append $logFile
}

Write-Output "Desktop & Start Menu shortcut deletion successful or not needed." | Out-File -Append $logFile

# ********************************
# *   CREATE DESKTOP SHORTCUTS   *
# ********************************

Write-Output "Begin Create desktop shortcuts" | Out-File -Append $logFile

try {
    $Shortcut = $WshShell.CreateShortcut($desktopShortcut)
    $Shortcut.TargetPath = $executablePath
    $Shortcut.Arguments = $executableArg
    $Shortcut.WorkingDirectory = $workingDirectory
    $Shortcut.IconLocation = $iconPath
    $Shortcut.Save()
    Write-Output "Desktop shortcut created" | Out-File -Append $logFile
}
catch {
    Write-Output "'All Users' Desktop shortcut creation failed: $_" | Out-File -Append $logFile
    Write-Output "Attempting to create for current user instead.." | Out-File -Append $logFile
    try {
        $currentUserDesktopShortcut = Join-Path -Path "$env:USERPROFILE\Desktop" -ChildPath "Epicor Live.lnk"
        Write-Output "Defined the paths for the shortcuts.." | Out-File -Append $logFile

        Write-Output "Attempting to create shortcut.." | Out-File -Append $logFile

        $Shortcut = $WshShell.CreateShortcut($currentUserDesktopShortcut)
        $Shortcut.TargetPath = $executablePath
        $Shortcut.Arguments = $executableArg
        $Shortcut.WorkingDirectory = $workingDirectory
        $Shortcut.IconLocation = $iconPath
        $Shortcut.Save()
        Write-Output "Current User Desktop shortcut created" | Out-File -Append $logFile
    }
    catch {
        Write-Output "Current User Desktop shortcut creation failed: $_" | Out-File -Append $logFile
    }    
}

Write-Output "End Create desktop shortcuts" | Out-File -Append $logFile

# ***********************************
# *   CREATE START MENU SHORTCUTS   *
# ***********************************

Write-Output "Begin Create Start Menu shortcuts" | Out-File -Append $logFile

try {
    $Shortcut = $WshShell.CreateShortcut($startMenuShortcut)
    $Shortcut.TargetPath = $executablePath
    $Shortcut.Arguments = $executableArg
    $Shortcut.WorkingDirectory = $workingDirectory
    $Shortcut.IconLocation = $iconPath
    $Shortcut.Save()
    Write-Output "Start menu shortcut created" | Out-File -Append $logFile
}
catch {
    Write-Output "'All Users' Start menu shortcut creation failed: $_" | Out-File -Append $logFile
    Write-Output "Attempting to create for current user instead.." | Out-File -Append $logFile

    try {
        $currentUserStartMenuShortcut = Join-Path -Path "$env:APPDATA\Microsoft\Windows\Start Menu\Programs" -ChildPath "Epicor Live.lnk"
        Write-Output "Defined the paths for the shortcuts.." | Out-File -Append $logFile

        Write-Output "Attempting to create shortcut.." | Out-File -Append $logFile
        $Shortcut = $WshShell.CreateShortcut($currentUserStartMenuShortcut)
        $Shortcut.TargetPath = $executablePath
        $Shortcut.Arguments = $executableArg
        $Shortcut.WorkingDirectory = $workingDirectory
        $Shortcut.IconLocation = $iconPath
        $Shortcut.Save()
        Write-Output "Start menu shortcut created" | Out-File -Append $logFile
    }
    catch {
        Write-Output "Current User Desktop shortcut creation failed: $_" | Out-File -Append $logFile
    }    
}


Write-Output "End Create Start Menu shortcuts" | Out-File -Append $logFile

# ********************
# *  Test SHORTCUTS  *
# ********************

# Test the StartMenu shortcut creation
if (!(Test-Shortcut -ShortcutPath $startMenuShortcut)) {
    Write-Output "StartMenu shortcut validation failed" | Out-File -Append $logFile
}

# Test the Desktop shortcut creation
if (!(Test-Shortcut -ShortcutPath $desktopShortcut)) {
    Write-Output "Desktop shortcut validation failed" | Out-File -Append $logFile
}

Write-Output "Desktop & Start Menu shortcut validation passed if both no failures directly above this line." | Out-File -Append $logFile

# *******************
# * REMOVE ZIP FILE *
# *******************

# Always clean up the temporary files
Write-Output "Removing temporary files" | Out-File -Append $logFile
if (Test-Path -Path $tempDir) {
    Remove-Item -Path $tempDir -Recurse -Force
}

# *************************
# *  SET SCRIPT END TIME  *
# *************************

# Get the current timestamp
$endTimeStamp = Get-Date

# Output the timestamp
Write-Output "$scriptName endTime: $endTimeStamp" | Out-File -Append $logFile

# *******************
# *  END OF SCRIPT  *
# *******************

Write-Output "End of script" | Out-File -Append $logFile
exit 0 # Success
```