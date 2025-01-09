# __      __ _     _           _____                           
# \ \    / /(_)   | |         /  ___\  _                       
#  \ \  / /  _  __| | ___  __ | (___  | |  __ _ _ __ ___  _ __   ___ _ __ 
#   \ \/ /  | |/ _  |/ _ \/  \\___  \[   ]/ _' | '_ ' _ \| '_ \ / _ \ '__|
#    \  /   | ||(_| || __/|()| ___) | | | |(_| | | | | | | |_) |  __/ |   
#     \/    |_|\___.|\___|\__/|_____/ |_| \___.|_| |_| |_| .__/ \___|_|   
#                                                        | |              
#                                                        |_|
# By Kasey M. Strube
# Version 0.1
#
# Utilizes ffmpeg to automatically add text and timestamps to videos.
# Used primary to convert iPhone .MOVs to MP4 for portability and add a
# camcorder style timestamp for viewing on TV

#Arguments
[CmdletBinding()]
Param (

)
$Verbose = $false
if ($PSBoundParameters.ContainsKey('Verbose')) { # Command line specifies -Verbose[:$false]
    $Verbose = $PsBoundParameters.Get_Item('Verbose')
}
if($Verbose) {
  Write-Host "Running VideoStamper in Verbose Mode"
}

# Function to read from INI file
function Get-IniValue($filePath, $section, $key) {
    if (Test-Path $filePath) {
        $ini = Get-IniContent $filePath
        return $ini[$section][$key]
    }
    return $null
}

# Function to write to INI file
function Set-IniValue($filePath, $section, $key, $value) {
    if (-Not (Test-Path $filePath)) {
        # Create the INI file if it doesn't exist
        $iniContent = @"
[$section]
$key=$value
"@
        $iniContent | Set-Content $filePath
    } else {
        $ini = Get-IniContent $filePath
        $ini[$section][$key] = $value
        $ini | Out-IniFile -FilePath $filePath
    }
}

function Get-IniContent ($filePath)
{
	$ini = @{}
	switch -regex -file $FilePath
	{
    	"^\[(.+)\]" # Section
    	{
        	$section = $matches[1]
        	$ini[$section] = @{}
        	$CommentCount = 0
    	}
    	"^(;.*)$" # Comment
    	{
        	$value = $matches[1]
        	$CommentCount = $CommentCount + 1
        	$name = "Comment" + $CommentCount
        	$ini[$section][$name] = $value
    	}
    	"(.+?)\s*=(.*)" # Key
    	{
        	$name,$value = $matches[1..2]
        	$ini[$section][$name] = $value
    	}
	}
	return $ini
}

function Out-IniFile($InputObject, $FilePath)
{
    $outFile = New-Item -ItemType file -Path $Filepath
    foreach ($i in $InputObject.keys)
    {
        if (!($($InputObject[$i].GetType().Name) -eq "Hashtable"))
        {
            #No Sections
            Add-Content -Path $outFile -Value "$i=$($InputObject[$i])"
        } else {
            #Sections
            Add-Content -Path $outFile -Value "[$i]"
            Foreach ($j in ($InputObject[$i].keys | Sort-Object))
            {
                if ($j -match "^Comment[\d]+") {
                    Add-Content -Path $outFile -Value "$($InputObject[$i][$j])"
                } else {
                    Add-Content -Path $outFile -Value "$j=$($InputObject[$i][$j])"
                }

            }
            Add-Content -Path $outFile -Value ""
        }
    }
}

# Load necessary .NET assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$localAppData = [Environment]::GetFolderPath("LocalApplicationData")
$iniFilePath = Join-Path $localAppData ".videoStamper.ini"


# Fancy ASCII art title screen for "Video Stamper"
$asciiArt = @"
 __      __ _     _           _____                           
 \ \    / /(_)   | |         /  ___\  _                       
  \ \  / /  _  __| | ___  __ | (___  | |  __ _ _ __ ___  _ __   ___ _ __ 
   \ \/ /  | |/ _  |/ _ \/  \\___  \[   ]/ _' | '_ ' _ \| '_ \ / _ \ '__|
    \  /   | ||(_| || __/|()| ___) | | | |(_| | | | | | | |_) |  __/ |   
     \/    |_|\___.|\___|\__/|_____/ |_| \___.|_| |_| |_| .__/ \___|_|   
                                                        | |              
                                                        |_|                                     
"@

# Clear the screen
Clear-Host

# Display the ASCII art title
Write-Output $asciiArt

# Display instructions
Write-Output ""
Write-Output "Welcome to Video Stamper!"
Write-Output "This tool lets you add date and time stamps to your videos with ease."
Write-Output ""
Write-Output "Press any key to begin..."
Write-Output ""

# Pause for any key press
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

# Continue with the rest of the script
Write-Verbose "Starting the Video Stamper script..."

Write-Output "Checking for FFMPEG..."

# Path to ffprobe
$ffprobe = ".\ffprobe.exe"
$ffprobeDir = (Get-Location).Path
# Path to ffmpeg
$ffmpeg = ".\ffmpeg.exe"
$ffmpegDir = (Get-Location).Path

if(-Not (Test-Path $ffprobe)) {
 $validPath = $false
 if(Test-Path $iniFilePath) {
    $ffprobeDir = Get-IniValue -filePath $iniFilePath -section "FFMPEG" -key "ffprobe_path"
    Write-Verbose "stored path: $ffprobeDir"
    if($ffprobeDir) {
      $ffprobe = Join-Path $ffprobeDir "ffprobe.exe"
    }
 }
 
 if(Test-Path $ffprobe) {
    $validPath = $true;
 }

 while(-Not $validPath){
    Write-Output "ffprobe.exe not found. Press any key to browse for ffprobe.exe..."
    # Pause for any key press
    $x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

    # Create and configure the OpenFileDialog
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.InitialDirectory = $ffprobeDir  # Default folder (optional)
    $openFileDialog.Filter = "Executable Files (*.exe)|*.exe|All Files (*.*)|*.*" # Only executable files
    $openFileDialog.Title = "Locate ffprobe.exe"

    # Show the dialog and get the selected file path
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $ffprobe = $openFileDialog.FileName
        $ffprobeItem = (Get-Item $ffprobe)
        $ffprobeName = $ffprobeItem.Name
        $ffprobeDir = $ffprobeItem.DirectoryName
        Write-Output "Current directiory: $ffprobeDir"
    } else {
        Write-Error "ffmprobe.exe was not found, and no file was selected. Exiting"
        exit;
    }
    # Check if the selected file is ffprobe.exe
    if ($ffprobeName -eq "ffprobe.exe") {
        # Save the valid path to the .ini file
        Set-IniValue -filePath $iniFilePath -section "FFMPEG" -key "ffprobe_path" -value $ffprobeDir
        $validPath = $true
    }
  }
  Write-Output "ffprobe.exe found in the current directory. Press any key to continue..."
  # Pause for any key press
  $x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
} else {
  Write-Verbose "ffprobe.exe found in the current directory"
}

if(-Not (Test-Path $ffmpeg)) {
 $validPath = $false

 if(Test-Path $iniFilePath) {
    $ffmpegDir = Get-IniValue -filePath $iniFilePath -section "FFMPEG" -key "ffmpeg_path"
    Write-Verbose "stored path: $ffmpegDir"
    if($ffmpegDir) {
       $ffmpeg = Join-Path $ffmpegDir "ffmpeg.exe"
    }
 } else {
    $ffmpegDir = $ffprobeDir
 }
 
 if(Test-Path $ffmpeg) {
    $validPath = $true;
 }

 while(-Not $validPath){
    Write-Output "ffmpeg.exe not found. Press any key to browse for ffmpeg.exe..."
    # Pause for any key press
    $x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

    # Create and configure the OpenFileDialog
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.InitialDirectory = $ffmpegDir  # Default folder (optional)
    $openFileDialog.Filter = "Executable Files (*.exe)|*.exe|All Files (*.*)|*.*" # Only executable files
    $openFileDialog.Title = "Locate ffmpeg.exe"

    # Show the dialog and get the selected file path
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $ffmpeg = $openFileDialog.FileName
        $ffmpegItem = (Get-Item $ffmpeg)
        $ffmpegName = $ffmpegItem.Name
        $ffmpegDir = $ffmpegItem.DirectoryName
        Write-Output "Current directiory: $ffmpegDir"
    } else {
        Write-Error "ffmpeg.exe was not found, and no file was selected. Exiting"
        exit;
    }
    # Check if the selected file is ffprobe.exe
    if ((Get-Item $ffmpeg).Name -eq "ffmpeg.exe") {
        # Save the valid path to the .ini file
        Set-IniValue -filePath $iniFilePath -section "FFMPEG" -key "ffmpeg_path" -value $ffmpegDir
        $validPath = $true
    }
  }
  Write-Output "ffmpeg.exe found in the current directory. Press any key to continue..."
  # Pause for any key press
  $x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
} else {
  Write-Verbose "ffmpeg.exe found in the current directory"
}

Write-Output "Press any key to select a video..."
# Pause for any key press
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

# Create and configure the OpenFileDialog
$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")  # Default folder (optional)
$openFileDialog.Filter = "Video Files (*.mov;*.mp4)|*.mov;*.mp4|All Files (*.*)|*.*" # File type filter
$openFileDialog.Title = "Select a Video File"

# Show the dialog and get the selected file path
if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $inputFile = $openFileDialog.FileName
    Write-Verbose "Selected File: $inputFile"
} else {
    Write-Error "No file was selected."
    exit
}

# $inputFile now contains the selected file path
Write-Verbose "Input File Path: $inputFile"

Write-Verbose 'FFProbe command: ffprobe -v quiet -print_format json -show_entries stream=width,height -show_entries format_tags "$inputFile"'

# Initialize ffprobe command to extract metadata
$ffprobeCmd = &$ffprobe -v quiet -print_format json -show_entries stream=width,height -show_entries format_tags "$inputFile"

# parse Json
$jsonObject = $ffprobeCmd | ConvertFrom-Json

Write-Verbose "Video Metadata:"
Write-Verbose $jsonObject

# Extract the creation date
$creationDateString = $jsonObject.format.tags.'com.apple.quicktime.creationdate'

Write-Verbose "Apple DateTime Metadata: $creationDateString"

# Check if creation date string is set to something
if(-Not $creationDateString){
   Write-Verbose "Invalid Apple Metadata"
   $creationDateString = $jsonObject.format.tags.'creation_time'
   Write-Verbose "Other DateTime Metadata: $creationDateString"
}

# Check if creation date string still isn't set
if(-Not $creationDateString){
   $creationDate = Get-Date
   Write-Verbose "No valid DateTime Metadata. Setting Time Stamp to $creationDate"
} else {
   # Parse the date string into a DateTime object
   $creationDate = [datetime]::Parse($creationDateString, $null, [System.Globalization.DateTimeStyles]::AssumeUniversal)
   Write-Verbose "Parsed DateTime:"
   Write-Verbose $creationDate
}
# Extract the time zone offset
$offset = $creationDate.ToString("zzz") # Extracts the offset in the format "+/-HH:mm"
Write-Verbose "Timezone Offset $offset"
$direction = [int]($offset[0] + 1) # First char is '+' or '-'
Write-Verbose "Offset direction $direction"
$hours = [int]$offset.Substring(1, 2)
Write-Verbose "Offset by $hours hours"
$minutes = [int]$offset.Substring(4,2)
Write-Verbose "Offset by $minutes minutes"
$seconds = 0

# Calculate the total offset in seconds
$offsetInSeconds = ($hours * 3600) + ($minutes * 60) + $seconds
Write-Verbose "Total Offset in seconds $offsetInSeconds"

# Convert to UTC and calculate the Unix epoch time
$epoch = [int64](Get-Date($creationDate.ToUniversalTime()) -UFormat "%s")

$epoch = $epoch + ($offsetInSeconds * $direction)
Write-Verbose "Unix epoch time $epoch"

# Get the width and height
$width = $jsonObject.streams[0].width
$height = $jsonObject.streams[0].height

Write-Verbose "Video Width: $width"
Write-Verbose "Video Height: $height"

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Video Stamper Settings"
$form.Width = 500
$form.Height = 150
$form.StartPosition = "CenterScreen"

# Create the TableLayoutPanel
$tableLayout = New-Object System.Windows.Forms.TableLayoutPanel
$tableLayout.RowCount = 3
$tableLayout.ColumnCount = 4
$tableLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
$tableLayout.AutoSize = $true
$tableLayout.CellBorderStyle = [System.Windows.Forms.TableLayoutPanelCellBorderStyle]::None
$form.Controls.Add($tableLayout)

# Set fixed column widths to match input widths
$tableLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 230))) | Out-Null  # Font dropdown
$tableLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 50))) | Out-Null  # Size input
$tableLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 90))) | Out-Null  # Color dropdown
$tableLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 100))) | Out-Null # Border checkbox

# Set row styles
$tableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null # Labels
$tableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null # Inputs
$tableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null # OK Button

# Helper function to add controls to TableLayoutPanel
function Add-Control {
    param (
        [System.Windows.Forms.TableLayoutPanel]$panel,
        [System.Windows.Forms.Control]$control,
        [int]$row,
        [int]$column
    )
    $panel.Controls.Add($control, $column, $row)
}

# Create Labels
$fontLabel = New-Object System.Windows.Forms.Label
$fontLabel.Text = "Font"
$fontLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft

$sizeLabel = New-Object System.Windows.Forms.Label
$sizeLabel.Text = "Size"
$sizeLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft

$colorLabel = New-Object System.Windows.Forms.Label
$colorLabel.Text = "Color"
$colorLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft

$borderLabel = New-Object System.Windows.Forms.Label
$borderLabel.Text = "Border"
$borderLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft

# Add Labels to Row 0
Add-Control -panel $tableLayout -control $fontLabel -row 0 -column 0
Add-Control -panel $tableLayout -control $sizeLabel -row 0 -column 1
Add-Control -panel $tableLayout -control $colorLabel -row 0 -column 2
Add-Control -panel $tableLayout -control $borderLabel -row 0 -column 3

# Load Fonts
$fonts = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts'
$fontList = @()

foreach ($font in $fonts.PSObject.Properties) {
  if($font.Name -ne "PSPath" -And $font.Name -ne "PSParentPath" -And $font.Name -ne "PSChildName" -And $font.Name -ne "PSDrive" -And $font.Name -ne "PSProvider"){
      $name = $font.Name -replace '\s*\(TrueType\)$', '' -replace '^\s+|\s+$', ''
      $fontList += [PSCustomObject]@{
          Name      = $name
          FileName = $font.Value
      }
  }
}

$fontList = $fontList | Sort-Object Name

# Create Inputs
$fontDropdown = New-Object System.Windows.Forms.ComboBox
$fontDropdown.Width = 230
$fontDropdown.DropDownStyle = "DropDownList"
$fontList | ForEach-Object {
    $fontDropdown.Items.Add($_.Name)
} | Out-Null
$fontDropdown.SelectedIndex = 0

$sizeInput = New-Object System.Windows.Forms.NumericUpDown
$sizeInput.Width = 50
$sizeInput.Minimum = 2
$sizeInput.Maximum = 256
$sizeInput.Value = 12
$sizeInput.Increment = 2

$colorDropdown = New-Object System.Windows.Forms.ComboBox
$colorDropdown.Width = 90
$colorDropdown.DropDownStyle = "DropDownList"
$ffmpegColors = @("black", "white", "red", "green", "blue", "yellow", "magenta", "cyan", "gray", "darkgray")
$ffmpegColors | ForEach-Object { $colorDropdown.Items.Add($_) } | Out-Null
$colorDropdown.SelectedIndex = 0

$borderCheckbox = New-Object System.Windows.Forms.CheckBox
$borderCheckbox.Text = "Enable"
$borderCheckbox.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter

# Add Inputs to Row 1
Add-Control -panel $tableLayout -control $fontDropdown -row 1 -column 0
Add-Control -panel $tableLayout -control $sizeInput -row 1 -column 1
Add-Control -panel $tableLayout -control $colorDropdown -row 1 -column 2
Add-Control -panel $tableLayout -control $borderCheckbox -row 1 -column 3

# Create and Add OK Button to Row 2
$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "OK"
$okButton.Width = 100
$okButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom
$tableLayout.SetColumnSpan($okButton, 4)  # Span the button across all columns
Add-Control -panel $tableLayout -control $okButton -row 2 -column 0

# Handle OK Button Click
$okButton.Add_Click({
    $form.Tag = @{
        Font = $fontDropdown.SelectedItem
        Size = $sizeInput.Value
        Color = $colorDropdown.SelectedItem
        Border = $borderCheckbox.Checked
    }
    $form.Close() | Out-Null
})

# Show the form
$form.ShowDialog()

# Process the results
if ($form.Tag -eq $null) {
 Write-Output "Operation canceled by the user."
 exit
}
$userInput = $form.Tag
$font = $fontList | Where-Object { $_.Name -eq $($userInput.Font) }
Write-Verbose "Selected Font: $font.Name"
$fontfile = $font.FileName
$fontfile = "C\:/Windows/Fonts/" + $fontfile
Write-Verbose "Font File: $fontfile"
$fontsize = $($userInput.Size)
Write-Verbose "Font Size: $fontsize"
$fontcolor = $($userInput.Color)
Write-Verbose "Font Color: $fontcolor"
if( $($userInput.Border)){
  $borderw = "borderw=1:bordercolor=black:"
  Write-Verbose "Border Settings: $borderw"
} else {
  $borderw = ""
}

# Determine orientation
if ($height -gt $width) {
# Portrait
    Write-Output "Video is Portrait Orientation"
    if($Verbose){ 
      &$ffmpeg -y -i "$inputFile" -vf "drawtext=fontfile='${fontfile}':fontsize=${fontsize}:fontcolor=${fontcolor}:${borderw}x=w-(w/2):y=h-(h/8):text='%{pts\:gmtime\:$epoch\:%#m\\/%#d\\/%Y}',drawtext=fontfile='${fontfile}':fontsize=${fontsize}:fontcolor=${fontcolor}:${borderw}x=w-(w/2):y=(h-(h/8))+text_h+10:text='%{pts\:gmtime\:$epoch\:%#I\\\:%M\\\:%S %p}'" -crf 27 -movflags use_metadata_tags -map_metadata 0 -preset ultrafast -f mp4 output.mp4
    } else {
     &$ffmpeg -hide_banner -loglevel error -y -i "$inputFile" -vf "drawtext=fontfile='${fontfile}':fontsize=${fontsize}:fontcolor=${fontcolor}:${borderw}x=w-(w/2):y=h-(h/8):text='%{pts\:gmtime\:$epoch\:%#m\\/%#d\\/%Y}',drawtext=fontfile='${fontfile}':fontsize=${fontsize}:fontcolor=${fontcolor}:${borderw}x=w-(w/2):y=(h-(h/8))+text_h+10:text='%{pts\:gmtime\:$epoch\:%#I\\\:%M\\\:%S %p}'" -crf 27 -movflags use_metadata_tags -map_metadata 0 -preset ultrafast -f mp4 output.mp4
    }
} else {
# Landscape
    Write-Output "Video is Landscape Orientation"
    if($Verbose){
      &$ffmpeg -y -i "$inputFile" -vf "drawtext=fontfile='${fontfile}':fontsize=${fontsize}:fontcolor=${fontcolor}:${borderw}x=w-(w/4):y=h-(h/8):text='%{pts\:gmtime\:$epoch\:%#m\\/%#d\\/%Y}',drawtext=fontfile='${fontfile}':fontsize=${fontsize}:fontcolor=${fontcolor}:${borderw}x=w-(w/4):y=(h-(h/8))+text_h+10:text='%{pts\:gmtime\:$epoch\:%#I\\\:%M\\\:%S %p}'" -crf 27 -movflags use_metadata_tags -map_metadata 0 -preset ultrafast -f mp4 output.mp4
    } else {
      &$ffmpeg -hide_banner -loglevel error -y -i "$inputFile" -vf "drawtext=fontfile='${fontfile}':fontsize=${fontsize}:fontcolor=${fontcolor}:${borderw}x=w-(w/4):y=h-(h/8):text='%{pts\:gmtime\:$epoch\:%#m\\/%#d\\/%Y}',drawtext=fontfile='${fontfile}':fontsize=${fontsize}:fontcolor=${fontcolor}:${borderw}x=w-(w/4):y=(h-(h/8))+text_h+10:text='%{pts\:gmtime\:$epoch\:%#I\\\:%M\\\:%S %p}'" -crf 27 -movflags use_metadata_tags -map_metadata 0 -preset ultrafast -f mp4 output.mp4
    }
}
Write-Host "Video has been Stamped!"