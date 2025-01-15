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

# Function to convert time format strings
function Convert-DateFormat {
    param (
        [string]$FormatString
    )
	
    Write-Verbose "Input format is: $FormatString"
	
    # Mapping of ISO 8601 components to strftime format codes
    $strftimeMap = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([System.StringComparer]::Ordinal)
    $strftimeMap["yyyy"] = "%Y"  # Year (4 digits)
    $strftimeMap["yy"] = "%y"    # Year (2 digits)
    $strftimeMap["MMMM"] = "%B"  # Full month name
    $strftimeMap["MMM"] = "%b"   # Abbreviated month name
    $strftimeMap["MM"] = "%m"    # Month (2 digits)
    $strftimeMap["M"] = "%#m"    # Month (no leading 0)
    $strftimeMap["dddd"] = "%A"  # Full weekday name
    $strftimeMap["ddd"] = "%a"   # Abbreviated weekday name
    $strftimeMap["dd"] = "%d"    # Day of the month (2 digits)
    $strftimeMap["d"] = "%#d"    # Day of the month (no leading 0)
    
    # Replace components with their strftime equivalents
    $strftimeFormat = ""
    $l = $FormatString.Length
    $string = ""
    $lastChar = $null

    for ($i = 0; $i -lt $l; $i++) {
        $currChar = $FormatString[$i]

        if ($string -and $currChar -eq $lastChar) {
            $string += $currChar
        } else {
            if ($string -and $strftimeMap.ContainsKey($string)) {
                Write-Verbose "Match found for '$string' -> '$($strftimeMap[$string])'"
                $strftimeFormat += $strftimeMap[$string]
            } elseif ($string) {
                Write-Verbose "No match for '$string'. Keeping as is."
                $strftimeFormat += $string
            }
            $string = $currChar
        }
        $lastChar = $currChar
    }

    # Final replacement for the last collected string
    if ($strftimeMap.ContainsKey($string)) {
        Write-Verbose "Final match for '$string' -> '$($strftimeMap[$string])'"
        $strftimeFormat += $strftimeMap[$string]
    } else {
        Write-Verbose "No match for final '$string'. Keeping as is."
        $strftimeFormat += $string
    }

    Write-Verbose "New datetime format string is: $strftimeFormat"
    return $strftimeFormat
}

# Function to convert time format strings
function Convert-TimeFormat {
    param (
        [string]$FormatString
    )
	
    Write-Verbose "Input format is: $FormatString"
	
    # Mapping of ISO 8601 components to strftime format codes
    $strftimeMap = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([System.StringComparer]::Ordinal)
    $strftimeMap["HH"] = "%H"    # Hour (24-hour format)
    $strftimeMap["H"] = "%#H"    # Hour (24-hour format no leading 0)
    $strftimeMap["hh"] = "%I"    # Hour (12-hour format)
    $strftimeMap["h"] = "%#I"    # Hour (12-hour format no leading 0)
    $strftimeMap["mm"] = "%M"    # Minute
    $strftimeMap["m"] = "%#M"    # Minute (no leading 0)
    $strftimeMap["ss"] = "%S"    # Second
    $strftimeMap["s"] = "%#S"    # Second (no leading 0)
    $strftimeMap["tt"] = "%p"    # AM/PM
    
    # Replace components with their strftime equivalents
    $strftimeFormat = ""
    $l = $FormatString.Length
    $string = ""
    $lastChar = $null

    for ($i = 0; $i -lt $l; $i++) {
        $currChar = $FormatString[$i]

        if ($string -and $currChar -eq $lastChar) {
            $string += $currChar
        } else {
            if ($string -and $strftimeMap.ContainsKey($string)) {
                Write-Verbose "Match found for '$string' -> '$($strftimeMap[$string])'"
                $strftimeFormat += $strftimeMap[$string]
            } elseif ($string) {
                Write-Verbose "No match for '$string'. Keeping as is."
                $strftimeFormat += $string
            }
            $string = $currChar
        }
        $lastChar = $currChar
    }

    # Final replacement for the last collected string
    if ($strftimeMap.ContainsKey($string)) {
        Write-Verbose "Final match for '$string' -> '$($strftimeMap[$string])'"
        $strftimeFormat += $strftimeMap[$string]
    } else {
        Write-Verbose "No match for final '$string'. Keeping as is."
        $strftimeFormat += $string
    }

    Write-Verbose "New datetime format string is: $strftimeFormat"
    return $strftimeFormat
}

# Function to escaoe ffmpeg option text
function Escape-FFmpegOption {
    param (
        [string]$UnescapedString
    )
	Write-Verbose "Unescaped String is: $UnescapedString"
    $escaped = $UnescapedString -replace "'", "\\'" -replace "\\", "\\\\" -replace ":", "\\\:" -replace ",", "\\," -replace "/", "\\/"
	Write-Verbose "Escaped String is: $escaped"
    return $escaped
}

# Function to read value from INI file
function Get-IniValue($filePath, $section, $key) {
    if (Test-Path $filePath) {
        $ini = Get-IniContent -filePath $filePath

        # Check if the section exists
        if ($ini.ContainsKey($section)) {
            # Check if the key exists within the section
            if ($ini[$section].ContainsKey($key)) {
                return $ini[$section][$key]
            } else {
                Write-Verbose "Key '$key' does not exist in section '$section'."
                return $null
            }
        } else {
            Write-Verbose "Section '$section' does not exist in the INI file."
            return $null
        }
    } else {
        Write-Verbose "INI file '$filePath' does not exist."
        return $null
    }
}

# Function to write value to INI file
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
		# Check if the section exists
        if (-Not $ini.ContainsKey($section)) {
            # Add the section if it doesn't exist
            $ini[$section] = @{}
            Write-Verbose "Section '$section' added to INI."
        }
        $ini[$section][$key] = $value
        Out-IniFile -InputObject $ini -FilePath $filePath
    }
}

# Fuction that loads INI to variable
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

# Function that outputs a properly formatted object to INI
function Out-IniFile {
    param (
        [Parameter(Mandatory)]
        [hashtable]$InputObject,

        [Parameter(Mandatory)]
        [string]$FilePath
    )

    # Safely handle overwriting the file
    if (Test-Path $FilePath) {
        Remove-Item -Path $FilePath -Force
    }

    # Write content to the file
    foreach ($section in $InputObject.Keys) {
        if ($InputObject[$section] -isnot [hashtable]) {
            # No sections, just key-value pairs
            Add-Content -Path $FilePath -Value "$section=$($InputObject[$section])"
        } else {
            # Sections with key-value pairs
            Add-Content -Path $FilePath -Value "[$section]"
            foreach ($key in ($InputObject[$section].Keys | Sort-Object)) {
                if ($key -match "^Comment[\d]+") {
                    # Handle comments
                    Add-Content -Path $FilePath -Value "$($InputObject[$section][$key])"
                } else {
                    # Write key-value pairs
                    Add-Content -Path $FilePath -Value "$key=$($InputObject[$section][$key])"
                }
            }
            Add-Content -Path $FilePath -Value "" # Add a blank line after the section
        }
    }
}


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

# Function to process .NET form inputs into a valid ffmpeg command
function Process-Video {
	param(
		[System.Windows.Forms.Form]$form,
		[int]$height,
		[int]$width
		)
	# Process Date Stamp
	$userInput = $form.Tag
	$font = $fontList | Where-Object { $_.Name -eq $($userInput.Font) }
	Write-Verbose "Datestamp Settings-"
	Write-Verbose "Selected Font: $font.Name"
	$fontfile = $font.FileName
	$fontfile = "C\:/Windows/Fonts/" + $fontfile
	Write-Verbose "Font File: $fontfile"
	$fontsize = $($userInput.Size)
	Write-Verbose "Font Size: $fontsize"
	$fontcolor = $($userInput.Color)
	Write-Verbose "Font Color: $fontcolor"
	$borderColor = $($userInput.BorderColor)
	if($borderColor -ne "none"){
	  Write-Verbose "Border Color: $borderColor"
	  $borderSize = $($userInput.BorderSize)
	  Write-Verbose "Border Size: $borderSize"
	  $borderw = "borderw=${borderSize}:bordercolor=${borderColor}:"
	} else {
	  $borderw = ""
	}
	$position = $($userInput.Position)
	$month = $($userInput.Month)
	$day = $($userInput.Day)
	$year = $($userInput.Year)
	$hour = $($userInput.Hour)
	$minute = $($userInput.Minute)
	$second = $($userInput.Second)
	$ampmText = $($userInput.AMPM)
	$dateFormat = $($userInput.DateFormat)
	$timeFormat = $($userInput.TimeFormat)
	Write-Verbose "Text Position: $position"
	$xPadding = $($userInput.XPadding)
	$yPadding = $($userInput.YPadding)

	# Convert the hour to 24-hour format based on AM/PM
	if ($ampmText -eq "PM" -and $hour -lt 12) {
		$hour += 12
	} elseif ($ampmText -eq "AM" -and $hour -eq 12) {
		$hour = 0
	}

	# Create the DateTime object
	$date = [datetime]::new($year, $month, $day, $hour, $minute, $second)

	# Format the DateTime object using -UFormat
	$epoch = Get-Date -Date $date -UFormat "%s"
	Write-Verbose "Unix epoch time $epoch"

	$escapedText = Escape-FFmpegOption -UnescapedString (Convert-DateFormat -FormatString $dateFormat)
	if($($userInput.TimePosition) -eq "Single Line" -And $($userInput.EnableTime)){
		if($($userInput.Enable)){
			$text = "%{pts\:gmtime\:" + $epoch + "\:" + $escapedText 
			$escapedText = Escape-FFmpegOption -UnescapedString (Convert-TimeFormat -FormatString $timeFormat)
			$text += " " + $escapedText + "}"
		} else {
			$userInput.TimePosition = $position
		}
	} else {
		$text = "%{pts\:gmtime\:" + $epoch + "\:" + $escapedText + "}"
	}

	# Init drawtext hashtable
	$drawcalls = @{
		"Top Left"     = @{ "Count" = 0; "drawtexts" = @() }
		"Top Middle"   = @{ "Count" = 0; "drawtexts" = @() }
		"Top Right"    = @{ "Count" = 0; "drawtexts" = @() }
		"Middle Left"  = @{ "Count" = 0; "drawtexts" = @() }
		"Middle"       = @{ "Count" = 0; "drawtexts" = @() }
		"Middle Right" = @{ "Count" = 0; "drawtexts" = @() }
		"Bottom Left"  = @{ "Count" = 0; "drawtexts" = @() }
		"Bottom Middle" = @{ "Count" = 0; "drawtexts" = @() }
		"Bottom Right" = @{ "Count" = 0; "drawtexts" = @() }
	}

	switch($position) {
	  "Top Left" { 
		$xCoor = "w-((w/100)*(100-${xPadding}))"
		$yCoor = $height-((($height/100)*(100-$yPadding)))
	  }
	  "Top Middle" {
		$xCoor = "w-(w/2)-(text_w/2)"
		$yCoor = $height-((($height/100)*(100-$yPadding)))	
	  }
	  "Top Right" {
		$xCoor = "w-(w*(${xPadding}/100))-text_w"
		$yCoor = $height-((($height/100)*(100-$yPadding)))
	  }
	  "Middle Left" {
		$xCoor = "w-((w/100)*(100-${xPadding}))"
		$yCoor = $height-($height/2)-($fontsize/2)
	  }
	  "Middle" {
		$xCoor = "w-(w/2)-(text_w/2)"
		$yCoor = $height-($height/2)-($fontsize/2)
	  }
	  "Middle Right" {
		$xCoor = "w-(w*(${xPadding}/100))-text_w"
		$yCoor = $height-($height/2)-($fontsize/2)
	  }
	  "Bottom Left" {
		$xCoor = "w-((w/100)*(100-${xPadding}))"
		$yCoor = $height-($height*($yPadding/100))-$fontsize
	  }
	  "Bottom Middle" {
		$xCoor = "w-(w/2)-(text_w/2)"
		$yCoor = $height-($height*($yPadding/100))-$fontsize
	  }
	  "Bottom Right" {
		$xCoor = "w-(w*(${xPadding}/100))-text_w"
		$yCoor = $height-($height*($yPadding/100))-$fontsize
	  }
	  default {
		$xCoor = "w-(w*(${xPadding}/100))-text_w"
		$yCoor = $height-($height*($yPadding/100))-$fontsize
	  }
	}

	# add to drawcalls
	$c = $drawcalls[$position]["Count"]
	$drawprops = @{
		"fontfile" = $fontfile
		"fontsize" = $fontsize
		"fontcolor" = $fontcolor
		"borderw" = $borderw
		"xCoor" = $xCoor
		"yCoor" = $yCoor
		"text" = $text
	}
	if($($userInput.Enable)){
		$drawcalls[$position]["drawtexts"] += $drawprops
		$drawcalls[$position]["Count"] += 1
	}

	# Process Time Stamp

	$font = $fontList | Where-Object { $_.Name -eq $($userInput.TimeFont) }
	Write-Verbose "Timestamp Settings-"
	Write-Verbose "Selected Font: $font.Name"
	$fontfile = $font.FileName
	$fontfile = "C\:/Windows/Fonts/" + $fontfile
	Write-Verbose "Font File: $fontfile"
	$fontsize = $($userInput.TimeSize)
	Write-Verbose "Font Size: $fontsize"
	$fontcolor = $($userInput.TimeColor)
	Write-Verbose "Font Color: $fontcolor"
	$borderColor = $($userInput.TimeBorderColor)
	if($borderColor -ne "none"){
	  Write-Verbose "Border Color: $borderColor"
	  $borderSize = $($userInput.TimeBorderSize)
	  Write-Verbose "Border Size: $borderSize"
	  $borderw = "borderw=${borderSize}:bordercolor=${borderColor}:"
	} else {
	  $borderw = ""
	}
	$position = $($userInput.TimePosition)

	if($position -ne "Single Line"){
		Write-Verbose "Text Position: $position"
		$xPadding = $($userInput.TimeXPadding)
		$yPadding = $($userInput.TimeYPadding)

		$escapedText = Escape-FFmpegOption -UnescapedString (Convert-TimeFormat -FormatString $timeFormat)
		$text = "%{pts\:gmtime\:" + $epoch + "\:" + $escapedText + "}"

		switch($position) {
		  "Top Left" { 
			$xCoor = "w-((w/100)*(100-${xPadding}))"
			$yCoor = $height-((($height/100)*(100-$yPadding)))
		  }
		  "Top Middle" {
			$xCoor = "w-(w/2)-(text_w/2)"
			$yCoor = $height-((($height/100)*(100-$yPadding)))	
		  }
		  "Top Right" {
			$xCoor = "w-(w*(${xPadding}/100))-text_w"
			$yCoor = $height-((($height/100)*(100-$yPadding)))
		  }
		  "Middle Left" {
			$xCoor = "w-((w/100)*(100-${xPadding}))"
			$yCoor = $height-($height/2)-($fontsize/2)
		  }
		  "Middle" {
			$xCoor = "w-(w/2)-(text_w/2)"
			$yCoor = $height-($height/2)-($fontsize/2)
		  }
		  "Middle Right" {
			$xCoor = "w-(w*(${xPadding}/100))-text_w"
			$yCoor = $height-($height/2)-($fontsize/2)
		  }
		  "Bottom Left" {
			$xCoor = "w-((w/100)*(100-${xPadding}))"
			$yCoor = $height-($height*($yPadding/100))-$fontsize
		  }
		  "Bottom Middle" {
			$xCoor = "w-(w/2)-(text_w/2)"
			$yCoor = $height-($height*($yPadding/100))-$fontsize
		  }
		  "Bottom Right" {
			$xCoor = "w-(w*(${xPadding}/100))-text_w"
			$yCoor = $height-($height*($yPadding/100))-$fontsize
		  }
		  default {
			$xCoor = "w-(w*(${xPadding}/100))-text_w"
			$yCoor = $height-($height*($yPadding/100))-$fontsize
		  }
		}

		# add to drawcalls
		$c = $drawcalls[$position]["Count"]
		$drawprops = @{
			"fontfile" = $fontfile
			"fontsize" = $fontsize
			"fontcolor" = $fontcolor
			"borderw" = $borderw
			"xCoor" = $xCoor
			"yCoor" = $yCoor
			"text" = $text
		}

		if($($userInput.EnableTime)){
			$drawcalls[$position]["drawtexts"] += $drawprops
			$drawcalls[$position]["Count"] += 1
		}
	}

	# Build drawtext commands
	$drawtext = ""

	$positions | ForEach-Object { 

		$pos = $_
		$c = $drawcalls[$pos]["Count"]
		Write-Verbose "$c number of texts at position $pos"

		for ($i = 0; $i -lt $c; $i++) {
			$currText = $drawcalls[$pos]["drawtexts"][$i]
			$adjusted = $false

			# Check against all previous texts in the same position
			do {
				$overlapFound = $false
				for ($j = 0; $j -lt $i; $j++) {
					$prevText = $drawcalls[$pos]["drawtexts"][$j]

					# Check for overlap with previous text
					Write-Verbose "Checking for text overlap. Text #$i yCoor: ${$currText.yCoor}. Text #$j yCoor: $($prevText.yCoor)"
					if ([math]::abs($currText.yCoor - $prevText.yCoor) -lt ($currText.fontsize + 10)) {
						if ($pos -match "Bottom") {
							# Move the previous text up if position contains "Bottom"
							$prevText.yCoor -= ($currText.fontsize + 10) - [math]::abs($currText.yCoor - $prevText.yCoor)
							Write-Verbose "Adjusted previous text #$j at $pos position upward. New yCoor: $($prevText.yCoor)"
						} else {
							# Move the current text down otherwise
							$currText.yCoor += ($currText.fontsize + 10) - [math]::abs($currText.yCoor - $prevText.yCoor)
							Write-Verbose "Adjusted current text #$i at $pos position downward. New yCoor: $($currText.yCoor)"
						}
						$overlapFound = $true
						$adjusted = $true
					}
				}
			} while ($overlapFound) # Repeat until no overlaps are found

			if ($adjusted) {
				Write-Verbose "Adjustments made for text #$i at $pos position."
			} else {
				Write-Verbose "No adjustment needed for text #$i at $pos position."
			}
		}
	}
	$positions | ForEach-Object { 

		$pos = $_
		$c = $drawcalls[$pos]["Count"]
		Write-Verbose "$c number of texts at position $pos"

		for ($i = 0; $i -lt $c; $i++) {
			$currText = $drawcalls[$pos]["drawtexts"][$i]
			if($drawtext){
				$drawtext = $drawtext + "," + "drawtext=fontfile='" + $currText.fontfile + "':fontsize=" + $currText.fontsize + ":fontcolor=" + $currText.fontcolor + ":" + $currText.borderw + "x=" + $currText.xCoor + ":y=" + $currText.yCoor + ":text='" + $currText.text + "'"
			} else {
				$drawtext = "drawtext=fontfile='" + $currText.fontfile + "':fontsize=" + $currText.fontsize + ":fontcolor=" + $currText.fontcolor + ":" + $currText.borderw + "x=" + $currText.xCoor + ":y=" + $currText.yCoor + ":text='" + $currText.text + "'"
			}
		}
	}
	return $drawtext
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

if($Verbose){ 
	Write-Output "Press any key to begin..."
	Write-Output ""

	# Pause for any key press
	$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# Continue with the rest of the script
Write-Verbose "Starting the Video Stamper script..."

Write-Output "Checking for FFMPEG..."

# Path to ffplay
$ffplay = ".\ffplay.exe"
$ffplayDir = (Get-Location).Path
# Path to ffprobe
$ffprobe = ".\ffprobe.exe"
$ffprobeDir = (Get-Location).Path
# Path to ffmpeg
$ffmpeg = ".\ffmpeg.exe"
$ffmpegDir = (Get-Location).Path

if(-Not (Test-Path $ffplay)) {
 $validPath = $false

 if(Test-Path $iniFilePath) {
    $ffplayDir = Get-IniValue -filePath $iniFilePath -section "FFMPEG" -key "ffplay_path"
    Write-Verbose "stored path: $ffplayDir"
    if($ffplayDir) {
       $ffplay = Join-Path $ffplayDir "ffplay.exe"
    }
 } 
 
 if(Test-Path $ffplay) {
    $validPath = $true;
 }

 while(-Not $validPath){
    Write-Output "ffplay.exe not found. Press any key to browse for ffplay.exe..."
    # Pause for any key press
    $x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

    # Create and configure the OpenFileDialog
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.InitialDirectory = $ffplayDir  # Default folder (optional)
    $openFileDialog.Filter = "Executable Files (*.exe)|*.exe|All Files (*.*)|*.*" # Only executable files
    $openFileDialog.Title = "Locate ffplay.exe"

    # Show the dialog and get the selected file path
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $ffplay = $openFileDialog.FileName
        $ffplayItem = (Get-Item $ffplay)
        $ffplayName = $ffplayItem.Name
        $ffplayDir = $ffplayItem.DirectoryName
        Write-Output "Current directiory: $ffplayDir"
    } else {
        Write-Error "ffplay.exe was not found, and no file was selected. Exiting"
        exit;
    }
    # Check if the selected file is ffprobe.exe
    if ((Get-Item $ffplay).Name -eq "ffplay.exe") {
        # Save the valid path to the .ini file
        Set-IniValue -filePath $iniFilePath -section "FFMPEG" -key "ffplay_path" -value $ffplayDir
        $validPath = $true
    }
  }
  if($Verbose){ 
	Write-Output "ffplay.exe found in the current directory. Press any key to continue..."
	# Pause for any key press
	$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
  }
} else {
  Write-Verbose "ffplay.exe found in the current directory"
}

if(-Not (Test-Path $ffprobe)) {
 $validPath = $false
 if(Test-Path $iniFilePath) {
    $ffprobeDir = Get-IniValue -filePath $iniFilePath -section "FFMPEG" -key "ffprobe_path"
    Write-Verbose "stored path: $ffprobeDir"
    if($ffprobeDir) {
      $ffprobe = Join-Path $ffprobeDir "ffprobe.exe"
    } else {
		$ffprobeDir = $ffplayDir
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
  if($Verbose){ 
	Write-Output "ffprobe.exe found in the current directory. Press any key to continue..."
	# Pause for any key press
	$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
  }
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
  if($Verbose){ 
	Write-Output "ffmpeg.exe found in the current directory. Press any key to continue..."
	# Pause for any key press
	$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
  }
} else {
  Write-Verbose "ffmpeg.exe found in the current directory"
}

Write-Output "Press any key to select a video..."
# Pause for any key press
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

# Create and configure the OpenFileDialog
$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
#$openFileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")  # Default folder (optional)
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

# Get the width and height
$width = $jsonObject.streams[0].width
$height = $jsonObject.streams[0].height

#get default preview sizeInput
if(Test-Path $iniFilePath) {
    [int]$previewMaxWidth = Get-IniValue -filePath $iniFilePath -section "Settings" -key "preview_width"
	if($previewMaxWidth){
		Write-Verbose "stored preview width: $previewMaxWidth"
	} else {
		$previewMaxWidth = 800
		Set-IniValue -filePath $iniFilePath -section "Settings" -key "preview_width" -value $previewMaxWidth
	}
}
if(Test-Path $iniFilePath) {
    [int]$previewMaxHeight = Get-IniValue -filePath $iniFilePath -section "Settings" -key "preview_height"
	if($previewMaxHeight){
		Write-Verbose "stored preview height: $previewMaxHeight"
	} else {
		$previewMaxHeight = 600
		Set-IniValue -filePath $iniFilePath -section "Settings" -key "preview_height" -value $previewMaxHeight
	}
}

$previewWidth = $previewMaxWidth
$previewHeight = $previewMaxHeight

Write-Verbose "Video Width: $width"
Write-Verbose "Video Height: $height"

# Create the form
$form = New-Object System.Windows.Forms.Form
# form settings
$rows = 11
$columns = 14

$monthNum = [int](Get-Date -Date $CreationDate -UFormat "%m")
$dayNum = [int](Get-Date -Date $CreationDate -UFormat "%d")
$yearNum = [int](Get-Date -Date $CreationDate -UFormat "%Y")

$hourNum = [int](Get-Date -Date $CreationDate -UFormat "%I")
$minuteNum = [int](Get-Date -Date $CreationDate -UFormat "%M")
$secondNum = [int](Get-Date -Date $CreationDate -UFormat "%S")
$seconds = [int](Get-Date -Date $CreationDate -UFormat "%s")
$ampm = [string](Get-Date -Date $CreationDate -UFormat "%p")

#Build Form
$form.Text = "Video Stamper Settings"
$form.Width = (50 * $columns) + 20
$form.Height = (30 * $rows)
$form.StartPosition = "CenterScreen"

# Create the TableLayoutPanel
$tableLayout = New-Object System.Windows.Forms.TableLayoutPanel
$tableLayout.RowCount = $rows
$tableLayout.ColumnCount = $columns
$tableLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
$tableLayout.AutoSize = $true
$tableLayout.CellBorderStyle = [System.Windows.Forms.TableLayoutPanelCellBorderStyle]::None
$form.Controls.Add($tableLayout)

# Set fixed column widths
$i = 0
while( $i -lt $columns ){ 
  $tableLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 50))) | Out-Null
  $i++
}

# Set row styles
$i = 0
while( $i -lt $rows) {
	$tableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null 
	$i++
}

# Create Labels
$dateStamp = New-Object System.Windows.Forms.Label
$dateStamp.Text = "Datestamp"
$dateStamp.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
$tableLayout.SetColumnSpan($dateStamp, $columns)  # Span the label across 5 columns

$enableLabel = New-Object System.Windows.Forms.Label
$enableLabel.Text = "Enable"
$enableLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft

$fontLabel = New-Object System.Windows.Forms.Label
$fontLabel.Text = "Font"
$fontLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
$tableLayout.SetColumnSpan($fontLabel, 5)  # Span the label across 5 columns

$sizeLabel = New-Object System.Windows.Forms.Label
$sizeLabel.Text = "Size"
$sizeLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft

$colorLabel = New-Object System.Windows.Forms.Label
$colorLabel.Text = "Color"
$colorLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
$tableLayout.SetColumnSpan($colorLabel, 2)  # Span the label across 2 columns

$borderLabel = New-Object System.Windows.Forms.Label
$borderLabel.Text = "Border Color"
$borderLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
$tableLayout.SetColumnSpan($borderLabel, 2)  # Span the label across 2 columns

$borderSizeLabel = New-Object System.Windows.Forms.Label
$borderSizeLabel.Text = "Size"
$borderSizeLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft

$positionLabel = New-Object System.Windows.Forms.Label
$positionLabel.Text = "Position"
$positionLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
$tableLayout.SetColumnSpan($positionLabel, 2)  # Span the label across 2 columns

# Add Labels to Row 0
Add-Control -panel $tableLayout -control $dateStamp -row 0 -column 0
Add-Control -panel $tableLayout -control $enableLabel -row 1 -column 0
Add-Control -panel $tableLayout -control $fontLabel -row 1 -column 1
Add-Control -panel $tableLayout -control $sizeLabel -row 1 -column 6
Add-Control -panel $tableLayout -control $colorLabel -row 1 -column 7
Add-Control -panel $tableLayout -control $borderLabel -row 1 -column 9
Add-Control -panel $tableLayout -control $borderSizeLabel -row 1 -column 11
Add-Control -panel $tableLayout -control $positionLabel -row 1 -column 12

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
$enableCheckbox = New-Object System.Windows.Forms.CheckBox
$enableCheckbox.Width = 50
$enableCheckbox.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter

$fontDropdown = New-Object System.Windows.Forms.ComboBox
$fontDropdown.Width = 250
$fontDropdown.DropDownStyle = "DropDownList"
$fontList | ForEach-Object {
	$fontDropdown.Items.Add($_.Name)
} | Out-Null
$fontDropdown.SelectedIndex = 0
$tableLayout.SetColumnSpan($fontDropdown, 5)  # Span the input across 2 columns

$sizeInput = New-Object System.Windows.Forms.NumericUpDown
$sizeInput.Width = 50
$sizeInput.Minimum = 2
$sizeInput.Maximum = 256
$sizeInput.Value = 36
$sizeInput.Increment = 2

$colorDropdown = New-Object System.Windows.Forms.ComboBox
$colorDropdown.Width = 100
$colorDropdown.DropDownStyle = "DropDownList"
$ffmpegColors = @("black", "white", "red", "green", "blue", "yellow", "magenta", "cyan", "gray", "darkgray")
$ffmpegColors | ForEach-Object { $colorDropdown.Items.Add($_) } | Out-Null
$colorDropdown.SelectedIndex = 0
$tableLayout.SetColumnSpan($colorDropdown, 2)  # Span the input across 2 columns

$borderDropdown = New-Object System.Windows.Forms.ComboBox
$borderDropdown.Width = 100
$borderDropdown.DropDownStyle = "DropDownList"
$ffmpegBorderColors = @("none", "black", "white", "red", "green", "blue", "yellow", "magenta", "cyan", "gray", "darkgray")
$ffmpegBorderColors | ForEach-Object { $borderDropdown.Items.Add($_) } | Out-Null
$borderDropdown.SelectedIndex = 0
$tableLayout.SetColumnSpan($borderDropdown, 2)  # Span the input across 2 columns

$borderSizeInput = New-Object System.Windows.Forms.NumericUpDown
$borderSizeInput.Width = 50
$borderSizeInput.Minimum = 1
$borderSizeInput.Maximum = 10
$borderSizeInput.Value = 1
$borderSizeInput.Increment = 1

$positionDropdown = New-Object System.Windows.Forms.ComboBox
$positionDropdown.Width = 100
$positionDropdown.DropDownStyle = "DropDownList"
$positions = @("Top Left", "Top Middle", "Top Right", "Middle Left", "Middle", "Middle Right", "Bottom Left", "Bottom Middle", "Bottom Right")
$positions | ForEach-Object { $positionDropdown.Items.Add($_) } | Out-Null
# Determine orientation
if ($height -gt $width) {
	# Portrait
	Write-Output "Video is Portrait Orientation"
	$positionDropdown.SelectedIndex = 7
} else {
	# Landscape
	Write-Output "Video is Landscape Orientation"
	$positionDropdown.SelectedIndex = 8
}
$tableLayout.SetColumnSpan($positionDropdown, 2)  # Span the input across 2 columns

# Add Inputs to Row 2
Add-Control -panel $tableLayout -control $enableCheckbox -row 2 -column 0
Add-Control -panel $tableLayout -control $fontDropdown -row 2 -column 1
Add-Control -panel $tableLayout -control $sizeInput -row 2 -column 6
Add-Control -panel $tableLayout -control $colorDropdown -row 2 -column 7
Add-Control -panel $tableLayout -control $borderDropdown -row 2 -column 9
Add-Control -panel $tableLayout -control $borderSizeInput -row 2 -column 11
Add-Control -panel $tableLayout -control $positionDropdown -row 2 -column 12

# Datestamp Labels
# Month Label
$monthLabel = New-Object System.Windows.Forms.Label
$monthLabel.Text = "Month"
$monthLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
$tableLayout.SetColumnSpan($monthLabel, 2)  # Span the label across 2 columns

# Day Label
$dayLabel = New-Object System.Windows.Forms.Label
$dayLabel.Text = "Day"
$dayLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
$tableLayout.SetColumnSpan($dayLabel, 2)  # Span the label across 2 columns

#Year Label
$yearLabel = New-Object System.Windows.Forms.Label
$yearLabel.Text = "Year"
$yearLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
$tableLayout.SetColumnSpan($yearLabel, 2)  # Span the label across 2 columns

#DateFormatString Label
$dateFormatLabel = New-Object System.Windows.Forms.Label
$dateFormatLabel.Text = "Date Format String"
$dateFormatLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
$tableLayout.SetColumnSpan($dateFormatLabel, 4)  # Span the label across 2 columns

#XPadding Label
$xpadLabel = New-Object System.Windows.Forms.Label
$xpadLabel.Text = "XPad%"
$xpadLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft

#YPadding Label
$ypadLabel = New-Object System.Windows.Forms.Label
$ypadLabel.Text = "YPad%"
$ypadLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft

# Add Labels to Row 3
Add-Control -panel $tableLayout -control $monthLabel -row 3 -column 0
Add-Control -panel $tableLayout -control $dayLabel -row 3 -column 2
Add-Control -panel $tableLayout -control $yearLabel -row 3 -column 4
Add-Control -panel $tableLayout -control $dateFormatLabel -row 3 -column 8
Add-Control -panel $tableLayout -control $xpadLabel -row 3 -column 12
Add-Control -panel $tableLayout -control $ypadLabel -row 3 -column 13

# Datestamp Inputs
# Month Input
$monthInput = New-Object System.Windows.Forms.NumericUpDown
$monthInput.Width = 100
$monthInput.Minimum = 1
$monthInput.Maximum = 12
$monthInput.Value = $monthNum
$monthInput.Increment = 1
$tableLayout.SetColumnSpan($monthInput, 2)  # Span the input across 2 columns

# Day Input
$dayInput = New-Object System.Windows.Forms.NumericUpDown
$dayInput.Width = 100
$dayInput.Minimum = 1
$dayInput.Maximum = 31
$dayInput.Value = $dayNum
$dayInput.Increment = 1
$tableLayout.SetColumnSpan($dayInput, 2)  # Span the input across 2 columns

# Year Input
$yearInput = New-Object System.Windows.Forms.NumericUpDown
$yearInput.Width = 100
$yearInput.Minimum = 1
$yearInput.Maximum = 9999
$yearInput.Value = $yearNum
$yearInput.Increment = 1
$tableLayout.SetColumnSpan($yearInput, 2)  # Span the input across 2 columns

# Format Input
$dateFormatInput = New-Object System.Windows.Forms.TextBox
$dateFormatInput.Width = 200
$dateFormatInput.Text = "M/d/yyyy"
$tableLayout.SetColumnSpan($dateFormatInput, 4)  # Span the input across 2 columns

# XPad Input
$xpadInput = New-Object System.Windows.Forms.NumericUpDown
$xpadInput.Width = 100
$xpadInput.Minimum = 0
$xpadInput.Maximum = 100
$xpadInput.Value = 5
$xpadInput.Increment = 1

# YPad Input
$ypadInput = New-Object System.Windows.Forms.NumericUpDown
$ypadInput.Width = 100
$ypadInput.Minimum = 0
$ypadInput.Maximum = 100
$ypadInput.Value = 5
$ypadInput.Increment = 1

# Add Labels to Row 3
Add-Control -panel $tableLayout -control $monthInput -row 4 -column 0
Add-Control -panel $tableLayout -control $dayInput -row 4 -column 2
Add-Control -panel $tableLayout -control $yearInput -row 4 -column 4
Add-Control -panel $tableLayout -control $dateFormatInput -row 4 -column 8
Add-Control -panel $tableLayout -control $xpadInput -row 4 -column 12
Add-Control -panel $tableLayout -control $ypadInput -row 4 -column 13

# Timestamp options
# Create Labels
$timeStamp = New-Object System.Windows.Forms.Label
$timeStamp.Text = "Timestamp"
$timeStamp.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
$tableLayout.SetColumnSpan($timeStamp, $columns)  # Span the label across 5 columns

$enableTimeLabel = New-Object System.Windows.Forms.Label
$enableTimeLabel.Text = "Enable"
$enableTimeLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft

$fontTimeLabel = New-Object System.Windows.Forms.Label
$fontTimeLabel.Text = "Font"
$fontTimeLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
$tableLayout.SetColumnSpan($fontTimeLabel, 5)  # Span the label across 5 columns

$sizeTimeLabel = New-Object System.Windows.Forms.Label
$sizeTimeLabel.Text = "Size"
$sizeTimeLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft

$colorTimeLabel = New-Object System.Windows.Forms.Label
$colorTimeLabel.Text = "Color"
$colorTimeLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
$tableLayout.SetColumnSpan($colorTimeLabel, 2)  # Span the label across 2 columns

$borderTimeLabel = New-Object System.Windows.Forms.Label
$borderTimeLabel.Text = "Border Color"
$borderTimeLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
$tableLayout.SetColumnSpan($borderTimeLabel, 2)  # Span the label across 2 columns

$borderSizeTimeLabel = New-Object System.Windows.Forms.Label
$borderSizeTimeLabel.Text = "Size"
$borderSizeTimeLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft

$positionTimeLabel = New-Object System.Windows.Forms.Label
$positionTimeLabel.Text = "Position"
$positionTimeLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
$tableLayout.SetColumnSpan($positionTimeLabel, 2)  # Span the label across 2 columns

# Add Labels to Row 6
Add-Control -panel $tableLayout -control $timeStamp -row 5 -column 0
Add-Control -panel $tableLayout -control $enableTimeLabel -row 6 -column 0
Add-Control -panel $tableLayout -control $fontTimeLabel -row 6 -column 1
Add-Control -panel $tableLayout -control $sizeTimeLabel -row 6 -column 6
Add-Control -panel $tableLayout -control $colorTimeLabel -row 6 -column 7
Add-Control -panel $tableLayout -control $borderTimeLabel -row 6 -column 9
Add-Control -panel $tableLayout -control $borderSizeTimeLabel -row 6 -column 11
Add-Control -panel $tableLayout -control $positionTimeLabel -row 6 -column 12

# Create Inputs
$enableTimeCheckbox = New-Object System.Windows.Forms.CheckBox
$enableTimeCheckbox.Width = 50
$enableTimeCheckbox.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter

$fontTimeDropdown = New-Object System.Windows.Forms.ComboBox
$fontTimeDropdown.Width = 250
$fontTimeDropdown.DropDownStyle = "DropDownList"
$fontList | ForEach-Object {
	$fontTimeDropdown.Items.Add($_.Name)
} | Out-Null
$fontTimeDropdown.SelectedIndex = 0
$tableLayout.SetColumnSpan($fontTimeDropdown, 5)  # Span the input across 2 columns

$sizeTimeInput = New-Object System.Windows.Forms.NumericUpDown
$sizeTimeInput.Width = 50
$sizeTimeInput.Minimum = 2
$sizeTimeInput.Maximum = 256
$sizeTimeInput.Value = 36
$sizeTimeInput.Increment = 2

$colorTimeDropdown = New-Object System.Windows.Forms.ComboBox
$colorTimeDropdown.Width = 100
$colorTimeDropdown.DropDownStyle = "DropDownList"
$ffmpegColors | ForEach-Object { $colorTimeDropdown.Items.Add($_) } | Out-Null
$colorTimeDropdown.SelectedIndex = 0
$tableLayout.SetColumnSpan($colorTimeDropdown, 2)  # Span the input across 2 columns

$borderTimeDropdown = New-Object System.Windows.Forms.ComboBox
$borderTimeDropdown.Width = 100
$borderTimeDropdown.DropDownStyle = "DropDownList"
$ffmpegBorderColors | ForEach-Object { $borderTimeDropdown.Items.Add($_) } | Out-Null
$borderTimeDropdown.SelectedIndex = 0
$tableLayout.SetColumnSpan($borderTimeDropdown, 2)  # Span the input across 2 columns

$borderSizeTimeInput = New-Object System.Windows.Forms.NumericUpDown
$borderSizeTimeInput.Width = 50
$borderSizeTimeInput.Minimum = 1
$borderSizeTimeInput.Maximum = 10
$borderSizeTimeInput.Value = 1
$borderSizeTimeInput.Increment = 1

$positionTimeDropdown = New-Object System.Windows.Forms.ComboBox
$positionTimeDropdown.Width = 100
$positionTimeDropdown.DropDownStyle = "DropDownList"
$positionsTime = @("Single Line", "Top Left", "Top Middle", "Top Right", "Middle Left", "Middle", "Middle Right", "Bottom Left", "Bottom Middle", "Bottom Right")
$positionsTime | ForEach-Object { $positionTimeDropdown.Items.Add($_) } | Out-Null
$positionTimeDropdown.SelectedIndex = 0

$tableLayout.SetColumnSpan($positionTimeDropdown, 2)  # Span the input across 2 columns

# Add Inputs to Row 7
Add-Control -panel $tableLayout -control $enableTimeCheckbox -row 7 -column 0
Add-Control -panel $tableLayout -control $fontTimeDropdown -row 7 -column 1
Add-Control -panel $tableLayout -control $sizeTimeInput -row 7 -column 6
Add-Control -panel $tableLayout -control $colorTimeDropdown -row 7 -column 7
Add-Control -panel $tableLayout -control $borderTimeDropdown -row 7 -column 9
Add-Control -panel $tableLayout -control $borderSizeTimeInput -row 7 -column 11
Add-Control -panel $tableLayout -control $positionTimeDropdown -row 7 -column 12

# Timestamp Labels
# Hour Label
$hourLabel = New-Object System.Windows.Forms.Label
$hourLabel.Text = "Hour"
$hourLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
$tableLayout.SetColumnSpan($hourLabel, 2)  # Span the label across 2 columns

# Minute Label
$minuteLabel = New-Object System.Windows.Forms.Label
$minuteLabel.Text = "Minute"
$minuteLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
$tableLayout.SetColumnSpan($minuteLabel, 2)  # Span the label across 2 columns

#Second Label
$secondLabel = New-Object System.Windows.Forms.Label
$secondLabel.Text = "Second"
$secondLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
$tableLayout.SetColumnSpan($secondLabel, 2)  # Span the label across 2 columns

#ampm Label
$ampmLabel = New-Object System.Windows.Forms.Label
$ampmLabel.Text = "AM/PM"
$ampmLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
$tableLayout.SetColumnSpan($ampmLabel, 2)  # Span the label across 2 columns

#TimeFormatString Label
$timeFormatLabel = New-Object System.Windows.Forms.Label
$timeFormatLabel.Text = "Tim Format String"
$timeFormatLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
$tableLayout.SetColumnSpan($timeFormatLabel, 4)  # Span the label across 2 columns

#XPadding Label
$xpadLabel = New-Object System.Windows.Forms.Label
$xpadLabel.Text = "XPad%"
$xpadLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft

#YPadding Label
$ypadLabel = New-Object System.Windows.Forms.Label
$ypadLabel.Text = "YPad%"
$ypadLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft

# Add Labels to Row 8
Add-Control -panel $tableLayout -control $hourLabel -row 8 -column 0
Add-Control -panel $tableLayout -control $minuteLabel -row 8 -column 2
Add-Control -panel $tableLayout -control $secondLabel -row 8 -column 4
Add-Control -panel $tableLayout -control $ampmLabel -row 8 -column 6
Add-Control -panel $tableLayout -control $timeFormatLabel -row 8 -column 8
Add-Control -panel $tableLayout -control $xpadLabel -row 8 -column 12
Add-Control -panel $tableLayout -control $ypadLabel -row 8 -column 13

# Timestamp Inputs
# Hour Input
$hourInput = New-Object System.Windows.Forms.NumericUpDown
$hourInput.Width = 100
$hourInput.Minimum = 1
$hourInput.Maximum = 12
$hourInput.Value = $hourNum
$hourInput.Increment = 1
$tableLayout.SetColumnSpan($hourInput, 2)  # Span the input across 2 columns

# Minute Input
$minuteInput = New-Object System.Windows.Forms.NumericUpDown
$minuteInput.Width = 100
$minuteInput.Minimum = 0
$minuteInput.Maximum = 59
$minuteInput.Value = $minuteNum
$minuteInput.Increment = 1
$tableLayout.SetColumnSpan($minuteInput, 2)  # Span the input across 2 columns

# Second Input
$secondInput = New-Object System.Windows.Forms.NumericUpDown
$secondInput.Width = 100
$secondInput.Minimum = 1
$secondInput.Maximum = 59
$secondInput.Value = $secondNum
$secondInput.Increment = 1
$tableLayout.SetColumnSpan($secondInput, 2)  # Span the input across 2 columns

$ampmDropdown = New-Object System.Windows.Forms.ComboBox
$ampmDropdown.Width = 100
$ampmDropdown.DropDownStyle = "DropDownList"
$ampmOptions = @("AM", "PM")
$ampmOptions | ForEach-Object { $ampmDropdown.Items.Add($_) } | Out-Null
if($ampm -eq "AM") {
  $ampmDropdown.SelectedIndex = 0
} else {
  $ampmDropdown.SelectedIndex = 1
}
$tableLayout.SetColumnSpan($ampmDropdown, 2)  # Span the input across 2 columns

# Format Input
$timeFormatInput = New-Object System.Windows.Forms.TextBox
$timeFormatInput.Width = 200
$timeFormatInput.Text = "h:mm:ss tt"
$tableLayout.SetColumnSpan($timeFormatInput, 4)  # Span the input across 2 columns

# XPad Input
$xpadTimeInput = New-Object System.Windows.Forms.NumericUpDown
$xpadTimeInput.Width = 100
$xpadTimeInput.Minimum = 0
$xpadTimeInput.Maximum = 100
$xpadTimeInput.Value = 5
$xpadTimeInput.Increment = 1

# YPad Input
$ypadTimeInput = New-Object System.Windows.Forms.NumericUpDown
$ypadTimeInput.Width = 100
$ypadTimeInput.Minimum = 0
$ypadTimeInput.Maximum = 100
$ypadTimeInput.Value = 5
$ypadTimeInput.Increment = 1

# Add Labels to Row 9
Add-Control -panel $tableLayout -control $hourInput -row 9 -column 0
Add-Control -panel $tableLayout -control $minuteInput -row 9 -column 2
Add-Control -panel $tableLayout -control $secondInput -row 9 -column 4
Add-Control -panel $tableLayout -control $ampmDropdown -row 9 -column 6
Add-Control -panel $tableLayout -control $timeFormatInput -row 9 -column 8
Add-Control -panel $tableLayout -control $xpadTimeInput -row 9 -column 12
Add-Control -panel $tableLayout -control $ypadTimeInput -row 9 -column 13

# Create and Add OK Button to Row 10
$previewButton = New-Object System.Windows.Forms.Button
$previewButton.Text = "Preview"
$previewButton.Width = 100
$previewButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom
$tableLayout.SetColumnSpan($previewButton, 2)  # Span the button across all columns
Add-Control -panel $tableLayout -control $previewButton -row 10 -column 5

# Create and Add OK Button to Row 10
$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "Save"
$okButton.Width = 100
$okButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom
$tableLayout.SetColumnSpan($okButton, 2)  # Span the button across all columns
Add-Control -panel $tableLayout -control $okButton -row 10 -column 7

# Handle Preview Button Click
$previewButton.Add_Click({
	$form.Tag = @{
		Enable = $enableCheckbox.Checked
		Font = $fontDropdown.SelectedItem
		Size = $sizeInput.Value
		Color = $colorDropdown.SelectedItem
		BorderColor = $borderDropdown.SelectedItem
		BorderSize = $borderSizeInput.Value
		Position = $positionDropdown.SelectedItem
		Month = $monthInput.Value
		Day = $dayInput.Value
		Year = $yearInput.Value
		DateFormat = $dateFormatInput.Text
		XPadding = $xpadInput.Value
		YPadding = $ypadInput.Value
		EnableTime = $enableTimeCheckbox.Checked
		TimeFont = $fontTimeDropdown.SelectedItem
		TimeSize = $sizeTimeInput.Value
		TimeColor = $colorTimeDropdown.SelectedItem
		TimeBorderColor = $borderTimeDropdown.SelectedItem
		TimeBorderSize = $borderSizeTimeInput.Value
		TimePosition = $positionTimeDropdown.SelectedItem
		Hour = $hourInput.Value
		Minute = $minuteInput.Value
		Second = $secondInput.Value
		AMPM = $ampmDropdown.SelectedItem
		TimeFormat = $timeFormatInput.Text
		TimeXPadding = $xpadTimeInput.Value
		TimeYPadding = $ypadTimeInput.Value
	}
	$drawtext = Process-Video -form $form -height $height -width $width
	
	if($drawtext){
		if($previewMaxWidth -gt $width){ $previewWidth = $width }
		if($previewMaxHeight -gt $height){ $previewHeight = $height }
		if ($Verbose) {
			# Build the arguments array
			$args = @(
				"-autoexit",
				"-i", $inputFile,
				"-vf", $drawtext,
				"-x", $previewWidth,
				"-y", $previewHeight,
				"-preset", "ultrafast",
				"-f", "mp4"
			)

			# Write the command for debugging
			$cmdString = "$ffplay " + ($args -join " ")
			Write-Verbose "Command to be run: $cmdString"

			# Execute the FFmpeg command
			& $ffplay @args
		}
		 else {
			# Build the arguments array
			$args = @(
				"-autoexit",
				"-hide_banner",
				"-loglevel", "error",
				"-i", $inputFile,
				"-vf", $drawtext,
				"-x", $previewWidth,
				"-y", $previewHeight,
				"-preset", "ultrafast",
				"-f", "mp4"
			)
			# Execute the FFmpeg command
			& $ffplay @args
		}
	} else {
		Write-Host "No text to draw to video! Please make sure the time or date stamp is enabled or that there is valid text to write!"
	}
})

# Handle OK Button Click
$okButton.Add_Click({
	$form.Tag = @{
		Enable = $enableCheckbox.Checked
		Font = $fontDropdown.SelectedItem
		Size = $sizeInput.Value
		Color = $colorDropdown.SelectedItem
		BorderColor = $borderDropdown.SelectedItem
		BorderSize = $borderSizeInput.Value
		Position = $positionDropdown.SelectedItem
		Month = $monthInput.Value
		Day = $dayInput.Value
		Year = $yearInput.Value
		DateFormat = $dateFormatInput.Text
		XPadding = $xpadInput.Value
		YPadding = $ypadInput.Value
		EnableTime = $enableTimeCheckbox.Checked
		TimeFont = $fontTimeDropdown.SelectedItem
		TimeSize = $sizeTimeInput.Value
		TimeColor = $colorTimeDropdown.SelectedItem
		TimeBorderColor = $borderTimeDropdown.SelectedItem
		TimeBorderSize = $borderSizeTimeInput.Value
		TimePosition = $positionTimeDropdown.SelectedItem
		Hour = $hourInput.Value
		Minute = $minuteInput.Value
		Second = $secondInput.Value
		AMPM = $ampmDropdown.SelectedItem
		TimeFormat = $timeFormatInput.Text
		TimeXPadding = $xpadTimeInput.Value
		TimeYPadding = $ypadTimeInput.Value
	}
	$form.Close()
})

# Show the form
$form.ShowDialog() | Out-Null

# Process the results
if ($form.Tag -eq $null) {
 Write-Output "Operation canceled by the user."
 exit
}

$drawtext = Process-Video -form $form -height $height -width $width

if(-Not $drawtext){
	Write-Host "No text to draw to video! Operation canceled. Video has not been written"
	exit
}

# Create a SaveFileDialog object
$saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog

# Configure the SaveFileDialog
$saveFileDialog.Title = "Save Stamped Video As"
$saveFileDialog.Filter = "MP4 (*.mp4)|*.mp4|All Files (*.*)|*.*"
$saveFileDialog.FilterIndex = 1
$saveFileDialog.DefaultExt = "mp4"
$saveFileDialog.OverwritePrompt = $true
# Set the initial directory to the directory of the input file
if (-not [string]::IsNullOrEmpty($inputFile) -and (Test-Path $inputFile)) {
    $saveFileDialog.InitialDirectory = (Get-Item $inputFile).DirectoryName
} else {
    # Fallback to Desktop if $inputFile is not valid
    $saveFileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
}

# Show the dialog and check if the user selected a file
if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
  $output = $saveFileDialog.FileName
} else {
  Write-Output "Operation canceled by the user."
  exit
}

if ($Verbose) {
    # Build the arguments array
    $args = @(
        "-y",
        "-i", $inputFile,
        "-vf", $drawtext,
        "-crf", "27",
        "-movflags", "use_metadata_tags",
        "-map_metadata", "0",
        "-preset", "ultrafast",
        "-f", "mp4",
        $output
    )

    # Write the command for debugging
    $cmdString = "$ffmpeg " + ($args -join " ")
    Write-Verbose "Command to be run: $cmdString"

    # Execute the FFmpeg command
    & $ffmpeg @args
}
 else {
 # Build the arguments array
    $args = @(
		"-hide_banner",
		"-loglevel", "error",
        "-y",
        "-i", $inputFile,
        "-vf", $drawtext,
        "-crf", "27",
        "-movflags", "use_metadata_tags",
        "-map_metadata", "0",
        "-preset", "ultrafast",
        "-f", "mp4",
        $output
    )
 # Execute the FFmpeg command
    & $ffmpeg @args
}

Write-Host "Video has been Stamped!"