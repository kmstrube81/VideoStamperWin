# __      __ _     _           _____                           
# \ \    / /(_)   | |         /  ___\  _                       
#  \ \  / /  _  __| | ___  __ | (___  | |  __ _ _ __ ___  _ __   ___ _ __ 
#   \ \/ /  | |/ _  |/ _ \/  \\___  \[   ]/ _' | '_ ' _ \| '_ \ / _ \ '__|
#    \  /   | ||(_| || __/|()| ___) | | | |(_| | | | | | | |_) |  __/ |   
#     \/    |_|\___.|\___|\__/|_____/ |_| \___.|_| |_| |_| .__/ \___|_|   
#                                                        | |              
#                                                        |_|
# By Kasey M. Strube
# Version 0.4
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

# Helper function to remove controls to TableLayoutPanel
function Remove-Control {
	param (
		[System.Windows.Forms.TableLayoutPanel]$panel,
		[System.Windows.Forms.Control]$control
	)
	$panel.Controls.Remove($control)
}

# Function to process .NET form inputs into a valid ffmpeg command
function Process-Video {
	param(
		[System.Windows.Forms.Form]$form,
		[object[]]$texts,
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
	if($($userInput.Enable)){
		$between = ""
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
			"start" = -1
			"end" = -1
			"between" = $between
		}
	
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
		
		if($($userInput.EnableTime)){
			$between = ""
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
				"start" = -1
				"end" = 100000
				"between" = $between
			}

			$drawcalls[$position]["drawtexts"] += $drawprops
			$drawcalls[$position]["Count"] += 1
		}
	}
	
	# Process Texts to add
	foreach($text in $texts) {
		
		#TextFont = $textFont
		#	TextSize = $textSize
		#	TextColor = $textColor
		#	TextBorderColor = $textBorderColor
		#	TextBorderSize = $textBorderSize
		#	TextPosition = $textPosition
		#	Text = $textText = $textText.Substring(0, [math]::Min(64, $textText.Length))
		#	TextStart = $textStart
		#	TextDuration = $textDuration
		#	TextXPad = $textXPad
		#	TextYPad = $textYPad
		
		$font = $fontList | Where-Object { $_.Name -eq $text.TextFont }
		Write-Verbose "Text Settings-"
		Write-Verbose "Selected Font: $font.Name"
		$fontfile = $font.FileName
		$fontfile = "C\:/Windows/Fonts/" + $fontfile
		Write-Verbose "Font File: $fontfile"
		$fontsize = $text.TextSize
		Write-Verbose "Font Size: $fontsize"
		$fontcolor = $text.TextColor
		Write-Verbose "Font Color: $fontcolor"
		$borderColor = $text.TextBorderColor
		if($borderColor -ne "none"){
		  Write-Verbose "Border Color: $borderColor"
		  $borderSize = $text.TextBorderSize
		  Write-Verbose "Border Size: $borderSize"
		  $borderw = "borderw=${borderSize}:bordercolor=${borderColor}:"
		} else {
		  $borderw = ""
		}
		$position = $text.TextPosition
		Write-Verbose "Text Position: $position"
		$xPadding = $text.TextXPad
		$yPadding = $text.TextYPad
		
		$textStart = $text.TextStart
		$textDuration = $text.TextDuration

		$text = Escape-FFmpegOption -UnescapedString $text.Text

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
		$between = ":enable='between(t," + $textStart + "," + $textDuration + ")'"
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
			"start" = $textStart
			"end" = $textDuration
			"between" = $between
		}
		$drawcalls[$position]["drawtexts"] += $drawprops
		$drawcalls[$position]["Count"] += 1
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
					if (([math]::abs($currText.yCoor - $prevText.yCoor) -lt ($currText.fontsize + 10)) -And ($currText.start -eq -1 -Or $prevText.start -eq -1 -Or ($currText.start -ge $prevText.start -And $currText.start -le $prevText.end)) ) {
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
				$drawtext = $drawtext + "," + "drawtext=fontfile='" + $currText.fontfile + "':fontsize=" + $currText.fontsize + ":fontcolor=" + $currText.fontcolor + ":" + $currText.borderw + "x=" + $currText.xCoor + ":y=" + $currText.yCoor + ":text='" + $currText.text + "'" + $currText.between
			} else {
				$drawtext = "drawtext=fontfile='" + $currText.fontfile + "':fontsize=" + $currText.fontsize + ":fontcolor=" + $currText.fontcolor + ":" + $currText.borderw + "x=" + $currText.xCoor + ":y=" + $currText.yCoor + ":text='" + $currText.text + "'" + $currText.between
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
$openFileDialog.Multiselect = $true

# Show the dialog and get the selected file path
if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $inputFiles = $openFileDialog.FileNames
    Write-Verbose "Selected Files: $($inputFiles -join ', ')" 
} else {
    Write-Error "No file was selected."
    exit
}

#init addedText Array
$global:addedTextElements = @() 

$prevSettings = $null
$tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ('VideoStamper_' + [guid]::NewGuid().Guid)
New-Item -ItemType Directory -Path $tempDir | Out-Null
$stampedClips = @()
foreach ($inputFile in $inputFiles) {
    Write-Verbose "Processing $inputFile"
	# Initialize ffprobe command to extract metadata
	$ffprobeCmd = &$ffprobe -v quiet -print_format json -show_entries stream=width,height,duration -show_entries format_tags "$inputFile"

	# parse Json
	$jsonObject = $ffprobeCmd | ConvertFrom-Json

	Write-Verbose "Video Metadata:"
	if($Verbose){
		Write $jsonObject
	}

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
		# If no metadata but previous settings exist, reuse previous timestamp
		if ($prevSettings -ne $null) {
			$creationDate = $prevSettings.CreationDate
			Write-Verbose "No valid DateTime Metadata. Using previous Video DateTime: $creationDate"
		} else {
			$creationDate = Get-Date
			Write-Verbose "No valid DateTime Metadata. Setting Time Stamp to $creationDate"
		}
	} else {
		# Parse the date string into a DateTime object
		$creationDate = [datetime]::Parse($creationDateString, $null, [System.Globalization.DateTimeStyles]::AssumeUniversal)
		Write-Verbose "Parsed DateTime:"
		Write-Verbose $creationDate
	}

	# Get the width and height
	$width = $jsonObject.streams[0].width
	$height = $jsonObject.streams[0].height
	$metaDur = $jsonObject.streams[0].duration

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
	$addedTexts = 0

	$monthNum = [int](Get-Date -Date $CreationDate -UFormat "%m")
	$dayNum = [int](Get-Date -Date $CreationDate -UFormat "%d")
	$yearNum = [int](Get-Date -Date $CreationDate -UFormat "%Y")

	$hourNum = [int](Get-Date -Date $CreationDate -UFormat "%I")
	$minuteNum = [int](Get-Date -Date $CreationDate -UFormat "%M")
	$secondNum = [int](Get-Date -Date $CreationDate -UFormat "%S")
	$seconds = [int](Get-Date -Date $CreationDate -UFormat "%s")
	$ampm = [string](Get-Date -Date $CreationDate -UFormat "%p")

	$titleBarText = ([IO.Path]::GetFileName($inputFile))
	#Build Form
	$form.Text = "Video Stamper - ${titleBarText}"
$form.AutoSize = $true
$form.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$form.AutoScroll = $true
$form.add_SizeChanged({ $form.StartPosition = "CenterScreen" })
	$form.Width = (50 * $columns) + 20
# 	$form.Height = (30 * $rows)
	$form.StartPosition = "CenterScreen"

	# Create the TableLayoutPanel
	$tableLayout = New-Object System.Windows.Forms.TableLayoutPanel
	$tableLayout.RowCount = $rows
	$tableLayout.ColumnCount = $columns
for ($i = 0; $i -lt $columns; $i++) {
    $tableLayout.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Absolute, 50)) | Out-Null
} 
for ($i = 0; $i -lt $rows; $i++) {
    $tableLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 30)) | Out-Null
}
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

	#load added Text array
	$addedTextElements = @()

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
	if($prevSettings -ne $null){
		$enableCheckbox.Checked = $prevSettings.Enable
	}
	
	$fontDropdown = New-Object System.Windows.Forms.ComboBox
	$fontDropdown.Width = 250
	$fontDropdown.DropDownStyle = "DropDownList"
	$fontList | ForEach-Object {
		$fontDropdown.Items.Add($_.Name)
	} | Out-Null
	$tableLayout.SetColumnSpan($fontDropdown, 5)  # Span the input across 2 columns
	if($prevSettings -ne $null){
		$fontDropdown.SelectedIndex = $prevSettings.FontIndex
	} else {
		$fontDropdown.SelectedIndex = 0
	}
	
	$sizeInput = New-Object System.Windows.Forms.NumericUpDown
	$sizeInput.Width = 50
	$sizeInput.Minimum = 2
	$sizeInput.Maximum = 256
	$sizeInput.Increment = 2
	if($prevSettings -ne $null){
		$sizeInput.Value = $prevSettings.Size
	} else {
		$sizeInput.Value = 36
	}
	
	$colorDropdown = New-Object System.Windows.Forms.ComboBox
	$colorDropdown.Width = 100
	$colorDropdown.DropDownStyle = "DropDownList"
	$ffmpegColors = @("black", "white", "red", "green", "blue", "yellow", "magenta", "cyan", "gray", "darkgray")
	$ffmpegColors | ForEach-Object { $colorDropdown.Items.Add($_) } | Out-Null
	$tableLayout.SetColumnSpan($colorDropdown, 2)  # Span the input across 2 columns
	if($prevSettings -ne $null){
		$colorDropdown.SelectedIndex = $prevSettings.ColorIndex
	} else {
		$colorDropdown.SelectedIndex = 0
	}
	
	$borderDropdown = New-Object System.Windows.Forms.ComboBox
	$borderDropdown.Width = 100
	$borderDropdown.DropDownStyle = "DropDownList"
	$ffmpegBorderColors = @("none", "black", "white", "red", "green", "blue", "yellow", "magenta", "cyan", "gray", "darkgray")
	$ffmpegBorderColors | ForEach-Object { $borderDropdown.Items.Add($_) } | Out-Null
	$tableLayout.SetColumnSpan($borderDropdown, 2)  # Span the input across 2 columns
	if($prevSettings -ne $null){
		$borderDropdown.SelectedIndex = $prevSettings.BorderColorIndex
	} else {
		$borderDropdown.SelectedIndex = 0
	}

	$borderSizeInput = New-Object System.Windows.Forms.NumericUpDown
	$borderSizeInput.Width = 50
	$borderSizeInput.Minimum = 1
	$borderSizeInput.Maximum = 10
	$borderSizeInput.Increment = 1
	if($prevSettings -ne $null){
		$borderSizeInput.Value = $prevSettings.BorderSize
	} else {
		$borderSizeInput.Value = 1
	}

	$positionDropdown = New-Object System.Windows.Forms.ComboBox
	$positionDropdown.Width = 100
	$positionDropdown.DropDownStyle = "DropDownList"
	$positions = @("Top Left", "Top Middle", "Top Right", "Middle Left", "Middle", "Middle Right", "Bottom Left", "Bottom Middle", "Bottom Right")
	$positions | ForEach-Object { $positionDropdown.Items.Add($_) } | Out-Null
	if($prevSettings -ne $null){
		$positionDropdown.SelectedIndex = $prevSettings.PositionIndex
	} else {
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
	$dateFormatLabel.Text = "Date Format"
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
	if ($prevSettings -ne $null -and $prevSettings.DateFormat) {
		$dateFormatInput.Text = $prevSettings.DateFormat
	} else {
		$dateFormatInput.Text = "M/d/yyyy"
	}
	$tableLayout.SetColumnSpan($dateFormatInput, 4)  # Span the input across 2 columns

	# Create a ToolTip object
	$dateToolTip = New-Object System.Windows.Forms.ToolTip

	# Set tooltip properties
	#$dateToolTip.AutoPopDelay = 5000         # Tooltip remains visible for 5 seconds
	$dateToolTip.InitialDelay = 500          # Delay before showing tooltip (ms)
	$dateToolTip.ReshowDelay = 200           # Delay before showing again (ms)
	#$dateToolTip.ShowAlways = $true          # Always show the tooltip, even if form is inactive

	$dateToolTipText = @"
Use these rules to build your format string
yyyy = Year (4 digits)
yy = Year (2 digits)
MMMM = Full month name
MMM = Abbreviated month name
MM = Month (2 digits)
M = Month (no leading 0)
dddd = Full weekday name
ddd = Abbreviated weekday name
dd = Day of the month (2 digits)
d = Day of the month (no leading 0)
"@

	# Associate the tooltip with the button
	$dateToolTip.SetToolTip($dateFormatInput, $dateToolTipText)

	# XPad Input
	$xpadInput = New-Object System.Windows.Forms.NumericUpDown
	$xpadInput.Width = 100
	$xpadInput.Minimum = 0
	$xpadInput.Maximum = 100
	$xpadInput.Increment = 1
	if($prevSettings -ne $null){
		$xpadInput.Value = $prevSettings.XPadding
	} else {
		$xpadInput.Value = 5
	}
	
	# YPad Input
	$ypadInput = New-Object System.Windows.Forms.NumericUpDown
	$ypadInput.Width = 100
	$ypadInput.Minimum = 0
	$ypadInput.Maximum = 100
	$ypadInput.Increment = 1
	if($prevSettings -ne $null){
		$ypadInput.Value = $prevSettings.YPadding
	} else {
		$ypadInput.Value = 5
	}
	
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
	if($prevSettings -ne $null){
		$enableTimeCheckbox.Checked = $prevSettings.EnableTime
	}

	$fontTimeDropdown = New-Object System.Windows.Forms.ComboBox
	$fontTimeDropdown.Width = 250
	$fontTimeDropdown.DropDownStyle = "DropDownList"
	$fontList | ForEach-Object {
		$fontTimeDropdown.Items.Add($_.Name)
	} | Out-Null
	$tableLayout.SetColumnSpan($fontTimeDropdown, 5)  # Span the input across 2 columns
	if($prevSettings -ne $null){
		$fontTimeDropdown.SelectedIndex = $prevSettings.TimeFontIndex
	} else {
		$fontTimeDropdown.SelectedIndex = 0
	}
	
	$sizeTimeInput = New-Object System.Windows.Forms.NumericUpDown
	$sizeTimeInput.Width = 50
	$sizeTimeInput.Minimum = 2
	$sizeTimeInput.Maximum = 256
	$sizeTimeInput.Increment = 2
	if($prevSettings -ne $null){
		$sizeTimeInput.Value = $prevSettings.TimeSize
	} else {
		$sizeTimeInput.Value = 36
	}

	$colorTimeDropdown = New-Object System.Windows.Forms.ComboBox
	$colorTimeDropdown.Width = 100
	$colorTimeDropdown.DropDownStyle = "DropDownList"
	$ffmpegColors | ForEach-Object { $colorTimeDropdown.Items.Add($_) } | Out-Null
	$tableLayout.SetColumnSpan($colorTimeDropdown, 2)  # Span the input across 2 columns
	if($prevSettings -ne $null){
		$colorTimeDropdown.SelectedIndex = $prevSettings.TimeColorIndex
	} else {
		$colorTimeDropdown.SelectedIndex = 0
	}
	
	$borderTimeDropdown = New-Object System.Windows.Forms.ComboBox
	$borderTimeDropdown.Width = 100
	$borderTimeDropdown.DropDownStyle = "DropDownList"
	$ffmpegBorderColors | ForEach-Object { $borderTimeDropdown.Items.Add($_) } | Out-Null
	$tableLayout.SetColumnSpan($borderTimeDropdown, 2)  # Span the input across 2 columns
	if($prevSettings -ne $null){
		$borderTimeDropdown.SelectedIndex = $prevSettings.TimeBorderColorIndex
	} else {
		$borderTimeDropdown.SelectedIndex = 0
	}

	$borderSizeTimeInput = New-Object System.Windows.Forms.NumericUpDown
	$borderSizeTimeInput.Width = 50
	$borderSizeTimeInput.Minimum = 1
	$borderSizeTimeInput.Maximum = 10
	$borderSizeTimeInput.Increment = 1
	if($prevSettings -ne $null){
		$borderSizeTimeInput.Value = $prevSettings.TimeBorderSize
	} else {
		$borderSizeTimeInput.Value = 1
	}

	$positionTimeDropdown = New-Object System.Windows.Forms.ComboBox
	$positionTimeDropdown.Width = 100
	$positionTimeDropdown.DropDownStyle = "DropDownList"
	$positionsTime = @("Single Line", "Top Left", "Top Middle", "Top Right", "Middle Left", "Middle", "Middle Right", "Bottom Left", "Bottom Middle", "Bottom Right")
	$positionsTime | ForEach-Object { $positionTimeDropdown.Items.Add($_) } | Out-Null
	if($prevSettings -ne $null) {
		$positionTimeDropdown.SelectedIndex = $prevSettings.TimePositionIndex
	} else {
		$positionTimeDropdown.SelectedIndex = 0
	}

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
	$timeFormatLabel.Text = "Time Format"
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
	if ($prevSettings -ne $null -and $prevSettings.TimeFormat) {
		$timeFormatInput.Text = $prevSettings.TimeFormat
	} else {
		$timeFormatInput.Text = "h:mm:ss tt"
	}
	$tableLayout.SetColumnSpan($timeFormatInput, 4)  # Span the input across 2 columns

	# Create a ToolTip object
	$timeToolTip = New-Object System.Windows.Forms.ToolTip

	# Set tooltip properties
	#$timeToolTip.AutoPopDelay = 5000         # Tooltip remains visible for 5 seconds
	$timeToolTip.InitialDelay = 500          # Delay before showing tooltip (ms)
	$timeToolTip.ReshowDelay = 200           # Delay before showing again (ms)
	#$timeToolTip.ShowAlways = $true          # Always show the tooltip, even if form is inactive

	$timeToolTipText = @"
Use these rules to build your format string
HH = Hour (24-hour format)
H = Hour (24-hour format no leading 0)
hh = Hour (12-hour format)
h = Hour (12-hour format no leading 0)
mm = Minute
m = Minute (no leading 0)
ss = Second
s = Second (no leading 0)
tt = AM/PM
"@

	# Associate the tooltip with the button
	$timeToolTip.SetToolTip($timeFormatInput, $timeToolTipText)

	# XPad Input
	$xpadTimeInput = New-Object System.Windows.Forms.NumericUpDown
	$xpadTimeInput.Width = 100
	$xpadTimeInput.Minimum = 0
	$xpadTimeInput.Maximum = 100
	$xpadTimeInput.Increment = 1
	if($prevSettings -ne $null) {
		$xpadTimeInput.Value = $prevSettings.TimeXPadding
	} else {
		$xpadTimeInput.Value = 5
	}
	
	# YPad Input
	$ypadTimeInput = New-Object System.Windows.Forms.NumericUpDown
	$ypadTimeInput.Width = 100
	$ypadTimeInput.Minimum = 0
	$ypadTimeInput.Maximum = 100
	$ypadTimeInput.Increment = 1
	if($prevSettings -ne $null) {
		$ypadTimeInput.Value = $prevSettings.TimeYPadding
	} else {		
		$ypadTimeInput.Value = 5
	}
	
	# Add Labels to Row 9
	Add-Control -panel $tableLayout -control $hourInput -row 9 -column 0
	Add-Control -panel $tableLayout -control $minuteInput -row 9 -column 2
	Add-Control -panel $tableLayout -control $secondInput -row 9 -column 4
	Add-Control -panel $tableLayout -control $ampmDropdown -row 9 -column 6
	Add-Control -panel $tableLayout -control $timeFormatInput -row 9 -column 8
	Add-Control -panel $tableLayout -control $xpadTimeInput -row 9 -column 12
	Add-Control -panel $tableLayout -control $ypadTimeInput -row 9 -column 13
	
	# Create and Add preview Button to Row 10
	$addTextButton = New-Object System.Windows.Forms.Button
	$addTextButton.Text = "Add Text"
	$addTextButton.Width = 100
	$addTextButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom
	$tableLayout.SetColumnSpan($addTextButton, 2)  # Span the button across all columns
	
	$textNon = 0
	$bros = 10
	
	$global:newAddedTextElements = @()
	$global:addedTextElements | ForEach-Object {
	
		$textNon = $_.ID
		Write-Verbose "Adding Existing Text Stamp #$textNon"
		#Build form
		$_.Elements | ForEach-Object {
			$labelvalue = 	$_.Label
			$value = 		$_.Value
			$textvalue =	$_.Text
			$index = 		$_.SelectedIndex
			# Add Labels to Row 0,1
			switch($labelvalue){
				"Text Title"		{	$Label = New-Object System.Windows.Forms.Label
										$Label.Text = "Text #$textNon"
										Add-Member -InputObject $Label -MemberType NoteProperty -Name "Label" -Value "Text Title"
										$Label.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
										Add-Control -panel $tableLayout -control $Label -row $bros -column 0 }
				"Enable Label"		{ 	$TextEnableLabel = New-Object System.Windows.Forms.Label
										$TextEnableLabel.Text = "Enable"
										Add-Member -InputObject $TextEnableLabel -MemberType NoteProperty -Name "Label" -Value "Enable Label"
										$TextEnableLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
										Add-Control -panel $tableLayout -control $TextEnableLabel -row ($bros +1) -column 0 }
				"Font Label"		{	$TextFontLabel = New-Object System.Windows.Forms.Label
										$TextFontLabel.Text = "Font"
										Add-Member -InputObject $TextFontLabel -MemberType NoteProperty -Name "Label" -Value "Font Label"
										$TextFontLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
										$tableLayout.SetColumnSpan($TextFontLabel, 2)  # Span the input across 2 columns
										Add-Control -panel $tableLayout -control $TextFontLabel -row ($bros +1) -column 1 }
				"Size Label"		{	$TextSizeLabel = New-Object System.Windows.Forms.Label
										$TextSizeLabel.Text = "Size"
										Add-Member -InputObject $TextSizeLabel -MemberType NoteProperty -Name "Label" -Value "Size Label"
										$TextSizeLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
										Add-Control -panel $tableLayout -control $TextSizeLabel -row ($bros +1) -column 6 }
				"Color Label"		{	$TextColorLabel = New-Object System.Windows.Forms.Label
										$TextColorLabel.Text = "Color"
										Add-Member -InputObject $TextColorLabel -MemberType NoteProperty -Name "Label" -Value "Color Label"
										$TextColorLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
										$tableLayout.SetColumnSpan($TextColorLabel, 2)  # Span the input across 2 columns
										Add-Control -panel $tableLayout -control $TextColorLabel -row ($bros +1) -column 7 }
				"Border Color Label"{	$TextBorderLabel = New-Object System.Windows.Forms.Label
										$TextBorderLabel.Text = "Border Color"
										Add-Member -InputObject $TextBorderLabel -MemberType NoteProperty -Name "Label" -Value "Border Color Label"
										$TextBorderLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
										$tableLayout.SetColumnSpan($TextBorderLabel, 2)  # Span the input across 2 columns
										Add-Control -panel $tableLayout -control $TextBorderLabel -row ($bros +1) -column 9 }
				"Border Size Label"	{	$TextBorderSizeLabel = New-Object System.Windows.Forms.Label
										$TextBorderSizeLabel.Text = "Size"
										Add-Member -InputObject $TextBorderSizeLabel -MemberType NoteProperty -Name "Label" -Value "Border Size Label"
										$TextBorderSizeLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
										Add-Control -panel $tableLayout -control $TextBorderSizeLabel -row ($bros +1) -column 11 }
				"Position Label"	{	$TextPositionLabel = New-Object System.Windows.Forms.Label
										$TextPositionLabel.Text = "Position"
										Add-Member -InputObject $TextPositionLabel -MemberType NoteProperty -Name "Label" -Value "Position Label"
										$TextPositionLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
										$tableLayout.SetColumnSpan($TextPositionLabel, 2)  # Span the input across 2 columns
										Add-Control -panel $tableLayout -control $TextPositionLabel -row ($bros +1) -column 12 }
				# Add Inputs to Row 2
				"Delete Button"		{	$deleteButton = New-Object System.Windows.Forms.Button
										$deleteButton.Width = 50
										$deleteButton.Text = "x"
										Add-Member -InputObject $deleteButton -MemberType NoteProperty -Name "Label" -Value "Delete Button"
										$deleteButton.Font = New-Object System.Drawing.Font("Wingdings", 12)
										$deleteButton.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
										
										$deleteButton.Tag = @{
											ID = $textNon
										}
										Add-Control -panel $tableLayout -control $deleteButton -row ($bros + 2) -column 0
										$deleteButton.Add_Click({
											$brows = $bros
											$addTextButton.Tag.RowHeight = $brows
											$id = $this.Tag.ID
											
# 											$form.Height = (30 * $brows )
											$controlsToRemove = $global:addedTextElements | Where-Object { $_.ID -eq $id }
											$tableLayout.SuspendLayout()
											foreach ($control in $controlsToRemove.Elements) {
												Remove-Control -panel $tableLayout -control $control
												$control.Dispose()
											}
											$tableLayout.ResumeLayout()
											#Move Buttons
											#Add-Control -panel $tableLayout -control $addTextButton -row ($bros - 1) -column 4
											#Add-Control -panel $tableLayout -control $previewButton -row ($bros - 1) -column 6
											#Add-Control -panel $tableLayout -control $okButton -row ($bros - 1) -column 8
											# Remove the textbox from the global array
											$newElements = @()
											$global:addedTextElements | ForEach-Object { 
												if($_.ID -ne $id){
												$newElements += $_
												}
											}
											$global:addedTextElements = $newElements
										})}
				"Font Value"		{	$fontDropdown = New-Object System.Windows.Forms.ComboBox
										$fontDropdown.Width = 250
										$fontDropdown.DropDownStyle = "DropDownList"
										$fontList | ForEach-Object { $fontDropdown.Items.Add($_.Name) } | Out-Null
										$fontDropdown.SelectedIndex = $index
										$fontDropdown.Name = "Font SelectedItem"
										Add-Member -InputObject $fontDropdown -MemberType NoteProperty -Name "Label" -Value "Font Value"
										$tableLayout.SetColumnSpan($fontDropdown, 5)  # Span the input across 2 columns
										Add-Control -panel $tableLayout -control $fontDropdown -row ($bros + 2) -column 1 }
				"Size Value"		{	$sizeInput = New-Object System.Windows.Forms.NumericUpDown
										$sizeInput.Width = 50
										$sizeInput.Minimum = 2
										$sizeInput.Maximum = 256
										$sizeInput.Value = $value
										$sizeInput.Increment = 2
										$sizeInput.Name = "Size Value"
										Add-Member -InputObject $sizeInput -MemberType NoteProperty -Name "Label" -Value "Size Value"
										Add-Control -panel $tableLayout -control $sizeInput -row ($bros + 2) -column 6 }
				"Color Value"		{	$colorDropdown = New-Object System.Windows.Forms.ComboBox
										$colorDropdown.Width = 100
										$colorDropdown.DropDownStyle = "DropDownList"
										$ffmpegColors = @("black", "white", "red", "green", "blue", "yellow", "magenta", "cyan", "gray", "darkgray")
										$ffmpegColors | ForEach-Object { $colorDropdown.Items.Add($_) } | Out-Null
										$colorDropdown.SelectedIndex = $index
										$colorDropdown.Name = "Color SelectedItem"
										Add-Member -InputObject $colorDropdown -MemberType NoteProperty -Name "Label" -Value "Color Value"
										$tableLayout.SetColumnSpan($colorDropdown, 2)  # Span the input across 2 columns
										Add-Control -panel $tableLayout -control $colorDropdown -row ($bros + 2) -column 7 }
				"Border Color Value"{ 	$borderDropdown = New-Object System.Windows.Forms.ComboBox
										$borderDropdown.Width = 100
										$borderDropdown.DropDownStyle = "DropDownList"
										$ffmpegBorderColors = @("none", "black", "white", "red", "green", "blue", "yellow", "magenta", "cyan", "gray", "darkgray")
										$ffmpegBorderColors | ForEach-Object { $borderDropdown.Items.Add($_) } | Out-Null
										$borderDropdown.SelectedIndex = $index
										$borderDropdown.Name = "Border Color SelectedItem"
										Add-Member -InputObject $borderDropdown -MemberType NoteProperty -Name "Label" -Value "Border Color Value"
										$tableLayout.SetColumnSpan($borderDropdown, 2)  # Span the input across 2 columns
										Add-Control -panel $tableLayout -control $borderDropdown -row ($bros + 2) -column 9 }
				"Border Size Value"	{	$borderSizeInput = New-Object System.Windows.Forms.NumericUpDown
										$borderSizeInput.Width = 50
										$borderSizeInput.Minimum = 1
										$borderSizeInput.Maximum = 10
										$borderSizeInput.Value = $value
										$borderSizeInput.Increment = 1
										$borderSizeInput.Name = "Border Size Value"
										Add-Member -InputObject $borderSizeInput -MemberType NoteProperty -Name "Label" -Value "Border Size Value"
										Add-Control -panel $tableLayout -control $borderSizeInput -row ($bros + 2) -column 11 }
				"Position Value"	{ 	$positionDropdown = New-Object System.Windows.Forms.ComboBox
										$positionDropdown.Width = 100
										$positionDropdown.DropDownStyle = "DropDownList"
										$positions = @("Top Left", "Top Middle", "Top Right", "Middle Left", "Middle", "Middle Right", "Bottom Left", "Bottom Middle", "Bottom Right")
										$positions | ForEach-Object { $positionDropdown.Items.Add($_) } | Out-Null
										$positionDropdown.SelectedIndex = $index
										$positionDropdown.Name = "Position SelectedItem"
										Add-Member -InputObject $positionDropdown -MemberType NoteProperty -Name "Label" -Value "Position Value"
										$tableLayout.SetColumnSpan($positionDropdown, 2)  # Span the input across 2 columns
										Add-Control -panel $tableLayout -control $positionDropdown -row ($bros + 2) -column 12 }
				# Add Labels to next Row
				"Text Label"		{ 	$TextLabel = New-Object System.Windows.Forms.Label
										$TextLabel.Text = "Text"
										Add-Member -InputObject $TextLabel -MemberType NoteProperty -Name "Label" -Value "Text Label"
										$TextLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
										$tableLayout.SetColumnSpan($TextLabel,8)  # Span the label across 2 columns
										Add-Control -panel $tableLayout -control $TextLabel -row ($bros + 3) -column 0 }
				"Start Label"		{	$StartLabel = New-Object System.Windows.Forms.Label
										$StartLabel.Text = "Start (s)"
										Add-Member -InputObject $StartLabel -MemberType NoteProperty -Name "Label" -Value "Start Label"
										$StartLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
										$tableLayout.SetColumnSpan($StartLabel,2)  # Span the label across 2 columns
										Add-Control -panel $tableLayout -control $StartLabel -row ($bros + 3) -column 8 }
				"Duration Label"	{	$DurationLabel = New-Object System.Windows.Forms.Label
										$DurationLabel.Text = "Dur (s)"
										Add-Member -InputObject $DurationLabel -MemberType NoteProperty -Name "Label" -Value "Duration Label"
										$DurationLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
										$tableLayout.SetColumnSpan($DurationLabel,2)  # Span the label across 2 columns
										Add-Control -panel $tableLayout -control $DurationLabel -row ($bros + 3) -column 10 }
				"XPad Label"		{	$xpadLabel = New-Object System.Windows.Forms.Label
										$xpadLabel.Text = "XPad%"
										$xpadLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
										Add-Member -InputObject $xpadLabel -MemberType NoteProperty -Name "Label" -Value "XPad Label"
										Add-Control -panel $tableLayout -control $xpadLabel -row ($bros + 3) -column 12 }
				"YPad Label"		{ 	$ypadLabel = New-Object System.Windows.Forms.Label
										$ypadLabel.Text = "YPad%"
										$ypadLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
										Add-Member -InputObject $ypadLabel -MemberType NoteProperty -Name "Label" -Value "YPad Label"
										Add-Control -panel $tableLayout -control $ypadLabel -row ($bros + 3) -column 13 }
				# Add Inputs to next Row
				"Text Value"		{	$TextInput = New-Object System.Windows.Forms.TextBox
										$TextInput.Width = 500
										$TextInput.Name = "Text Value"
										$TextInput.Text = $textvalue
										Add-Member -InputObject $TextInput -MemberType NoteProperty -Name "Label" -Value "Text Value"
										$tableLayout.SetColumnSpan($TextInput, 7)  # Span the input across 2 columns
										Add-Control -panel $tableLayout -control $TextInput -row ($bros + 4) -column 0 }
				"Start Value"		{	$StartInput = New-Object System.Windows.Forms.NumericUpDown
										$StartInput.Width = 75
										$StartInput.Minimum = 0
										$StartInput.Maximum = 100
										$StartInput.Value = $value
										$StartInput.DecimalPlaces = 3
										$StartInput.Increment = 0.1
										$StartInput.Name = "Start Value"
										Add-Member -InputObject $StartInput -MemberType NoteProperty -Name "Label" -Value "Start Value"
										$tableLayout.SetColumnSpan($StartInput, 2)  # Span the input across 2 columns
										Add-Control -panel $tableLayout -control $StartInput -row ($bros + 4) -column 8 }
				"Duration Value"	{ 	$DurationInput = New-Object System.Windows.Forms.NumericUpDown
										$DurationInput.Width = 75
										$DurationInput.Minimum = 0
										$DurationInput.Maximum = 99999
										$DurationInput.Value = $value
										$DurationInput.DecimalPlaces = 3
										$DurationInput.Increment = 0.1
										$DurationInput.Name = "Duration Value"
										Add-Member -InputObject $DurationInput -MemberType NoteProperty -Name "Label" -Value "Duration Value"
										$tableLayout.SetColumnSpan($DurationInput, 2)  # Span the input across 2 columns
										Add-Control -panel $tableLayout -control $DurationInput -row ($bros + 4) -column 10 }
				"XPad Value"		{	$xpadInput = New-Object System.Windows.Forms.NumericUpDown
										$xpadInput.Width = 100
										$xpadInput.Minimum = 0
										$xpadInput.Maximum = 100
										$xpadInput.Value = $value
										$xpadInput.Increment = 1
										$xpadInput.Name = "XPad Value"
										Add-Member -InputObject $xpadInput -MemberType NoteProperty -Name "Label" -Value "XPad Value"
										Add-Control -panel $tableLayout -control $xpadInput -row ($bros + 4) -column 12 }
				"YPad Value"		{ 	$ypadInput = New-Object System.Windows.Forms.NumericUpDown
										$ypadInput.Width = 100
										$ypadInput.Minimum = 0
										$ypadInput.Maximum = 100
										$ypadInput.Value = $value
										$ypadInput.Increment = 1
										$ypadInput.Name = "YPad Value"
										Add-Member -InputObject $ypadInput -MemberType NoteProperty -Name "Label" -Value "YPad Value"
										Add-Control -panel $tableLayout -control $ypadInput -row ($bros + 4) -column 13 }
			}
		} | Out-Null
		$bros = $bros + 5
		$tableLayout.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Absolute, 30))
		#load objects into array
		$global:newAddedTextElements += [PSCustomObject]@{
			ID = $textNon
			Elements = @(
			$Label,
			$TextEnableLabel,
			$TextFontLabel,
			$TextSizeLabel,
			$TextColorLabel,
			$TextBorderLabel,
			$TextBorderSizeLabel,
			$TextPositionLabel,
			$deleteButton,
			$fontDropdown,
			$sizeInput,
			$colorDropdown,
			$borderDropdown,
			$borderSizeInput,
			$positionDropdown,
			$TextLabel,
			$StartLabel,
			$DurationLabel,
			$xpadLabel,
			$ypadLabel,
			$TextInput,
			$StartInput,
			$DurationInput,
			$xpadInput,
			$ypadInput
			)
		}
	} | Out-Null
	$global:addedTextElements = $global:newAddedTextElements
	Add-Control -panel $tableLayout -control $addTextButton -row $bros -column 4
	$textNo = $textNon + 1
	$addTextButton.Tag = @{
		Rows = ($bros + 1)
		RowHeight = ($bros + 1)
		TextNo = $textNo
	}

	# Create and Add preview Button to Row 10
	$previewButton = New-Object System.Windows.Forms.Button
	$previewButton.Text = "Preview"
	$previewButton.Width = 100
	$previewButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom
	$tableLayout.SetColumnSpan($previewButton, 2)  # Span the button across all columns
	Add-Control -panel $tableLayout -control $previewButton -row $bros -column 6

	# Create and Add OK Button to Row 10
	$okButton = New-Object System.Windows.Forms.Button
	$okButton.Text = "Save"
	$okButton.Width = 100
	$okButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom
	$tableLayout.SetColumnSpan($okButton, 2)  # Span the button across all columns
	Add-Control -panel $tableLayout -control $okButton -row $bros -column 8
	Write-Verbose "Number of Rows is: $bros"
# 	$form.Height = (30 * ($bros + 1) )

	# Handle Add Text Button Click
	$addTextButton.Add_Click({
		if($global:addedTextElements.Length -lt 3){
			$tableLayout.suspendLayout()
			$bros = $addTextButton.Tag.Rows
			$brows = $addTextButton.Tag.RowHeight
			$textNon = $addTextButton.Tag.TextNo
			# Create Labels
			$Label = New-Object System.Windows.Forms.Label
			$Label.Text = "Text #$textNon"
			Add-Member -InputObject $Label -MemberType NoteProperty -Name "Label" -Value "Text Title"
			$Label.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft

			$TextEnableLabel = New-Object System.Windows.Forms.Label
			$TextEnableLabel.Text = "Enable"
			Add-Member -InputObject $TextEnableLabel -MemberType NoteProperty -Name "Label" -Value "Enable Label"
			$TextEnableLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft

			$TextFontLabel = New-Object System.Windows.Forms.Label
			$TextFontLabel.Text = "Font"
			Add-Member -InputObject $TextFontLabel -MemberType NoteProperty -Name "Label" -Value "Font Label"
			$TextFontLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft

			$TextSizeLabel = New-Object System.Windows.Forms.Label
			$TextSizeLabel.Text = "Size"
			Add-Member -InputObject $TextSizeLabel -MemberType NoteProperty -Name "Label" -Value "Size Label"
			$TextSizeLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft

			$TextColorLabel = New-Object System.Windows.Forms.Label
			$TextColorLabel.Text = "Color"
			Add-Member -InputObject $TextColorLabel -MemberType NoteProperty -Name "Label" -Value "Color Label"
			$TextColorLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft

			$TextBorderLabel = New-Object System.Windows.Forms.Label
			$TextBorderLabel.Text = "Border Color"
			Add-Member -InputObject $TextBorderLabel -MemberType NoteProperty -Name "Label" -Value "Border Color Label"
			$TextBorderLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
			
			$TextBorderSizeLabel = New-Object System.Windows.Forms.Label
			$TextBorderSizeLabel.Text = "Size"
			Add-Member -InputObject $TextBorderSizeLabel -MemberType NoteProperty -Name "Label" -Value "Border Size Label"
			$TextBorderSizeLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft

			$TextPositionLabel = New-Object System.Windows.Forms.Label
			$TextPositionLabel.Text = "Position"
			Add-Member -InputObject $TextPositionLabel -MemberType NoteProperty -Name "Label" -Value "Position Label"
			$TextPositionLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft		

			$tableLayout.SetColumnSpan($Label, $columns)  # Span the label across 5 columns
			$tableLayout.SetColumnSpan($TextFontLabel, 5)  # Span the label across 5 columns
			$tableLayout.SetColumnSpan($TextColorLabel, 2)  # Span the label across 2 columns
			$tableLayout.SetColumnSpan($TextBorderLabel, 2)  # Span the label across 2 columns
			$tableLayout.SetColumnSpan($TextPositionLabel, 2)  # Span the label across 2 columns

			# Create Inputs
			$deleteButton = New-Object System.Windows.Forms.Button
			$deleteButton.Width = 50
			$deleteButton.Text = "x"
			Add-Member -InputObject $deleteButton -MemberType NoteProperty -Name "Label" -Value "Delete Button"
			$deleteButton.Font = New-Object System.Drawing.Font("Wingdings", 12)
			$deleteButton.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
			
			$deleteButton.Tag = @{
				ID = $textNon
			}

			$fontDropdown = New-Object System.Windows.Forms.ComboBox
			$fontDropdown.Width = 250
			$fontDropdown.DropDownStyle = "DropDownList"
			$fontList | ForEach-Object {
				$fontDropdown.Items.Add($_.Name)
			} | Out-Null
			$fontDropdown.SelectedIndex = 0
			$fontDropdown.Name = "Font SelectedItem"
			Add-Member -InputObject $fontDropdown -MemberType NoteProperty -Name "Label" -Value "Font Value"
			$tableLayout.SetColumnSpan($fontDropdown, 5)  # Span the input across 2 columns

			$sizeInput = New-Object System.Windows.Forms.NumericUpDown
			$sizeInput.Width = 50
			$sizeInput.Minimum = 2
			$sizeInput.Maximum = 256
			$sizeInput.Value = 36
			$sizeInput.Increment = 2
			$sizeInput.Name = "Size Value"
			Add-Member -InputObject $sizeInput -MemberType NoteProperty -Name "Label" -Value "Size Value"

			$colorDropdown = New-Object System.Windows.Forms.ComboBox
			$colorDropdown.Width = 100
			$colorDropdown.DropDownStyle = "DropDownList"
			$ffmpegColors = @("black", "white", "red", "green", "blue", "yellow", "magenta", "cyan", "gray", "darkgray")
			$ffmpegColors | ForEach-Object { $colorDropdown.Items.Add($_) } | Out-Null
			$colorDropdown.SelectedIndex = 0
			$colorDropdown.Name = "Color SelectedItem"
			Add-Member -InputObject $colorDropdown -MemberType NoteProperty -Name "Label" -Value "Color Value"
			$tableLayout.SetColumnSpan($colorDropdown, 2)  # Span the input across 2 columns

			$borderDropdown = New-Object System.Windows.Forms.ComboBox
			$borderDropdown.Width = 100
			$borderDropdown.DropDownStyle = "DropDownList"
			$ffmpegBorderColors = @("none", "black", "white", "red", "green", "blue", "yellow", "magenta", "cyan", "gray", "darkgray")
			$ffmpegBorderColors | ForEach-Object { $borderDropdown.Items.Add($_) } | Out-Null
			$borderDropdown.SelectedIndex = 0
			$borderDropdown.Name = "Border Color SelectedItem"
			Add-Member -InputObject $borderDropdown -MemberType NoteProperty -Name "Label" -Value "Border Color Value"
			$tableLayout.SetColumnSpan($borderDropdown, 2)  # Span the input across 2 columns

			$borderSizeInput = New-Object System.Windows.Forms.NumericUpDown
			$borderSizeInput.Width = 50
			$borderSizeInput.Minimum = 1
			$borderSizeInput.Maximum = 10
			$borderSizeInput.Value = 1
			$borderSizeInput.Increment = 1
			$borderSizeInput.Name = "Border Size Value"
			Add-Member -InputObject $borderSizeInput -MemberType NoteProperty -Name "Label" -Value "Border Size Value"

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
			$positionDropdown.Name = "Position SelectedItem"
			Add-Member -InputObject $positionDropdown -MemberType NoteProperty -Name "Label" -Value "Position Value"
			$tableLayout.SetColumnSpan($positionDropdown, 2)  # Span the input across 2 columns
			
			#Text Label
			$TextLabel = New-Object System.Windows.Forms.Label
			$TextLabel.Text = "Text"
			Add-Member -InputObject $TextLabel -MemberType NoteProperty -Name "Label" -Value "Text Label"
			$TextLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
			$tableLayout.SetColumnSpan($TextLabel,8)  # Span the label across 2 columns
			
			#Start Label
			$StartLabel = New-Object System.Windows.Forms.Label
			$StartLabel.Text = "Start (s)"
			Add-Member -InputObject $StartLabel -MemberType NoteProperty -Name "Label" -Value "Start Label"
			$StartLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
			$tableLayout.SetColumnSpan($StartLabel,2)  # Span the label across 2 columns
			
			#Duration Label
			$DurationLabel = New-Object System.Windows.Forms.Label
			$DurationLabel.Text = "Dur (s)"
			Add-Member -InputObject $DurationLabel -MemberType NoteProperty -Name "Label" -Value "Duration Label"
			$DurationLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
			$tableLayout.SetColumnSpan($DurationLabel,2)  # Span the label across 2 columns

			#XPadding Label
			$xpadLabel = New-Object System.Windows.Forms.Label
			$xpadLabel.Text = "XPad%"
			$xpadLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
			Add-Member -InputObject $xpadLabel -MemberType NoteProperty -Name "Label" -Value "XPad Label"

			#YPadding Label
			$ypadLabel = New-Object System.Windows.Forms.Label
			$ypadLabel.Text = "YPad%"
			$ypadLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft
			Add-Member -InputObject $ypadLabel -MemberType NoteProperty -Name "Label" -Value "YPad Label"
			
			# Text Input
			$Input = New-Object System.Windows.Forms.TextBox
			$Input.Width = 500
			$Input.Name = "Text Value"
			Add-Member -InputObject $Input -MemberType NoteProperty -Name "Label" -Value "Text Value"
			$tableLayout.SetColumnSpan($Input, 7)  # Span the input across 2 columns
			
			# Start Input
			$StartInput = New-Object System.Windows.Forms.NumericUpDown
			$StartInput.Width = 75
			$StartInput.Minimum = 0
			$StartInput.Maximum = 100
			$StartInput.Value = 0
			$StartInput.DecimalPlaces = 3
			$StartInput.Increment = 0.1
			$StartInput.Name = "Start Value"
			Add-Member -InputObject $StartInput -MemberType NoteProperty -Name "Label" -Value "Start Value"
			$tableLayout.SetColumnSpan($StartInput, 2)  # Span the input across 2 columns

			# Duration Input
			$DurationInput = New-Object System.Windows.Forms.NumericUpDown
			$DurationInput.Width = 75
			$DurationInput.Minimum = 0
			$DurationInput.Maximum = 99999
			$DurationInput.Value = 0
			$DurationInput.DecimalPlaces = 3
			$DurationInput.Increment = 0.1
			$DurationInput.Name = "Duration Value"
			Add-Member -InputObject $DurationInput -MemberType NoteProperty -Name "Label" -Value "Duration Value"
			$tableLayout.SetColumnSpan($DurationInput, 2)  # Span the input across 2 columns

			# XPad Input
			$xpadInput = New-Object System.Windows.Forms.NumericUpDown
			$xpadInput.Width = 100
			$xpadInput.Minimum = 0
			$xpadInput.Maximum = 100
			$xpadInput.Value = 5
			$xpadInput.Increment = 1
			$xpadInput.Name = "XPad Value"
			Add-Member -InputObject $xpadInput -MemberType NoteProperty -Name "Label" -Value "XPad Value"

			# YPad Input
			$ypadInput = New-Object System.Windows.Forms.NumericUpDown
			$ypadInput.Width = 100
			$ypadInput.Minimum = 0
			$ypadInput.Maximum = 100
			$ypadInput.Value = 5
			$ypadInput.Increment = 1
			$ypadInput.Name = "YPad Value"
			Add-Member -InputObject $ypadInput -MemberType NoteProperty -Name "Label" -Value "YPad Value"

			#load objects into array
			$global:addedTextElements += [PSCustomObject]@{
				ID = $textNon
				Elements = @(
				$Label,
				$TextEnableLabel,
				$TextFontLabel,
				$TextSizeLabel,
				$TextColorLabel,
				$TextBorderLabel,
				$TextBorderSizeLabel,
				$TextPositionLabel,
				$deleteButton,
				$fontDropdown,
				$sizeInput,
				$colorDropdown,
				$borderDropdown,
				$borderSizeInput,
				$positionDropdown,
				$TextLabel,
				$StartLabel,
				$DurationLabel,
				$xpadLabel,
				$ypadLabel,
				$Input,
				$StartInput,
				$DurationInput,
				$xpadInput,
				$ypadInput
				)
			}
			
			$deleteButton.Add_Click({
				$brows = $addTextButton.Tag.RowHeight
				$brows = $brows - 5
				$addTextButton.Tag.RowHeight = $brows
				$id = $this.Tag.ID
				
# 				$form.Height = (30 * $brows )
				$controlsToRemove = $global:addedTextElements | Where-Object { $_.ID -eq $id }
				$tableLayout.SuspendLayout()
				foreach ($control in $controlsToRemove.Elements) {
					Remove-Control -panel $tableLayout -control $control
					$control.Dispose()
				}
				$tableLayout.ResumeLayout()
				#Move Buttons
				#Add-Control -panel $tableLayout -control $addTextButton -row ($bros - 1) -column 4
				#Add-Control -panel $tableLayout -control $previewButton -row ($bros - 1) -column 6
				#Add-Control -panel $tableLayout -control $okButton -row ($bros - 1) -column 8
				# Remove the textbox from the global array
				$newElements = @()
				$global:addedTextElements | ForEach-Object { 
					if($_.ID -ne $id){
					$newElements += $_
					}
				} | Out-Null
				$global:addedTextElements = $newElements
			})
			
			#add rows
			$addTextButton.Tag.TextNo = $textNon + 1
# 			$form.Height = (30 * ($bros + 4) )
			#Build form
			#Move Buttons
			Add-Control -panel $tableLayout -control $addTextButton -row ($bros + 4) -column 4
			Add-Control -panel $tableLayout -control $previewButton -row ($bros + 4) -column 6
			Add-Control -panel $tableLayout -control $okButton -row ($bros + 4) -column 8
			# Add Labels to Row 0
			Add-Control -panel $tableLayout -control $Label -row ($bros - 1) -column 0
			Add-Control -panel $tableLayout -control $TextEnableLabel -row $bros -column 0
			Add-Control -panel $tableLayout -control $TextFontLabel -row $bros -column 1
			Add-Control -panel $tableLayout -control $TextSizeLabel -row $bros -column 6
			Add-Control -panel $tableLayout -control $TextColorLabel -row $bros -column 7
			Add-Control -panel $tableLayout -control $TextBorderLabel -row $bros -column 9
			Add-Control -panel $tableLayout -control $TextBorderSizeLabel -row $bros -column 11
			Add-Control -panel $tableLayout -control $TextPositionLabel -row $bros -column 12
			# Add Inputs to Row 2
			Add-Control -panel $tableLayout -control $deleteButton -row ($bros + 1) -column 0
			Add-Control -panel $tableLayout -control $fontDropdown -row ($bros + 1) -column 1
			Add-Control -panel $tableLayout -control $sizeInput -row ($bros + 1) -column 6
			Add-Control -panel $tableLayout -control $colorDropdown -row ($bros + 1) -column 7
			Add-Control -panel $tableLayout -control $borderDropdown -row ($bros + 1) -column 9
			Add-Control -panel $tableLayout -control $borderSizeInput -row ($bros + 1) -column 11
			Add-Control -panel $tableLayout -control $positionDropdown -row ($bros + 1) -column 12
			# Add Labels to next Row
			Add-Control -panel $tableLayout -control $TextLabel -row ($bros + 2) -column 0
			Add-Control -panel $tableLayout -control $StartLabel -row ($bros + 2) -column 8
			Add-Control -panel $tableLayout -control $DurationLabel -row ($bros + 2) -column 10
			Add-Control -panel $tableLayout -control $xpadLabel -row ($bros + 2) -column 12
			Add-Control -panel $tableLayout -control $ypadLabel -row ($bros + 2) -column 13
			# Add Labels to next Row
			Add-Control -panel $tableLayout -control $Input -row ($bros + 3) -column 0
			Add-Control -panel $tableLayout -control $StartInput -row ($bros + 3) -column 8
			Add-Control -panel $tableLayout -control $DurationInput -row ($bros + 3) -column 10
			Add-Control -panel $tableLayout -control $xpadInput -row ($bros + 3) -column 12
			Add-Control -panel $tableLayout -control $ypadInput -row ($bros + 3) -column 13
			$tableLayout.ResumeLayout()
			$bros = $bros + 5
			$brows = $brows + 5
			$addTextButton.Tag.Rows = $bros
			$addTextButton.Tag.RowHeight = $brows
		}
	})

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
		$processedTexts = @()

		$global:addedTextElements | ForEach-Object {
		
			$_.Elements | ForEach-Object {
				$field = $_
				switch($field.Name){
					"Font SelectedItem" { $textFont = $field.SelectedItem }
					"Size Value" { $textSize = $field.Value }
					"Color SelectedItem" { $textColor = $field.SelectedItem }
					"Border Color SelectedItem" { $textBorderColor = $field.SelectedItem }
					"Border Size Value" { $textBorderSize = $field.Value }
					"Position SelectedItem" { $textPosition = $field.SelectedItem }
					"Text Value" { $textText = $field.Text.Substring(0, [math]::Min(64, $field.Text.Length)) }
					"Start Value" { $textStart = [float]$field.Value }
					"Duration Value" { if($field.Value -gt 0) { $textDuration = [float]$field.Value + $textStart } else { $textDuration = [float]$metaDur } }
					"XPad Value" { $textXPad = $field.Value }
					"YPad Value" { $textYPad = $field.Value }
				}
			} | Out-Null
			$settings = @{
				TextFont = $textFont
				TextSize = $textSize
				TextColor = $textColor
				TextBorderColor = $textBorderColor
				TextBorderSize = $textBorderSize
				TextPosition = $textPosition
				Text = $textText
				TextStart = $textStart
				TextDuration = $textDuration
				TextXPad = $textXPad
				TextYPad = $textYPad
			}
			$processedTexts += $settings
		} | Out-Null

		$drawtext = Process-Video -form $form -texts $processedTexts -height $height -width $width
		
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
			DateFormat = $dateFormatInput.Text.Substring(0, [math]::Min(32, $dateFormatInput.Text.Length))
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
			TimeFormat = $timeFormatInput.Text.Substring(0, [math]::Min(32, $timeFormatInput.Text.Length))
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

	$processedTexts = @()

	$global:addedTextElements | ForEach-Object {
		
		$_.Elements | ForEach-Object {
			$field = $_
			switch($field.Name){
				"Font SelectedItem" { $textFont = $field.SelectedItem }
				"Size Value" { $textSize = $field.Value }
				"Color SelectedItem" { $textColor = $field.SelectedItem }
				"Border Color SelectedItem" { $textBorderColor = $field.SelectedItem }
				"Border Size Value" { $textBorderSize = $field.Value }
				"Position SelectedItem" { $textPosition = $field.SelectedItem }
				"Text Value" { $textText = $field.Text.Substring(0, [math]::Min(64, $field.Text.Length)) }
				"Start Value" { $textStart = [float]$field.Value }
				"Duration Value" { if($field.Value -gt 0) { $textDuration = [float]$field.Value + $textStart } else { $textDuration = [float]$metaDur } }
				"XPad Value" { $textXPad = $field.Value }
				"YPad Value" { $textYPad = $field.Value }
			}
		} | Out-Null
		$settings = @{
			TextFont = $textFont
			TextSize = $textSize
			TextColor = $textColor
			TextBorderColor = $textBorderColor
			TextBorderSize = $textBorderSize
			TextPosition = $textPosition
			Text = $textText
			TextStart = $textStart
			TextDuration = $textDuration
			TextXPad = $textXPad
			TextYPad = $textYPad
		}
		$processedTexts += $settings
	} | Out-Null

	$drawtext = Process-Video -form $form -texts $processedTexts -height $height -width $width


	if(-Not $drawtext){
		Write-Host "No text to draw to video! Operation canceled. Video has not been written"
		exit
	}

	# ----- convert the 12-hour value to 24-hour first -----
	$hour24 = $hourInput.Value
	if ($ampm -eq 'PM' -and $hourInput.Value -lt 12) { $hour24 += 12 }
	if ($ampm -eq 'AM' -and $hourInput.Value -eq 12) { $hour24 = 0 }   # midnight edge-case

	# ----- build the DateTime -----
	$creationDate = Get-Date `
					-Year   $yearInput.Value  `
					-Month  $monthInput.Value `
					-Day    $dayInput.Value   `
					-Hour   $hour24   `
					-Minute $minuteInput.Value`
					-Second $secondInput.Value`
	# Auto-generate output filename in same folder
	$base = [IO.Path]::GetFileNameWithoutExtension($inputFile)
	$ext = [IO.Path]::GetExtension($inputFile)
	$output = Join-Path $tempDir "$base`_stamped$ext"
	$prevSettings = [PSCustomObject]@{
		Enable = $enableCheckbox.Checked
		FontIndex = $fontDropdown.SelectedIndex
		Size = $sizeInput.Value
		ColorIndex = $colorDropdown.SelectedIndex
		BorderColorIndex = $borderDropdown.SelectedIndex
		BorderSize = $borderSizeInput.Value
		PositionIndex = $positionDropdown.SelectedIndex
		XPadding = $xpadInput.Value
		YPadding = $ypadInput.Value
		EnableTime = $enableTimeCheckbox.Checked
		TimeFontIndex = $fontTimeDropdown.SelectedIndex
		TimeSize = $sizeTimeInput.Value
		TimeColorIndex = $colorTimeDropdown.SelectedIndex
		TimeBorderColorIndex = $borderTimeDropdown.SelectedIndex
		TimeBorderSize = $borderSizeTimeInput.Value
		TimePositionIndex = $positionTimeDropdown.SelectedIndex
		TimeXPadding = $xpadTimeInput.Value
		TimeYPadding = $ypadTimeInput.Value
		CreationDate = $creationDate
		DateFormat = $dateFormatInput.Text
		TimeFormat = $timeFormatInput.Text
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
	$stampedClips += $output
}




if ($stampedClips.Count -gt 1) {
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.InitialDirectory = (Split-Path $inputFiles[0])
    $saveFileDialog.Filter = "MP4 Video (*.mp4)|*.mp4"
    $saveFileDialog.Title  = "Save stitched video as..."
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $finalOut = $saveFileDialog.FileName
        $listFile = Join-Path $tempDir "concat_list.txt"
        $stampedClips | ForEach-Object { "file '$($_)'" } | Out-File -FilePath $listFile -Encoding ascii
        & $ffmpeg -hide_banner -loglevel error -f concat -safe 0 -i $listFile -c copy -y $finalOut
        Write-Output "Finished - stitched file saved to:  $finalOut"
    } else {
        Write-Warning "Stitching cancelled - leaving stamped clips intact in $tempDir"
    }
} else {
    # Only one stamped clip – ask where to save it
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.InitialDirectory = (Split-Path $inputFiles[0])
    $saveFileDialog.Filter = "MP4 Video (*.mp4)|*.mp4"
    $saveFileDialog.Title  = "Save stamped video as…"
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $finalOut = $saveFileDialog.FileName
        Move-Item -LiteralPath $stampedClips[0] -Destination $finalOut -Force
        Write-Output "Finished - stamped file saved to:  $finalOut"
    } else {
        Write-Warning "Save cancelled - leaving stamped clip in $tempDir"
    }
}

# Clean up temporary folder
if (Test-Path $tempDir) {
    Remove-Item -LiteralPath $tempDir -Recurse -Force
}
