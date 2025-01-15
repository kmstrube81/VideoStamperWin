# VideoStamperWin
By Kasey M. Strube

Version 0.2


Utilizes ffmpeg to automatically add text and timestamps to videos.
Used primary to convert iPhone .MOVs to MP4 for portability and add a
camcorder style timestamp for viewing on TV


Requirements-

Windows

Powershell

FFMPEG binaries (including ffprobe and ffplay)


TODO Version 0.3 Features-

*add arbitrary text to video

*add tooltip for format strings

*give arbitrary text custom start time and duration.

Features-

*set border color and border size

*position stamps in Top Left, Top Middle, Top Right, Middle Left, 
Middle, Middle Right, Bottom Left, Bottom Middle, Bottom Right

*Manually set date and time on timestamp

*Independentally position Date Stamp and Time Stamp using position and horizontal (xPad%)
and vertical (yPad%) padding values

*Preview before saving changes

*Specify a save location and file name

*Stamps timestamp from metadata of .MOV or .MP4 video files

*Allows the choice of using any installed font, font size, color, and border.


Acknowledgements-

INI Read Write modules by Oliver Lipkau 

https://devblogs.microsoft.com/scripting/use-powershell-to-work-with-any-ini-file/
