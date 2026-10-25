-- usage: osascript export.applescript /abs/deck.pptx /abs/outdir
-- Exports every non-skipped slide as PNG and prints presenter notes separated by ASCII RS.
on run argv
	set inFile to POSIX file (item 1 of argv)
	set outDir to POSIX file (item 2 of argv)
	set sep to (ASCII character 30)
	with timeout of 180 seconds
	tell application "Keynote"
		set doc to open inFile
		delay 1
		export doc to outDir as slide images with properties {image format:PNG, skipped slides:false}
		set notesList to {}
		repeat with s in slides of doc
			if skipped of s is false then set end of notesList to (presenter notes of s as text)
		end repeat
		close doc saving no
	end tell
	end timeout
	set AppleScript's text item delimiters to sep
	set out to notesList as text
	set AppleScript's text item delimiters to ""
	return out
end run
