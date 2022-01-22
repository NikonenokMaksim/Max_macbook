property theList : {"doc", "docx"}

on run {input, parameters}
	tell application "Microsoft Word" to set theOldDefaultPath to get default file path file path type documents path
	--tell application "Finder"
	--set input to selection
	--end tell
	
	--tell application id "com.microsoft.Word"
	--activate
	--set view type of view of active window to draft view
	
	--repeat with aFile in input
	--open aFile
	--set theOutputPath to ((aFile as text) & ".pdf")
	--repeat while not (active document is not missing value)
	--delay 0.5
	--end repeat
	--set activeDoc to active document
	--save as activeDoc file name theOutputPath file format format PDF
	--close active document saving no
	--end repeat
	
	--end tell
	
	repeat with x in input
		try
			set theDoc to contents of x
			--tell application "Finder"
			tell application "Finder"
				
				set theFilePath to container of theDoc as text
				
				set ext to name extension of theDoc
				if ext is in theList then
					set theName to name of theDoc
					copy length of theName to l
					copy length of ext to exl
					
					set n to l - exl - 1
					copy characters 1 through n of theName as string to theFilename
					
					set theFilename to theFilename & ".pdf"
					
					tell application "Microsoft Word"
						
						set default file path file path type documents path path theFilePath
						open theDoc
						set theActiveDoc to the active document
						
						save as theActiveDoc file format format PDF file name theFilename
						copy (POSIX path of (theFilePath & theFilename as string)) to end of output
						
						
						--close active document saving no
						closetheActiveDoc
					end tell
				end if
			end tell
		end try
	end repeat
	
	tell application "Microsoft Word"
		set default file path file path type documents path path theOldDefaultPath
		quit
	end tell
end run