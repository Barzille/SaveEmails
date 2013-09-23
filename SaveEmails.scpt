property homefolder : path to home folder
-- set any folder as anchor point
property dropboxFolder : (homefolder as string) & "Dropbox:Gewerbe:Steuerberater:Buchhaltung:"

-- select any number of messages in your mail client and run the app
tell application "Mail"
	try
		set selectedMessages to selection
		repeat with msg in selectedMessages
			set msgName to subject of msg
			set msgDate to date sent of msg as string
			
			
			set oldDelimiters to AppleScript's text item delimiters
			set AppleScript's text item delimiters to " "
			set dateArray to every text item of msgDate
			set sentDay to 2nd item in dateArray
			set sentMonth to 3rd item in dateArray
			set sentYear to 4th item in dateArray
			
			-- set different save folders
			-- each mail is save to YEAR/EINGANG/QUARTER/MONTH/
			if sentMonth is in {"Januar", "Februar", "März"} then
				set quarter to "I. Quartal"
			else if sentMonth is in {"April", "Mai", "Juni"} then
				set quarter to "II. Quartal"
			else if sentMonth is in {"Juli", "August", "September"} then
				set quarter to "III. Quartal"
			else if sentMonth is in {"Oktober", "November", "Dezember"} then
				set quarter to "IV. Quartal"
			end if
			
			-- replace : with _ in email subjects
			set {ASTID, AppleScript's text item delimiters} to {AppleScript's text item delimiters, ":"}
			set msgName to text items of msgName
			set AppleScript's text item delimiters to "_"
			set msgName to msgName as rich text
			set AppleScript's text item delimiters to oldDelimiters
			set msgContent to content of msg
			
			-- check folders and create if neccessary
			my createFolderIfNotExists(dropboxFolder, sentYear)
			my createFolderIfNotExists((dropboxFolder & sentYear & ":") as string, "Eingang")
			my createFolderIfNotExists((dropboxFolder & sentYear & ":Eingang:") as string, quarter)
			my createFolderIfNotExists((dropboxFolder & sentYear & ":Eingang:" & quarter & ":") as string, sentMonth & " " & sentYear)
			
			-- create text file
			set fileName to (dropboxFolder & sentYear & ":Eingang:" & quarter & ":" & sentMonth & " " & sentYear & ":" & msgName & ".txt") as string
			set newFileID to open for access file fileName with write permission
			write msgContent to newFileID
			close access newFileID
		end repeat
		display dialog "E-Mails wurden erfolgreich gespeichert"
	on error line number num
		display dialog "Error on line number " & num
	end try
	
end tell

on createFolderIfNotExists(atFolder, folderName)
	tell application "Finder"
		if (exists folder (atFolder & folderName)) then
			-- do nothing
		else
			make new folder at atFolder with properties {name:folderName}
		end if
	end tell
end createFolderIfNotExists
