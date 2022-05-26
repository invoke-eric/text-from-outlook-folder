tell application "Microsoft Outlook"
	
	set theContent to ""
	set topFolder to folder "Inbox" of default account
	set subFolder to folder "subFoldername" of topFolder
	set subFolder2 to folder "subFolder2name" of subFolder
	set theMessages to messages of subFolder2
	repeat with theMessage in theMessages
		if subject of theMessage contains "Subject line I care about" then
			set theContent to theContent & plain text content of theMessage
		end if
	end repeat
	
	
end tell

do shell script "echo  " & quoted form of theContent & " >  /Users/[username]/Documents/output.txt"
