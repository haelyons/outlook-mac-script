tell application "Microsoft Outlook"
	-- List of folder IDs to process
	set folderIDs to {157, 162, 136, 151}
	
	-- Path to save attachments
	set savePath to (path to documents folder as string) & "NUS Attachments:"
	
	-- Check if NUS Attachments folder exists, if not create it
	tell application "Finder"
		set documentsFolderPath to (path to documents folder as string)
		set nusAttachmentsFolderPath to documentsFolderPath & "NUS Attachments"
		
		-- Check if the NUS Attachments folder already exists
		if not (exists folder nusAttachmentsFolderPath) then
			make new folder at documentsFolderPath with properties {name:"NUS Attachments"}
		end if
	end tell
	
	-- Iterate through each specified folder ID
	repeat with folderID in folderIDs
		set currentFolder to get mail folder id folderID
		
		-- Check if there are messages with attachments
		repeat with msg in currentFolder's messages
			if (count of msg's attachments) > 0 then
				repeat with att in msg's attachments
					-- Generate unique file name
					set originalName to att's name
					set uniqueName to originalName
					set counter to 1
					set saveLocation to savePath & uniqueName
					
					-- Check if file exists and rename if necessary
					tell application "Finder"
						repeat while (exists file saveLocation)
							set uniqueName to (text 1 through ((length of originalName) - (length of (name extension of originalName)) - 1) of originalName) & "-" & counter & "." & (name extension of originalName)
							set saveLocation to savePath & uniqueName
							set counter to counter + 1
						end repeat
					end tell
					
					-- Save the attachment
					save att in saveLocation
				end repeat
			end if
		end repeat
	end repeat
end tell
