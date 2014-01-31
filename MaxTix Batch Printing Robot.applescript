-- Environment variables/settings
property setupPage : {accessKey:"d", pageTitle:"Max 1.3.6 (901:1)"}
property reportPage : {accessKey:"r", pageTitle:"Max 1.3.6 (902:7)"}
property MaxTixURL : "https://maxtix.bard.edu/scripts/max/2000/max.exe"
property numberOfSingleBatchCopies : "3"

local setupPageSource

tell application "Safari"
	activate
	set activeWindow to current tab of window 1
	set setupPageSource to source of activeWindow
	-- Check to make sure we're in the right place.
	if URL of activeWindow is not MaxTixURL or (do JavaScript "document.title" in activeWindow) is not pageTitle of setupPage or setupPageSource does not contain "<TH COLSPAN=2>Batch Report</TH>" then
		my failWith("This script will only run when the Batch Report settings page is open.")
	end if
	try
		set dialogResult to display dialog "Enter starting and ending batches separated by a hyphen." buttons {"Cancel", "Print Batches", "Print with Total"} default button "Print with Total" cancel button "Cancel" default answer "####-####"
	on error number -128
		return "User cancelled."
	end try
end tell

-- TO-DO: Add UI Scripting to click "Batch Range" radio button.
-- Validate the user-inputted batch range.
try
	set batchRange to text returned of dialogResult
	if batchRange does not contain "-" then
		failWith("Please enter a batch range in the required format: ####-####.")
	end if
	set ASTID to AppleScript's text item delimiters
	set AppleScript's text item delimiters to "-"
	set startBatch to first text item of batchRange
	set endBatch to second text item of batchRange
	set AppleScript's text item delimiters to ASTID
	if batchIDIsValid(startBatch) is true and batchIDIsValid(endBatch) is true then
		if startBatch is greater than endBatch then
			failWith("The start batch cannot be greater than the end batch.")
		end if
		if endBatch is greater than findCurrentBatchID(setupPageSource) then
			failWith("The end batch cannot be greater than the current batch.")
		end if
		if endBatch - startBatch is greater than 30 and button returned of dialogResult is "Print with Total" then
			failWith("The batch range too big for printing with a grand total.")
		end if
		if endBatch - startBatch is greater than 99 then
			tell application "Safari"
				display dialog "Are you sure you want to print " & endBatch - startBatch & " batches?"
			end tell
		end if
	else
		failWith("Please enter a batch range in the required format: ####-####.")
	end if
on error number -128
	return "Batch range was invalid."
end try

-- Print each batch, starting with a single copy of the grand total, if requested.
if button returned of dialogResult is "Print with Total" and startBatch is not equal to endBatch then
	printBatchReport(startBatch, endBatch, 1)
end if
repeat with currentBatch from startBatch to endBatch
	printBatchReport(currentBatch, currentBatch, numberOfSingleBatchCopies)
end repeat


-- ::HANDLERS:: --
-- Manual validation that batchIDs only contain numbers.
on batchIDIsValid(theBatch)
	set legalCharacters to {"1", "2", "3", "4", "5", "6", "7", "8", "9", "0"}
	set badBatchID to false
	repeat with i from 1 to (get count of characters in theBatch)
		set theChar to character i of theBatch
		if theChar is not in legalCharacters then
			set badBatchID to true
		end if
	end repeat
	if badBatchID is false then
		return true
	end if
	return false
end batchIDIsValid

-- This is the AppleScript way to sed
on findCurrentBatchID(pageSource)
	set ASTID to AppleScript's text item delimiters
	set AppleScript's text item delimiters to "<TD>Batch Id:</TD>
<TD><B>"
	set sourceChunk to second text item of pageSource
	set AppleScript's text item delimiters to "</B></TD>"
	set batchid to first text item of sourceChunk
	set AppleScript's text item delimiters to ASTID
	if batchIDIsValid(batchid) then
		return batchid
	else
		failWith("MaxTix Batch Printing Robot has experienced an error in findCurrentBatchID.")
	end if
end findCurrentBatchID

-- Run a batch report and print it
on printBatchReport(startBatch, endBatch, copies)
	tell application "System Events"
		tell application process "Safari"
			keystroke tab
			keystroke startBatch as text
			keystroke tab
			keystroke endBatch as text
			my loadPage(reportPage)
			my printCopiesOfCurrentPage(copies)
			my loadPage(setupPage)
		end tell
	end tell
end printBatchReport

-- This is a bug in Safari that causes a hang.
--on BROKENprintCopiesOfCurrentPage(copies)
--	tell application "Safari"
--		print document 1 with properties {copies:copies, collating:true} without print dialog
--	end tell
--end BROKENprintCopiesOfCurrentPage

-- Instead, do this the old fashioned way.
on printCopiesOfCurrentPage(copies)
	tell application "System Events"
		tell application process "Safari"
			keystroke "p" using command down
			-- OR THIS -- click menu item "Print…" of menu "File" of menu bar 1
			delay 1
			keystroke copies as text
			keystroke return
			delay 3
		end tell
	end tell
end printCopiesOfCurrentPage

-- Load the requested page, exiting after loading is complete or retrying access key.
on loadPage(requestedPage)
	set isComplete to false
	repeat while isComplete is false
		tell application "System Events"
			tell application process "Safari"
				keystroke accessKey of requestedPage using {control down, option down}
			end tell
		end tell
		tell application "Safari"
			set activeWindow to current tab of window 1
			repeat 20 times
				if (do JavaScript "document.title" in activeWindow) is pageTitle of requestedPage then
					repeat
						if (do JavaScript "document.readyState" in activeWindow) is "complete" then exit repeat
						delay 0.1
					end repeat
					set isComplete to true
					exit repeat
				end if
				delay 0.1
			end repeat
		end tell
	end repeat
	return isComplete
end loadPage

-- There are a lot of ways this can go wrong…
on failWith(reason)
	tell application "Safari"
		display dialog reason
		error number -128
	end tell
end failWith
