--==============================
-- OmniFocus > Prepare Task Completion Report
-- Version 2.0.1
-- Written By: Ben Waldie <https://about.me/benwaldie>
-- Updated By: Barry Sampson <https://barrysampson.com/>

-- Description: This script retrieves a list of OmniFocus tasks completed today, yesterday, this week, last week, or this month. It then summarizes the tasks in a new DevonThink note. This script is a copy of the one written by Ben Waldie and posted to Engadget - https://www.engadget.com/2013/02/18/applescripting-omnifocus-send-completed-task-report-to-evernot/. All I've done is modified it to send the compled note to DevonThink rather than Evernote. I've made some very minor tweaks to the HTML to suit my own preference.  
-- Version History:
-- 1.0.0 - Initial release
-- 2.0.0 - Added support for including full project paths, context names, estimate time, start dates, modification dates, completion dates, and notes in task reports.
-- 2.0.1 - Modified the script to send the report to DevonThink Pro rather than Evernote.
--==============================

-- This property controls whether full project paths (including parent folders) are displayed
property includeFullProjectPaths : true

-- These properties control whether additional task content is displayed
property includeTaskContext : true
property includeTaskEstimatedTime : false
property includeTaskStartDate : false
property includeTaskModificationDate : false
property includeTaskCompletionDate : true
property includeTaskNotes : false

-- This setting specifies a name for the new note
property theNoteName : "OmniFocus Completed Task Report"

-- Prompt the user to choose a scope for the report
activate
set theReportScope to choose from list {"Today", "Yesterday", "This Week", "Last Week", "This Month", "Last Month"} default items {"Yesterday"} with prompt "Generate a report for:" with title "OmniFocus Completed Task Report"
if theReportScope = false then return
set theReportScope to item 1 of theReportScope

-- Calculate the task start and end dates, based on the specified scope
set theStartDate to current date
set hours of theStartDate to 0
set minutes of theStartDate to 0
set seconds of theStartDate to 0
set theEndDate to theStartDate + (23 * hours) + (59 * minutes) + 59

if theReportScope = "Today" then
	set theDateRange to date string of theStartDate
else if theReportScope = "Yesterday" then
	set theStartDate to theStartDate - 1 * days
	set theEndDate to theEndDate - 1 * days
	set theDateRange to date string of theStartDate
else if theReportScope = "This Week" then
	repeat until (weekday of theStartDate) = Sunday
		set theStartDate to theStartDate - 1 * days
	end repeat
	repeat until (weekday of theEndDate) = Saturday
		set theEndDate to theEndDate + 1 * days
	end repeat
	set theDateRange to (date string of theStartDate) & " through " & (date string of theEndDate)
else if theReportScope = "Last Week" then
	set theStartDate to theStartDate - 7 * days
	set theEndDate to theEndDate - 7 * days
	repeat until (weekday of theStartDate) = Sunday
		set theStartDate to theStartDate - 1 * days
	end repeat
	repeat until (weekday of theEndDate) = Saturday
		set theEndDate to theEndDate + 1 * days
	end repeat
	set theDateRange to (date string of theStartDate) & " through " & (date string of theEndDate)
else if theReportScope = "This Month" then
	repeat until (day of theStartDate) = 1
		set theStartDate to theStartDate - 1 * days
	end repeat
	repeat until (month of theEndDate) is not equal to (month of theStartDate)
		set theEndDate to theEndDate + 1 * days
	end repeat
	set theEndDate to theEndDate - 1 * days
	set theDateRange to (date string of theStartDate) & " through " & (date string of theEndDate)
end if

-- Begin preparing the task list as HTML.
set theProgressDetail to "<html><body><h1>Completed Tasks</h1><b>" & theDateRange & "</b><hr><br>"

-- Retrieve a list of projects modified within the specified scope
set modifiedTasksDetected to false
tell application "OmniFocus"
	tell front document
		set theModifiedProjects to every flattened project where its modification date is greater than theStartDate and modification date is less than theEndDate
		
		-- Loop through any detected projects
		repeat with a from 1 to length of theModifiedProjects
			set theCurrentProject to item a of theModifiedProjects
			
			-- Retrieve any project tasks modified within the specified scope
			set theCompletedTasks to (every flattened task of theCurrentProject where its completed = true and modification date is greater than theStartDate and modification date is less than theEndDate and number of tasks = 0)
			
			-- Process the project if tasks were found
			if theCompletedTasks is not equal to {} then
				set modifiedTasksDetected to true
				
				-- Append the project folder path to the project name
				set theProjectFolderPath to ""
				if includeFullProjectPaths = true then
					set theProjectFolderPath to getProjectFolderPath(theCurrentProject) of me
					if theProjectFolderPath is not equal to "" then set theProjectFolderPath to theProjectFolderPath & " > "
				end if
				
				-- Append the project name to the report
				set theProgressDetail to theProgressDetail & "<h2>" & theProjectFolderPath & name of theCurrentProject & "</h2>" & return & "<ul>"
				
				-- Loop through the detected tasks for the project
				repeat with b from 1 to length of theCompletedTasks
					set theCurrentTask to item b of theCompletedTasks
					
					-- Append the tasks's name to the task list
					set theProgressDetail to theProgressDetail & "<li>" & name of theCurrentTask
					
					-- Set up a variable for the task detail, if relevant
					set theTaskDetail to ""
					
					-- Append the context to the task detail
					if includeTaskContext = true then
						set theContext to context of theCurrentTask
						if theContext is not equal to missing value then set theTaskDetail to appendTaskDetail(theTaskDetail, name of theContext, "Context", "") of me
					end if
					
					-- Append the estimated time to the task detail
					if includeTaskEstimatedTime = true then set theTaskDetail to appendTaskDetail(theTaskDetail, estimated minutes of theCurrentTask, "Estimated Time", " minutes") of me
					
					-- Append the start date to the task detail
					if includeTaskStartDate = true then set theTaskDetail to appendTaskDetail(theTaskDetail, defer date of theCurrentTask, "Start Date", "") of me
					
					-- Append the modification date to the task detail
					if includeTaskStartDate = true then set theTaskDetail to appendTaskDetail(theTaskDetail, modification date of theCurrentTask, "Modification Date", "") of me
					
					-- Append the completion date to the task detail
					if includeTaskStartDate = true then set theTaskDetail to appendTaskDetail(theTaskDetail, completion date of theCurrentTask, "Completion Date", "") of me
					
					-- Append the task's notes to the task detail
					if includeTaskNotes = true then set theTaskDetail to appendTaskDetail(theTaskDetail, note of theCurrentTask, "Note", "") of me
					
					-- Append the task detail to the task list
					if theTaskDetail is not equal to "" then
						set theProgressDetail to theProgressDetail & "<br><p style=\"color: gray\">" & theTaskDetail & "</p>"
					end if
					
					-- Finish adding the task's HTML to the list
					set theProgressDetail to theProgressDetail & "</li>" & return
				end repeat
				set theProgressDetail to theProgressDetail & "</ul>" & return
			end if
		end repeat
	end tell
end tell
set theProgressDetail to theProgressDetail & "</body></html>"

-- Notify the user if no projects or tasks were found
if modifiedTasksDetected = false then
	display alert "OmniFocus Completed Task Report" message "No modified tasks were found for " & theReportScope & "."
	return
end if

-- Create the note in DevonThink.
tell application id "com.devon-technologies.thinkpro2"
	activate
	create record with {type:html, source:theProgressDetail, name:theNoteName} in "/Inbox"
end tell

-- This handler gets the folder path for a project
on getProjectFolderPath(theProject)
	tell application "OmniFocus"
		set theFolderPath to ""
		if folder of theProject exists then
			set theFolder to folder of theProject
			repeat
				if theFolderPath is not equal to "" then set theFolderPath to " : " & theFolderPath
				set theFolderPath to name of theFolder & theFolderPath
				if class of container of theFolder = folder then
					set theFolder to container of theFolder
				else
					exit repeat
				end if
			end repeat
		end if
		if theFolderPath = "" then set theFolderPath to ""
		return theFolderPath
	end tell
end getProjectFolderPath

-- This handler appends a value to the task detail
on appendTaskDetail(theTaskDetail, theValue, thePrefix, theSuffix)
	if theTaskDetail is not equal to "" then set theTaskDetail to theTaskDetail & "<br>"
	if theValue = missing value or theValue = "" then
		set theValue to "N/A"
	else
		set theValue to theValue & theSuffix
	end if
	return theTaskDetail & thePrefix & ": " & theValue
end appendTaskDetail
