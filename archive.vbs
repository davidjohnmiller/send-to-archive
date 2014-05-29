' **************************************************************************************
' NAME: 		Archiver 
' PURPOSE:		Copies selected objects (files and/or folders) to a folder
' There are a few options that may be useful for groups, not so much for an individual
' E.g. Putting the username in the name of each subfolder
' Some may also not be a fan of getting prompted to include a comment with each archive
' Either way, just comment out the sections you don't need - indicated below
' **************************************************************************************
sDateTime = ""
sUserName = ""
sComment = ""

Const OverWrite = True
Base_Folder = Replace (Wscript.ScriptName, ".vbs", "")
GetDateTime sDateTime

' COMMENT OUT THE FOLLOWING TWO LINES IF YOU DON'T WANT USERNAME OR COMMENTS
'GetUserName sUserName
sComment = InputBox("Enter a Comment For the Folder Name - Spaces Will Be Chomped")'
sComment = Replace (sComment, " ", "") 'I hate spaces in FolderNames

Folder_Detail = sDateTime
If sUserName = "" Then
	'Nothing
Else
	Folder_Detail = Folder_Detail & "_" & sUserName
End If

If sComment = "" Then
	'Nothing
Else
	Folder_Detail = Folder_Detail & "_" & sComment
End If


Set objArgs = WScript.Arguments 
 For I = 0 to objArgs.Count - 1 
 'WScript.Echo objArgs(I)
 
 Set objFSO = CreateObject("Scripting.FileSystemObject")
 
	If objFSO.FolderExists(objArgs(I)) Then
		
		'Wscript.Echo objArgs(I) + " is a folder"
		Set objFolder = objFSO.GetFolder(objArgs(I))

		'Can't set the parent folder from right-click initiated script
		'If the first item is a folder:
		'Can find the parent folder by string replacement (no parent item)
		'In the full path, replace the actual folder name with nothing
		'Leaves the parent path
		If I = 0 Then
			Parent_Folder = Replace (objArgs(I), objFolder.Name, "")
			Archive_Folder = Parent_Folder & Base_Folder & "\"
			Destination_Folder = Archive_Folder & Folder_Detail & "\"
			'Wscript.Echo "First is Folder and will be created at " & Destination_Folder
		End If

		from_this_folder = objArgs(I)
		to_this_folder = Destination_Folder & objFolder.Name
		
		'Make sure folders exists before attempting to copy
		If NOT (objFSO.FolderExists(Archive_Folder)) Then
			objFSO.CreateFolder(Archive_Folder)
		Else
			'Folder is Already there - good to go
		End If

		If NOT (objFSO.FolderExists(Destination_Folder)) Then
			objFSO.CreateFolder(Destination_Folder)
		Else
			'Folder is Already there - good to go
		End If
		
		objFSO.CopyFolder from_this_folder , to_this_folder , OverWrite
		
	Else
	
		'Wscript.Echo objArgs(I) + " is a file"
		Set objFile = objFSO.GetFile(objArgs(I))
		
		'Can't set the parent folder from right-click initiated script
		'If the first item is a file:
		'Can find the parent folder by inspection
		If I = 0 Then
			Parent_Folder = objFSO.GetParentFolderName(objFile)
			Archive_Folder = Parent_Folder & "\" & Base_Folder & "\"
			Destination_Folder = Archive_Folder & Folder_Detail & "\"
			'Wscript.Echo "First is File and will be created at " & Destination_Folder
		End If

		from_this_file = objArgs(I)
		to_this_file = Destination_Folder & objFSO.GetFileName(objFile)
		
		'Make sure folders exists before attempting to copy
		If NOT (objFSO.FolderExists(Archive_Folder)) Then
			objFSO.CreateFolder(Archive_Folder)
		Else
			'Folder is Already there - good to go
		End If

		If NOT (objFSO.FolderExists(Destination_Folder)) Then
			objFSO.CreateFolder(Destination_Folder)
		Else
			'Folder is Already there - good to go
		End If
		
		objFSO.CopyFile from_this_file , to_this_file, OverWrite
		
	End If
	
 Next
 
' **************************************************************************************
' NAME:			GetUserName
' PURPOSE:		Gets the Name of the User Currently Logged In
' **************************************************************************************
FUNCTION GetUserName(UserName)
DIM objNetwork
'
SET objNetwork = CreateObject("WScript.Network")
UserName = objNetwork.UserName
'
END FUNCTION
 
' **************************************************************************************
' NAME:			GetDateTime
' PURPOSE:		Gets the current date and time
' **************************************************************************************
FUNCTION GetDateTime(DateTime)
DIM       sDate
DIM       sTime
DIM       iYear
DIM       iMonth
DIM       iDay
DIM       iHour
DIM       iMin
DIM       iSec
'
' Get the date
iYear = Year(now)
iMonth = Month(now)
iDay = Day(now)
iHour = Hour(now)
iMin = Minute(now)
iSec  = Second(now)
'
' Format date and time in the format YYYYMMDDTHHMMSS
' Year will always be in the correct format
sDate = iYear
' Format the Month
IF iMonth < 10 THEN
    sDate = sDate & "0" & iMonth
ELSE
    sDate = sDate & iMonth
END IF
' Format the Day
IF iDay < 10 THEN
    sDate = sDate & "0" & iDay
ELSE
    sDate = sDate & iDay
END IF
' Format the Hour
IF iHour < 10 THEN
    sTime = "0" & iHour
ELSE
    sTime = iHour
END IF
' Format the Minute
IF iMin < 10 THEN
    sTime = sTime & "0" & iMin
ELSE
    sTime = sTime & iMin
END IF
' Format the Second
IF iSec < 10 THEN
    sTime = sTime & "0" & iSec
ELSE
    sTime = sTime & iSec
END IF
'
'Concatenate date and time in the required format
DateTime = sDate & "T" & sTime
'
END FUNCTION