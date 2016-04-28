' **************************************************************************************
' NAME: 		Archiver 
' PURPOSE:		Copies selected objects (files and/or folders) to a folder
' There are a few options that may be useful for groups, not so much for an individual
' e.g. putting the username in the name of each subfolder, just comment out the sections
' you don't need
' **************************************************************************************
sDateTime = ""
sUserName = ""
sComment = ""

Const OverWrite = True
Base_Folder = Replace (Wscript.ScriptName, ".vbs", "")
GetDateTime sDateTime

Folder_Detail = sDateTime

'GetUserName sUserName
If sUserName <> "" Then
	Folder_Detail = Folder_Detail & "_" & sUserName
End If

sComment = InputBox("Enter a comment for the archive")'
sComment = Replace(sComment, " ", "")
If sComment <> "" Then
	Folder_Detail = Folder_Detail & "_" & sComment
End If

Set objArgs = WScript.Arguments 
	For I = 0 to objArgs.Count - 1 
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		If objFSO.FolderExists(objArgs(I)) Then
			Set objFolder = objFSO.GetFolder(objArgs(I))
			'Can't set the parent folder from right-click initiated script
			'If the first item is a folder:
			'Can find the parent folder by string replacement (no parent item)
			'In the full path, replace the actual folder name with Nothing
			'Leaves the parent path
			If I = 0 Then
				Parent_Folder = Replace (objArgs(I), objFolder.Name, "")
				Archive_Folder = Parent_Folder & Base_Folder & "\"
				Destination_Folder = Archive_Folder & Folder_Detail & "\"
			End If
			from_this_folder = objArgs(I)
			to_this_folder = Destination_Folder & objFolder.Name
			'Make sure folders exists before attempting to copy
			If Not (objFSO.FolderExists(Archive_Folder)) Then
				objFSO.CreateFolder(Archive_Folder)
			End If
			If Not (objFSO.FolderExists(Destination_Folder)) Then
				objFSO.CreateFolder(Destination_Folder)
			End If
			objFSO.CopyFolder from_this_folder , to_this_folder , OverWrite
		Else
			Set objFile = objFSO.GetFile(objArgs(I))
			'Can't set the parent folder from right-click initiated script
			'If the first item is a file:
			'Can find the parent folder by inspection
			If I = 0 Then
				Parent_Folder = objFSO.GetParentFolderName(objFile)
				Archive_Folder = Parent_Folder & "\" & Base_Folder & "\"
				Destination_Folder = Archive_Folder & Folder_Detail & "\"
			End If
			from_this_file = objArgs(I)
			to_this_file = Destination_Folder & objFSO.GetFileName(objFile)
			'Make sure folders exists before attempting to copy
			If Not (objFSO.FolderExists(Archive_Folder)) Then
				objFSO.CreateFolder(Archive_Folder)
			End If
			If Not (objFSO.FolderExists(Destination_Folder)) Then
				objFSO.CreateFolder(Destination_Folder)
			End If
			objFSO.CopyFile from_this_file , to_this_file, OverWrite
		End If
 Next
 
' **************************************************************************************
' NAME:			GetUserName
' PURPOSE:		Gets the Name of the User Currently Logged In
' **************************************************************************************
Function GetUserName(UserName)
	Dim objNetwork
	Set objNetwork = CreateObject("WScript.Network")
	UserName = objNetwork.UserName
End Function
 
' **************************************************************************************
' NAME:			GetDateTime
' PURPOSE:		Gets the current date and time
' **************************************************************************************
Function GetDateTime(DateTime)
	Dim sDate
	Dim sTime
	Dim iYear
	Dim iMonth
	Dim iDay
	Dim iHour
	Dim iMin
	Dim iSec
' Get the date
	iYear = Year(now)
	iMonth = Month(now)
	iDay = Day(now)
	iHour = Hour(now)
	iMin = Minute(now)
	iSec  = Second(now)
' Format date and time in the format YYYYMMDDTHHMMSS
' Year will always be in the correct format
	sDate = iYear
' Format the month
	sDate = sDate & Right("0" & iMonth, 2)
' Format the day
	sDate = sDate & Right("0" & iDay, 2)
' Format the hour
	sTime = Right("0" & iHour, 2)
' Format the minutes
	sTime = sTime & Right("0" & iMin, 2)
' Format the seconds
	sTime = sTime & Right("0" & iSec, 2)
'Concatenate date and time in the required format
	DateTime = sDate & "T" & sTime
End Function