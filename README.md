Send To Archive
===============

## Intended Use

1. Select the required file/s and/or folder/s
2. Right-click > Send to > Archive.vbs
3. Selection copied to "Archive/dated-subfolder"


## Intended Platform

Uses the built-in right-click send-to capability in Windows 7/8.
Script written in VBS so that no other files/dependencies are required.


## Installation

1. [Download Archive.vbs][1]

### Right-Click

2. Put the Archive.vbs file (or shortcut) into `%APPDATA%\Microsoft\Windows\SendTo`

### Taskbar

2. Put the Archive.vbs file (or shortcut) into a new folder
3. Right-click the taskbar and select Toolbars > New toolbar...
4. Choose the folder from Step 2
5. You can now drag and drop files and folders onto the taskbar shortcut to archive them

## Customisation

### Change Folder Name
The folder name mimics the name of the VBS file, therefore if you want to change the name of the main archive folder just change the name of the script.
For example, change Archive.vbs to something like Backup.vbs and the path will be "Backup/dated-subfolder".

### Include Current Username in the Subfolder Name
The name of the user can be added to the folder name by removing the comment in the script.
Having this on may be useful for a multi-user setting, not so much for a single user.

### Remove Option to Ask/Add a Comment to the Subfolder Name
The option to ask/add a comment can be commented out without breaking anything.

### Other Modifications?
Feel free to go crazy modifying the script. If you make something useful let me know.

[1]: https://github.com/chrisfreiberg/send-to-archive/archive/master.zip
