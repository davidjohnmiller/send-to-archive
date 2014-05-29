# send-to-archive


## Intended Use

1. Select the Required File/s and/or Folder/s
2. Right-Click > Send-To > Archive.vbs
3. Selection Copied to "archive/dated-subfolder"


## Intended Platform

Uses the built-in right-click send-to capability in Windows 7/8.
Script written in vbs so that no other files/dependencies are required.


## Installation

1. Put the archive.vbs file into your Windows Send-To folder

### Tell Me More...

For Windows 7/8 the Send-To folder is located at:
C:\Users\{your-username}\AppData\Roaming\Microsoft\Windows\SendTo

You may have to show hidden files/folders to see the folders.

The easiest way to "get" this script is by selecting the "Download Zip" on the lower right of the github page.
Alternatively you could just [click here][1].

The better way is to fork the repo to get your own version controlled copy to play with.
If you make some useful changes let me know.


## Customisation

### Change Folder Name
The folder name mimics the name of the vbs file, therefore if you want to change the name of the main archive folder just change the name of the script.
E.g. Change archive.vbs to something like backup.vbs and the path will be "backup/dated-subfolder".

### Include Current Username in the SubFolder Name
The name of the user can be added to the folder name by removing the comment in the script.
Having this on may be useful for a multi-user setting, no so much for a single user.

### Remove Option to Ask/Add a Comment to the Subfolder Name
The option to ask/add a comment can be commented out without breaking anything.

### Other Modifications?
Feel free to go crazy modifying the script. If you make something useful let me know.

[1]: https://github.com/chrisfreiberg/send-to-archive/archive/master.zip