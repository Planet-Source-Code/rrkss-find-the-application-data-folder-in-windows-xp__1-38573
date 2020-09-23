<div align="center">

## Find the Application Data Folder in Windows XP


</div>

### Description

This article uses the SHGetFolderPath API to find the Application Data Folder In "C:\Documents and Settings\UserName\Application Data\" so you can save program data files and allow limited users in windows xp to use you program easier.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[RRKSS](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/rrkss.md)
**Level**          |Advanced
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/rrkss-find-the-application-data-folder-in-windows-xp__1-38573/archive/master.zip)





### Source Code

Since the SHGetFolderPath API is not in the api viewer here are the declarations you need.<BR><BR>
Declare Function SHGetFolderPath Lib "shell32.dll" Alias "SHGetFolderPathA" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal lpszPath As String) As Long<BR>
Public Const CSIDL_APPDATA = &H1A<BR>
Public Const SHGFP_TYPE_CURRENT = 0<BR>
Public Const SHGFP_TYPE_DEFAULT = 1<BR>
Public DataFolder as String<BR><BR>
Add a module to your project and copy and paste those declarations into the module. Now create a function called GetDataFolder and paste this code into it.<BR><BR>
Public Function GetDataFolder()<BR>
 On Error Goto GenericFolder<BR>
 Dim ReturnVal as long<BR>
 Dim PathName as long<BR>
 pathname = Space(260)<BR>
 retval = SHGetFolderPath(Form1.hWnd, CSIDL_APPDATA, 0, SHGFP_TYPE_CURRENT, pathname)
 pathname = Left(pathname, InStr(pathname, vbNullChar) - 1)<BR>
 DataFolder = PathName<BR>
 exit function<BR>
GenericFolder:<BR>
 'Since Windows XP\2000 is not installed we don't have this api so just use the App.path<BR>
 if err.number = 453 then<BR>
  datafolder = app.path<BR>
 end if<BR>
end function<BR><BR>
With that code you should be able to make your programs more XP compatible and still be able to run it on windows 9x. Please leave your comments and vote.

