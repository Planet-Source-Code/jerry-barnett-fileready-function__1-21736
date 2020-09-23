<div align="center">

## FileReady Function


</div>

### Description

Here is a useful function I created to solve a

problem I was having in processing files for

one of my applications.

This function works better than trying to use

the VB 'OPEN' command because it will always

return the correct state of the file (even on a

file that is being FTP'd at the time of the

test.) See remarks in the FileReady function on

it's use, as well as a sample at the end of this

note.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jerry Barnett](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jerry-barnett.md)
**Level**          |Intermediate
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jerry-barnett-fileready-function__1-21736/archive/master.zip)

### API Declarations

```
Public Const SHARE_EXCLUSIVE = &H0
Public Const INVALID_HANDLE_VALUE = -1
Public Const ERROR_ALREADY_EXISTS = 183&
Public Const OPEN_EXISTING = 3
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const GENERIC_READ = &H80000000
Public Type SECURITY_ATTRIBUTES
 nLength As Long
 lpSecurityDescriptor As Long
 bInheritHandle As Long
End Type
Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Function CreateFile _
 Lib "kernel32" Alias "CreateFileA" _
 (ByVal lpFileName As String, _
 ByVal dwDesiredAccess As Long, _
 ByVal dwShareMode As Long, _
 lpSecurityAttributes As SECURITY_ATTRIBUTES, _
 ByVal dwCreationDisposition As Long, _
 ByVal dwFlagsAndAttributes As Long, _
 ByVal hTemplateFile As Long) As Long
Public Declare Function CloseHandle _
 Lib "kernel32" (ByVal hObject As Long) As Long
```


### Source Code

```
Private Function FileReady(strFileName As String) As Boolean
'***********************************************
' * Programmer Name : Jerry Barnett
' * Procedure Name : FileReady
' * Parameters : strFileName As String - 
' * Filename to check
' * Returns : TRUE - if the file exists and
' *   is not in use
' *   by any other process.
' *  FALSE - if the file is in use by
' *   another process, or does
' *   not exist.
'***********************************************
' * Comments : This function checks to
' * see if a file is ready for use. It
' * tries to open the file for
' * exclusive use.
' * 
' * NOTE - An example of where this
' * function would be used is as
' * follows:
' * You have an application that needs
' * to process files as they are
' * created in a directory. However
' * since they could be large files
' * you don't want to start
' * processing the file before it
' * is completely copied (or FTP'd)
' * into the directory. This function
' * will determine if the copy or FTP
' * is complete so that you can then
' * open the file for processing.
'************************************
' * The following Constants and
' * Declares must be placed in the
' * Module DECLARES section.
'************************************
' *
' * Public Const SHARE_EXCLUSIVE = &H0
' * Public Const INVALID_HANDLE_VALUE = -1
' * Public Const ERROR_ALREADY_EXISTS = 183&
' * Public Const OPEN_EXISTING = 3
' * Public Const FILE_ATTRIBUTE_NORMAL = &H80
' * Public Const GENERIC_READ = &H80000000
' *
' * Public Type SECURITY_ATTRIBUTES
' *   nLength As Long
' *   lpSecurityDescriptor As Long
' *   bInheritHandle As Long
' * End Type
' *
' * Public Declare Function GetLastError _
' * Lib "kernel32" () As Long
' *
' * Public Declare Function CreateFile Lib _
' * "kernel32" Alias "CreateFileA" _
' * (ByVal lpFileName As String, _
' *  ByVal dwDesiredAccess As Long, _
' *  ByVal dwShareMode As Long, _
' *  pSecurityAttributes As SECURITY_ATTRIBUTES, _
' *  ByVal dwCreationDisposition As Long, _
' *  ByVal dwFlagsAndAttributes As Long, _
' *  ByVal hTemplateFile As Long) As Long
' *
' * Public Declare Function CloseHandle Lib _
' * "kernel32" (ByVal hObject As Long) As Long
' *
'************************************************
 Dim lReturnCode As Long
 Dim typAtrib As SECURITY_ATTRIBUTES
 ' Try to open the file for exclusive use
 lReturnCode = CreateFile(strFileName, _
    GENERIC_READ, _
    SHARE_EXCLUSIVE, _
    typAtrib, _
    OPEN_EXISTING, _
    FILE_ATTRIBUTE_NORMAL, 0)
 If lReturnCode = INVALID_HANDLE_VALUE Then
 ' Failed exclusive use of file (File not ready)
 FileReady = False
 Exit Function ' Exit function
 End If
 ' File exists and is ready, so close the file
 lReturnCode = CloseHandle(lReturnCode)
 ' Return True (File is Ready)
 FileReady = True
End Function
'************************************************
' A Sample of how to use this function:
Private Sub Main()
 Dim lCount as Long
 Dim Const MAXCOUNT = 5 ' Actually this would be in
 ' the module declares section
 Do While Not FileReady("FileToCheckFor.txt") Then
 lCount = lCount + 1
 ' ...... wait some predetermined amount
 ' of time .....
 If lCount = MAXCOUNT Then
  Msg "File Not Ready! Maximum try's exceeded!"
  End
 End If
 Loop
 Msg "File can now pe processed!"
 ' .... Do your processing code to work
 ' with the file.
End Sub
```

