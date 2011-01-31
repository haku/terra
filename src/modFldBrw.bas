Attribute VB_Name = "modFldBrw"
'mostly taken from some example i found ages ago.
'this code has passed through many projects.

Option Explicit
Option Compare Binary
Option Base 0

Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Private Const DRIVE_REMOVABLE = 2
Private Const DRIVE_FIXED = 3
Private Const DRIVE_REMOTE = 4
Private Const DRIVE_CDROM = 5
Private Const DRIVE_RAMDISK = 6

' Structures
Private Type SHFILEOPSTRUCT
    hWND As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As Long
End Type

' API
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

' Contants
Private Const FO_DELETE = &H3
Private Const FOF_ALLOWUNDO = &H40
Private Const FOF_NOCONFIRMATION = &H10 ' Responds with "yes to all" for any dialog box that is displayed.
Private Const FOF_SILENT = &H4

Private Const BIF_RETURNONLYFSDIRS As Long = &H1
Private Const BIF_NEWDIALOGSTYLE As Long = &H40
Private Const BIF_EDITBOX As Long = &H10
Private Const BIF_SHAREABLE As Long = &H8000
Private Const BIF_STATUSTEXT As Long = &H4

Public Type BROWSEINFOTYPE
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBROWSEINFOTYPE As BROWSEINFOTYPE) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWND As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

Private Const WM_USER = &H400

Private Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
Private Const BFFM_SETSELECTIONW As Long = (WM_USER + 103)
Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Const LPTR = (&H0 Or &H40)
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWND As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Function BrowseCallbackProcStr(ByVal hWND As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
If uMsg = 1 Then
    Call SendMessage(hWND, BFFM_SETSELECTIONA, True, ByVal lpData)
End If
End Function

Private Function FunctionPointer(FunctionAddress As Long) As Long
FunctionPointer = FunctionAddress
End Function

Public Function BrowseForFolder(startPath As String, Optional ByVal strTitle As String) As String
Dim Browse_for_folder As BROWSEINFOTYPE
Dim itemID As Long
Dim selectedPathPointer As Long
Dim tmpPath As String * 256
Dim selectedPath As String

selectedPath = startPath ' Take the selected path from startPath
If Len(selectedPath) > 0 Then
    If Not right$(selectedPath, 1) <> "\" Then selectedPath = Left$(selectedPath, Len(selectedPath) - 1) ' Remove "\" if the user added
End If

With Browse_for_folder
    .hOwner = 0 ' Window Handle
    If Len(strTitle) > 0 Then
        .lpszTitle = strTitle ' Dialog Title
    Else
        .lpszTitle = "Select A Folder"
    End If
    .lpfn = FunctionPointer(AddressOf BrowseCallbackProcStr) ' Dialog callback function that preselectes the folder specified
    selectedPathPointer = LocalAlloc(LPTR, Len(selectedPath) + 1) ' Allocate a string
    CopyMemory ByVal selectedPathPointer, ByVal selectedPath, Len(selectedPath) + 1 ' Copy the path to the string
    .lParam = selectedPathPointer ' The folder to preselect
    
    .ulFlags = BIF_RETURNONLYFSDIRS + BIF_NEWDIALOGSTYLE
End With
itemID = SHBrowseForFolder(Browse_for_folder) ' Execute the BrowseForFolder API
If itemID Then
    If SHGetPathFromIDList(itemID, tmpPath) Then ' Get the path for the selected folder in the dialog
        BrowseForFolder = Left$(tmpPath, InStr(tmpPath, vbNullChar) - 1)  ' Take only the path without the nulls
    End If
    Call CoTaskMemFree(itemID) ' Free the itemID
End If
Call LocalFree(selectedPathPointer) ' Free the string from the memory
End Function
