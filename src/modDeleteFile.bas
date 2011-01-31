Attribute VB_Name = "modDeleteFile"

Option Explicit
Option Compare Binary
Option Base 0

Private Declare Function SHFileOperation Lib "shell32.dll" (lpFileOp As SHFILEOPSTRUCT) As Long
    Private Const FO_DELETE = &H3
    Private Const FOF_ALLOWUNDO = &H40
    Private Const FOF_NOCONFIRMATION = &H10


Private Type SHFILEOPSTRUCT
    hWND As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As String
End Type

Public Function MoveFileToRecycle(f As String) As Boolean
On Error GoTo MoveFileToRecycle_err

Dim sfoStruct As SHFILEOPSTRUCT, lRet As Long

With sfoStruct
    '.hwnd = lHwnd
    .pFrom = f
    .wFunc = FO_DELETE
    .fFlags = FOF_ALLOWUNDO Or FOF_NOCONFIRMATION
End With

lRet = SHFileOperation(sfoStruct)

MoveFileToRecycle = Not (lRet > 0)

Exit Function
MoveFileToRecycle_err:
Main_Err "MoveFileToRecycle_err."
Err.Clear
End Function
