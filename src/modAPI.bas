Attribute VB_Name = "modAPI"

Option Explicit
Option Compare Binary
Option Base 0

'cpu timer=====================================================================
Public Declare Function GetTickCount Lib "kernel32" () As Long

'window api====================================================================
Public Declare Function GetWindowPlacement Lib "user32.dll" (ByVal hWND As Long, ByRef lpwndpl As WINDOWPLACEMENT) As Long
Public Type WINDOWPLACEMENT
    Length As Long
    flags As Long
    showCmd As Long
    ptMinPosition As POINTAPI
    ptMaxPosition As POINTAPI
    rcNormalPosition As RECT
End Type
'mouse stuff ==================================================================
Public Declare Function LoadCursor Lib "User32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Public Declare Function SetCursor Lib "User32" (ByVal hCursor As Long) As Long
Public Declare Function GetCursor Lib "User32" () As Long
Public Const IDC_HAND = 32649&
Public Const IDC_ARROW = 32512&

Public Declare Function GetCursorPos Lib "User32" (lpPoint As POINTAPI) As Long
'==============================================================================
Public Sub ShellExecuteURL(ByVal sURL As String)
On Error GoTo ShellExecuteURL_Err
Dim lRet As Long

lRet = ShellExecute(0, "open", sURL, "", vbNull, 1)
If lRet <> 0 Then Exit Sub
ShellExecuteURL_Err:
Main_Err "Error Opening Link to " & sURL & "."
Err.Clear
End Sub
