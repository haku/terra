VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSysTray"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' bits of this probably came from an example, but has changed substantially as it has
' passed through numerous projects.

Option Explicit
Option Compare Binary
Option Base 0

Private Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowRect Lib "User32" (ByVal hWND As Long, lpRect As RECT) As Long
Private Declare Function GetClassName Lib "User32" Alias "GetClassNameA" (ByVal hWND As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindow Lib "User32" (ByVal hWND As Long, ByVal wCmd As Long) As Long
Private Declare Function DrawAnimatedRects Lib "User32" (ByVal hWND As Long, ByVal idAni As Long, lprcFrom As RECT, lprcTo As RECT) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private Declare Function SetWindowPos Lib "User32" (ByVal hWND As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Type NOTIFYICONDATA
    cbSize As Long
    hWND As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Type RECT
    Left As Long
    Top As Long
    right As Long
    bottom As Long
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Const GW_CHILD = 5
Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDLAST = 1
Private Const GW_HWNDNEXT = 2
Private Const GW_HWNDPREV = 3
Private Const GWL_WNDPROC = (-4)
Private Const IDANI_OPEN = &H1
Private Const IDANI_CLOSE = &H2
Private Const IDANI_CAPTION = &H3
Private Const WM_USER = &H400

Private stData As NOTIFYICONDATA
Private LastTrayHWND As Long

Private Sub SetLastTray()
LastTrayHWND = FindWindow("Shell_TrayWnd", vbNullString)
End Sub

Public Sub CheckStillInTray()
Dim tmp As Long

tmp = FindWindow("Shell_TrayWnd", vbNullString)

If (tmp <> LastTrayHWND) And (tmp > 0) Then
    RemoveIcon
    Shell_NotifyIcon 0, stData
    SetLastTray
End If
End Sub

Public Sub PlaceIcon(hWND As Long, Icon As Object, Caption As String)
stData.cbSize = Len(stData)
stData.hWND = hWND
stData.uId = vbNull
stData.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
stData.uCallBackMessage = WM_MOUSEMOVE
stData.hIcon = Icon
stData.szTip = Caption + Chr(0)

Shell_NotifyIcon 0, stData
End Sub

Public Sub ChangeIcon(Icon As Object)
stData.hIcon = Icon
Shell_NotifyIcon 1, stData
End Sub

Public Sub ChangeCaption(Caption As String)
stData.szTip = Caption + Chr(0)
Shell_NotifyIcon 1, stData
End Sub

Public Sub RemoveIcon()
Shell_NotifyIcon 2, stData
End Sub

Public Function FlyingWindow(Frm As Form, OpenClose As Integer)
Dim rctFrom As RECT, rctTo As RECT
Dim lngTrayHand As Long
Dim lngStartMenuHand As Long, lngChildHand As Long
Dim strClass As String * 255

lngStartMenuHand = FindWindow("Shell_TrayWnd", vbNullString)
lngChildHand = GetWindow(lngStartMenuHand, GW_CHILD)

Do
    GetClassName lngChildHand, strClass, Len(strClass)
    If InStr(1, strClass, "TrayNotifyWnd") Then
        lngTrayHand = lngChildHand
        Exit Do
    End If
    lngChildHand = GetWindow(lngChildHand, GW_HWNDNEXT)
Loop

GetWindowRect Frm.hWND, rctTo
GetWindowRect lngTrayHand, rctFrom

Select Case OpenClose
    Case 0: DrawAnimatedRects Frm.hWND, IDANI_OPEN Or IDANI_CAPTION, rctFrom, rctTo
    Case 1: DrawAnimatedRects Frm.hWND, IDANI_CLOSE Or IDANI_CAPTION, rctTo, rctFrom
End Select
End Function

Private Sub Class_Initialize()
SetLastTray
End Sub
