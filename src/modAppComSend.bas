Attribute VB_Name = "modAppComSend"

Option Explicit
Option Compare Binary
Option Base 0

Private Type COPYDATASTRUCT
     dwData As Long
     cbData As Long
     lpData As Long
End Type

Private Const WM_COPYDATA = &H4A
Private Const WM_COMMAND = &H111

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWND As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function EnumWindows Lib "User32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowText Lib "User32" Alias "GetWindowTextA" (ByVal hWND As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "User32" Alias "GetWindowTextLengthA" (ByVal hWND As Long) As Long

'because VB6 sucks and does not allow custom window names,
'we must resort to identifing the window by its caption.  *sob*
Public Const sMainWindowID As String = "//terra"

'tracking of finding the window
Private Const m_MaxSearchTime   As Long = 1000
Public m_TerraWinH              As Long
Private m_bTerraWinFound        As Boolean

'internal commands
Public Const AppCmd_PlayPause   As Long = 1000
Public Const AppCmd_Stop        As Long = 1001
Public Const AppCmd_Next        As Long = 1002
'

'################################################
'#  find main window long  ######################
'################################################

Public Sub AppCom_WinH_Update()
Dim lTimer As Long

m_TerraWinH = 0
m_bTerraWinFound = False
EnumWindows AddressOf AppCom_WinH_CallBackProc, 1

lTimer = GetTickCount
Do Until m_bTerraWinFound
    DoEvents
    If GetTickCount - lTimer >= m_MaxSearchTime Then Exit Do
Loop
End Sub

Public Function AppCom_WinH_CallBackProc(ByVal hWND As Long, ByVal lParam As Long) As Long
If InStr(1, WindowTitleFromHwnd(hWND), sMainWindowID) > 0 Then
    m_TerraWinH = hWND
    m_bTerraWinFound = True
    AppCom_WinH_CallBackProc = False
Else
    AppCom_WinH_CallBackProc = True
End If
End Function

'################################################
'#  send message  ###############################
'################################################

Public Sub AppCom_Send_Str(m As String)
AppCom_WinH_Update
If m_TerraWinH = 0 Then Exit Sub

Dim cds As COPYDATASTRUCT, buf(1 To 255) As Byte

Call CopyMemory(buf(1), ByVal m, Len(m))
cds.dwData = 3
cds.cbData = Len(m) + 1
cds.lpData = VarPtr(buf(1))
SendMessage m_TerraWinH, WM_COPYDATA, 0, cds
End Sub

Public Sub AppCom_Send_Long(l As Long)
AppCom_WinH_Update
If m_TerraWinH = 0 Then Exit Sub

SendMessage m_TerraWinH, WM_COMMAND, l, 0
End Sub

'################################################
'#  extra functions  ############################
'################################################

Public Function WindowTitleFromHwnd(ByVal lhWnd As Long) As String
Dim lLen As Long, sBuf As String

lLen = GetWindowTextLength(lhWnd)
If (lLen > 0) Then
    sBuf = String$(lLen + 1, 0)
    lLen = GetWindowText(lhWnd, sBuf, lLen + 1)
    
    WindowTitleFromHwnd = Left$(sBuf, lLen)
End If
End Function
