Attribute VB_Name = "modAppCom"

Option Explicit
Option Compare Binary
Option Base 0

Private Type COPYDATASTRUCT
     dwData As Long
     cbData As Long
     lpData As Long
End Type

Private Const GWL_WNDPROC = (-4)
Private Const WM_COPYDATA = &H4A
Private Const WM_COMMAND = &H111
Private Const WM_HOTKEY = &H312

Private Const WM_APPCOMMAND As Integer = 793 'Monitor Multimedia events
Private Const WM_SYSCOMMAND = &H112 'Monitor For close/kill

Global lpPrevWndProc As Long
Global gHW As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function CallWindowProc Lib "User32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWND As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWND As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWND As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Sub AppCom_Hook(hWND As Long)
On Error GoTo AppCom_Hook_err

gHW = hWND
lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, AddressOf WindowProc)

Exit Sub
AppCom_Hook_err:
Main_Err "AppCom_Hook_err."
Err.Clear
End Sub

Public Sub AppCom_Unhook()
On Error GoTo AppCom_Unhook_err

Dim temp As Long
temp = SetWindowLong(gHW, GWL_WNDPROC, lpPrevWndProc)

Exit Sub
AppCom_Unhook_err:
Main_Err "AppCom_Unhook_err."
Err.Clear
End Sub

Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
If uMsg = WM_COPYDATA Then 'internal messages
    ProcessMessage lParam
ElseIf uMsg = WM_COMMAND Then 'commands
    ProcessCommand wParam
ElseIf uMsg = WM_HOTKEY Then
    frmMain.wMsgProc_Hotkey lParam
ElseIf uMsg = &H20A Then 'mouse wheel events
    frmMain.wMsgProc_MouseWheel wParam
ElseIf uMsg = WM_APPCOMMAND Then
    ProcessMmCommand lParam
End If
WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
End Function

Sub ProcessCommand(wParam As Long)
'SendMessage lhwnd,WM_COMMAND, messagelong,0

Select Case wParam
    Case AppCmd_PlayPause
        frmMain.cmdPlayState_Click 1
    
    Case AppCmd_Stop
        frmMain.cmdPlayState_Click 0
    
    Case AppCmd_Next
        frmMain.pb_NextFile
    
End Select
End Sub

Sub ProcessMmCommand(lParam As Long)
Select Case lParam
    Case MMkey_Play, MMkey_Play2, MMkey_Pause
        frmMain.cmdPlayState_Click 1
    
    Case MMkey_Stop
        frmMain.cmdPlayState_Click 0
    
    Case MMkey_Prev_Item, MMkey_Prev_Track, MMkey_Prev2
        
    
    Case MMkey_Next_Item, MMkey_Next_Track, MMkey_Next2
        frmMain.pb_NextFile
    
    Case Else
        Debug.Print "unknown WM_APPCOMMAND:" & lParam
    
End Select
End Sub

Sub ProcessMessage(lParam As Long)
On Error GoTo ProcessMessage_err:

Dim cds As COPYDATASTRUCT, buf(1 To 255) As Byte, a$, _
    sPrefix As String, sBody As String, c As Long

CopyMemory cds, ByVal lParam, Len(cds)
Select Case cds.dwData
    Case 3
        CopyMemory buf(1), ByVal cds.lpData, cds.cbData
        a$ = StrConv(buf, vbUnicode)
        a$ = Left$(a$, InStr(1, a$, Chr$(0)) - 1)
        
        If a$ = "" Then Exit Sub
        
        sPrefix = Mid$(a$, 1, 1)
        If Len(a$) > 1 Then sBody = Mid$(a$, 2)
        
        Select Case sPrefix
            Case "p" 'play / pause
                frmMain.cmdPlayState_Click 1
            
            Case "a" 'set focus
                On Error Resume Next
                'MsgBox "poke recieved!"
                If frmMain.m_bInTray Then
                    frmMain.tray_Rest
                Else
                    frmMain.Show
                End If
                On Error GoTo ProcessMessage_err
            
            Case "b" 'file
                'MsgBox "file recieved '" & sBody & "'."
                If fsoMain.FileExists(sBody) Then
                    If IsMediaFile(sBody) Then
                        frmMain.cue_AddTo -1, sBody
                    End If
                End If
            
            Case "x" 'request for playback progress
                'todo: clean up this code
                Dim cdsSend As COPYDATASTRUCT, bufSend(1 To 255) As Byte, m As String, lH As Long
                lH = Val(sBody)
                'Debug.Print "request from " & lH
                
                If lH <> 0 Then 'should contain hwnd
                    m = "X" & Trim$(Str$(frmMain.sldPlay.Value))
                    
                    Call CopyMemory(bufSend(1), ByVal m, Len(m))
                    cdsSend.dwData = 3
                    cdsSend.cbData = Len(m) + 1
                    cdsSend.lpData = VarPtr(bufSend(1))
                    SendMessage lH, WM_COPYDATA, 0, cdsSend
                End If
                
        End Select
End Select

Exit Sub
ProcessMessage_err:
Main_Err "ProcessMessage_err."
Err.Clear
End Sub
