Attribute VB_Name = "modGlobalHotkey"

Option Explicit
Option Compare Binary
Option Base 0

Public Declare Function RegisterHotKey Lib "User32" ( _
    ByVal hWND As Long, _
    ByVal id As Long, _
    ByVal fsModifiers As Long, _
    ByVal vk As Long _
    ) As Long

Public Declare Function UnregisterHotKey Lib "User32" ( _
    ByVal hWND As Long, _
    ByVal id As Long _
    ) As Long

Public Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" ( _
    ByVal nAtom As Integer, _
    ByVal lpBuffer As String, _
    ByVal nSize As Long _
    ) As Long

Public Declare Function GlobalAddAtom Lib "kernel32.dll" Alias "GlobalAddAtomA" ( _
    ByVal lpString As String _
    ) As Integer

Public Declare Function GlobalDeleteAtom Lib "kernel32" ( _
    ByVal nAtom As Integer _
    ) As Integer

Public Declare Function PeekMessage Lib "User32" Alias "PeekMessageA" ( _
    lpMsg As Msg, _
    ByVal hWND As Long, _
    ByVal wMsgFilterMin As Long, _
    ByVal wMsgFilterMax As Long, _
    ByVal wRemoveMsg As Long _
    ) As Long

Public Declare Function WaitMessage Lib "User32" () As Long

'Public Type POINTAPI
'    x As Long
'    y As Long
'End Type

Public Type Msg
    hWND As Long
    Message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Public Const MOD_ALT = &H1
Public Const MOD_CONTROL = &H2
Public Const MOD_SHIFT = &H4
Public Const MOD_WIN = &H8
Public Const PM_REMOVE = &H1
'Public Const WM_HOTKEY = &H312

Public m_HotKeyAtomNameCount As Long

Public Function HiWord(dw As Long) As Integer
If dw And &H80000000 Then
    HiWord = (dw \ 65535) - 1
Else
    HiWord = dw \ 65535
End If
End Function

Public Function LoWord(dw As Long) As Integer
If dw And &H8000& Then
    LoWord = &H8000 Or (dw And &H7FFF&)
Else
    LoWord = dw And &HFFFF&
End If
End Function
