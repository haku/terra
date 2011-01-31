VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHotKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Option Compare Binary
Option Base 0

Private m_lHkAtom As Long, m_bHkActive As Boolean, _
    m_lhWnd As Long, m_lKey As Long, m_lShift As Long
'

Public Function SetKey(lhWnd As Long, lKey As Long, lShift As Long) As Boolean
SetKey = False

Dim lRet As Long

UnregHotkey

lRet = RegisterHotKey(lhWnd, m_lHkAtom, lShift, lKey)

If lRet = 0 Then
    Debug.Print "error: unable to register hot key."
Else
    'Debug.Print "hotkey registered; '" & lRet; "'."
    m_lhWnd = lhWnd
    m_bHkActive = True
    SetKey = True
End If

m_lKey = lKey
m_lShift = lShift
End Function

'accepts date in the form genereated by GetKeyDataString
Public Function SetKeyFromData(lhWnd As Long, sData As String)
Dim lKey As Long, lShift As Long, arrlKeys() As String

arrlKeys = Split(sData, "|")

lKey = Val(arrlKeys(0))
lShift = _
    IIf(Val(arrlKeys(1)) > 0, 1, 0) * MOD_SHIFT + _
    IIf(Val(arrlKeys(2)) > 0, 1, 0) * MOD_CONTROL + _
    IIf(Val(arrlKeys(3)) > 0, 1, 0) * MOD_ALT

SetKey lhWnd, lKey, lShift
End Function

Public Sub ClearKey()
UnregHotkey
End Sub

Public Sub ClearData()
m_lKey = 0
m_lShift = 0
End Sub

Public Function IsKey(lParam As Long) As Boolean
IsKey = (HiWord(lParam) = m_lKey) And (LoWord(lParam) = m_lShift) And IsSet
End Function

Public Function IsSet() As Boolean
IsSet = m_bHkActive
End Function

Public Function HasData() As Boolean
HasData = Not (m_lKey = 0 And m_lShift = 0)
End Function

Private Sub UnregHotkey()
Dim lRet As Long

If m_bHkActive Then
    lRet = UnregisterHotKey(m_lhWnd, m_lHkAtom)
    
    If lRet = 0 Then
        Debug.Print "error: unable to unregister hot key." '; '" & GetErrDes & "'."
    Else
        'Debug.Print "hotkey '" & m_lHkAtom & "' unregistered."
        m_bHkActive = False
    End If
End If
End Sub

Public Function GetKeyDes() As String
GetKeyDes = _
    IIf(m_lShift And MOD_SHIFT, "shift+", "") & _
    IIf(m_lShift And MOD_CONTROL, "ctrl+", "") & _
    IIf(m_lShift And MOD_ALT, "alt+", "") & _
    Trim$(Str$(m_lKey))
End Function

Public Function GetKeyDataString() As String
If Not HasData Then
    GetKeyDataString = ""
Else
    GetKeyDataString = _
        Trim$(Str$(m_lKey)) & "|" & _
        IIf(m_lShift And MOD_SHIFT, "1", "0") & "|" & _
        IIf(m_lShift And MOD_CONTROL, "1", "0") & "|" & _
        IIf(m_lShift And MOD_ALT, "1", "0") & "|"
End If
End Function

Private Sub Class_Initialize()
If IsEmpty(m_HotKeyAtomNameCount) Then m_HotKeyAtomNameCount = 0
m_lHkAtom = GlobalAddAtom("terrahotkey" & Trim$(Str$(m_HotKeyAtomNameCount)))
m_HotKeyAtomNameCount = m_HotKeyAtomNameCount + 1

If m_lHkAtom = 0 Then
    Debug.Print "error: unable to create hotkey atom."
Else
    Dim a As String
    a = String$(255, 0)
    GlobalGetAtomName m_lHkAtom, a, Len(a)
    'Debug.Print "atom: '" & a & "', '" & m_lHkAtom & "' created."
End If

m_bHkActive = False
End Sub

Private Sub Class_Terminate()
'remove hotkey:
UnregHotkey

'remove atom:
SetLastError ERROR_SUCCESS
GlobalDeleteAtom m_lHkAtom
If GetLastError <> ERROR_SUCCESS Then
    Debug.Print "error: unable to delete atom."
Else
    'Debug.Print "deleted atom '" & m_lHkAtom & "'."
    m_lHkAtom = 0
End If
End Sub
