Attribute VB_Name = "modMain"

Option Explicit
Option Compare Binary
Option Base 0

Public pref_NoHook As Boolean
Public pref_Debug As Boolean

Public cMedia As clsMedia
Public m_cM As New cMonitors
'

Sub Main()
On Error GoTo MainError

Dim e As String, _
    s As String, arr() As String, i As Long, _
    cFiles As Collection, sNoFileCmd As String

Set fsoMain = New FileSystemObject

pref_NoHook = False
pref_Debug = False

e = "processing command '" & Command & "'."

'valid switches:
'   -p = play / pause
'   --nohook = no window hooking
'   --debug = debug menu visible

sNoFileCmd = "a"
Set cFiles = New Collection

'play / pause
s = Trim$(Command)
If InStr(1, LCase$(Command), "-p") > 0 Then
    sNoFileCmd = "p"
End If

'disable system hooks
If InStr(1, LCase$(Command), "--nohook") > 0 Then
    pref_NoHook = True
End If

'enable debug menu
If InStr(1, LCase$(Command), "--debug") > 0 Then
    pref_Debug = True
End If

'debug logging
If pref_Debug Then
    file_DebugINI = App.Path & "\debug.ini"
    WriteToIni "main", "cmd", Command, file_DebugINI
End If

'open files
If Len(Replace(s, Chr(34), "")) > 0 Then
    'quotes only used if spaces in file names
    If InStr(1, s, Chr(34)) > 0 Then
        arr = Split(s, Chr(34))
    Else
        arr = Split(s, " ")
    End If
    
    For i = LBound(arr) To UBound(arr)
        arr(i) = Trim$(arr(i))
        If Len(arr(i)) > 0 Then
            If fsoMain.FileExists(arr(i)) Then
                cFiles.Add arr(i)
            End If
        End If
    Next i
End If

e = "checking for prev inst;" & cFiles.count & "."
If App.PrevInstance Then
    If cFiles.count < 1 Then
        AppCom_Send_Str sNoFileCmd '"a"
    Else
        For i = 1 To cFiles.count
            AppCom_Send_Str "b" & cFiles(i)
        Next i
    End If
    
    End 'todo: this line should probabbly be replaced with something more elegant.
End If

e = "init. unicode drawing."
'for unicode drawing
VerInitialise


e = "init. sqlite db."
'connected to the db
mdb_connect

e = "init. media player."
'init media player
Set cMedia = New clsMedia

e = "init. general settings."
'general settings
file_INI = App.Path & "\terra.ini"
folder_Playlists = App.Path & "\playlists\"
If Not fsoMain.FolderExists(folder_Playlists) Then MkDir folder_Playlists
file_PlayListIndex = folder_Playlists & "index.txt"

e = "loading frmMain."
Load frmMain

e = "adding command line files;" & cFiles.count & "."
If cFiles.count > 0 Then
    For i = 1 To cFiles.count
        frmMain.cue_AddTo -1, cFiles(i)
    Next i
End If

e = "showing frmMain."
frmMain.Show

e = "showing minivid."
'minivid
If GetFromIniEx("minivid", "active", "0", file_INI) = "1" Then
    frmMain.cmdVidMinivid_Click
End If

'done loading everything; allow pref auto saves.
pref_CanSave = True

'these 2 lines gen a run-time error
'On Error GoTo 0
'MsgBox (1 / 0)

Exit Sub
MainError:
Main_Err "sub Main - " & e
If pref_Debug Then
    WriteToIni "main", "error", e, file_DebugINI
End If

err.Clear
End Sub

'types:
'0-critical (an error that is my fault.)
'1-not critical (an error that is not my fault.  e.g. file not found.)
Sub Main_Err(e As String, Optional iType As Integer = 0)
Dim a As String, sSysErr As String
sSysErr = err.Number & "; " & err.Description & "."

On Error GoTo Main_Err_err

Select Case iType
    Case 0
        a = "//terra has encountered a critical error." & vbNewLine & vbNewLine & _
            "if this is the first time you have seen this message you may be " & _
            "able to contine using terra by pressing the 'continue' button.  if not " & _
            "then you may have to end terra by pressing the 'force end' button.  this " & _
            "will result in the loss of all unsaved data." & vbNewLine & vbNewLine & _
            "if this problem persists and you require supprt please visit the wiki " & _
            "at terra.aefaradien.net or contact terra@aefaradien.net quoting the " & _
            "following data and giving a brief description of what terra was doing " & _
            "at the time." & vbNewLine & vbNewLine & _
            "internal description: " & vbNewLine & _
            e & vbNewLine & vbNewLine & _
            "system description: " & vbNewLine & _
            sSysErr & vbNewLine & vbNewLine & _
            "application version: " & App.Major & "." & App.Minor & "(" & App.Revision & ")"
    
    Case 1
        a = "//terra has encountered a non-critical error." & vbNewLine & vbNewLine & _
            "internal description: " & vbNewLine & _
            e & vbNewLine & vbNewLine & _
            "if you need help dealing with this error, please visit the wiki at " & _
            "terra.aefaradien.net for further information."
    
End Select

Dim frmE As New frmErr
frmE.SetType iType
frmE.txtOutput.Text = a
frmE.Show 1
Unload frmE
Set frmE = Nothing

Exit Sub
Main_Err_err:
MsgBox "unable to show error window." & vbNewLine & a
err.Clear
End Sub

Sub Main_About()
MsgBox "//terra is © copyright Alex Hutter (aefaradien) 2006 to 2008." & vbNewLine & _
    "this is version: " & App.Major & "." & right$("0" & App.Minor, 2) & _
    " (build " & App.Revision & ")." & vbNewLine & _
    "this software is freeware." & vbNewLine & _
    "for full details and references, please refer to readme.txt." & vbNewLine & _
    vbNewLine & _
    "for support and other info please see the wiki at terra.aefaradien.net."
End Sub

Sub main_end(Optional bForceEnd As Boolean = False)
On Error Resume Next

Dim f As Form

For Each f In Forms
    Unload f 'destroy GUI component
    Set f = Nothing 'destroy code component
Next f

Set cMedia = Nothing
mdb_disconnect
gdi_Main_DeleteObjects

If bForceEnd Then End
End Sub
