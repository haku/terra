Attribute VB_Name = "modGenFnc"

Option Explicit
Option Compare Binary
Option Base 0

Sub AddToLog(txtLog As TextBox, ByVal a As String, Optional bNewLine As Boolean = True)
If bNewLine = True Then
    a = IIf(Len(txtLog.Text) > 0, vbNewLine, "") & Format(Now, "hh:mm:ss") & vbTab & a
End If
txtLog.SelStart = Len(txtLog.Text)
txtLog.SelText = a
txtLog.SelStart = Len(txtLog.Text)
End Sub

Function IsMediaFile(ByVal f As String) As Boolean
Dim arrExts() As String, i As Long, b As Boolean
'arrExts = Split(".mp3|.ogg|.wma|.wmv|.avi|.mpg|.mpeg|.ac3|.mp4|.wav|.ra|.mpga|.mkv|.ogm", "|")
arrExts = Split(file_ext_list, "|")

f = LCase$(f)
b = False

For i = LBound(arrExts) To UBound(arrExts)
    If right$(f, Len(arrExts(i))) = arrExts(i) Then
        'Debug.Assert i < 1
        b = True
        Exit For
    End If
Next i

IsMediaFile = b
End Function

Function ConvertSecToMin(ByVal s As Long) As String
On Error GoTo ConvertSecToMin_err

Dim a As Long

a = (s Mod 60) 'get the seconds
ConvertSecToMin = Trim$(Str$(((s - a) / 60))) & ":" & right$("0" & Trim$(Str$(a)), 2)

Exit Function
ConvertSecToMin_err:
err.Clear
ConvertSecToMin = ""
End Function

Function ConvertSecToHours(ByVal t As Long) As String
On Error GoTo ConvertSecToHours_err

Dim h As Long, m As Long, s As Long

h = t \ 3600
t = t Mod 3600
m = t \ 60
t = t Mod 60
s = t

ConvertSecToHours = _
    IIf(h > 0, Trim$(Str$(h)) & ":" & right$("0" & Trim$(Str$(m)), 2), Trim$(Str$(m))) _
    & ":" & right$("0" & Trim$(Str$(s)), 2)

Exit Function
ConvertSecToHours_err:
err.Clear
ConvertSecToHours = ""
End Function

'Function TrncFilePath(f As String, n As Long, bRev As Boolean) As String
'Dim i As Long, j As Long, X As Long
'
'If bRev Then
'    X = 0
'    For i = 1 To Len(f)
'        If Mid$(f, i, 1) = "\" Then
'            X = X + 1
'            If X >= n Then
'                TrncFilePath = Mid$(f, i + 1)
'                Exit For
'            End If
'        End If
'    Next i
'
'Else
'    X = 1
'    i = -1
'
'    For j = n - 1 To 0 Step -1
'        i = InStrRev(f, "\", IIf(j = n - 1, -1, i - 1))
'        If i > 0 Then X = i Else Exit For
'    Next j
'
'    TrncFilePath = Mid$(f, X + 1)
'End If
'End Function

Function RemExtFromPath(f As String) As String
Dim X As Long
X = InStrRev(f, ".")
If X > 0 Then
    RemExtFromPath = Mid$(f, 1, X - 1)
Else
    RemExtFromPath = f
End If
End Function

Function FileNameFromPath(f As String) As String
FileNameFromPath = Mid$(f, InStrRev(f, "\") + 1)
End Function

Function FileFolderFromPath(f As String) As String
FileFolderFromPath = Mid$(f, 1, InStrRev(f, "\"))
End Function

Public Function GetFreePlaylistFileName() As String
Dim i As String, f As String

i = 0
Do
    'f = folder_Playlists & "list" & Right$("0000" & i, 4) & file_ext_playlist
    f = folder_Playlists & _
        Trim$(Str$(Year(Now))) & "-" & _
        right$("0" & Trim$(Str$(Month(Now))), 2) & "-" & _
        right$("0" & Trim$(Str$(Day(Now))), 2) & "-" & _
        right$("00" & i, 3) & file_ext_playlist
    If Not fsoMain.FileExists(f) Then Exit Do
    i = i + 1
Loop

GetFreePlaylistFileName = f
End Function

Function IsCompiled() As Boolean
Dim b As Boolean
b = False
Debug.Assert IsCompiled_check(b)
IsCompiled = Not b
End Function

Private Function IsCompiled_check(b As Boolean) As Boolean
b = True
IsCompiled_check = True
End Function
