VERSION 5.00
Begin VB.Form frmMiniVid 
   ClientHeight    =   780
   ClientLeft      =   40
   ClientTop       =   40
   ClientWidth     =   1320
   ControlBox      =   0   'False
   FillColor       =   &H8000000F&
   Icon            =   "frmMiniVid.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   78
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   132
   ShowInTaskbar   =   0   'False
   Begin terra.ctlXpThemeBtn cmdPlayPause 
      Height          =   170
      Left            =   720
      Tag             =   "4"
      Top             =   120
      Width           =   170
      _ExtentX        =   300
      _ExtentY        =   300
      DrawMode        =   0
      Caption         =   ""
      BtnType         =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   6
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      BackColour      =   -2147483633
   End
   Begin terra.ctlSldSmall sldSeek 
      Height          =   170
      Left            =   360
      Top             =   600
      Width           =   850
      _ExtentX        =   1499
      _ExtentY        =   300
      Enabled         =   -1  'True
   End
   Begin terra.ctlFrame fraVid 
      Height          =   490
      Left            =   0
      Top             =   0
      Width           =   610
      _ExtentX        =   1076
      _ExtentY        =   864
      Caption         =   ""
      Backcolour      =   0
   End
   Begin terra.ctlXpThemeBtn cmdNext 
      Height          =   170
      Left            =   960
      Tag             =   ":"
      Top             =   120
      Width           =   170
      _ExtentX        =   300
      _ExtentY        =   300
      DrawMode        =   0
      Caption         =   ""
      BtnType         =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   6
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      BackColour      =   -2147483633
   End
   Begin terra.ctlXpThemeBtn cmdClose 
      Height          =   170
      Left            =   960
      Tag             =   "r"
      Top             =   360
      Width           =   170
      _ExtentX        =   300
      _ExtentY        =   300
      DrawMode        =   0
      Caption         =   ""
      BtnType         =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Marlett"
         Size            =   6
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      BackColour      =   -2147483633
   End
End
Attribute VB_Name = "frmMiniVid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub SetPosition()
Dim l As Long, t As Long, a
l = (m_cM.Monitor(1).WorkLeft + m_cM.Monitor(1).WorkWidth) * Screen.TwipsPerPixelX - Width
t = ((m_cM.Monitor(1).WorkTop + m_cM.Monitor(1).WorkHeight) / 2) * Screen.TwipsPerPixelY + (Height / 2)

'load position data
Move _
    GetFromIniEx("minivid", "normalleft", Trim$(Str$(l)), file_INI), _
    GetFromIniEx("minivid", "normaltop", Trim$(Str$(t)), file_INI), _
    GetFromIniEx("minivid", "normalwidth", Width, file_INI), _
    GetFromIniEx("minivid", "normalheight", Height, file_INI)
a = GetFromIniEx("minivid", "showcmd", "1", file_INI)
Select Case a
    Case "1": WindowState = vbNormal
    Case "2": WindowState = vbMinimized
    Case "3": WindowState = vbMaximized
End Select

Show
End Sub

Public Sub UpdatedatePlaystate()
cmdPlayPause.Tag = IIf(frmMain.pb_GetPlayState = 1, "1", "0")
cmdPlayPause.ForceRedraw
End Sub

Public Sub UpdatePogbar()
If cMedia.Duration > 0 Then
    sldSeek.Value = cMedia.Position / cMedia.Duration
End If
End Sub

Private Sub Draw3PxArrow(lHdc As Long, l As Long, t As Long, lForeColour As Long)
Dim pt As POINTAPI

MoveToEx lHdc, l, t, pt
LineTo lHdc, l, t + 5

MoveToEx lHdc, l + 1, t + 1, pt
LineTo lHdc, l + 1, t + 4

SetPixel lHdc, l + 2, t + 2, lForeColour
End Sub

Private Sub cmdClose_AfterRedraw(lHdc As Long, lW As Long, lH As Long, lForeColour As Long)
Dim x As Long, Y As Long, pt As POINTAPI, lPen As Long
x = Fix(lW / 2)
Y = Fix(lH / 2)

lPen = CreatePen(PS_SOLID, 1, lForeColour)
SelectObject lHdc, lPen

MoveToEx lHdc, x - 2, Y - 2, pt
LineTo lHdc, x + 3, Y + 3

MoveToEx lHdc, x - 2, Y + 2, pt
LineTo lHdc, x + 3, Y - 3

DeleteObject lPen
End Sub

Private Sub cmdClose_Click()
frmMain.cmdVidMinivid_Click
End Sub

Private Sub cmdNext_AfterRedraw(lHdc As Long, lW As Long, lH As Long, lForeColour As Long)
Dim x As Long, Y As Long, pt As POINTAPI, lPen As Long
x = Fix(lW / 2)
Y = Fix(lH / 2)

lPen = CreatePen(PS_SOLID, 1, lForeColour)
SelectObject lHdc, lPen

Draw3PxArrow lHdc, x - 3, Y - 2, lForeColour
Draw3PxArrow lHdc, x, Y - 2, lForeColour

MoveToEx lHdc, x + 3, Y - 2, pt
LineTo lHdc, x + 3, Y + 3

DeleteObject lPen
End Sub

Private Sub cmdNext_Click()
frmMain.pb_NextFile
End Sub

Private Sub cmdPlayPause_AfterRedraw(lHdc As Long, lW As Long, lH As Long, lForeColour As Long)
Dim x As Long, Y As Long, iX As Long, pt As POINTAPI, lPen As Long
x = Fix(lW / 2)
Y = Fix(lH / 2)

lPen = CreatePen(PS_SOLID, 1, lForeColour)
SelectObject lHdc, lPen

Select Case cmdPlayPause.Tag
    Case "0" 'play
            Draw3PxArrow lHdc, x - 1, Y - 2, lForeColour
        
    Case "1" 'pause
        For iX = -2 To 2
            If iX <> 0 Then
                MoveToEx lHdc, x + iX, Y - 2, pt
                LineTo lHdc, x + iX, Y + 3
            End If
        Next iX
    
End Select

DeleteObject lPen
End Sub

Private Sub cmdPlayPause_Click()
frmMain.cmdPlayState_Click 1
End Sub

Private Sub Form_Activate()
KeepOnTop Me, True
UpdatedatePlaystate
cmdNext.ForceRedraw
cmdClose.ForceRedraw
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyF2, vbKeyEscape
            cmdClose.SimClick
        
    End Select
End If
End Sub

Private Sub Form_Load()
Icon = frmMain.Icon
fraVid.SetDropmode
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button <> 1 Then Exit Sub
ReleaseCapture
SendMessage Me.hWND, &HA1, 2, 0&
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Select Case UnloadMode
    Case vbFormCode, vbAppWindows, vbAppTaskManager
        
    Case Else 'user indused
        frmMain.cmdVidMinivid_Click
End Select
End Sub

Private Sub Form_Resize()
On Error GoTo Form_Resize_err

fraVid.Move 0, 0, ScaleWidth, ScaleHeight - sldSeek.Height
cmdPlayPause.Move 0, fraVid.Height
cmdNext.Move cmdPlayPause.Width, cmdPlayPause.Top
sldSeek.Move cmdNext.Left + cmdNext.Width, fraVid.Height, _
    ScaleWidth - cmdPlayPause.Width - cmdNext.Width - cmdClose.Width
cmdClose.Move sldSeek.Left + sldSeek.Width, sldSeek.Top

Exit Sub
Form_Resize_err:
err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim wp As WINDOWPLACEMENT
GetWindowPlacement hWND, wp
WriteToIni "minivid", "showcmd", Trim$(Str$(wp.showCmd)), file_INI
WriteToIni "minivid", "normalleft", Trim$(Str$(wp.rcNormalPosition.Left * Screen.TwipsPerPixelX)), file_INI
WriteToIni "minivid", "normaltop", Trim$(Str$(wp.rcNormalPosition.Top * Screen.TwipsPerPixelY)), file_INI
WriteToIni "minivid", "normalwidth", Trim$(Str$((wp.rcNormalPosition.right - wp.rcNormalPosition.Left) * Screen.TwipsPerPixelX)), file_INI
WriteToIni "minivid", "normalheight", Trim$(Str$((wp.rcNormalPosition.bottom - wp.rcNormalPosition.Top) * Screen.TwipsPerPixelY)), file_INI
End Sub

Private Sub fraVid_AfterRedraw(lHdc As Long, lW As Long, lH As Long)
frmMain.pb_DrawCurrentItemSmall lHdc, lW, lH, &HFFFFFF
End Sub

Private Sub fraVid_AfterResize(lW As Long, lH As Long)
If Len(cMedia.FileName) > 0 And cMedia.Width > 0 And cMedia.HasVideo Then
    cMedia.ResizeWindow
End If
End Sub

Private Sub fraVid_DblClick()
frmMain.gen_GoFullscreen m_cM.MonitorForWindow(hWND)
End Sub

Private Sub fraVid_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Form_MouseMove Button, Shift, x, Y
End Sub

Private Sub fraVid_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim i As Long

If Not Data.GetFormat(vbCFFiles) Then Exit Sub

For i = 1 To Data.Files.count
    If fsoMain.FileExists(Data.Files(i)) Then
        If IsMediaFile(Data.Files(i)) Then
            frmMain.cue_AddTo -1, Data.Files(i)
        End If
    ElseIf fsoMain.FolderExists(Data.Files(i)) Then
        
        Dim j As Long, cFlds As Collection, cFiles As Collection
    
        Set cFlds = New Collection
        BuildDirCollection Data.Files(i), cFlds
        
        Set cFiles = New Collection
        For j = 1 To cFlds.count
            AddFilesToCollection cFlds(j), cFiles
        Next j
        
        If cFiles.count < 1 Then Exit Sub
        
        For j = 1 To cFiles.count
            If IsMediaFile(cFiles(j)) Then
                frmMain.cue_AddTo -1, cFiles(j)
            End If
        Next j
        
    End If
Next i
End Sub

Private Sub sldSeek_ValueChanged(v As Double)
frmMain.pb_SetPlaybackPosition cMedia.Duration * v
End Sub
