VERSION 5.00
Begin VB.Form frmDisp 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "fullscreen | terra"
   ClientHeight    =   2020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3120
   FillColor       =   &H8000000F&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDisp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   202
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "//terra"
      ForeColor       =   &H00E0E0E0&
      Height          =   180
      Left            =   840
      TabIndex        =   0
      Top             =   840
      Width           =   400
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmDisp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Compare Binary
Option Base 0

Public Sub SetDisp(cMon As cMonitor)
WindowState = vbNormal
Move cMon.WorkLeft * Screen.TwipsPerPixelX + 1, _
    cMon.WorkTop * Screen.TwipsPerPixelY + 1

Caption = "terra//" & cMon.Name
lblCaption.Caption = "//terra on " & _
    cMon.Name & " (" & cMon.Width & "×" & _
    cMon.Height & ").  press escape to remove full screen."

Show
WindowState = vbMaximized
End Sub

Public Sub UpdateDisplay()
Dim b As Boolean
If frmMain.pb_GetPlayState > 0 Then
    b = Not cMedia.HasVideo
Else
    b = True
End If
lblCaption.Visible = b
Form_Paint
End Sub

Private Sub Form_Activate()
KeepOnTop Me, True
UpdateDisplay
End Sub

Private Sub Form_DblClick()
frmMain.cmdVidFull_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyEscape
        frmMain.cmdVidFull_Click
    
    Case vbKeySpace
        frmMain.cmdPlayState_Click 1
    
End Select
End Sub

Private Sub Form_Load()
Icon = frmMain.Icon
End Sub

Private Sub Form_Paint()
Cls
frmMain.pb_DrawCurrentItem hDC, ScaleWidth, ScaleHeight, &HFFFFFF
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Select Case UnloadMode
    Case vbFormCode, vbAppWindows, vbAppTaskManager
        
    Case Else 'user indused
        frmMain.cmdVidFull_Click
End Select
End Sub

Private Sub Form_Resize()
lblCaption.Move 0, 0, ScaleWidth
End Sub
