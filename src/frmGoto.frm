VERSION 5.00
Begin VB.Form frmGoto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "jump to track | terra"
   ClientHeight    =   3700
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   4940
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGoto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3700
   ScaleWidth      =   4940
   Begin terra.ctlXpThemeBtn cmdCancel 
      Cancel          =   -1  'True
      Height          =   300
      Left            =   120
      Top             =   3360
      Width           =   1090
      _ExtentX        =   1923
      _ExtentY        =   529
      DrawMode        =   0
      Caption         =   "cancel"
      BtnType         =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      BackColour      =   -2147483633
   End
   Begin VB.PictureBox picList 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   2290
      Left            =   120
      ScaleHeight     =   229
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   469
      TabIndex        =   1
      Top             =   720
      Width           =   4690
   End
   Begin VB.TextBox txtSearch 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   250
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4690
   End
   Begin terra.ctlXpThemeBtn cmdOk 
      Height          =   300
      Index           =   1
      Left            =   2400
      Top             =   3360
      Width           =   1090
      _ExtentX        =   1923
      _ExtentY        =   529
      DrawMode        =   0
      Caption         =   "add to cue"
      BtnType         =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      BackColour      =   -2147483633
   End
   Begin terra.ctlXpThemeBtn cmdOk 
      Height          =   300
      Index           =   0
      Left            =   3720
      Top             =   3360
      Width           =   1090
      _ExtentX        =   1923
      _ExtentY        =   529
      DrawMode        =   0
      Caption         =   "play now"
      BtnType         =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      BackColour      =   -2147483633
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "[info]"
      Height          =   180
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   330
   End
   Begin VB.Label lblStat 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "[...]"
      Height          =   180
      Left            =   2280
      TabIndex        =   2
      Top             =   480
      Width           =   210
   End
End
Attribute VB_Name = "frmGoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private plResults   As typPl
Private m_PlMdc     As pcMemDC
Private m_PlTopI    As Long

Public m_bLocked    As Boolean
'

'Public m_sTrack     As String
'Public m_lMethod    As Long
'

Public Sub CloseMe()
Unload Me
End Sub

Private Sub UpdateResults(Optional bRedraw As Boolean = True)
Dim sTrm As String

If Len(txtSearch.Text) >= 3 Then
    sTrm = txtSearch.Text
    
    sTrm = Replace(sTrm, "'", "''")
    sTrm = Replace(sTrm, "\", "\\")
    sTrm = Replace(sTrm, "%", "\%")
    sTrm = Replace(sTrm, "_", "\_")
    sTrm = Replace(sTrm, "*", "%")
    
    mdb_SearchToPL sTrm, plResults, "\"
    
    lblStat.Caption = plResults.lCnt & " results."
    plResults.lIndex = 0
    m_PlTopI = 0
Else
    pl_SetAsNew plResults
    lblStat.Caption = "to search, enter three or more characters.  use * for wildcard."
End If

If bRedraw Then RedrawResults
End Sub

Private Sub RedrawResults()
pl_Draw plResults, m_PlTopI, _
    m_PlMdc.hDC, m_PlMdc.Width, m_PlMdc.Height, _
    True, _
    "up and down keys: select item" & vbNewLine & _
    "enter: play item" & vbNewLine & _
    "ctrl+enter: add to cue" & vbNewLine & _
    "escape: cancel" & vbNewLine & _
    "alt: lock window open" _
    , "0000", False, 1, False

picList_Paint
End Sub

Private Sub UpdateInfo()
lblInfo.Caption = IIf(m_bLocked, "window locked open.  press alt to release.", "")
lblInfo.Visible = Len(lblInfo.Caption) > 0

Form_Resize
End Sub

Private Sub cmdCancel_Click()
CloseMe
End Sub

Private Sub cmdOk_Click(Index As Integer)
Dim i As Long

If plResults.lCnt < 1 Then Exit Sub
If plResults.lIndex < 0 Then Exit Sub

i = frmMain.mdbl_FindIndexFromFile(plResults.Items(plResults.lIndex).sFile)

Select Case Index
    Case 0 'play
        frmMain.pb_StartPlayback 0, i
    
    Case 1 'cue
        frmMain.cue_AddTo 0, _
            mdb_List.Items(i).sFile, , _
            mdb_List.Items(i).lMD5, _
            mdb_List.Items(i).lDuration
    
End Select

If Not m_bLocked Then CloseMe
End Sub

Private Sub Form_Activate()
KeepOnTop Me, True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Long

Select Case KeyCode
    Case vbKeyDown, vbKeyUp
        If plResults.lCnt < 1 Then Exit Sub
        
        i = plResults.lIndex + IIf(KeyCode = vbKeyUp, -1, 1)
        If i < 0 Then
            i = 0
        ElseIf i > plResults.lCnt - 1 Then
            i = plResults.lCnt - 1
        End If
        
        If i < m_PlTopI Then
            m_PlTopI = i
        ElseIf i > m_PlTopI + Fix(picList.ScaleHeight / mdbl_ItmH) - 1 Then
            m_PlTopI = i - Fix(picList.ScaleHeight / mdbl_ItmH) + 1
        End If
        
        If plResults.lIndex <> i Then
            plResults.lIndex = i
            RedrawResults
        End If
        
        KeyCode = 0
    
    Case vbKeyMenu
        KeyCode = 0
        m_bLocked = Not m_bLocked
        UpdateInfo
    
    Case vbKeyReturn
        'If Shift = vbCtrlMask Or Shift = vbShiftMask Or Shift = 18 Then
        If Shift > 0 Then
            cmdOk_Click 1 'cue
        Else
            cmdOk_Click 0 'play now
        End If
    
End Select
End Sub

Private Sub Form_Load()
m_bGotoVis = True 'mark that this form is open
Icon = frmMain.Icon 'copy icon
Set m_PlMdc = New pcMemDC 'setup drawing canvas
m_bLocked = False
UpdateInfo

If Len(m_GotoText) > 0 Then
    txtSearch.Text = m_GotoText 'remember last search and trigger search
Else
    UpdateResults False 'init list
End If
End Sub

Private Sub Form_Resize()
On Error GoTo frmResize_err

Dim lInfoH As Long

lInfoH = IIf(lblInfo.Visible, lblInfo.Height + lGp, 0)

lblStat.Move 0, 0, ScaleWidth
txtSearch.Move lGp, lblStat.Top + lblStat.Height + lGp, ScaleWidth - lGp * 2

picList.Move 0, txtSearch.Top + txtSearch.Height + lGp, ScaleWidth, _
    ScaleHeight - txtSearch.Height - lblStat.Height - cmdCancel.Height - lInfoH - lGp * 4

lblInfo.Move lGp, picList.Top + picList.Height + lGp, ScaleWidth - lGp * 2

cmdCancel.Move lGp, picList.Top + picList.Height + lInfoH + lGp
cmdOk(0).Move ScaleWidth - cmdOk(0).Width - lGp, cmdCancel.Top
cmdOk(1).Move cmdOk(0).Left - cmdOk(1).Width - lGp, cmdOk(0).Top

m_PlMdc.Width = picList.ScaleWidth
m_PlMdc.Height = picList.ScaleHeight

RedrawResults

Exit Sub
frmResize_err:
err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
m_GotoText = txtSearch.Text
Set m_PlMdc = Nothing
m_bGotoVis = False
End Sub

Private Sub picList_Paint()
On Error GoTo picPl_err
m_PlMdc.Draw picList.hDC, 0, 0, m_PlMdc.Width, m_PlMdc.Height, 0, 0

Exit Sub
picPl_err:
err.Clear
End Sub

Private Sub txtSearch_Change()
UpdateResults
End Sub

Private Sub txtSearch_GotFocus()
txtSearch.SelStart = 0
txtSearch.SelLength = Len(txtSearch.Text)
End Sub
