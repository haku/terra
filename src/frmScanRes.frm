VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmScanRes 
   Caption         =   "results of media library scan | terra"
   ClientHeight    =   3300
   ClientLeft      =   110
   ClientTop       =   350
   ClientWidth     =   9920
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmScanRes.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   9920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMenu 
      Caption         =   "change action..."
      Height          =   300
      Left            =   6840
      TabIndex        =   6
      Top             =   720
      Width           =   1330
   End
   Begin VB.CheckBox chkCountMerge 
      Caption         =   "transfer meta data based on crc."
      Height          =   300
      Left            =   1560
      TabIndex        =   4
      Top             =   2880
      Value           =   1  'Checked
      Width           =   3730
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&save list..."
      Height          =   300
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   1330
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "con&firm"
      Height          =   300
      Index           =   1
      Left            =   5400
      TabIndex        =   2
      Top             =   2880
      Width           =   1330
   End
   Begin VB.CommandButton cmdAction 
      Cancel          =   -1  'True
      Caption         =   "&cancel"
      Height          =   300
      Index           =   0
      Left            =   6840
      TabIndex        =   1
      Top             =   2880
      Width           =   1330
   End
   Begin MSComctlLib.ListView lvRes 
      Height          =   1690
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   8050
      _ExtentX        =   14199
      _ExtentY        =   2981
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "file"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "stored crc"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "issue"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "action"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.Label lblCap 
      AutoSize        =   -1  'True
      Caption         =   $"frmScanRes.frx":000C
      Height          =   360
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   8190
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuAct 
      Caption         =   "change action"
      Visible         =   0   'False
      Begin VB.Menu mnuActI 
         Caption         =   "[]"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmScanRes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Compare Binary
Option Base 0

Public lRes As Long

Public Function AddItem(ByVal f As String, lCRC As Long, _
    sIss As String, sAct As trrFileAction, _
    Optional bReplace As Boolean = False) _
    As Boolean

Dim l As MSComctlLib.ListItem, i As Long, b As Boolean, x As Long

'todo: this is a bug fix, as the list view does not support unicode.
f = us_Encode(f)

b = False
If lvRes.ListItems.count > 0 Then
    For i = 1 To lvRes.ListItems.count
        If lvRes.ListItems(i).Text = f Then
            b = True
            x = i
            Exit For
        End If
    Next i
End If

If b = False Then
    Set l = lvRes.ListItems.Add(, , f)
    l.SubItems(1) = Hex$(lCRC)
    l.SubItems(2) = sIss
    l.Tag = sAct
    l.SubItems(3) = GettrrFileActDes(sAct)
ElseIf bReplace Then
    Set l = lvRes.ListItems(x)
    l.SubItems(1) = Hex$(lCRC)
    l.SubItems(2) = sIss
    l.Tag = sAct
    l.SubItems(3) = GettrrFileActDes(sAct)
End If

AddItem = Not b
End Function

Private Sub cmdAction_Click(Index As Integer)
Dim i As Long, x As Long, cItems As Collection, bFound As Boolean

If Index = 1 Then
    'move function does not work, so make sure its not used (a temp fix).
    For i = 1 To lvRes.ListItems.count
        If lvRes.ListItems(i).Tag = trrMoveFile Then
            MsgBox "the move file feature is not yet finished, please make another choice."
            Exit Sub
        End If
    Next i
    
    'check that we are not going to merge data for items that are not being
    ' removed from the library.
    Set cItems = New Collection
    If chkCountMerge.Value = 1 Then
        'step 1: build a list of the crcs that are staying the library
        For i = 1 To lvRes.ListItems.count
            If lvRes.ListItems(i).Tag = trrDoNothing _
                Or lvRes.ListItems(i).Tag = trrMarkMissing Then
            
                cItems.Add lvRes.ListItems(i).SubItems(1)
                
            End If
        Next i
        
        'step 2: check for duplicates
        bFound = False
        For i = 1 To cItems.count
            For x = IIf(i + 1 > cItems.count, cItems.count, i + 1) To cItems.count
                If cItems(i) = cItems(x) And i <> x Then
                    bFound = True
                    Exit For
                End If
            Next x
        Next i
        
        If bFound Then
            MsgBox "In order to merge metadata by CRC, you must ensure that all " & _
                "CRC values that remain in the library (after processing) are unique.  " & _
                "Please adjust the item processing options or disable this feature, " & _
                "then try again."
            Exit Sub
        End If
    End If
    
    If MsgBox("sure?", vbYesNo) <> vbYes Then Exit Sub
End If

lRes = Index
Hide
End Sub

Private Sub cmdMenu_Click()
PopupMenu mnuAct, vbPopupMenuRightAlign, cmdMenu.Left + cmdMenu.Width, cmdMenu.Top + cmdMenu.Height
End Sub

Private Sub cmdSave_Click()
Dim sBuf As New cStringBuilder, i As Long, f As String

If lvRes.ListItems.count < 1 Then Exit Sub

f = SaveDialog(Me, "text files (*.txt)|*.txt", "save report", "")
If f = "" Then Exit Sub

With lvRes.ListItems
    For i = 1 To .count
        sBuf.Append us_Decode(.Item(i).Text) & ", " & _
        .Item(i).SubItems(1) & ", " & _
        .Item(i).SubItems(2) & ", " & _
        .Item(i).SubItems(3) & vbNewLine
    Next i
End With

UnicodeFile_Write_FSO f, sBuf.ToString
End Sub

Private Sub Form_Load()
Dim i As Long

For i = 0 To 3
    If i > mnuActI.count - 1 Then Load mnuActI(i)
    mnuActI(i).Caption = GettrrFileActDes(i)
Next i

Icon = frmMain.Icon
End Sub

Private Sub Form_Resize()
On Error GoTo frmResize_err

lblCap.Move lGp, lGp, ScaleWidth - lGp * 2
lblCap.AutoSize = True
cmdMenu.Move ScaleWidth - cmdMenu.Width - lGp, lblCap.Top + lblCap.Height + lGp
lvRes.Move 0, cmdMenu.Top + cmdMenu.Height + lGp, ScaleWidth, _
    ScaleHeight - cmdAction(0).Height - lblCap.Height - cmdMenu.Height - lGp * 5
cmdSave.Move lSp, ScaleHeight - cmdSave.Height - lGp
cmdAction(0).Move ScaleWidth - cmdAction(0).Width - lSp, ScaleHeight - cmdAction(0).Height - lGp
cmdAction(1).Move cmdAction(0).Left - cmdAction(1).Width - lSp, _
    ScaleHeight - cmdAction(0).Height - lGp
chkCountMerge.Move cmdSave.Left + cmdSave.Width + lSp, cmdSave.Top, _
    cmdAction(1).Left - chkCountMerge.Left - lGp

Exit Sub
frmResize_err:
err.Clear
End Sub

Private Sub lvRes_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 2 And x >= 0 And Y >= 0 And x <= lvRes.Width And Y <= lvRes.Height Then
    PopupMenu mnuAct
End If
End Sub

Private Sub mnuActI_Click(Index As Integer)
Dim i As Long
If lvRes.ListItems.count < 1 Then Exit Sub
For i = 1 To lvRes.ListItems.count
    If lvRes.ListItems(i).Selected Then
        lvRes.ListItems(i).SubItems(3) = GettrrFileActDes((Index))
        lvRes.ListItems(i).Tag = Index
    End If
Next i
End Sub
