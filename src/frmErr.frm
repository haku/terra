VERSION 5.00
Begin VB.Form frmErr 
   Caption         =   "error report | terra"
   ClientHeight    =   3010
   ClientLeft      =   40
   ClientTop       =   340
   ClientWidth     =   5170
   Icon            =   "frmErr.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3010
   ScaleWidth      =   5170
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEnd 
      Caption         =   "force &end"
      Height          =   330
      Left            =   3000
      TabIndex        =   1
      Top             =   2520
      Width           =   970
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "copy &message to clipboard"
      Height          =   330
      Left            =   600
      TabIndex        =   2
      Top             =   2520
      Width           =   2290
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&continue"
      Height          =   330
      Left            =   4080
      TabIndex        =   0
      Top             =   2520
      Width           =   970
   End
   Begin VB.TextBox txtOutput 
      BorderStyle     =   0  'None
      Height          =   2410
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   0
      Width           =   5170
   End
End
Attribute VB_Name = "frmErr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Compare Binary
Option Base 0

Public Sub SetType(i As Integer)
Select Case i
    Case 0
        cmdEnd.Visible = True
    
    Case 1
        cmdEnd.Visible = False
    
End Select

Form_Resize
End Sub

Private Sub cmdCopy_Click()
Clipboard.Clear
Clipboard.SetText txtOutput.Text, vbCFRTF
Clipboard.SetText txtOutput.Text, vbCFText
End Sub

Private Sub cmdEnd_Click()
main_end True
End Sub

Private Sub cmdOk_Click()
Unload Me
End Sub

Private Sub Form_Resize()
On Error GoTo Form_Resize_err

txtOutput.Move 0, 0, ScaleWidth, ScaleHeight - cmdOk.Height - lSp * 2
cmdOk.Move ScaleWidth - cmdOk.Width - lSp, ScaleHeight - cmdOk.Height - lSp

If cmdEnd.Visible Then
    cmdEnd.Move cmdOk.Left - cmdEnd.Width - lSp, cmdOk.Top
    cmdCopy.Move cmdEnd.Left - cmdCopy.Width - lSp, cmdOk.Top
Else
    cmdCopy.Move cmdOk.Left - cmdCopy.Width - lSp, cmdOk.Top
End If

Form_Resize_err:
err.Clear
End Sub
