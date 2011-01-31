VERSION 5.00
Begin VB.Form frmOutput 
   Caption         =   "output"
   ClientHeight    =   4620
   ClientLeft      =   40
   ClientTop       =   340
   ClientWidth     =   5920
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOutput.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4620
   ScaleWidth      =   5920
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtOutput 
      BorderStyle     =   0  'None
      Height          =   730
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   1090
   End
End
Attribute VB_Name = "frmOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Compare Binary
Option Base 0

Private Sub Form_Load()
Icon = frmMain.Icon
End Sub

Private Sub Form_Resize()
txtOutput.Move 0, 0, ScaleWidth, ScaleHeight
End Sub
