VERSION 5.00
Begin VB.UserControl ctlSysTray 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   CanGetFocus     =   0   'False
   ClientHeight    =   260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   260
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   26
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   26
End
Attribute VB_Name = "ctlSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Option Compare Binary
Option Base 0

Event MouseMove(x As Single)

Public Function GetHwnd() As Long
GetHwnd = UserControl.hWND
End Function

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent MouseMove(x)
End Sub
