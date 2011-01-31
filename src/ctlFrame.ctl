VERSION 5.00
Begin VB.UserControl ctlFrame 
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   360
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   480
End
Attribute VB_Name = "ctlFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Option Compare Binary
Option Base 0

Private Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long

Const lCornerSize As Long = 5
Const lTopBrd As Long = 6
Const lTextTop As Long = 0

Private m_Caption As String

Private m_cMemDc As pcMemDC

Event DblClick()
Event AfterRedraw(lHdc As Long, lW As Long, lH As Long)
Event AfterResize(lW As Long, lH As Long)
Event AfterPaint()
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = -518
Caption = m_Caption
End Property
Public Property Let Caption(a As String)
m_Caption = a
UserControl_Resize
PropertyChanged "Caption"
End Property

Public Property Get BackColour() As OLE_COLOR
BackColour = UserControl.BackColor
End Property
Public Property Let BackColour(c As OLE_COLOR)
UserControl.BackColor = c
UserControl_Resize
PropertyChanged "Backcolour"
End Property

Function hWND() As Long
hWND = UserControl.hWND
End Function

Sub SetBackColour(c As Long)
BackColor = c

UserControl_Resize
End Sub

Sub ForceRedraw()
UserControl_Resize
End Sub

Sub ForceResize()
UserControl_Resize
End Sub

Sub SetDropmode()
OLEDropMode = 1

End Sub

Private Sub UserControl_DblClick()
RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()
Set m_cMemDc = New pcMemDC

UserControl_Resize
End Sub

Private Sub UserControl_InitProperties()
m_Caption = "[]"

UserControl_Resize
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub UserControl_Paint()
m_cMemDc.Draw hDC, 0, 0, m_cMemDc.Width, m_cMemDc.Height, 0, 0
End Sub

Private Sub UserControl_Resize()
Dim rcMain As RECT, lBrushBg As Long, lPen As Long, _
    rc As RECT, c As Long, X As Long

m_cMemDc.Width = ScaleWidth
m_cMemDc.Height = ScaleHeight

Dim hFontNormal As Long, nFontHeight As Long
nFontHeight = -MulDiv(UserControl.Font.Size, GetDeviceCaps(m_cMemDc.hDC, 90), 72)
hFontNormal = CreateFont(nFontHeight, _
    0, 0, 0, _
    FW_DONTCARE, 0, 0, 0, _
    DEFAULT_CHARSET, _
    0, 0, 0, 0, UserControl.Font.Name)
SelectObject m_cMemDc.hDC, hFontNormal

rcMain.Left = 0
rcMain.Top = 0
rcMain.right = ScaleWidth
rcMain.bottom = ScaleHeight

OleTranslateColor BackColor, 0, c
lBrushBg = CreateSolidBrush(c)
FillRect m_cMemDc.hDC, rcMain, lBrushBg

SetBkMode m_cMemDc.hDC, TRANSPARENT
SetTextColor m_cMemDc.hDC, GetSysColor(COLOR_BTNTEXT)

If m_Caption <> "" Then
    lPen = CreatePen(PS_SOLID, 1, GetSysColor(COLOR_BTNSHADOW))
    SelectObject m_cMemDc.hDC, lPen
    SelectObject m_cMemDc.hDC, lBrushBg
    RoundRect m_cMemDc.hDC, 0, lTopBrd, ScaleWidth, ScaleHeight, lCornerSize, lCornerSize
    DeleteObject lPen
    
    rc.Top = 0
    
    DrawText m_cMemDc.hDC, m_Caption, Len(m_Caption), rc, dt_left + DT_CALCRECT
    
    X = rc.right - rc.Left
    
    rc.Left = (ScaleWidth / 2) - (X / 2)
    rc.right = rc.Left + X
    
    FillRect m_cMemDc.hDC, rc, lBrushBg
    
    DrawText m_cMemDc.hDC, m_Caption, Len(m_Caption), rc, dt_left
End If

DeleteObject lBrushBg
DeleteObject hFontNormal

RaiseEvent AfterRedraw(m_cMemDc.hDC, ScaleWidth, ScaleHeight)
RaiseEvent AfterResize(ScaleWidth * Screen.TwipsPerPixelX, ScaleHeight * Screen.TwipsPerPixelY)

UserControl_Paint

RaiseEvent AfterPaint
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
m_Caption = PropBag.ReadProperty("Caption", "[]")
UserControl.BackColor = PropBag.ReadProperty("Backcolour", UserControl.BackColor)

UserControl_Resize
End Sub

Private Sub UserControl_Terminate()
Set m_cMemDc = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "Caption", m_Caption
PropBag.WriteProperty "Backcolour", UserControl.BackColor
End Sub
