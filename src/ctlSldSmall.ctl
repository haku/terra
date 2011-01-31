VERSION 5.00
Begin VB.UserControl ctlSldSmall 
   CanGetFocus     =   0   'False
   ClientHeight    =   2400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3200
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Timer tmrMouse 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "ctlSldSmall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Binary
Option Base 0

Private Declare Function GetCursorPos Lib "User32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "User32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
'Private Type POINTAPI
'    x As Long
'    y As Long
'End Type

Dim m_Value As Double, m_TempValue As Double

Dim c_State As Long, c_Button As Long, c_OldState As Long, _
    c_MouseX As Single, c_MouseY As Single
'   0=normal
'   1=hover over
'   2=depressed

Private m_cMemDc As pcMemDCSld

Private Const gdi_lHndNrmWidth      As Long = 5
Private Const gdi_lHndNrmHeight     As Long = 11
Private Const gdi_lEndSpace         As Long = 5

Event ValueChanged(v As Double)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event AfterRedraw(lHdc As Long, lW As Long, lH As Long)
Event AfterResize(lW As Long, lH As Long)

Sub gdi_Clear()
Dim b As Long, rc As RECT
rc.Left = 0
rc.Top = 0
rc.right = ScaleWidth
rc.bottom = ScaleHeight
b = CreateSolidBrush(GetSysColor(COLOR_BTNFACE))
FillRect m_cMemDc.hDC, rc, b
DeleteObject b
End Sub

Sub Redraw(Optional bRePaint As Boolean = True, Optional bCheckOldState As Boolean = False)
If c_State = c_OldState And bCheckOldState Then
    Exit Sub
End If
c_OldState = c_State

Dim rc As RECT, lBrush0 As Long, lBrush1 As Long, lBrush2 As Long

Dim rcSlide As RECT  'central bar
Dim rcFrame As RECT  'hover frame
Dim rcProg  As RECT  'progress
Dim rcSeek  As RECT  'seek to marker

gdi_Clear

lBrush0 = CreateSolidBrush(GetSysColor(COLOR_BTNFACE)) 'gray
lBrush1 = CreateSolidBrush(GetSysColor(COLOR_BTNTEXT)) 'black
lBrush2 = CreateSolidBrush(GetSysColor(COLOR_HIGHLIGHT)) 'blue

With m_cMemDc
    'static centre track
    rcSlide.Left = gdi_lEndSpace
    rcSlide.right = ScaleWidth - gdi_lEndSpace - 1
    rcSlide.Top = Fix(ScaleHeight / 2)
    rcSlide.bottom = rcSlide.Top + 1
    FillRect .hDC, rcSlide, lBrush1
    
    If UserControl.Enabled Then 'if not disabled...
        SelectObject .hDC, lBrush0
        
        'hover frame
        If c_State > 0 Then
            rcFrame = rcSlide
            ShrinkRect rcFrame, -4
            FrameRect .hDC, rcFrame, lBrush2
        End If
        
        'draw progress track
        rcProg = rcSlide
        rcProg.Top = rcProg.Top - 2
        rcProg.bottom = rcProg.bottom + 2
        rcProg.right = rcProg.Left + (rcProg.right - rcProg.Left) * m_Value
        FillRect .hDC, rcProg, lBrush1
        
        If UserControl.Enabled And c_Button = 1 And _
            c_MouseX >= 0 And c_MouseY >= 0 And _
            c_MouseX <= ScaleWidth - 1 And c_MouseY <= ScaleHeight - 1 Then
            
            SetRectHW rcSeek, _
                gdi_lEndSpace + (ScaleWidth - gdi_lEndSpace * 2) * m_TempValue - (gdi_lHndNrmWidth / 2), _
                rcSlide.Top + ((rcSlide.bottom - rcSlide.Top) / 2) - (gdi_lHndNrmHeight / 2), gdi_lHndNrmWidth, gdi_lHndNrmHeight
            
            FillRect .hDC, rcSeek, lBrush2
        End If
    End If
End With

DeleteObject lBrush0
DeleteObject lBrush1
DeleteObject lBrush2

If bRePaint Then UserControl_Paint
End Sub

Function UnderMouse() As Boolean
Dim ptMouse As POINTAPI
GetCursorPos ptMouse
If WindowFromPoint(ptMouse.X, ptMouse.Y) = UserControl.hWND Then
   UnderMouse = True
Else
   UnderMouse = False
End If
End Function

Public Property Get Enabled() As Boolean
Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(e As Boolean)
UserControl.Enabled = e
Redraw
PropertyChanged "Enabled"
End Property

Public Property Get Value() As Double
Value = m_Value
End Property
Public Property Let Value(v As Double)
m_Value = v
Redraw
PropertyChanged "Value"
End Property

Function hWND() As Long
hWND = UserControl.hWND
End Function

Sub ForceRedraw()
Redraw
End Sub

Sub ForceResize()
UserControl_Resize
End Sub

Private Sub tmrMouse_Timer()
If Not UnderMouse Then
    Select Case c_Button
        Case 0
            c_State = 0
            tmrMouse.Enabled = False
        
        Case Else
            c_State = 1
        
    End Select
Else
    Select Case c_Button
        Case 0
            c_State = 1
        
        Case Else
            c_State = 2
        
    End Select
End If

Redraw True, True
End Sub

Private Sub UserControl_DblClick()
c_State = 0
End Sub

Private Sub UserControl_Initialize()
Set m_cMemDc = New pcMemDCSld

m_Value = 0

c_State = 0
c_OldState = -1
c_Button = 0

Redraw False
End Sub

Private Sub UserControl_InitProperties()
UserControl_Resize
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
c_State = 2
UserControl_MouseMove Button, Shift, X, Y

RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
c_Button = Button

c_MouseX = X
c_MouseY = Y

If Button = 1 And X >= 0 And Y >= 0 And _
    X <= ScaleWidth - 1 And Y <= ScaleHeight - 1 Then
    
    m_TempValue = (X - gdi_lEndSpace) / (ScaleWidth - gdi_lEndSpace * 2 - 1)
    
    If m_TempValue < 0 Then
        m_TempValue = 0
    ElseIf m_TempValue > 1 Then
        m_TempValue = 1
    End If
    
    Redraw
End If

If tmrMouse.Enabled <> True Then tmrMouse.Enabled = True
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
UserControl_MouseMove 0, Shift, X, Y

If Button = 1 And X >= 0 And Y >= 0 And _
    X <= ScaleWidth - 1 And Y <= ScaleHeight - 1 Then
    
    RaiseEvent ValueChanged(m_TempValue)
End If
End Sub

Private Sub UserControl_Paint()
m_cMemDc.Draw hDC, 0, 0, m_cMemDc.Width, m_cMemDc.Height, 0, 0
End Sub

Private Sub UserControl_Resize()
m_cMemDc.Width = ScaleWidth
m_cMemDc.Height = ScaleHeight

Redraw False

RaiseEvent AfterRedraw(hDC, ScaleWidth, ScaleHeight)
RaiseEvent AfterResize(ScaleWidth * Screen.TwipsPerPixelX, ScaleHeight * Screen.TwipsPerPixelY)

UserControl_Paint
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
Redraw
End Sub

Private Sub UserControl_Terminate()
Set m_cMemDc = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "Enabled", UserControl.Enabled
End Sub


