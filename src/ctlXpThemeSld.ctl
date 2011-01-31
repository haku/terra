VERSION 5.00
Begin VB.UserControl ctlXpThemeSld 
   CanGetFocus     =   0   'False
   ClientHeight    =   1330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3270
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   133
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   327
   Begin VB.Timer tmrMouse 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Image imgHndDwn 
      Height          =   1050
      Left            =   960
      Picture         =   "ctlXpThemeSld.ctx":0000
      Top             =   120
      Visible         =   0   'False
      Width           =   110
   End
   Begin VB.Image imgHndNrm 
      Height          =   1050
      Left            =   720
      Picture         =   "ctlXpThemeSld.ctx":0F08
      Top             =   120
      Visible         =   0   'False
      Width           =   110
   End
   Begin VB.Image imgTrack 
      Height          =   50
      Left            =   600
      Picture         =   "ctlXpThemeSld.ctx":1E10
      Top             =   120
      Visible         =   0   'False
      Width           =   50
   End
End
Attribute VB_Name = "ctlXpThemeSld"
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

Dim m_Value As Double, m_TempValue As Double, m_GapX As Long, _
    m_DrawMode As expDrawMode

Dim c_State As Long, c_Button As Long, c_OldState As Long, _
    c_MouseX As Single, c_MouseY As Single
'   0=normal
'   1=hover over
'   2=depressed

Event ValueChanged(v As Double)
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event AfterRedraw(lHdc As Long, lW As Long, lH As Long)
Event AfterResize(lW As Long, lH As Long)

'Private m_cUxTheme As cUxTheme
Private m_cMemDc As pcMemDCSld, _
    m_Dib_Track As cDIBSection, m_Dib_HndNrm As cDIBSection, m_Dib_HndDwn As cDIBSection
'

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

Dim iState As Integer, _
    lHndNrmLeft As Long, lHndNrmWidth As Long, lHndNrmHeight As Long, _
    rc As RECT, lBrush0 As Long, lBrush1 As Long, lBrush2 As Long, bPen0 As Long

c_OldState = c_State
'Debug.Print Format(Now, "hh:mm:ss") & " draw"

gdi_Clear

'If m_DrawMode = xpUxTheme Then
'    With m_cUxTheme
'        .hdc = m_cMemDc.hdc
'        .Class = "Trackbar"
'        .Part = 3
'        lHndNrmWidth = .PartWidth
'        lHndNrmHeight = .PartHeight
'    End With
'Else
    lHndNrmWidth = m_Dib_HndNrm.Width
    lHndNrmHeight = m_Dib_HndNrm.Height / 5
'End If

lHndNrmLeft = ((ScaleWidth - m_GapX * 2) * m_Value) - (lHndNrmWidth / 2) + m_GapX

If UserControl.Enabled Then
    Select Case c_State
        Case 0: iState = 1 'normal
        Case 1
            If c_MouseX >= lHndNrmLeft And c_MouseX <= lHndNrmLeft + lHndNrmWidth And _
                c_MouseY >= 0 And c_MouseY <= ScaleHeight Then
                iState = 2 'hover
            Else
                iState = 1 'normal
            End If
        
        Case 2: iState = 3 'depressed
    End Select
Else
    iState = 5 'disabled
End If

Select Case m_DrawMode
'    Case xpUxTheme
'        With m_cUxTheme
'            .Part = 1
'            .Left = m_GapX
'            .Top = (ScaleHeight / 2) - (.PartHeight / 2)
'            .Width = ScaleWidth - m_GapX * 2
'            .Height = .PartHeight
'            .Draw
'
'            .Part = 3
'            .State = iState
'            .Left = lHndNrmLeft
'            .Top = (ScaleHeight / 2) - (.PartHeight / 2)
'            .Width = .PartWidth
'            .Height = .PartHeight
'            .Draw
'
'            If UserControl.Enabled And c_Button = 1 And _
'                c_MouseX >= m_GapX And c_MouseY >= 0 And _
'                c_MouseX <= ScaleWidth - m_GapX And c_MouseY <= ScaleHeight Then
'
'                .Part = 4
'                .Left = ((ScaleWidth - m_GapX * 2) * m_TempValue) - (.PartWidth / 2) + m_GapX
'                .Top = (ScaleHeight / 2) - (.PartHeight / 2)
'                .Width = .PartWidth
'                .Height = .PartHeight
'                .State = 3
'                .Draw
'            End If
'        End With
    
    Case xpBitmap
        With m_Dib_Track
            BitBlt m_cMemDc.hDC, m_GapX, (ScaleHeight / 2) - (.Height / 2), _
                2, .Height, .hDC, 0, 0, SRCCOPY
            StretchBlt m_cMemDc.hDC, m_GapX + 2, (ScaleHeight / 2) - (.Height / 2), _
                ScaleWidth - ((m_GapX + 2) * 2), .Height, .hDC, _
                2, 0, 1, .Height, SRCCOPY
            BitBlt m_cMemDc.hDC, ScaleWidth - m_GapX - 2, _
                (ScaleHeight / 2) - (.Height / 2), _
                2, .Height, .hDC, 3, 0, SRCCOPY
        End With
        
        With m_Dib_HndNrm
            BitBlt m_cMemDc.hDC, _
                lHndNrmLeft, _
                (ScaleHeight / 2) - (lHndNrmHeight / 2), _
                .Width, lHndNrmHeight, .hDC, 0, lHndNrmHeight * (iState - 1), SRCCOPY
        End With
        
        If UserControl.Enabled And c_Button = 1 And _
            c_MouseX >= m_GapX And c_MouseY >= 0 And _
            c_MouseX <= ScaleWidth - m_GapX And c_MouseY <= ScaleHeight Then
            
            With m_Dib_HndDwn
                BitBlt m_cMemDc.hDC, _
                    ((ScaleWidth - m_GapX * 2) * m_TempValue) - (.Width / 2) + m_GapX, _
                    (ScaleHeight / 2) - lHndNrmHeight / 2, _
                    .Width, lHndNrmHeight, .hDC, 0, lHndNrmHeight * (iState - 1), SRCCOPY
            End With
        End If
    
    Case xpInternal
        bPen0 = CreatePen(PS_SOLID, 1, GetSysColor(COLOR_BTNTEXT))
        lBrush0 = CreateSolidBrush(GetSysColor(COLOR_BTNFACE))
        lBrush1 = CreateSolidBrush(GetSysColor(COLOR_BTNTEXT))
        lBrush2 = CreateSolidBrush(GetSysColor(COLOR_HIGHLIGHT))
        
        With m_cMemDc
            'static centre track
            rc.Left = m_GapX + 2
            rc.right = ScaleWidth - m_GapX - 2
            rc.Top = (ScaleHeight / 2) - 1
            rc.bottom = rc.Top + 2
            FillRect .hDC, rc, lBrush1
            
            If iState <> 5 Then 'if not disabled...
                SelectObject .hDC, lBrush0
                SelectObject .hDC, bPen0
                
                'hover frame
                rc.Left = m_GapX
                rc.right = ScaleWidth - m_GapX
                rc.Top = ScaleHeight / 4
                rc.bottom = rc.Top + (ScaleHeight / 2)
                If c_State > 0 Then FrameRect .hDC, rc, lBrush2
                
                'draw progress track
                rc.Left = rc.Left + 2
                rc.right = 2 + ((ScaleWidth - 4 - m_GapX * 2) * m_Value) + m_GapX
                rc.Top = rc.Top + 2
                rc.bottom = rc.bottom - 2
                FillRect .hDC, rc, lBrush1
                
                'draw track-to markers.
                If UserControl.Enabled And c_Button = 1 And _
                    c_MouseX >= m_GapX And c_MouseY >= 0 And _
                    c_MouseX <= ScaleWidth - m_GapX And c_MouseY <= ScaleHeight Then
                    
                    rc.Top = (ScaleHeight / 2) - (lHndNrmHeight / 2)
                    rc.bottom = rc.Top + lHndNrmHeight
                    rc.Left = ((ScaleWidth - m_GapX * 2) * m_TempValue) - (lHndNrmWidth / 2) + m_GapX
                    rc.right = rc.Left + lHndNrmWidth
                    
                    Dim pt(0 To 2) As POINTAPI
                    pt(0).x = rc.Left
                    pt(0).y = rc.Top
                    pt(1).x = rc.right - 1
                    pt(1).y = rc.Top
                    pt(2).x = rc.Left + (rc.right - rc.Left) / 2 - 1
                    pt(2).y = rc.Top + (rc.bottom - rc.Top) / 3
                    Polygon .hDC, pt(0), UBound(pt) + 1
                    
                    pt(0).x = rc.Left
                    pt(0).y = rc.bottom - 1
                    pt(1).x = rc.right - 1
                    pt(1).y = rc.bottom - 1
                    pt(2).x = rc.Left + (rc.right - rc.Left) / 2 - 1
                    pt(2).y = rc.bottom - (rc.bottom - rc.Top) / 3 - 1
                    Polygon .hDC, pt(0), UBound(pt) + 1
                End If
            End If
        
        End With
        
        DeleteObject lBrush0
        DeleteObject lBrush1
        DeleteObject lBrush2
        DeleteObject bPen0
    
End Select

If bRePaint Then UserControl_Paint
End Sub

'Public Function DebugInfo() As String
'DebugInfo = _
'    "usercontrol.hWnd=" & UserControl.hwnd & vbNewLine & _
'    "usercontrol.hDC=" & UserControl.hdc & vbNewLine & _
'    "m_cUxTheme.Hwnd=" & m_cUxTheme.hwnd & vbNewLine & _
'    "m_cUxTheme.hdc=" & m_cUxTheme.hdc & vbNewLine
'End Function
'
'Public Sub SetRaiseErr(e As Boolean)
'm_cUxTheme.RaiseErrors = e
'End Sub

Function UnderMouse() As Boolean
Dim ptMouse As POINTAPI
GetCursorPos ptMouse
If WindowFromPoint(ptMouse.x, ptMouse.y) = UserControl.hWND Then
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

Public Property Get DrawMode() As expDrawMode
DrawMode = m_DrawMode
End Property
Public Property Let DrawMode(d As expDrawMode)
m_DrawMode = d
Redraw
PropertyChanged "DrawMode"
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

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
UserControl_Click
End Sub

Private Sub UserControl_Click()
'tmrMouse.Enabled = False
'c_State = 0
'RaiseEvent Click
'tmrMouse.Enabled = True
End Sub

Private Sub UserControl_DblClick()
c_State = 0
End Sub

Private Sub UserControl_Initialize()
'Set m_cUxTheme = New cUxTheme
'With m_cUxTheme
'    .hwnd = UserControl.hwnd
'    .TextAlign = DT_CENTERABS
'End With

Set m_cMemDc = New pcMemDCSld

Set m_Dib_Track = New cDIBSection
Set m_Dib_HndNrm = New cDIBSection
Set m_Dib_HndDwn = New cDIBSection
m_Dib_Track.CreateFromPicture imgTrack.Picture
m_Dib_HndNrm.CreateFromPicture imgHndNrm.Picture
m_Dib_HndDwn.CreateFromPicture imgHndDwn.Picture

m_GapX = 100 / Screen.TwipsPerPixelX

m_Value = 0
m_DrawMode = xpInternal

c_State = 0
c_OldState = -1
c_Button = 0

Redraw False
End Sub

Private Sub UserControl_InitProperties()
UserControl_Resize
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
c_Button = Button

Select Case Button
    Case 1
        c_State = 2
        m_TempValue = (x - m_GapX) / (ScaleWidth - m_GapX * 2)
        Redraw
End Select

RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
c_Button = Button

c_MouseX = x
c_MouseY = y

If Button > 0 Then
    m_TempValue = (x - m_GapX) / (ScaleWidth - m_GapX * 2)
    Redraw
End If

If tmrMouse.Enabled <> True Then tmrMouse.Enabled = True
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
c_Button = 0

If Button = 1 And x >= 0 And y >= 0 And x <= ScaleWidth And y <= ScaleHeight Then
    'just in case (and to help fix a bug where maximizing by double-clicking
    '   the title bar results in an accidental "mouse-up" event for the slider.)
    m_TempValue = (x - m_GapX) / (ScaleWidth - m_GapX * 2)
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
m_DrawMode = PropBag.ReadProperty("DrawMode", xpInternal)

Redraw
End Sub

Private Sub UserControl_Terminate()
'Set m_cUxTheme = Nothing
Set m_cMemDc = Nothing
Set m_Dib_Track = Nothing
Set m_Dib_HndNrm = Nothing
Set m_Dib_HndDwn = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "Enabled", UserControl.Enabled
PropBag.WriteProperty "DrawMode", m_DrawMode
End Sub
