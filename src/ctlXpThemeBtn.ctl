VERSION 5.00
Begin VB.UserControl ctlXpThemeBtn 
   CanGetFocus     =   0   'False
   ClientHeight    =   1690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3250
   DefaultCancel   =   -1  'True
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   169
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   325
   Begin VB.Timer tmrMouse 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Image imgBtnSplitDrop 
      Height          =   460
      Left            =   1200
      Picture         =   "ctlXpThemeBtn.ctx":0000
      Top             =   120
      Visible         =   0   'False
      Width           =   110
   End
   Begin VB.Image imgBtnSplit 
      Height          =   460
      Left            =   840
      Picture         =   "ctlXpThemeBtn.ctx":066A
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgBtnNrm 
      Height          =   460
      Left            =   480
      Picture         =   "ctlXpThemeBtn.ctx":0EFC
      Top             =   120
      Visible         =   0   'False
      Width           =   200
   End
End
Attribute VB_Name = "ctlXpThemeBtn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Option Compare Binary
Option Base 0

Private Const arrow_W As Long = 7

Private Declare Function GetCursorPos Lib "User32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "User32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
'Private Type POINTAPI
'    x As Long
'    y As Long
'End Type

Dim m_Caption As String, _
    m_Type As eXpTmBtType, _
    m_DrawMode As expDrawMode

Dim c_State As Long
'   0=normal
'   1=hover over
'   2=depressed
Dim c_Button As Long '= is mouse down?
Dim c_OldState As Long
Dim c_MouseOverPart As Long '0 = main, 1 = side
Dim c_MouseOverPartOld As Long

Const gdi_lRectCorner As Long = 5
Const gdi_lBtmpCorner As Long = 4
Dim lSidePartWidth As Long

Dim m_Dib_BtnNrm As cDIBSection, _
    m_Dib_BtnSplit As cDIBSection, _
    m_Dib_BtnSplitDrop As cDIBSection

Private m_cMemDc As pcMemDC

Private gdi_hNrmFilB As Long
Private gdi_hHigFilB As Long
Private gdi_hLinePen As Long
Private gdi_hArrowPenNrm As Long
Private gdi_hArrowPenHig As Long
Private gdi_hFontNormal As Long
Private gdi_lFontHeight As Long

Event Click()
Event ClickSide()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single, lW As Long, lH As Long)
Event AfterRedraw(lHdc As Long, lW As Long, lH As Long, lForeColour As Long)
Event AfterResize(lW As Long, lH As Long)

Private Sub gdi_Make()
Dim c As Long
OleTranslateColor BackColor, 0, c
gdi_hNrmFilB = CreateSolidBrush(c)
gdi_hHigFilB = CreateSolidBrush(GetSysColor(COLOR_HIGHLIGHT))
gdi_hLinePen = CreatePen(PS_SOLID, 1, GetSysColor(COLOR_HIGHLIGHT))
gdi_hArrowPenNrm = CreatePen(PS_SOLID, 1, GetSysColor(COLOR_BTNTEXT))
gdi_hArrowPenHig = CreatePen(PS_SOLID, 1, GetSysColor(COLOR_HIGHLIGHTTEXT))
gdi_lFontHeight = -MulDiv(UserControl.Font.Size, GetDeviceCaps(m_cMemDc.hDC, 90), 72)
gdi_hFontNormal = CreateFont(gdi_lFontHeight, _
    0, 0, 0, _
    IIf(Font.Bold, FW_BOLD, FW_DONTCARE), 0, 0, 0, _
    DEFAULT_CHARSET, _
    0, 0, 0, 0, UserControl.Font.Name)
End Sub

Private Sub gdi_Destroy()
DeleteObject gdi_hNrmFilB
DeleteObject gdi_hHigFilB
DeleteObject gdi_hLinePen
DeleteObject gdi_hArrowPenNrm
DeleteObject gdi_hArrowPenHig
DeleteObject gdi_hFontNormal
End Sub

Sub gdi_Clear()
Dim b As Long, rc As RECT, c As Long
rc.Left = 0
rc.Top = 0
rc.right = ScaleWidth
rc.bottom = ScaleHeight
'b = CreateSolidBrush(GetSysColor(COLOR_BTNFACE))
OleTranslateColor BackColor, 0, c
b = CreateSolidBrush(c)
FillRect m_cMemDc.hDC, rc, b
DeleteObject b
End Sub

Sub ProcClick()
tmrMouse.Enabled = False
c_State = 0
If c_MouseOverPart = 1 And m_Type = btyp_DropDown Then
    RaiseEvent ClickSide
Else
    RaiseEvent Click
End If
tmrMouse.Enabled = True
End Sub

Sub SimClick(Optional DropDn As Boolean = False)
c_State = 2
c_MouseOverPart = IIf(DropDn, 1, 0)
Redraw
ProcClick
c_State = 0
Redraw
End Sub

Sub Redraw(Optional bRePaint As Boolean = True, Optional bCheckOldState As Boolean = False)
If c_State = c_OldState And c_MouseOverPart = c_MouseOverPartOld And bCheckOldState Then
    Exit Sub
Else
    c_OldState = c_State
    c_MouseOverPartOld = c_MouseOverPart
    'Debug.Print Format(Now, "hh:mm:ss") & " draw."
End If

Dim c As Long, _
    rc As RECT, pt As POINTAPI, _
    cDibSrc As cDIBSection, i As Long, _
    bMainDn As Boolean, bSideDn As Boolean, _
    lForeCol As Long

Dim rcWhole As RECT 'entire control.
Dim rcText As RECT 'text label location.
Dim rcMainBtn As RECT 'area for primary button.
Dim rcSideBtn As RECT 'area for drop-down button.

'some basic drawing stuff
SelectObject m_cMemDc.hDC, gdi_hFontNormal

'various areas need setting up
SetRectHW rcWhole, 0, 0, ScaleWidth, ScaleHeight

Select Case m_Type
    Case btyp_Normal
        Set cDibSrc = m_Dib_BtnNrm
        rcMainBtn = rcWhole
        rcText = rcWhole
    
    Case btyp_DropArrow
        Set cDibSrc = m_Dib_BtnNrm
        rcMainBtn = rcWhole
        rcText = rcWhole
        rcText.right = rcText.right - lSidePartWidth
        rcSideBtn = rcWhole
        rcSideBtn.Left = rcSideBtn.right - lSidePartWidth
    
    Case btyp_DropDown
        Set cDibSrc = m_Dib_BtnSplit
        rcMainBtn = rcWhole
        rcMainBtn.right = rcMainBtn.right - lSidePartWidth
        rcSideBtn = rcWhole
        rcSideBtn.Left = rcMainBtn.right
        rcText = rcMainBtn
    
End Select

'what is down?
bMainDn = False
bSideDn = False
If c_State = 2 Then
    If m_Type = btyp_DropDown Then
        bSideDn = (c_MouseOverPart = 1)
        bMainDn = Not bSideDn
    Else
        bMainDn = True
    End If
End If

'start drawing
Select Case m_DrawMode
    Case xpBitmap
        Select Case c_State
            Case 0
                gdi_Clear
            
            Case 1, 2
                'draw body;
                With cDibSrc
                    'draw the button body
                    DrawCtlBtm .hDC, m_cMemDc.hDC, _
                        0, (.Height / 2) * IIf(bMainDn, 1, 0), _
                        rcMainBtn.Left, rcMainBtn.Top, _
                        .Width, .Height / 2, _
                        rcMainBtn.right - rcMainBtn.Left, _
                        rcMainBtn.bottom - rcMainBtn.Top, _
                        gdi_lBtmpCorner
                End With
                
                'draw drop down end?
                If m_Type = btyp_DropArrow Then
                    'arrow is drawn later
                ElseIf m_Type = btyp_DropDown Then
                    With m_Dib_BtnSplitDrop
                        DrawCtlBtm .hDC, m_cMemDc.hDC, _
                            0, (.Height / 2) * IIf(bSideDn, 1, 0), _
                            rcSideBtn.Left, rcSideBtn.Top, _
                            .Width, (.Height / 2), _
                            rcSideBtn.right - rcSideBtn.Left, _
                            rcSideBtn.bottom - rcSideBtn.Top, _
                            gdi_lBtmpCorner
                    End With
                End If
                
                'corner pixels
                OleTranslateColor BackColor, 0, c
                SetPixel m_cMemDc.hDC, rcWhole.Left, rcWhole.Top, c
                SetPixel m_cMemDc.hDC, rcWhole.Left, rcWhole.bottom - rcWhole.Top - 1, c
                SetPixel m_cMemDc.hDC, rcWhole.right - rcWhole.Left - 1, 0, c
                SetPixel m_cMemDc.hDC, rcWhole.right - rcWhole.Left - 1, _
                    rcWhole.bottom - rcWhole.Top - 1, c
            
        End Select
        
        lForeCol = GetSysColor(COLOR_BTNTEXT)
        'SetTextColor m_cMemDc.hDC, GetSysColor(COLOR_BTNTEXT)
    
    Case xpInternal
        Select Case c_State
            Case 0
                gdi_Clear
            
            Case 1, 2
                With m_cMemDc
                    'set drawing style
                    SelectObject .hDC, gdi_hLinePen
                    
                    'draw side btn
                    If m_Type = btyp_DropDown Then
                        SelectObject .hDC, IIf(bSideDn, gdi_hHigFilB, gdi_hNrmFilB)
                        RoundRect .hDC, _
                                rcMainBtn.Left, rcSideBtn.Top, _
                                rcSideBtn.right, rcSideBtn.bottom, _
                                gdi_lRectCorner, gdi_lRectCorner
                    End If
                    
                    'draw main button
                    SelectObject .hDC, IIf(bMainDn, gdi_hHigFilB, gdi_hNrmFilB)
                    RoundRect .hDC, _
                        rcMainBtn.Left, rcMainBtn.Top, _
                        rcMainBtn.right, rcMainBtn.bottom, _
                        gdi_lRectCorner, gdi_lRectCorner
                End With
            
        End Select
        
        'decide on the main text colour.
        If (c_State <> 2) Or (m_Type = btyp_DropDown And c_MouseOverPart = 1) Then
            'SetTextColor m_cMemDc.hDC, GetSysColor(COLOR_BTNTEXT)
            lForeCol = GetSysColor(COLOR_BTNTEXT)
        Else
            'SetTextColor m_cMemDc.hDC, GetSysColor(COLOR_HIGHLIGHTTEXT)
            lForeCol = GetSysColor(COLOR_HIGHLIGHTTEXT)
        End If
    
End Select

'draw arrow and text.
If m_Type <> btyp_Normal Then
    'draw the arrow.
    rc.Left = rcSideBtn.Left + ((rcSideBtn.right - rcSideBtn.Left) / 2) - (arrow_W / 2) - 1
    rc.Top = ((rcSideBtn.bottom - rcSideBtn.Top) / 2) - 1
    rc.right = rc.Left + arrow_W
    rc.bottom = rc.Top + arrow_W - 1
    
    With m_cMemDc
        If (bSideDn Or (bMainDn And m_Type = btyp_DropArrow)) And m_DrawMode = xpInternal Then
        'If c_State = 2 And (c_MouseOverPart = 1 Or m_Type <> btyp_DropDown) Then
            SelectObject .hDC, gdi_hArrowPenHig
        Else
            SelectObject .hDC, gdi_hArrowPenNrm
        End If
        For i = 0 To Fix(arrow_W / 2)
            MoveToEx .hDC, rc.Left + i, rc.Top + i, pt
            LineTo .hDC, rc.right - i, rc.Top + i
        Next i
    End With
End If

'draw text:
SetTextColor m_cMemDc.hDC, lForeCol
SetBkMode m_cMemDc.hDC, TRANSPARENT
DrawText m_cMemDc.hDC, m_Caption, Len(m_Caption), rcText, DT_CENTERABS

'finish up.
'note: pass rcText as this is basicly the caption area.
RaiseEvent AfterRedraw(m_cMemDc.hDC, _
    rcText.right - rcText.Left, _
    rcText.bottom - rcText.Top, _
    lForeCol)
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

Private Sub SetAccessKey()
Dim lAmpPos As Long

If Len(m_Caption) < 2 Then
    UserControl.AccessKeys = ""
    Exit Sub
End If

lAmpPos = InStr(1, m_Caption, "&", vbTextCompare)

If (lAmpPos >= Len(m_Caption)) Or (lAmpPos <= 0) Then
    UserControl.AccessKeys = ""
End If

If Mid$(m_Caption, lAmpPos + 1, 1) <> "&" Then
    UserControl.AccessKeys = LCase$(Mid$(m_Caption, lAmpPos + 1, 1))
    Exit Sub
End If

lAmpPos = InStr(lAmpPos + 2, m_Caption, "&", vbTextCompare)

If Mid$(m_Caption, lAmpPos + 1, 1) <> "&" Then
    UserControl.AccessKeys = LCase$(Mid$(m_Caption, lAmpPos + 1, 1))
Else
    UserControl.AccessKeys = ""
End If
End Sub

Public Property Get DrawMode() As expDrawMode
DrawMode = m_DrawMode
End Property
Public Property Let DrawMode(d As expDrawMode)
m_DrawMode = d
Redraw
PropertyChanged "DrawMode"
End Property

Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = -518
Caption = m_Caption
End Property
Public Property Let Caption(a As String)
m_Caption = a
Redraw
SetAccessKey
PropertyChanged "Caption"
End Property

Public Property Get BtnType() As eXpTmBtType
BtnType = m_Type
End Property
Public Property Let BtnType(NewType As eXpTmBtType)
m_Type = NewType
Redraw
PropertyChanged "BtnType"
End Property

Public Property Get Font() As StdFont
Set Font = UserControl.Font
End Property
Public Property Set Font(ByVal NewFont As StdFont)
Set UserControl.Font = NewFont
gdi_Destroy
gdi_Make
Redraw
PropertyChanged "Font"
End Property

Public Property Get Enabled() As Boolean
Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(e As Boolean)
UserControl.Enabled = e
Redraw
PropertyChanged "Enabled"
End Property

Public Property Get BackColour() As OLE_COLOR
BackColour = BackColor
End Property
Public Property Let BackColour(c As OLE_COLOR)
UserControl.BackColor = c
gdi_Destroy
gdi_Make
Redraw
PropertyChanged "BackColour"
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
ProcClick
End Sub

Private Sub UserControl_DblClick()
c_State = 0
End Sub

Private Sub UserControl_Initialize()
Set m_cMemDc = New pcMemDC

'init internal bmp's
Set m_Dib_BtnNrm = New cDIBSection
Set m_Dib_BtnSplit = New cDIBSection
Set m_Dib_BtnSplitDrop = New cDIBSection
m_Dib_BtnNrm.CreateFromPicture imgBtnNrm.Picture
m_Dib_BtnSplit.CreateFromPicture imgBtnSplit.Picture
m_Dib_BtnSplitDrop.CreateFromPicture imgBtnSplitDrop.Picture

c_State = 0
c_OldState = -1
c_MouseOverPartOld = -1
c_Button = 0

lSidePartWidth = m_Dib_BtnSplitDrop.Width
m_DrawMode = xpInternal
gdi_Make
End Sub

Private Sub UserControl_InitProperties()
m_Caption = "[]"
m_Type = btyp_Normal
Set Font = Ambient.Font

UserControl_Resize
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
c_Button = Button

Select Case Button
    Case 1
        c_State = 2
        Redraw
End Select

RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
c_Button = Button

If X > ScaleWidth - lSidePartWidth Then
    c_MouseOverPart = 1
Else
    c_MouseOverPart = 0
End If

If tmrMouse.Enabled <> True Then tmrMouse.Enabled = True
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
c_Button = 0

RaiseEvent MouseUp(Button, Shift, X, Y, ScaleWidth, ScaleHeight)

If Button = 1 And X >= 0 And X <= ScaleWidth And Y >= 0 And Y <= ScaleHeight Then
    ProcClick
End If
End Sub

Private Sub UserControl_Paint()
m_cMemDc.Draw hDC, 0, 0, m_cMemDc.Width, m_cMemDc.Height, 0, 0
End Sub

Private Sub UserControl_Resize()
m_cMemDc.Width = ScaleWidth
m_cMemDc.Height = ScaleHeight

Redraw False

RaiseEvent AfterResize(ScaleWidth * Screen.TwipsPerPixelX, ScaleHeight * Screen.TwipsPerPixelY)

UserControl_Paint
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
m_DrawMode = PropBag.ReadProperty("DrawMode", xpInternal)
m_Caption = PropBag.ReadProperty("Caption", "[]")
BtnType = PropBag.ReadProperty("BtnType", 0)
Set Font = PropBag.ReadProperty("Font", Ambient.Font)
UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
UserControl.BackColor = PropBag.ReadProperty("BackColour", BackColor)

Redraw
End Sub

Private Sub UserControl_Terminate()
gdi_Destroy
Set m_cMemDc = Nothing
Set m_Dib_BtnNrm = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "DrawMode", m_DrawMode
PropBag.WriteProperty "Caption", m_Caption
PropBag.WriteProperty "BtnType", m_Type
PropBag.WriteProperty "Font", Font
PropBag.WriteProperty "Enabled", UserControl.Enabled
PropBag.WriteProperty "BackColour", UserControl.BackColor
End Sub
