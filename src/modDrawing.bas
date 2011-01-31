Attribute VB_Name = "modDrawing"

Option Explicit
Option Compare Binary
Option Base 0

'bitblt stuff
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Public Const SRCCOPY = &HCC0020
Public Const SRCAND = &H8800C6
Public Const SRCPAINT = &HEE0086
Public Const SRCINVERT = &H660046
Public Const SRCERASE = &H440328
Public Const NOTSRCCOPY = &H330008
Public Const NOTSRCERASE = &H1100A6
Public Const MERGECOPY = &HC000CA
Public Const MERGEPAINT = &HBB0226
Public Const PATCOPY = &HF00021
Public Const PATPAINT = &HFB0A09
Public Const PATINVERT = &H5A0049
Public Const DSTINVERT = &H550009
Public Const BLACKNESS = &H0
Public Const WHITENESS = &HFFFFFF

'pens
Public Const PS_DOT = 2
Public Const PS_SOLID = 0
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

'TextDrawing
Public Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkMode Lib "gdi32.dll" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
    Public Const OPAQUE As Long = 2
    Public Const TRANSPARENT As Long = 1
Public Declare Function SetBkColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long

Public Declare Function CreateFont Lib "gdi32" _
    Alias "CreateFontA" ( _
    ByVal h As Long, _
    ByVal w As Long, _
    ByVal e As Long, _
    ByVal O As Long, _
    ByVal w As Long, _
    ByVal i As Long, _
    ByVal u As Long, _
    ByVal s As Long, _
    ByVal c As Long, _
    ByVal OP As Long, _
    ByVal CP As Long, _
    ByVal Q As Long, _
    ByVal PAF As Long, _
    ByVal f As String) As Long

Public Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
'Const LOGPIXELSY = 90

Public Const FW_BOLD As Long = 700
Public Const FW_DONTCARE As Long = 0
Public Const FW_EXTRABOLD As Long = 800
Public Const FW_EXTRALIGHT As Long = 200
Public Const FW_HEAVY As Long = 900
Public Const FW_LIGHT As Long = 300
Public Const FW_MEDIUM As Long = 500
Public Const FW_NORMAL As Long = 400
Public Const FW_SEMIBOLD As Long = 600
Public Const FW_THIN As Long = 100

Public Const DEFAULT_CHARSET As Long = 1

'Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
'
'Public Const DT_LEFT = &H0
'Public Const DT_CENTER = &H1
'Public Const DT_RIGHT = &H2
'
'Public Const DT_TOP = &H0
'Public Const DT_VCENTER = &H4
'Public Const DT_BOTTOM = &H8
'
'Public Const DT_CENTERABS = &H65
'
'Public Const DT_WORDBREAK = &H10
'Public Const DT_SINGLELINE = &H20
'Public Const DT_NOCLIP = &H100
'Public Const DT_CALCRECT = &H400
'Public Const DT_NOPREFIX = &H800

'fill
Public Declare Function FillRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long

'Boxes
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function FillRect Lib "User32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function FrameRect Lib "User32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function DrawFocusRect Lib "User32" (ByVal hDC As Long, lpRect As RECT) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long

'lines
Public Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Public Type RECT
    Left As Long
    Top As Long
    right As Long
    bottom As Long
End Type
Public Type POINTAPI
    x As Long
    y As Long
End Type
'points
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" Alias "SetPixelV" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

'==============================================================================
Public Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal oleColor As OLE_COLOR, ByVal hPalette As Long, pColorRef As Long) As Long
'system colours================================================================
Public Declare Function GetSysColor Lib "User32" (ByVal nIndex As Long) As Long
Public Const COLOR_SCROLLBAR = 0
Public Const COLOR_BACKGROUND = 1
Public Const COLOR_ACTIVECAPTION = 2
Public Const COLOR_INACTIVECAPTION = 3
Public Const COLOR_MENU = 4
Public Const COLOR_WINDOW = 5
Public Const COLOR_WINDOWFRAME = 6
Public Const COLOR_MENUTEXT = 7
Public Const COLOR_WINDOWTEXT = 8
Public Const COLOR_CAPTIONTEXT = 9
Public Const COLOR_ACTIVEBORDER = 10
Public Const COLOR_INACTIVEBORDER = 11
Public Const COLOR_APPWORKSPACE = 12
Public Const COLOR_HIGHLIGHT = 13
Public Const COLOR_HIGHLIGHTTEXT = 14
Public Const COLOR_BTNFACE = 15
Public Const COLOR_BTNSHADOW = 16
Public Const COLOR_GRAYTEXT = 17
Public Const COLOR_BTNTEXT = 18
Public Const COLOR_INACTIVECAPTIONTEXT = 19
Public Const COLOR_BTNHIGHLIGHT = 20
Public Const COLOR_BTNDKSHADOW = 21
Public Const COLOR_BTNLIGHT = 22
'==============================================================================

Public Sub SetRect(rc As RECT, l As Long, t As Long, r As Long, b As Long)
rc.Left = l
rc.Top = t
rc.right = r
rc.bottom = b
End Sub

Public Sub SetRectHW(rc As RECT, l As Long, t As Long, w As Long, h As Long)
rc.Left = l
rc.Top = t
rc.right = l + w
rc.bottom = t + h
End Sub

Public Sub ShrinkRect(rc As RECT, x As Long)
rc.Left = rc.Left + x
rc.Top = rc.Top + x
rc.right = rc.right - x
rc.bottom = rc.bottom - x
End Sub

Public Sub DrawCtlBtm( _
    lSrcHdc As Long, lTgtHdc As Long, _
    lSrcX As Long, lSrcY As Long, _
    lTgtX As Long, lTgtY As Long, _
    lSrcW As Long, lSrcH As Long, _
    lTgtW As Long, lTgtH As Long, _
    lCornerSize As Long)

'Corners
    'top left
BitBlt lTgtHdc, lTgtX, lTgtY, lCornerSize, lCornerSize, lSrcHdc, lSrcX, lSrcY, SRCCOPY

    'bottom left
BitBlt lTgtHdc, lTgtX, lTgtY + lTgtH - lCornerSize, _
    lCornerSize, lCornerSize, _
    lSrcHdc, lSrcX, lSrcY + lSrcH - lCornerSize, SRCCOPY

    'top right
BitBlt lTgtHdc, lTgtX + lTgtW - lCornerSize, lTgtY, lCornerSize, lCornerSize, _
    lSrcHdc, lSrcW - lCornerSize, lSrcY, SRCCOPY

    'bottom right
BitBlt lTgtHdc, _
    lTgtX + lTgtW - lCornerSize, _
    lTgtY + lTgtH - lCornerSize, _
    lCornerSize, lCornerSize, _
    lSrcHdc, _
    lSrcX + lSrcW - lCornerSize, _
    lSrcY + lSrcH - lCornerSize, SRCCOPY

'edges:
    'top edge
StretchBlt lTgtHdc, lTgtX + lCornerSize, lTgtY, _
    lTgtW - (lCornerSize * 2), lCornerSize, _
    lSrcHdc, lCornerSize, _
    lSrcY, lSrcW - (lCornerSize * 2), lCornerSize, SRCCOPY

    'right edge
StretchBlt lTgtHdc, _
    lTgtX + lCornerSize, lTgtY + lTgtH - lCornerSize, _
    lTgtW - (lCornerSize * 2), lCornerSize, _
    lSrcHdc, _
    lSrcX + lCornerSize, lSrcY + lSrcH - lCornerSize, _
    lSrcW - (lCornerSize * 2), lCornerSize, SRCCOPY

    'left edge
StretchBlt lTgtHdc, _
    lTgtX, lTgtY + lCornerSize, _
    lCornerSize, lTgtH - (lCornerSize * 2), _
    lSrcHdc, _
    lSrcX, lSrcY + lCornerSize, _
    lCornerSize, lSrcH - (lCornerSize * 2), SRCCOPY

    'botton edge
StretchBlt _
    lTgtHdc, lTgtX + lTgtW - lCornerSize, lTgtY + lCornerSize, _
    lCornerSize, lTgtH - (lCornerSize * 2), _
    lSrcHdc, _
    lSrcW - lCornerSize, lSrcY + lCornerSize, _
    lCornerSize, lSrcH - (lCornerSize * 2), SRCCOPY

'middle
SetStretchBltMode lTgtHdc, vbPaletteModeNone
StretchBlt lTgtHdc, _
    lTgtX + lCornerSize, lTgtY + lCornerSize, _
    lTgtW - (lCornerSize * 2), lTgtH - (lCornerSize * 2), _
    lSrcHdc, _
    lSrcX + lCornerSize, lSrcY + lCornerSize, _
     lSrcW - (lCornerSize * 2), lSrcH - (lCornerSize * 2), SRCCOPY
End Sub
