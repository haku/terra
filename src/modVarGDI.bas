Attribute VB_Name = "modVarGdi"
Option Explicit

Public gdi_Main_Brush(0 To 2) As Long, _
    gdi_Main_hFontNormal As Long, _
    gdi_Main_hFontNormalStrike As Long, _
    gdi_Main_hFontNormalStrikeItalic As Long, _
    gdi_Main_hFontSel As Long, _
    gdi_Main_hFontSelStrike As Long, _
    gdi_Main_hFontDouble As Long, _
    gdi_Main_hFontTripple As Long
'

Public Sub gdi_Main_MakeObjects(hDC As Long, Font As StdFont)
Dim nFontHeight As Long

gdi_Main_Brush(0) = CreateSolidBrush(GetSysColor(COLOR_WINDOW))
gdi_Main_Brush(1) = CreateSolidBrush(GetSysColor(COLOR_HIGHLIGHT))
gdi_Main_Brush(2) = CreateSolidBrush(GetSysColor(COLOR_BTNTEXT))

nFontHeight = -MulDiv(Font.Size, GetDeviceCaps(hDC, 90), 72)

gdi_Main_hFontNormal = CreateFont(nFontHeight, _
    0, 0, 0, _
    FW_DONTCARE, 0, 0, 0, _
    DEFAULT_CHARSET, _
    0, 0, 0, 0, Font.Name)

gdi_Main_hFontNormalStrike = CreateFont(nFontHeight, _
    0, 0, 0, _
    FW_DONTCARE, 0, 0, 1, _
    DEFAULT_CHARSET, _
    0, 0, 0, 0, Font.Name)

gdi_Main_hFontNormalStrikeItalic = CreateFont(nFontHeight, _
    0, 0, 0, _
    FW_DONTCARE, 1, 0, 1, _
    DEFAULT_CHARSET, _
    0, 0, 0, 0, Font.Name)

gdi_Main_hFontSel = CreateFont(nFontHeight, _
    0, 0, 0, _
    FW_BOLD, 0, 0, 0, _
    DEFAULT_CHARSET, _
    0, 0, 0, 0, Font.Name)

gdi_Main_hFontSelStrike = CreateFont(nFontHeight, _
    0, 0, 0, _
    FW_BOLD, 0, 0, 1, _
    DEFAULT_CHARSET, _
    0, 0, 0, 0, Font.Name)

gdi_Main_hFontDouble = CreateFont(nFontHeight * 2, _
    0, 0, 0, _
    FW_DONTCARE, 0, 0, 0, _
    DEFAULT_CHARSET, _
    0, 0, 0, 0, Font.Name)

gdi_Main_hFontTripple = CreateFont(nFontHeight * 3, _
    0, 0, 0, _
    FW_DONTCARE, 0, 0, 0, _
    DEFAULT_CHARSET, _
    0, 0, 0, 0, Font.Name)
End Sub

Public Sub gdi_Main_DeleteObjects()
Dim i As Long
For i = 0 To UBound(gdi_Main_Brush)
    DeleteObject gdi_Main_Brush(i)
Next i

DeleteObject gdi_Main_hFontNormal
DeleteObject gdi_Main_hFontNormalStrike
DeleteObject gdi_Main_hFontNormalStrikeItalic
DeleteObject gdi_Main_hFontSel
DeleteObject gdi_Main_hFontSelStrike
DeleteObject gdi_Main_hFontDouble
DeleteObject gdi_Main_hFontTripple
End Sub

