Attribute VB_Name = "modUnicodeDraw"
'various bits of code probabbly from vbaccelerator.com again.
'can't think where else i would have found them.

Option Explicit
Option Compare Binary
Option Base 0

Private Type OSVERSIONINFO
   dwVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion(0 To 127) As Byte
End Type
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInfo As OSVERSIONINFO) As Long
Private Const VER_PLATFORM_WIN32_NT = 2
Private m_bIsNt As Boolean

Private Declare Function DrawTextA Lib "User32" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawTextW Lib "User32" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Public Const DT_CENTERABS = &H65

Public Enum DrawTextFlags
    DT_TOP = &H0
    dt_left = &H0
    DT_CENTER = &H1
    DT_RIGHT = &H2
    DT_VCENTER = &H4
    DT_BOTTOM = &H8
    DT_WORDBREAK = &H10
    DT_SINGLELINE = &H20
    DT_EXPANDTABS = &H40
    DT_TABSTOP = &H80
    DT_NOCLIP = &H100
    DT_EXTERNALLEADING = &H200
    DT_CALCRECT = &H400
    DT_NOPREFIX = &H800
    DT_INTERNAL = &H1000
    DT_EDITCONTROL = &H2000
    DT_PATH_ELLIPSIS = &H4000
    DT_END_ELLIPSIS = &H8000&
    DT_MODIFYSTRING = &H10000
    DT_RTLREADING = &H20000
    DT_WORD_ELLIPSIS = &H40000
    DT_NOFULLWIDTHCHARBREAK = &H80000
    DT_HIDEPREFIX = &H100000
    DT_PREFIXONLY = &H200000
End Enum

Public Sub VerInitialise()
Dim tOSV As OSVERSIONINFO
tOSV.dwVersionInfoSize = Len(tOSV)
GetVersionEx tOSV

m_bIsNt = ((tOSV.dwPlatformId And VER_PLATFORM_WIN32_NT) = VER_PLATFORM_WIN32_NT)
End Sub
Public Sub DrawText( _
    ByVal lhDC As Long, ByVal sText As String, ByVal lLength As Long, _
    tR As RECT, ByVal lFlags As Long)

Dim LPTR As Long

If (m_bIsNt) Then
   LPTR = StrPtr(sText)
   If Not (LPTR = 0) Then ' NT4 crashes with ptr = 0
      DrawTextW lhDC, LPTR, -1, tR, lFlags
   End If
Else
   DrawTextA lhDC, sText, -1, tR, lFlags
End If
End Sub

'Purpose:Returns True if string has a Unicode char.
Public Function IsUnicode(s As String) As Boolean
   Dim i As Long
   Dim bLen As Long
   Dim Map() As Byte

   If LenB(s) Then
      Map = s
      bLen = UBound(Map)
      For i = 1 To bLen Step 2
         If (Map(i) > 0) Then
            IsUnicode = True
            Exit Function
         End If
      Next
   End If
End Function
