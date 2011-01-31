Attribute VB_Name = "modSystemMenu"
'this adds some extra items to the form's system menu which would
'other wise clutter the GUI.
'note the use of InsertMenu and NOT AppendMenu.  adding items to
'the bottom of the system just winds people up with they
'right-click in the taskbar.

Option Explicit
Option Compare Binary
Option Base 0

Dim MenuHandle As Long

Dim Checked As Boolean

Public OldProc As Long

Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWND As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hWND As Long, ByVal nIndex As Long) As Long
Public Const GWL_WNDPROC = (-4)

Declare Function CallWindowProc Lib "User32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWND As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_SYSCOMMAND = &H112 'The window message to monitor

'menu API's
Declare Function GetSubMenu Lib "User32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemID Lib "User32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetSystemMenu Lib "User32" (ByVal hWND As Long, ByVal bRevert As Long) As Long
'Declare Function AppendMenu Lib "User32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Declare Function InsertMenu Lib "User32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Declare Function DrawMenuBar Lib "User32" (ByVal hWND As Long) As Long
Declare Function CheckMenuItem Lib "User32" (ByVal hMenu As Long, ByVal wIDCheckItem As Long, ByVal wCheck As Long) As Long
'Public Const MF_SEPARATOR = &H800&
Public Const MF_CHECKED = &H8&
'Public Const MF_UNCHECKED = &H0&

Const MF_APPEND = &H100&
Const MF_BYCOMMAND = &H0&
Const MF_BYPOSITION = &H400&
Const MF_DEFAULT = &H1000&
Const MF_DISABLED = &H2&
Const MF_ENABLED = &H0&
Const MF_GRAYED = &H1&
Const MF_MENUBARBREAK = &H20&
Const MF_MENUBREAK = &H40&
Const MF_OWNERDRAW = &H100&
Const MF_POPUP = &H10&
Const MF_REMOVE = &H1000&
Const MF_SEPARATOR = &H800&
Const MF_STRING = &H0&
Const MF_UNCHECKED = &H0&
Const MF_BITMAP = &H4&
Const MF_USECHECKBITMAPS = &H200&

'Window Positioning API
Declare Function SetWindowPos Lib "user32.dll" (ByVal hWND As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1

'Used to set window to always be on top or not
Const HWND_NOTOPMOST = -2
Const HWND_TOPMOST = -1

Public Function WndProc(ByVal hWND As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim retval As Long

'Is triggered if Always on top is clicked.
If wMsg = WM_SYSCOMMAND And wParam = 555 Then
    WndProc = 0
    If Checked Then
        'switch menu to unchecked
        retval = CheckMenuItem(MenuHandle, 555, MF_UNCHECKED)
        'set window to not top most window
        retval = SetWindowPos(hWND, HWND_NOTOPMOST, 0, 0, 1, 1, SWP_NOMOVE Or SWP_NOSIZE)
        'toggle checked
        Checked = Not Checked
    Else
        'switch menu to checked
        retval = CheckMenuItem(MenuHandle, 555, MF_CHECKED)
        'make window always on top
        retval = SetWindowPos(hWND, HWND_TOPMOST, 0, 0, 1, 1, SWP_NOMOVE Or SWP_NOSIZE)
        'toggle checked
        Checked = Not Checked
    End If
ElseIf wMsg = WM_SYSCOMMAND And wParam = 556 Then
    Main_About
ElseIf wMsg = WM_SYSCOMMAND And wParam = 557 Then
    frmMain.tray_Min
'ElseIf wMsg = WM_SYSCOMMAND And wParam = WM_HOTKEY Then
'    Debug.Print "hotkey!"
'ElseIf wMsg = &H20A Then
'    Debug.Print "wheel!"
Else
    'Pass on all the other unhandled messages
    WndProc = CallWindowProc(OldProc, hWND, wMsg, wParam, lParam)
End If
End Function

Public Sub AddMenuItem(hWND As Long)
Checked = False

'Get system menu handle
MenuHandle = GetSystemMenu(hWND, False)

'note that these are in reverse order.
InsertMenu MenuHandle, 0, MF_SEPARATOR Or MF_BYPOSITION, 0, ""
InsertMenu MenuHandle, 0, 0&, 556, "&about"
InsertMenu MenuHandle, 0, MF_UNCHECKED, 555, "keep on &top"
InsertMenu MenuHandle, 0, MF_UNCHECKED, 557, "minimize to tra&y"

'Redraw the menubar
DrawMenuBar hWND

'store the old message handler.
OldProc = GetWindowLong(hWND, GWL_WNDPROC)

'set the message handler to ours.
SetWindowLong hWND, GWL_WNDPROC, AddressOf WndProc
End Sub

Sub UnHookWindow(hWND As Long)
'Sets procedure for handling events back to the original.
SetWindowLong hWND, GWL_WNDPROC, OldProc
End Sub
