VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMedia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' i think this module originally came from planet-source-code.com/vb,
' though it has been mostly re-written.
' the function JumphWnd was conceived from scratch and was particually problematic.

Option Explicit
Option Compare Binary
Option Base 0

' Declarations
Private Declare Function GetWindowRect Lib "User32" (ByVal hWND As Long, lpRect As RECT) As Long

' Variables
Private m_objBasicAudio As IBasicAudio
Private m_objBasicVideo As IBasicVideo
Private m_objVideoWindow As IVideoWindow
Private m_objMediaControl As IMediaControl
Private m_objMediaPosition As IMediaPosition
Private m_hWnd As Long
Private m_FileName As String
Private m_HasAudio As Boolean
Private m_HasVideo As Boolean
Private m_ShowMouse As Boolean

Public m_RaiseErrors As Boolean

' Constants
Private Const WS_VISIBLE = &H10000000
Private Const WS_CHILD = &H40000000

'objVideoWindow.WindowState As Long
'# SW_SHOW
'# SW_HIDE
'# SW_MAXIMIZE
'# SW_MINIMIZE

' Types
Private Type RECT
    Left As Long
    Top As Long
    right As Long
    bottom As Long
End Type

' Enums
Public Enum enumState
    stStopped
    stPlaying
    stPaused
End Enum

'this function will move the video display from one hWnd to
'another without interupting playback of disturbing the
'play / pause state.
Public Function JumphWnd(lhWnd As Long) As Boolean
On Error GoTo JumphWnd_err

JumphWnd = False

Window = lhWnd

'if no video atm, just remember the hWnd
If Not HasVideo Then
    JumphWnd = True
    Exit Function
End If

m_objVideoWindow.Owner = Window 'fix to stop flash of un-owned video
m_objVideoWindow.WindowStyle = WS_VISIBLE
m_objVideoWindow.Owner = Window
ResizeWindow

'by calling this command with false, we *should*
'  a) prevent the video hiding behind something
'  b) prevent the video window steal focus
'  [fingers crossed]
m_objVideoWindow.SetWindowForeground False

m_objVideoWindow.HideCursor Not m_ShowMouse
m_objVideoWindow.MessageDrain = Window

JumphWnd = True

Exit Function
JumphWnd_err:
Debug.Print "error: unable to jump hWnd."
err.Clear
End Function

'debug and work-around function
Public Sub BumpVidWindow()
If Not HasVideo Then Exit Sub

'JumphWnd Window

m_objVideoWindow.Owner = Window
End Sub

Private Sub Update_HasAudio()
On Error GoTo Update_HasAudio_err
Dim t

HasAudio = False
t = m_objBasicAudio.Volume
HasAudio = True

Update_HasAudio_err:
err.Clear
End Sub

Private Sub Update_HasVideo()
On Error GoTo Update_HasVideo_err
Dim t

HasVideo = False
t = m_objBasicVideo.VideoWidth
HasVideo = True

Update_HasVideo_err:
err.Clear
End Sub

Public Property Get Width()
On Error GoTo Width_err

If HasVideo Then
    Width = m_objBasicVideo.SourceWidth
Else
    Width = 0
End If

Width_err:
err.Clear
End Property

Public Property Get Height()
On Error GoTo Height_err

If HasVideo Then
    Height = m_objBasicVideo.SourceHeight
Else
    Height = 0
End If

Height_err:
err.Clear
End Property

Public Property Let Window(ByVal hWND As Long)
m_hWnd = hWND
End Property

Public Property Get Window() As Long
Window = m_hWnd
End Property

Public Property Get FileName() As String
FileName = m_FileName
End Property

Public Property Let FileName(ByVal Value As String)
On Error GoTo FileName_err

Dim r As RECT, e As String

m_FileName = Value

e = "unloading old objects."
If Not (m_objMediaControl Is Nothing) Then
    m_objMediaControl.Stop
    Set m_objBasicAudio = Nothing
    Set m_objBasicVideo = Nothing
    Set m_objMediaControl = Nothing
    Set m_objVideoWindow = Nothing
    Set m_objMediaPosition = Nothing
End If

If Len(FileName) <= 0 Then Exit Property
If Not fsoMain.FileExists(FileName) Then Exit Property

e = "loading new FilgraphManager."
Set m_objMediaControl = New FilgraphManager
e = "RenderFile."
m_objMediaControl.RenderFile FileName
e = "loading new objects."
Set m_objBasicAudio = m_objMediaControl
Set m_objBasicVideo = m_objMediaControl
Set m_objVideoWindow = m_objMediaControl
Set m_objMediaPosition = m_objMediaControl

e = "checking for audio stream."
Update_HasAudio

e = "checking for video stream."
Update_HasVideo

e = "setup video window."
JumphWnd Window

Exit Property
FileName_err:
If m_RaiseErrors Then Main_Err "clsMedia.FileName - " & e & "  file='" & m_FileName & "'."
err.Clear
End Property

Public Property Get FullScreen() As Boolean
FullScreen = m_objVideoWindow.FullScreenMode
End Property

Public Property Let FullScreen(ByVal bFullScreen As Boolean)
m_objVideoWindow.FullScreenMode = bFullScreen
End Property

Public Sub ResizeWindow()
On Error GoTo ResizeWindow_err
Dim r As RECT, s1 As Single, s2 As Single, w As Long, h As Long

If HasVideo Then
    GetWindowRect Window, r 'note: for windows, this will INCLUDE the window border.
    
    w = (r.right - r.Left)
    h = (r.bottom - r.Top)
    
    s1 = (w / Width)
    s2 = (h / Height)
    
    If s1 < s2 Then
        'use s1
        m_objVideoWindow.Width = Width * s1
        m_objVideoWindow.Height = Height * s1
    Else
        'use s2
        m_objVideoWindow.Width = Width * s2
        m_objVideoWindow.Height = Height * s2
    End If
    
    m_objVideoWindow.Left = (w / 2) - (m_objVideoWindow.Width / 2)
    m_objVideoWindow.Top = (h / 2) - (m_objVideoWindow.Height / 2)
End If

Exit Sub
ResizeWindow_err:
Debug.Print "ResizeWindow_err"
err.Clear
'not really anthing to do here, it just fixes a bug:
' - if a video finishes in full-screen mode, crashes with no error.
End Sub

Public Sub Play()
On Error GoTo Play_err

If Len(m_FileName) = 0 Then
    Exit Sub
End If
If Position = Duration Then
    Position = 0
End If
m_objMediaControl.Run

Play_err:
Debug.Print "clsMedia.play error; " & err.Number & ":" & err.Description
err.Clear
End Sub

Public Sub Pause()
On Error GoTo Pause_err

m_objMediaControl.Pause

Pause_err:
err.Clear
End Sub

Public Sub StopPlaying()
On Error GoTo StopPlaying_err

m_objMediaControl.Stop

StopPlaying_err:
err.Clear
End Sub

Public Property Let Position(ByVal NewPos As Double)
On Error GoTo Position_err

With m_objMediaPosition
    If NewPos < 0 Then
        NewPos = 0
    ElseIf NewPos > .Duration Then
        NewPos = .Duration
    End If
    m_objMediaPosition.CurrentPosition = NewPos
    If HasVideo Then
        m_objMediaControl.Run
        Do Until m_objMediaPosition.CurrentPosition > NewPos
        Loop
        m_objMediaPosition.CurrentPosition = NewPos
        m_objMediaControl.Stop
    End If
End With

Position_err:
err.Clear
End Property

Public Property Get Position() As Double
On Error GoTo Position_err

Dim l As Long
l = m_objMediaPosition.CurrentPosition

Position = m_objMediaPosition.CurrentPosition

Exit Property
Position_err:
err.Clear
Position = Duration
End Property

Public Property Get State() As enumState
On Error GoTo State_err

Dim TimeOut As Long, s As Long

m_objMediaControl.GetState TimeOut, s
If s = 0 Then
    State = stStopped
ElseIf s = 1 Then
    State = stPaused
Else
    State = stPlaying
End If

State_err:
err.Clear
End Property

Public Property Get Speed() As Double
On Error GoTo Speed_err

Speed = m_objMediaPosition.Rate

Speed_err:
err.Clear
End Property

Public Property Let Speed(ByVal Value As Double)
On Error GoTo Speed_err

If Value < 0.1 Then
    Value = 0.1
End If
m_objMediaPosition.Rate = Value

Speed_err:
err.Clear
End Property

Public Property Get Duration() As Double
On Error GoTo Duration_err

Dim l As Long
l = m_objMediaPosition.Duration

Duration = m_objMediaPosition.Duration

Exit Property
Duration_err:
err.Clear
Duration = 1
End Function

Public Property Get Volume() As Long
On Error GoTo Volume_err

Volume = m_objBasicAudio.Volume

Volume_err:
err.Clear
End Property

Public Property Let Volume(ByVal Value As Long)
On Error GoTo Volume_err

If Value > 0 Then Value = 0
If Value < -10000 Then Value = -10000
m_objBasicAudio.Volume = Value

Volume_err:
err.Clear
End Property

Public Property Get Balance() As Long
On Error GoTo Balance_err

Balance = m_objBasicAudio.Balance

Balance_err:
err.Clear
End Property

Public Property Let Balance(ByVal Value As Long)
On Error GoTo Balance_err

If Value < -10000 Then Value = -10000
If Value > 10000 Then Value = 10000
m_objBasicAudio.Balance = Value

Balance_err:
err.Clear
End Property

Public Property Get HasAudio() As Boolean
HasAudio = m_HasAudio
End Property

Private Property Let HasAudio(ByVal Value As Boolean)
m_HasAudio = Value
End Property

Public Property Get HasVideo() As Boolean
HasVideo = m_HasVideo
End Property

Private Property Let HasVideo(ByVal Value As Boolean)
m_HasVideo = Value
End Property

Public Property Get ShowMouse() As Boolean
ShowMouse = m_ShowMouse
End Property
Public Property Let ShowMouse(b As Boolean)
m_ShowMouse = b
If HasVideo And Not (m_objVideoWindow Is Nothing) Then
    m_objVideoWindow.HideCursor Not m_ShowMouse
End If
End Property

Private Sub Class_Initialize()
m_RaiseErrors = True
m_ShowMouse = False
End Sub
