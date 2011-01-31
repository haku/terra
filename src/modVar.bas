Attribute VB_Name = "modVar"

Option Explicit
Option Compare Binary
Option Base 0

'enum's;
Public Enum expDrawMode
    xpInternal = 0
    xpBitmap = 1
End Enum

Public Enum eXpTmBtType
    btyp_Normal = 0
    btyp_DropArrow = 1
    btyp_DropDown = 2
End Enum

'media library stuff:
Public Type typMdbItem
    sFile As String
    dAdded As Date
    lStartCnt As Long
    lEndCnt As Long
    dLastPlay As Date
    bSel As Boolean
    lMD5 As Long
    lDuration As Long
    bEnabled As Boolean
    
    bMissing As Boolean
End Type

Public Type typMdb
    lCnt As Long
    lIndex As Long
    Items() As typMdbItem
    lTrunk As Long
    bTrunkReverse As Boolean
End Type

'playlist stuff
Public Type typPlItem
    sFile As String
    lMD5 As Long
    bSel As Boolean
    lStartCnt As Long
    lEndCnt As Long

    lDuration As Long
    dLastPlay As Date

    lSource(0 To 1) As Long
End Type

Public Type typPl
    Items() As typPlItem
    lCnt As Long
    lIndex As Long
    sName As String
    lCurrent As Long
    sFilePath As String
    bEnab As Boolean
    lScroll As Long
    bArc As Boolean
    
    lTotalDuration As Long
    bTotalDurationComplete As Boolean
End Type

Public mdb_List As typMdb, mdbl_ItmH As Long, m_ListDc As pcMemDC
Public mdb_PL() As typPl, pl_ListDc As pcMemDC
Public mdb_Cue As typPl, cue_ListDc As pcMemDC
Public mdb_History As typPl

Public Const lHistoryLimit          As Long = 12
Public Const lScrollWheelMovement   As Long = 8

'playback stuff:
Public pb_lItemSource As Long 'media libary; playlist; etc.
'   0= : media lib
'   0> : play list
Public pb_llItemIndex As Long 'index in the list
'   0 = stop
'   1 = play
'   2 = pause

'maint. stuff:
Public Enum trrFileAction
    trrDoNothing = 0
    trrDelLibRef = 1
    trrMoveFile = 2
    trrMarkMissing = 3
End Enum

'general files stuff:
Public file_INI As String, folder_Playlists As String, file_PlayListIndex As String
Public Const file_ext_playlist As String = ".trrpl"
Public file_DebugINI As String

'file types stuff:
Public Const file_ext_list_def As String = ".mp3|.ogg|.wma|.wmv|.avi|.mpg|.mpeg|.ac3|.mp4|.wav|.ra|.mpga|.mkv|.ogm"
Public file_ext_list As String

'preferences:
Public Const pref_AutoSaveTime = 60 'seconds
Public pref_CanSave As Boolean
Public pref_AutoSaveCounter As Long

'gui consts:
Public Const lSp As Long = 120
Public Const lGp As Long = 60
Public Const lGpBtwnLstAndSB As Long = 30
Public Const dListCellPaddingH As Double = 1.05 '1.1
Public Const sDateFormatString As String = "yyyy-mm-dd hh:mm"

'playback stuff:
Public Const pb_RetryCount_Max As Long = 10

'visible forms:
Public m_bGotoVis As Boolean

'saved strings:
Public m_GotoText As String
