VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cUxTheme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' To Report API errors:
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Private Declare Function FormatMessage Lib "kernel32" _
    Alias "FormatMessageA" ( _
    ByVal dwFlags As Long, lpSource As Any, _
    ByVal dwMessageId As Long, ByVal dwLanguageId As Long, _
    ByVal lpBuffer As String, ByVal nSize As Long, _
    Arguments As Long) As Long

'' //---------------------------------------------------------------------------------------
'' //   "Window" (i.e., non-client) Parts & States
'' //
'' //---------------------------------------------------------------------------------------
Public Enum UxThemeWindowParts
    WP_CAPTION = 1
    WP_SMALLCAPTION = 2
    WP_MINCAPTION = 3
    WP_SMALLMINCAPTION = 4
    WP_MAXCAPTION = 5
    WP_SMALLMAXCAPTION = 6
    WP_FRAMELEFT = 7
    WP_FRAMERIGHT = 8
    WP_FRAMEBOTTOM = 9
    WP_SMALLFRAMELEFT = 10
    WP_SMALLFRAMERIGHT = 11
    WP_SMALLFRAMEBOTTOM = 12
    '' //---- window frame buttons ----
    WP_SYSBUTTON = 13
    WP_MDISYSBUTTON = 14
    WP_MINBUTTON = 15
    WP_MDIMINBUTTON = 16
    WP_MAXBUTTON = 17
    WP_CLOSEBUTTON = 18
    WP_SMALLCLOSEBUTTON = 19
    WP_MDICLOSEBUTTON = 20
    WP_RESTOREBUTTON = 21
    WP_MDIRESTOREBUTTON = 22
    WP_HELPBUTTON = 23
    WP_MDIHELPBUTTON = 24
    '' //---- scrollbars
    WP_HORZSCROLL = 25
    WP_HORZTHUMB = 26
    WP_VERTSCROLL = 27
    WP_VERTTHUMB = 28
    '' //---- dialog ----
    WP_DIALOG = 29
    '' //---- hit-test templates ---
    WP_CAPTIONSIZINGTEMPLATE = 30
    WP_SMALLCAPTIONSIZINGTEMPLATE = 31
    WP_FRAMELEFTSIZINGTEMPLATE = 32
    WP_SMALLFRAMELEFTSIZINGTEMPLATE = 33
    WP_FRAMERIGHTSIZINGTEMPLATE = 34
    WP_SMALLFRAMERIGHTSIZINGTEMPLATE = 35
    WP_FRAMEBOTTOMSIZINGTEMPLATE = 36
    WP_SMALLFRAMEBOTTOMSIZINGTEMPLATE = 37
End Enum

Public Enum UxThemeFrameStates
    FS_ACTIVE = 1
    FS_INACTIVE = 2
End Enum

Public Enum UxThemeCaptionStates
    CS_ACTIVE = 1
    CS_INACTIVE = 2
    CS_DISABLED = 3
End Enum
    
Public Enum UxThemeMaxCaptionStates
    MXCS_ACTIVE = 1
    MXCS_INACTIVE = 2
    MXCS_DISABLED = 3
End Enum

Public Enum UxThemeMinCaptionStates
    MNCS_ACTIVE = 1
    MNCS_INACTIVE = 2
    MNCS_DISABLED = 3
End Enum

Public Enum UxThemeHorzScrollStates
    HSS_NORMAL = 1
    HSS_HOT = 2
    HSS_PUSHED = 3
    HSS_DISABLED = 4
End Enum

Public Enum UxThemeHorzThumbStates
    HTS_NORMAL = 1
    HTS_HOT = 2
    HTS_PUSHED = 3
    HTS_DISABLED = 4
End Enum

Public Enum UxThemeVertScrollStates
    VSS_NORMAL = 1
    VSS_HOT = 2
    VSS_PUSHED = 3
    VSS_DISABLED = 4
End Enum

Public Enum UxThemeVertThumbStates
    VTS_NORMAL = 1
    VTS_HOT = 2
    VTS_PUSHED = 3
    VTS_DISABLED = 4
End Enum

Public Enum UxThemeSysButtonStates
    SBS_NORMAL = 1
    SBS_HOT = 2
    SBS_PUSHED = 3
    SBS_DISABLED = 4
End Enum

Public Enum UxThemeMinButtonStates
    MINBS_NORMAL = 1
    MINBS_HOT = 2
    MINBS_PUSHED = 3
    MINBS_DISABLED = 4
End Enum

Public Enum UxThemeMaxButtonStates
    MAXBS_NORMAL = 1
    MAXBS_HOT = 2
    MAXBS_PUSHED = 3
    MAXBS_DISABLED = 4
End Enum

Public Enum UxThemeRestoreButtonStates
    RBS_NORMAL = 1
    RBS_HOT = 2
    RBS_PUSHED = 3
    RBS_DISABLED = 4
End Enum

Public Enum UxThemeHelpButtonStates
    HBS_NORMAL = 1
    HBS_HOT = 2
    HBS_PUSHED = 3
    HBS_DISABLED = 4
End Enum

Public Enum UxThemeCloseButtonStates
    CBS_NORMAL = 1
    CBS_HOT = 2
    CBS_PUSHED = 3
    CBS_DISABLED = 4
End Enum


'' //---------------------------------------------------------------------------------------
'' //   "Button" Parts & States
'' //--------------------------------------------------------------------------------------
Public Enum UxThemeButtonParts
    BP_PUSHBUTTON = 1
    BP_RADIOBUTTON = 2
    BP_CHECKBOX = 3
    BP_GROUPBOX = 4
    BP_USERBUTTON = 5
End Enum

Public Enum UxThemePushButtonStates
    PBS_NORMAL = 1
    PBS_HOT = 2
    PBS_PRESSED = 3
    PBS_DISABLED = 4
    PBS_DEFAULTED = 5
End Enum

Public Enum UxThemeRadioButtonStates
    RBS_UNCHECKEDNORMAL = 1
    RBS_UNCHECKEDHOT = 2
    RBS_UNCHECKEDPRESSED = 3
    RBS_UNCHECKEDDISABLED = 4
    RBS_CHECKEDNORMAL = 5
    RBS_CHECKEDHOT = 6
    RBS_CHECKEDPRESSED = 7
    RBS_CHECKEDDISABLED = 8
End Enum

Public Enum UxThemeCheckBoxStates
    CBS_UNCHECKEDNORMAL = 1
    CBS_UNCHECKEDHOT = 2
    CBS_UNCHECKEDPRESSED = 3
    CBS_UNCHECKEDDISABLED = 4
    CBS_CHECKEDNORMAL = 5
    CBS_CHECKEDHOT = 6
    CBS_CHECKEDPRESSED = 7
    CBS_CHECKEDDISABLED = 8
    CBS_MIXEDNORMAL = 9
    CBS_MIXEDHOT = 10
    CBS_MIXEDPRESSED = 11
    CBS_MIXEDDISABLED = 12
End Enum

Public Enum UxThemeGroupBoxStates
    GBS_NORMAL = 1
    GBS_DISABLED = 2
End Enum


'' //---------------------------------------------------------------------------------------
'' //   "Rebar" Parts & States
'' //---------------------------------------------------------------------------------------
Public Enum UxThemeRebarParts
    RP_GRIPPER = 1
    RP_GRIPPERVERT = 2
    RP_BAND = 3
    RP_CHEVRON = 4
    RP_CHEVRONVERT = 5
End Enum

Public Enum UxThemeChevronStates
    CHEVS_NORMAL = 1
    CHEVS_HOT = 2
    CHEVS_PRESSED = 3
End Enum


'' //---------------------------------------------------------------------------------------
'' //   "Toolbar" Parts & States
'' //---------------------------------------------------------------------------------------
Public Enum UxThemeToolBarParts
    TP_BUTTON = 1
    TP_DROPDOWNBUTTON = 2
    TP_SPLITBUTTON = 3
    TP_SPLITBUTTONDROPDOWN = 4
    TP_SEPARATOR = 5
    TP_SEPARATORVERT = 6
End Enum

Public Enum UxThemeToolBarStates
    TS_NORMAL = 1
    TS_HOT = 2
    TS_PRESSED = 3
    TS_DISABLED = 4
    TS_CHECKED = 5
    TS_HOTCHECKED = 6
End Enum

'' //---------------------------------------------------------------------------------------
'' //   "Status" Parts & States
'' //---------------------------------------------------------------------------------------
Public Enum UxThemeStatusParts
    SP_PANE = 1
    SP_GRIPPERPANE = 2
    SP_GRIPPER = 3
End Enum

'' //---------------------------------------------------------------------------------------
'' //   "Menu" Parts & States
'' //---------------------------------------------------------------------------------------
Public Enum UxThemeMenuParts
    MP_MENUITEM = 1
    MP_MENUDROPDOWN = 2
    MP_MENUBARITEM = 3
    MP_MENUBARDROPDOWN = 4
    MP_CHEVRON = 5
    MP_SEPARATOR = 6
End Enum

Public Enum UxThemeMenuStates
    MS_NORMAL = 1
    MS_SELECTED = 2
    MS_DEMOTED = 3
End Enum

' //---------------------------------------------------------------------------------------
' //   "ListView" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemeLISTVIEWParts
   LVP_LISTITEM = 1
   LVP_LISTGROUP = 2
   LVP_LISTDETAIL = 3
   LVP_LISTSORTEDDETAIL = 4
   LVP_EMPTYTEXT = 5
End Enum

Public Enum UxThemeLISTITEMStates
   LIS_NORMAL = 1
   LIS_HOT = 2
   LIS_SELECTED = 3
   LIS_DISABLED = 4
   LIS_SELECTEDNOTFOCUS = 5
End Enum

' //---------------------------------------------------------------------------------------
' //   "Header" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemeHEADERParts
   HP_HEADERITEM = 1
   HP_HEADERITEMLEFT = 2
   HP_HEADERITEMRIGHT = 3
   HP_HEADERSORTARROW = 4
End Enum

Public Enum UxThemeHEADERITEMStates
   HIS_NORMAL = 1
   HIS_HOT = 2
   HIS_PRESSED = 3
End Enum

Public Enum UxThemeHEADERITEMLEFTStates
   HILS_NORMAL = 1
   HILS_HOT = 2
   HILS_PRESSED = 3
End Enum

Public Enum UxThemeHEADERITEMRIGHTStates
   HIRS_NORMAL = 1
   HIRS_HOT = 2
   HIRS_PRESSED = 3
End Enum

Public Enum UxThemeHEADERSORTARROWStates
   HSAS_SORTEDUP = 1
   HSAS_SORTEDDOWN = 2
End Enum

' //---------------------------------------------------------------------------------------
' //   "Progress" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemePROGRESSParts
   PP_BAR = 1
   PP_BARVERT = 2
   PP_CHUNK = 3
   PP_CHUNKVERT = 4
End Enum

' //---------------------------------------------------------------------------------------
' //   "Tab" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UsxThemeTABParts
   TABP_TABITEM = 1
   TABP_TABITEMLEFTEDGE = 2
   TABP_TABITEMRIGHTEDGE = 3
   TABP_TABITEMBOTHEDGE = 4
   TABP_TOPTABITEM = 5
   TABP_TOPTABITEMLEFTEDGE = 6
   TABP_TOPTABITEMRIGHTEDGE = 7
   TABP_TOPTABITEMBOTHEDGE = 8
   TABP_PANE = 9
   TABP_BODY = 10
End Enum

Public Enum UxThemeTABITEMStates
   TIS_NORMAL = 1
   TIS_HOT = 2
   TIS_SELECTED = 3
   TIS_DISABLED = 4
   TIS_FOCUSED = 5
End Enum

Public Enum UxThemeTABITEMLEFTEDGEStates
   TILES_NORMAL = 1
   TILES_HOT = 2
   TILES_SELECTED = 3
   TILES_DISABLED = 4
   TILES_FOCUSED = 5
End Enum

Public Enum UxThemeTABITEMRIGHTEDGEStates
   TIRES_NORMAL = 1
   TIRES_HOT = 2
   TIRES_SELECTED = 3
   TIRES_DISABLED = 4
   TIRES_FOCUSED = 5
End Enum

Public Enum UxThemeTABITEMBOTHEDGESStates
   TIBES_NORMAL = 1
   TIBES_HOT = 2
   TIBES_SELECTED = 3
   TIBES_DISABLED = 4
   TIBES_FOCUSED = 5
End Enum

Public Enum UxThemeTOPTABITEMStates
   TTIS_NORMAL = 1
   TTIS_HOT = 2
   TTIS_SELECTED = 3
   TTIS_DISABLED = 4
   TTIS_FOCUSED = 5
End Enum

Public Enum UxThemeTOPTABITEMLEFTEDGEStates
   TTILES_NORMAL = 1
   TTILES_HOT = 2
   TTILES_SELECTED = 3
   TTILES_DISABLED = 4
   TTILES_FOCUSED = 5
End Enum

Public Enum UxThemeTOPTABITEMRIGHTEDGEStates
   TTIRES_NORMAL = 1
   TTIRES_HOT = 2
   TTIRES_SELECTED = 3
   TTIRES_DISABLED = 4
   TTIRES_FOCUSED = 5
End Enum

Public Enum UxThemeTOPTABITEMBOTHEDGESStates
   TTIBES_NORMAL = 1
   TTIBES_HOT = 2
   TTIBES_SELECTED = 3
   TTIBES_DISABLED = 4
   TTIBES_FOCUSED = 5
End Enum

' //---------------------------------------------------------------------------------------
' //   "Trackbar" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemeTRACKBARParts
   TKP_TRACK = 1
   TKP_TRACKVERT = 2
   TKP_THUMB = 3
   TKP_THUMBBOTTOM = 4
   TKP_THUMBTOP = 5
   TKP_THUMBVERT = 6
   TKP_THUMBLEFT = 7
   TKP_THUMBRIGHT = 8
   TKP_TICS = 9
   TKP_TICSVERT = 10
End Enum

Public Enum UxThemeTRACKBARStates
   TKS_NORMAL = 1
End Enum

Public Enum UxThemeTRACKStates
   TRS_NORMAL = 1
End Enum

Public Enum UxThemeTRACKVERTStates
   TRVS_NORMAL = 1
End Enum

Public Enum UxThemeTHUMBStates
   TUS_NORMAL = 1
   TUS_HOT = 2
   TUS_PRESSED = 3
   TUS_FOCUSED = 4
   TUS_DISABLED = 5
End Enum

Public Enum UxThemeTHUMBBOTTOMStates
   TUBS_NORMAL = 1
   TUBS_HOT = 2
   TUBS_PRESSED = 3
   TUBS_FOCUSED = 4
   TUBS_DISABLED = 5
End Enum

Public Enum UxThemeTHUMBTOPStates
   TUTS_NORMAL = 1
   TUTS_HOT = 2
   TUTS_PRESSED = 3
   TUTS_FOCUSED = 4
   TUTS_DISABLED = 5
End Enum

Public Enum UxThemeTHUMBVERTStates
   TUVS_NORMAL = 1
   TUVS_HOT = 2
   TUVS_PRESSED = 3
   TUVS_FOCUSED = 4
   TUVS_DISABLED = 5
End Enum

Public Enum UxThemeTHUMBLEFTStates
   TUVLS_NORMAL = 1
   TUVLS_HOT = 2
   TUVLS_PRESSED = 3
   TUVLS_FOCUSED = 4
   TUVLS_DISABLED = 5
End Enum

Public Enum UxThemeTHUMBRIGHTStates
   TUVRS_NORMAL = 1
   TUVRS_HOT = 2
   TUVRS_PRESSED = 3
   TUVRS_FOCUSED = 4
   TUVRS_DISABLED = 5
End Enum

Public Enum UxThemeTICSStates
   TSS_NORMAL = 1
End Enum

Public Enum UxThemeTICSVERTStates
   TSVS_NORMAL = 1
End Enum

' //---------------------------------------------------------------------------------------
' //   "Tooltips" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemeTOOLTIPParts
   TTP_STANDARD = 1
   TTP_STANDARDTITLE = 2
   TTP_BALLOON = 3
   TTP_BALLOONTITLE = 4
   TTP_CLOSE = 5
End Enum

Public Enum UxThemeCLOSEStates
   TTCS_NORMAL = 1
   TTCS_HOT = 2
   TTCS_PRESSED = 3
End Enum

Public Enum UxThemeSTANDARDStates
   TTSS_NORMAL = 1
   TTSS_LINK = 2
End Enum

Public Enum UxThemeBALLOONStates
   TTBS_NORMAL = 1
   TTBS_LINK = 2
End Enum

' //---------------------------------------------------------------------------------------
' //   "TreeView" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemeTREEVIEWParts
   TVP_TREEITEM = 1
   TVP_GLYPH = 2
   TVP_BRANCH = 3
End Enum

Public Enum UxThemeTREEITEMStates
   TREIS_NORMAL = 1
   TREIS_HOT = 2
   TREIS_SELECTED = 3
   TREIS_DISABLED = 4
   TREIS_SELECTEDNOTFOCUS = 5
End Enum

Public Enum UxThemeGLYPHStates
   GLPS_CLOSED = 1
   GLPS_OPENED = 2
End Enum

' //---------------------------------------------------------------------------------------
' //   "Spin" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemeSPINStates
   SPNP_UP = 1
   SPNP_DOWN = 2
   SPNP_UPHORZ = 3
   SPNP_DOWNHORZ = 4
End Enum

Public Enum UxThemeUPStates
   UPS_NORMAL = 1
   UPS_HOT = 2
   UPS_PRESSED = 3
   UPS_DISABLED = 4
End Enum

Public Enum UxThemeDOWNStates
   DNS_NORMAL = 1
   DNS_HOT = 2
   DNS_PRESSED = 3
   DNS_DISABLED = 4
End Enum

Public Enum UxThemeUPHORZStates
   UPHZS_NORMAL = 1
   UPHZS_HOT = 2
   UPHZS_PRESSED = 3
   UPHZS_DISABLED = 4
End Enum

Public Enum UxThemeDOWNHORZStates
   DNHZS_NORMAL = 1
   DNHZS_HOT = 2
   DNHZS_PRESSED = 3
   DNHZS_DISABLED = 4
End Enum

' //---------------------------------------------------------------------------------------
' //   "Page" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemePageParts
   PGRP_UP = 1
   PGRP_DOWN = 2
   PGRP_UPHORZ = 3
   PGRP_DOWNHORZ = 4
End Enum

' //--- Pager uses same states as Spin ---

' //---------------------------------------------------------------------------------------
' //   "Scrollbar" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemeSCROLLBARParts
   SBP_ARROWBTN = 1
   SBP_THUMBBTNHORZ = 2
   SBP_THUMBBTNVERT = 3
   SBP_LOWERTRACKHORZ = 4
   SBP_UPPERTRACKHORZ = 5
   SBP_LOWERTRACKVERT = 6
   SBP_UPPERTRACKVERT = 7
   SBP_GRIPPERHORZ = 8
   SBP_GRIPPERVERT = 9
   SBP_SIZEBOX = 10
End Enum



Public Enum UxThemeARROWBTNStates
   ABS_UPNORMAL = 1
   ABS_UPHOT = 2
   ABS_UPPRESSED = 3
   ABS_UPDISABLED = 4
   ABS_DOWNNORMAL = 5
   ABS_DOWNHOT = 6
   ABS_DOWNPRESSED = 7
   ABS_DOWNDISABLED = 8
   ABS_LEFTNORMAL = 9
   ABS_LEFTHOT = 10
   ABS_LEFTPRESSED = 11
   ABS_LEFTDISABLED = 12
   ABS_RIGHTNORMAL = 13
   ABS_RIGHTHOT = 14
   ABS_RIGHTPRESSED = 15
   ABS_RIGHTDISABLED = 16
End Enum

Public Enum UxThemeSCROLLBARStates
   SCRBS_NORMAL = 1
   SCRBS_HOT = 2
   SCRBS_PRESSED = 3
   SCRBS_DISABLED = 4
End Enum

Public Enum UxThemeSIZEBOXStates
   SZB_RIGHTALIGN = 1
   SZB_LEFTALIGN = 2
End Enum

' //---------------------------------------------------------------------------------------
' //   "Edit" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemeEDITParts
   EP_EDITTEXT = 1
   EP_CARET = 2
End Enum

Public Enum UxThemeEDITTEXTStates
   ETS_NORMAL = 1
   ETS_HOT = 2
   ETS_SELECTED = 3
   ETS_DISABLED = 4
   ETS_FOCUSED = 5
   ETS_READONLY = 6
   ETS_ASSIST = 7
End Enum

' //---------------------------------------------------------------------------------------
' //   "ComboBox" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemeComboBoxParts
   CP_DROPDOWNBUTTON = 1
End Enum

Public Enum UxThemeComboBoxStates
   CBXS_NORMAL = 1
   CBXS_HOT = 2
   CBXS_PRESSED = 3
   CBXS_DISABLED = 4
End Enum

' //---------------------------------------------------------------------------------------
' //   "Taskbar Clock" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemeCLOCKParts
   CLP_TIME = 1
End Enum

Public Enum UxThemeCLOCKStates
   CLS_NORMAL = 1
End Enum

' //---------------------------------------------------------------------------------------
' //   "Tray Notify" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemeTRAYNOTIFYParts
   TNP_BACKGROUND = 1
   TNP_ANIMBACKGROUND = 2
End Enum

' //---------------------------------------------------------------------------------------
' //   "TaskBar" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemeTASKBARParts
   TBP_BACKGROUNDBOTTOM = 1
   TBP_BACKGROUNDRIGHT = 2
   TBP_BACKGROUNDTOP = 3
   TBP_BACKGROUNDLEFT = 4
   TBP_SIZINGBARBOTTOM = 5
   TBP_SIZINGBARRIGHT = 6
   TBP_SIZINGBARTOP = 7
   TBP_SIZINGBARLEFT = 8
End Enum

' //---------------------------------------------------------------------------------------
' //   "TaskBand" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemeTASKBANDParts
   TDP_GROUPCOUNT = 1
   TDP_FLASHBUTTON = 2
   TDP_FLASHBUTTONGROUPMENU = 3
End Enum

' //---------------------------------------------------------------------------------------
' //   "StartPanel" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemeSTARTPANELParts
   SPP_USERPANE = 1
   SPP_MOREPROGRAMS = 2
   SPP_MOREPROGRAMSARROW = 3
   SPP_PROGLIST = 4
   SPP_PROGLISTSEPARATOR = 5
   SPP_PLACESLIST = 6
   SPP_PLACESLISTSEPARATOR = 7
   SPP_LOGOFF = 8
   SPP_LOGOFFBUTTONS = 9
   SPP_USERPICTURE = 10
   SPP_PREVIEW = 11
End Enum

Public Enum UxThemeMOREPROGRAMSARROWStates
   SPS_NORMAL = 1
   SPS_HOT = 2
   SPS_PRESSED = 3
End Enum

Public Enum UxThemeLOGOFFBUTTONSStates
   SPLS_NORMAL = 1
   SPLS_HOT = 2
   SPLS_PRESSED = 3
End Enum

' //---------------------------------------------------------------------------------------
' //   "ExplorerBar" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemeEXPLORERBARParts
   EBP_HEADERBACKGROUND = 1
   EBP_HEADERCLOSE = 2
   EBP_HEADERPIN = 3
   EBP_IEBARMENU = 4
   EBP_NORMALGROUPBACKGROUND = 5
   EBP_NORMALGROUPCOLLAPSE = 6
   EBP_NORMALGROUPEXPAND = 7
   EBP_NORMALGROUPHEAD = 8
   EBP_SPECIALGROUPBACKGROUND = 9
   EBP_SPECIALGROUPCOLLAPSE = 10
   EBP_SPECIALGROUPEXPAND = 11
   EBP_SPECIALGROUPHEAD = 12
End Enum

Public Enum UxThemeHEADERCLOSEStates
   EBHC_NORMAL = 1
   EBHC_HOT = 2
   EBHC_PRESSED = 3
End Enum

Public Enum UxThemeHEADERPINStates
   EBHP_NORMAL = 1
   EBHP_HOT = 2
   EBHP_PRESSED = 3
   EBHP_SELECTEDNORMAL = 4
   EBHP_SELECTEDHOT = 5
   EBHP_SELECTEDPRESSED = 6
End Enum

Public Enum UxThemeIEBARMENUStates
   EBM_NORMAL = 1
   EBM_HOT = 2
   EBM_PRESSED = 3
End Enum

Public Enum UxThemeNORMALGROUPCOLLAPSEStates
   EBNGC_NORMAL = 1
   EBNGC_HOT = 2
   EBNGC_PRESSED = 3
End Enum

Public Enum UxThemeNORMALGROUPEXPANDStates
   EBNGE_NORMAL = 1
   EBNGE_HOT = 2
   EBNGE_PRESSED = 3
End Enum

Public Enum UxThemeSPECIALGROUPCOLLAPSEStates
   EBSGC_NORMAL = 1
   EBSGC_HOT = 2
   EBSGC_PRESSED = 3
End Enum

Public Enum UxThemeSPECIALGROUPEXPANDStates
   EBSGE_NORMAL = 1
   EBSGE_HOT = 2
   EBSGE_PRESSED = 3
End Enum

' //---------------------------------------------------------------------------------------
' //   "TaskBand" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemeMENUBANDParts
   MDP_NEWAPPBUTTON = 1
   MDP_SEPERATOR = 2
End Enum

Public Enum UxThemeMENUBANDStates
   MDS_NORMAL = 1
   MDS_HOT = 2
   MDS_PRESSED = 3
   MDS_DISABLED = 4
   MDS_CHECKED = 5
   MDS_HOTCHECKED = 6
End Enum

Private Type SIZE
   cx As Long
   cy As Long
End Type

Private Type RECT
   left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Declare Function GetVersion Lib "kernel32" () As Long
Private Declare Function OpenThemeData Lib "uxtheme.dll" _
   (ByVal hwnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" _
   (ByVal hTheme As Long) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme.dll" _
   (ByVal hTheme As Long, ByVal lhDC As Long, _
    ByVal iPartId As Long, ByVal iStateId As Long, _
    pRect As RECT, pClipRect As RECT) As Long
Private Declare Function DrawThemeParentBackground Lib "uxtheme.dll" _
   (ByVal hwnd As Long, ByVal hdc As Long, prc As RECT) As Long
Private Declare Function GetThemeBackgroundContentRect Lib "uxtheme.dll" _
   (ByVal hTheme As Long, ByVal hdc As Long, _
    ByVal iPartId As Long, ByVal iStateId As Long, _
    pBoundingRect As RECT, pContentRect As RECT) As Long
Private Declare Function DrawThemeText Lib "uxtheme.dll" _
   (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, _
    ByVal iStateId As Long, ByVal pszText As Long, _
    ByVal iCharCount As Long, ByVal dwTextFlag As Long, _
    ByVal dwTextFlags2 As Long, pRect As RECT) As Long
Private Declare Function DrawThemeIcon Lib "uxtheme.dll" _
   (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, _
    ByVal iStateId As Long, pRect As RECT, _
    ByVal hIml As Long, ByVal iImageIndex As Long) As Long
Public Enum THEMESIZE
    TS_MIN             '// minimum size
    TS_TRUE            '// size without stretching
    TS_DRAW            '// size that theme mgr will use to draw part
End Enum
Private Declare Function GetThemePartSize Lib "uxtheme.dll" _
   (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, _
   ByVal iStateId As Long, prc As RECT, ByVal eSize As THEMESIZE, _
   psz As SIZE) As Long
Private Declare Function GetThemeTextExtent Lib "uxtheme.dll" _
   (ByVal hTheme As Long, ByVal hdc As Long, _
    ByVal iPartId As Long, ByVal iStateId As Long, _
    ByVal pszText As Long, ByVal iCharCount As Long, _
    ByVal dwTextFlags As DrawTextFlags, pBoundingRect As RECT, _
    pExtentRect As RECT) As Long
Private Declare Function DrawThemeEdge Lib "uxtheme.dll" _
   (ByVal hTheme As Long, ByVal hdc As Long, _
   ByVal iPartId As Long, ByVal iStateId As Long, _
   pDestRect As RECT, _
   ByVal uEdge As DrawEdgeEdgeTypes, ByVal uFlags As DrawEdgeBorderFlags, _
   pContentRect As RECT) As Long
Private Const S_OK = 0
Private Const HWND_DESKTOP = 0

'Public Enum DrawTextFlags
'    DT_TOP = &H0
'    DT_LEFT = &H0
'    DT_CENTER = &H1
'    DT_RIGHT = &H2
'    DT_VCENTER = &H4
'    DT_BOTTOM = &H8
'    DT_WORDBREAK = &H10
'    DT_SINGLELINE = &H20
'    DT_EXPANDTABS = &H40
'    DT_TABSTOP = &H80
'    DT_NOCLIP = &H100
'    DT_EXTERNALLEADING = &H200
'    DT_CALCRECT = &H400
'    DT_NOPREFIX = &H800
'    DT_INTERNAL = &H1000
'    DT_EDITCONTROL = &H2000
'    DT_PATH_ELLIPSIS = &H4000
'    DT_END_ELLIPSIS = &H8000
'    DT_MODIFYSTRING = &H10000
'    DT_RTLREADING = &H20000
'    DT_WORD_ELLIPSIS = &H40000
'    DT_NOFULLWIDTHCHARBREAK = &H80000
'    DT_HIDEPREFIX = &H100000
'    DT_PREFIXONLY = &H200000
'End Enum

Public Enum DrawEdgeEdgeTypes
    BDR_RAISEDOUTER = &H1
    BDR_SUNKENOUTER = &H2
    BDR_RAISEDINNER = &H4
    BDR_SUNKENINNER = &H8

    BDR_OUTER = (BDR_RAISEDOUTER Or BDR_SUNKENOUTER)
    BDR_INNER = (BDR_RAISEDINNER Or BDR_SUNKENINNER)
    BDR_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
    BDR_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)


    EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
    EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
    EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
    EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
End Enum

Public Enum DrawEdgeBorderFlags
    BF_LEFT = &H1
    BF_TOP = &H2
    BF_RIGHT = &H4
    BF_BOTTOM = &H8

    BF_TOPLEFT = (BF_TOP Or BF_LEFT)
    BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
    BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
    BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
    BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

    BF_DIAGONAL = &H10
    
    BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)
    BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
    BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
    BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)

    BF_MIDDLE = &H800         '/* Fill in the middle */
    BF_SOFT = &H1000          '/* For softer buttons */
    BF_ADJUST = &H2000        '/* Calculate the space left over */
    BF_FLAT = &H4000          '/* For flat rather than 3D borders */
    BF_MONO = &H8000          '/* For monochrome borders */
End Enum

Public Enum DrawTextAdditionalFlags
   DTT_GRAYED = &H1           '// draw a grayed-out string
End Enum

Private Declare Function ImageList_GetImageRect Lib "comctl32.dll" ( _
        ByVal hIml As Long, _
        ByVal i As Long, _
        prcImage As RECT _
    ) As Long


Private m_sClass As String
Private m_lPartId As Long
Private m_lStateId As Long
Private m_hDC As Long
Private m_hWnd As Long
Private m_lLeft As Long
Private m_lTop As Long
Private m_lWidth As Long
Private m_lHeight As Long
Private m_sText As String
Private m_eTextAlign As DrawTextFlags
Private m_lIconIndex As Long
Private m_hIml As Long
Private m_bRaiseError As Boolean
Private m_bUseThemeSize As Boolean

Public Property Get RaiseErrors() As Boolean
   RaiseErrors = m_bRaiseError
End Property
Public Property Let RaiseErrors(ByVal Value As Boolean)
   m_bRaiseError = Value
End Property
Public Property Get UseThemeSize() As Boolean
   UseThemeSize = m_bUseThemeSize
End Property
Public Property Let UseThemeSize(ByVal Value As Boolean)
   m_bUseThemeSize = Value
End Property
Public Property Get Class() As String
   Class = m_sClass
End Property
Public Property Let Class(ByVal Value As String)
   m_sClass = Value
End Property
Public Property Get Part() As Long
   Part = m_lPartId
End Property
Public Property Let Part(ByVal Value As Long)
   m_lPartId = Value
End Property
Public Property Get State() As Long
   State = m_lStateId
End Property
Public Property Let State(ByVal Value As Long)
   m_lStateId = Value
End Property
Public Property Get hIml() As Long
   hIml = m_hIml
End Property
Public Property Let hIml(ByVal Value As Long)
   m_hIml = Value
End Property
Public Property Get hdc() As Long
   hdc = m_hDC
End Property
Public Property Let hdc(ByVal Value As Long)
   m_hDC = Value
End Property
Public Property Get hwnd() As Long
   hwnd = m_hWnd
End Property
Public Property Let hwnd(ByVal Value As Long)
   m_hWnd = Value
End Property

Public Property Get Text() As String
   Text = m_sText
End Property
Public Property Let Text(ByVal Value As String)
   m_sText = Value
End Property
Public Property Get TextAlign() As DrawTextFlags
   TextAlign = m_eTextAlign
End Property
Public Property Let TextAlign(ByVal Value As DrawTextFlags)
   m_eTextAlign = Value
End Property
Public Property Get IconIndex() As Long
   IconIndex = m_lIconIndex
End Property
Public Property Let IconIndex(ByVal Value As Long)
   m_lIconIndex = Value
End Property
Public Property Get left() As Long
   left = m_lLeft
End Property
Public Property Let left(ByVal Value As Long)
   m_lLeft = Value
End Property
Public Property Get Top() As Long
   Top = m_lTop
End Property
Public Property Let Top(ByVal Value As Long)
   m_lTop = Value
End Property
Public Property Get Width() As Long
   Width = m_lWidth
End Property
Public Property Let Width(ByVal Value As Long)
   m_lWidth = Value
End Property
Public Property Get Height() As Long
   Height = m_lHeight
End Property
Public Property Let Height(ByVal Value As Long)
   m_lHeight = Value
End Property
Public Property Get PartWidth(Optional ByVal eWidthOptions As THEMESIZE = TS_TRUE) As Long
   Dim tSize As SIZE
   Dim tR As RECT
   Dim hTheme As Long
   Dim lR As Long
   hTheme = OpenThemeData(m_hWnd, StrPtr(m_sClass))
   If (hTheme) Then
      lR = GetThemePartSize(hTheme, m_hDC, m_lPartId, m_lStateId, tR, eWidthOptions, tSize)
      If (lR = S_OK) Then
         PartWidth = tSize.cx
      Else
         pFailed "Failed to read part size for class '" & m_sClass & "', partId=" & m_lPartId & ", stateId=" & m_lStateId, lR
      End If
      CloseThemeData hTheme
   Else
      pFailed "No theme data for class '" & m_sClass & "'", Err.LastDllError
   End If
End Property
Public Property Get PartHeight(Optional ByVal eWidthOptions As THEMESIZE = TS_TRUE) As Long
   Dim tSize As SIZE
   Dim tR As RECT
   Dim hTheme As Long
   Dim lR As Long
   hTheme = OpenThemeData(m_hWnd, StrPtr(m_sClass))
   If (hTheme) Then
      lR = GetThemePartSize(hTheme, m_hDC, m_lPartId, m_lStateId, tR, eWidthOptions, tSize)
      If (lR = S_OK) Then
         PartHeight = tSize.cy
      Else
         pFailed "Failed to read part size for class '" & m_sClass & "', partId=" & m_lPartId & ", stateId=" & m_lStateId, lR
      End If
      CloseThemeData hTheme
   Else
      pFailed "No theme data for class '" & m_sClass & "'", Err.LastDllError
   End If
End Property

Public Function Draw() As Boolean
Dim hTheme As Long
Dim bSuccess As Boolean
Dim lR As Long
Dim tTextR As RECT
Dim tContentR As RECT
Dim tImlR As RECT
   
   bSuccess = True
   hTheme = OpenThemeData(m_hWnd, StrPtr(m_sClass))
   If (hTheme) Then
      
      Dim tR As RECT
      Dim lWidthTaken As Long
      tR.left = m_lLeft
      tR.Top = m_lTop
      If (m_lIconIndex > -1) And (m_hIml) Then
         ImageList_GetImageRect m_hIml, m_lIconIndex, tImlR
         lWidthTaken = tImlR.Right - tImlR.left + 4
      End If
      
      If (m_bUseThemeSize) Then
         Dim tSize As SIZE
         lR = GetThemePartSize(hTheme, m_hDC, m_lPartId, m_lStateId, tR, TS_TRUE, tSize)
         tR.Right = tR.left + tSize.cx
         tR.Bottom = tR.Top + tSize.cy
         lR = GetThemeBackgroundContentRect(hTheme, m_hDC, m_lPartId, m_lStateId, tR, tContentR)
         If (m_lIconIndex > -1) And (m_hIml) Then
            If ((tContentR.Bottom - tContentR.Top) < (tImlR.Bottom - tImlR.Top + 4)) Then
               tR.Bottom = tR.Bottom + ((tImlR.Bottom - tImlR.Top + 4) - (tContentR.Bottom - tContentR.Top))
            End If
            If ((tContentR.Right - tContentR.left) < (tImlR.Right - tImlR.left + 4)) Then
               tR.Right = tR.Right + ((tImlR.Right - tImlR.left + 4) - (tContentR.Right - tContentR.left))
            End If
         End If
         If Len(m_sText) > 0 Then
            lR = GetThemeBackgroundContentRect(hTheme, m_hDC, m_lPartId, m_lStateId, tR, tContentR)
            lR = GetThemeTextExtent(hTheme, m_hDC, m_lPartId, m_lStateId, StrPtr(m_sText), -1, m_eTextAlign, tR, tTextR)
            If ((tContentR.Bottom - tContentR.Top) < (tTextR.Bottom - tTextR.Top)) Then
               tR.Bottom = tR.Bottom + ((tTextR.Bottom - tTextR.Top) - (tContentR.Bottom - tContentR.Top))
            End If
            If ((tContentR.Right - tContentR.left - lWidthTaken) < (tTextR.Right - tTextR.left + 8)) Then
               tR.Right = tR.Right + ((tTextR.Right - tTextR.left + 8) - (tContentR.Right - tContentR.left - lWidthTaken))
            End If
         End If
      Else
         tR.Right = m_lLeft + m_lWidth
         tR.Bottom = m_lTop + m_lHeight
      End If
      
'      lR = DrawThemeParentBackground( _
'         m_hWnd, _
'         m_hDC, _
'         tR)
'      If (lR <> S_OK) Then
'         bSuccess = False
'         pFailed "Failed to parent draw background for class '" & m_sClass & "', partId=" & m_lPartId & ", stateId=" & m_lStateId, lR
'      End If
      lR = DrawThemeBackground( _
         hTheme, _
         m_hDC, _
         m_lPartId, _
         m_lStateId, _
         tR, tR)
      If (lR <> S_OK) Then
         bSuccess = False
         pFailed "Failed to draw background for class '" & m_sClass & "', partId=" & m_lPartId & ", stateId=" & m_lStateId, lR
      End If
      If Len(m_sText) > 0 Then
         lR = GetThemeBackgroundContentRect( _
            hTheme, _
            m_hDC, _
            m_lPartId, _
            m_lStateId, _
            tR, _
            tTextR)
         If (lR <> S_OK) Then
            bSuccess = False
            pFailed "Failed to retrieve background content rectangle for class '" & m_sClass & "', partId=" & m_lPartId & ", stateId=" & m_lStateId, lR
         End If
         tTextR.left = tTextR.left + lWidthTaken
         lR = DrawThemeText( _
            hTheme, _
            m_hDC, _
            m_lPartId, _
            m_lStateId, _
            StrPtr(m_sText), _
            -1, _
            m_eTextAlign, _
            0, _
            tTextR)
         If (lR <> S_OK) Then
            bSuccess = False
            pFailed "Failed to draw theme text for class '" & m_sClass & "', partId=" & m_lPartId & ", stateId=" & m_lStateId, lR
         End If
      End If
      If (m_lIconIndex > -1) Then
         Dim tIconR As RECT
         lR = GetThemeBackgroundContentRect( _
            hTheme, _
            m_hDC, _
            m_lPartId, _
            m_lStateId, _
            tR, _
            tIconR)
         ImageList_GetImageRect m_hIml, m_lIconIndex, tImlR
         tIconR.left = tIconR.left + 2
         tIconR.Top = tIconR.Top + 2
         tIconR.Right = tIconR.left + tImlR.Right - tImlR.left
         tIconR.Bottom = tIconR.Top + tImlR.Bottom - tImlR.Top
         lR = DrawThemeIcon( _
            hTheme, _
            m_hDC, _
            m_lPartId, _
            m_lStateId, _
            tIconR, _
            m_hIml, _
            m_lIconIndex)
         If (lR <> S_OK) Then
            bSuccess = False
            pFailed "Failed to draw theme icon for class '" & m_sClass & "', partId=" & m_lPartId & ", stateId=" & m_lStateId, lR
         End If
      End If
      CloseThemeData hTheme
   Else
      pFailed "No theme data for class '" & m_sClass & "'", Err.LastDllError
      bSuccess = False
   End If
   Draw = bSuccess
End Function

Private Sub pFailed(ByVal sMsg As String, ByVal lR As Long)
Dim sError As String
   sError = WinAPIError(lR)
   If (m_bRaiseError) Then
      MsgBox vbObjectError + 1048 + 10 & "," & App.EXEName & ".cUxTheme" & "," & sMsg & vbCrLf & sError
      Err.Raise vbObjectError + 1048 + 10, App.EXEName & ".cUxTheme", sMsg & vbCrLf & sError
   Else
      'Debug.Print sMsg & vbCrLf & sError
   End If
End Sub

Private Function WinAPIError(ByVal lLastDLLError As Long) As String
Dim sBuff As String
Dim lCount As Long
    
    ' Return the error message associated with LastDLLError:
    sBuff = String$(256, 0)
    lCount = FormatMessage( _
      FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, _
       0, lLastDLLError, 0&, sBuff, Len(sBuff), ByVal 0)
    If lCount Then
       WinAPIError = left$(sBuff, lCount)
    End If

End Function

Private Sub Class_Initialize()
   m_sClass = "Button"
   m_lPartId = 1
   m_lStateId = 1
   m_lIconIndex = -1
End Sub

