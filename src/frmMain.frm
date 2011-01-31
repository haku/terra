VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "//terra"
   ClientHeight    =   7720
   ClientLeft      =   40
   ClientTop       =   340
   ClientWidth     =   9040
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7720
   ScaleWidth      =   9040
   Begin VB.Timer tmrVidBumper 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7680
      Top             =   1560
   End
   Begin MSComctlLib.ImageList imlTabs 
      Left            =   4200
      Top             =   240
      _ExtentX        =   670
      _ExtentY        =   670
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6042
            Key             =   "vid"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6133
            Key             =   "lib"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":61AA
            Key             =   "sys"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":629E
            Key             =   "pl"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":639C
            Key             =   "arc"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6496
            Key             =   "cue"
         EndProperty
      EndProperty
   End
   Begin terra.ctlFrame fraVidPrev 
      Height          =   610
      Left            =   8280
      Top             =   0
      Visible         =   0   'False
      Width           =   730
      _ExtentX        =   1288
      _ExtentY        =   1076
      Caption         =   ""
      Backcolour      =   0
   End
   Begin VB.PictureBox picIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   160
      Index           =   5
      Left            =   8160
      Picture         =   "frmMain.frx":6586
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   89
      Top             =   2400
      Visible         =   0   'False
      Width           =   160
   End
   Begin VB.PictureBox picIconMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   160
      Index           =   5
      Left            =   8400
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   88
      Top             =   2400
      Visible         =   0   'False
      Width           =   160
   End
   Begin VB.PictureBox picIconMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   160
      Index           =   4
      Left            =   8400
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   87
      Top             =   2160
      Visible         =   0   'False
      Width           =   160
   End
   Begin VB.PictureBox picIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   160
      Index           =   4
      Left            =   8160
      Picture         =   "frmMain.frx":666A
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   86
      Top             =   2160
      Visible         =   0   'False
      Width           =   160
   End
   Begin VB.PictureBox picIconMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   160
      Index           =   3
      Left            =   8400
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   85
      Top             =   1920
      Visible         =   0   'False
      Width           =   160
   End
   Begin VB.PictureBox picIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   160
      Index           =   3
      Left            =   8160
      Picture         =   "frmMain.frx":6713
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   84
      Top             =   1920
      Visible         =   0   'False
      Width           =   160
   End
   Begin VB.PictureBox picIconMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   160
      Index           =   2
      Left            =   8400
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   83
      Top             =   1680
      Visible         =   0   'False
      Width           =   160
   End
   Begin VB.PictureBox picIconMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   160
      Index           =   1
      Left            =   8400
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   82
      Top             =   1440
      Visible         =   0   'False
      Width           =   160
   End
   Begin VB.PictureBox picIconMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   160
      Index           =   0
      Left            =   8400
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   81
      Top             =   1200
      Visible         =   0   'False
      Width           =   160
   End
   Begin VB.PictureBox picIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   160
      Index           =   2
      Left            =   8160
      Picture         =   "frmMain.frx":680E
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   80
      Top             =   1680
      Visible         =   0   'False
      Width           =   160
   End
   Begin VB.PictureBox picIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   160
      Index           =   1
      Left            =   8160
      Picture         =   "frmMain.frx":68BB
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   79
      Top             =   1440
      Visible         =   0   'False
      Width           =   160
   End
   Begin VB.PictureBox picIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   160
      Index           =   0
      Left            =   8160
      Picture         =   "frmMain.frx":6967
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   78
      Top             =   1200
      Visible         =   0   'False
      Width           =   160
   End
   Begin terra.ctlSysTray sytMinimize 
      Left            =   8040
      Top             =   840
      _ExtentX        =   441
      _ExtentY        =   441
   End
   Begin terra.ctlFrame fraPref 
      Height          =   4210
      Left            =   120
      Top             =   2880
      Width           =   8650
      _ExtentX        =   15258
      _ExtentY        =   7426
      Caption         =   ""
      Backcolour      =   -2147483633
      Begin terra.ctlFrame fraPrefPage 
         Height          =   3250
         Index           =   5
         Left            =   1680
         Top             =   0
         Width           =   6010
         _ExtentX        =   10601
         _ExtentY        =   5733
         Caption         =   "media library"
         Backcolour      =   -2147483633
         Begin VB.CommandButton cmdPrefLibGotoFolders 
            Caption         =   "chose media foldes..."
            Height          =   300
            Left            =   360
            TabIndex        =   34
            Top             =   720
            Width           =   2170
         End
         Begin VB.CommandButton cmdPrefLibRunDefScan 
            Caption         =   "run default scan"
            Height          =   300
            Left            =   360
            TabIndex        =   35
            Top             =   1680
            Width           =   2170
         End
         Begin VB.CommandButton cmdPrefLibGotoScan 
            Caption         =   "advanced..."
            Height          =   300
            Left            =   360
            TabIndex        =   36
            Top             =   2040
            Width           =   2170
         End
         Begin VB.Label lblPrefLibFldCnt 
            AutoSize        =   -1  'True
            Caption         =   "### folders selected."
            Height          =   300
            Left            =   2760
            TabIndex        =   116
            Tag             =   " folders selected."
            Top             =   720
            Width           =   1450
         End
         Begin VB.Label lblPrefLibPage 
            AutoSize        =   -1  'True
            Caption         =   $"frmMain.frx":6A55
            Height          =   360
            Index           =   1
            Left            =   120
            TabIndex        =   112
            Top             =   1200
            Width           =   5770
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblPrefLibPage 
            AutoSize        =   -1  'True
            Caption         =   $"frmMain.frx":6AFE
            Height          =   360
            Index           =   0
            Left            =   120
            TabIndex        =   111
            Top             =   240
            Width           =   5770
            WordWrap        =   -1  'True
         End
      End
      Begin terra.ctlFrame fraPrefPage 
         Height          =   3970
         Index           =   0
         Left            =   1680
         Top             =   0
         Width           =   6010
         _ExtentX        =   10601
         _ExtentY        =   7003
         Caption         =   "display"
         Backcolour      =   -2147483633
         Begin VB.CheckBox chkOptDisplayMintray 
            Caption         =   "minimize to tray."
            Height          =   250
            Left            =   360
            TabIndex        =   6
            Top             =   1980
            Width           =   5530
         End
         Begin VB.CheckBox chkOptDisplayMinivid 
            Caption         =   "show preview-vid when not on display tab (letter-boxed to 4:3)."
            Height          =   250
            Left            =   360
            TabIndex        =   5
            Top             =   1720
            Width           =   5530
         End
         Begin VB.CheckBox chkOptDisplayRepDashNewLine 
            Caption         =   "when showing title in display tab, replce "" - "" with a carrige return."
            Height          =   250
            Left            =   360
            TabIndex        =   4
            Top             =   1460
            Width           =   5530
         End
         Begin VB.ComboBox cbbTabPosition 
            Height          =   260
            Left            =   360
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   480
            Width           =   4570
         End
         Begin VB.CheckBox chkOptCuebtm 
            Caption         =   "show cue at bottom of window."
            Height          =   250
            Left            =   360
            TabIndex        =   3
            Top             =   1200
            Value           =   1  'Checked
            Width           =   5530
         End
         Begin VB.Label lblPrefDisp 
            AutoSize        =   -1  'True
            Caption         =   "tab bar position:"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   76
            Top             =   240
            Width           =   1120
         End
         Begin VB.Label lblPrefDisp 
            AutoSize        =   -1  'True
            Caption         =   "miscellaneous display options."
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   74
            Top             =   960
            Width           =   2110
         End
      End
      Begin VB.CommandButton cmdPrefClose 
         Caption         =   "close tab"
         Height          =   250
         Left            =   0
         TabIndex        =   115
         Tag             =   "save pref. now"
         Top             =   3960
         Width           =   1450
      End
      Begin terra.ctlFrame fraPrefPage 
         Height          =   2050
         Index           =   1
         Left            =   1680
         Tag             =   "  > "
         Top             =   0
         Width           =   6010
         _ExtentX        =   10601
         _ExtentY        =   3616
         Caption         =   "lists"
         Backcolour      =   -2147483633
         Begin VB.CheckBox chkOptLibNum 
            Caption         =   "index number."
            Height          =   250
            Left            =   360
            TabIndex        =   7
            Top             =   480
            Value           =   1  'Checked
            Width           =   5530
         End
         Begin VB.CheckBox chkOptLibcols 
            Caption         =   "date last played."
            Height          =   250
            Index           =   1
            Left            =   360
            TabIndex        =   9
            Top             =   980
            Width           =   5530
         End
         Begin VB.CheckBox chkOptLibcols 
            Caption         =   "file duration (mmm:ss)."
            Height          =   250
            Index           =   4
            Left            =   360
            TabIndex        =   93
            Top             =   1730
            Value           =   1  'Checked
            Width           =   5530
         End
         Begin VB.CheckBox chkOptLibcols 
            Caption         =   "play counts."
            Height          =   250
            Index           =   0
            Left            =   360
            TabIndex        =   8
            Top             =   730
            Value           =   1  'Checked
            Width           =   5530
         End
         Begin VB.CheckBox chkOptLibcols 
            Caption         =   "date added to library."
            Height          =   250
            Index           =   2
            Left            =   360
            TabIndex        =   10
            Top             =   1230
            Width           =   5530
         End
         Begin VB.CheckBox chkOptLibcols 
            Caption         =   "hash code."
            Height          =   250
            Index           =   3
            Left            =   360
            TabIndex        =   11
            Top             =   1480
            Width           =   5530
         End
         Begin VB.Label lblPrefLists 
            AutoSize        =   -1  'True
            Caption         =   "columns to shown in lists (where data are available)."
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   92
            Top             =   240
            Width           =   3570
         End
      End
      Begin terra.ctlFrame fraPrefPage 
         Height          =   3010
         Index           =   2
         Left            =   1680
         Tag             =   "  > "
         Top             =   0
         Width           =   6010
         _ExtentX        =   10601
         _ExtentY        =   5309
         Caption         =   "names"
         Backcolour      =   -2147483633
         Begin VB.CommandButton cmdPrefDirHideRem 
            Caption         =   "remove selected"
            Height          =   250
            Left            =   2160
            TabIndex        =   107
            Top             =   480
            Width           =   1690
         End
         Begin VB.CommandButton cmdPrefDirHideAdd 
            Caption         =   "add a folder"
            Height          =   250
            Left            =   360
            TabIndex        =   106
            Top             =   480
            Width           =   1690
         End
         Begin VB.ListBox lstPrefDirHide 
            Height          =   850
            IntegralHeight  =   0   'False
            Left            =   360
            MultiSelect     =   2  'Extended
            TabIndex        =   94
            Top             =   840
            Width           =   5530
         End
         Begin VB.CheckBox chkOptHideExt 
            Caption         =   "hide file extensions."
            Height          =   250
            Left            =   360
            TabIndex        =   96
            Top             =   2400
            Width           =   5530
         End
         Begin VB.CheckBox chkOptDirLvlsReverse 
            Caption         =   "count from right (instead of left)."
            Height          =   250
            Left            =   360
            TabIndex        =   97
            Top             =   2650
            Width           =   5530
         End
         Begin MSComctlLib.Slider sldOptDirLvls 
            Height          =   250
            Left            =   960
            TabIndex        =   95
            Top             =   2040
            Width           =   4930
            _ExtentX        =   8696
            _ExtentY        =   441
            _Version        =   393216
            LargeChange     =   1
            Max             =   20
         End
         Begin VB.Label lblPrefNames 
            AutoSize        =   -1  'True
            Caption         =   "try and hide the following strings from the front of paths."
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   105
            Top             =   240
            Width           =   3760
         End
         Begin VB.Label lblOptDirLvls 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            Height          =   250
            Left            =   240
            TabIndex        =   104
            Top             =   2040
            Width           =   610
         End
         Begin VB.Label lblPrefNames 
            AutoSize        =   -1  'True
            Caption         =   "otherwise, number directory levels to show in lists."
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   103
            Top             =   1800
            Width           =   3420
         End
      End
      Begin terra.ctlFrame fraPrefPage 
         Height          =   3130
         Index           =   11
         Left            =   1680
         Tag             =   "  > "
         Top             =   0
         Width           =   6010
         _ExtentX        =   10601
         _ExtentY        =   5521
         Caption         =   "log"
         Backcolour      =   -2147483633
         Begin VB.TextBox txtLog 
            Height          =   2770
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   102
            Top             =   240
            Width           =   5770
         End
      End
      Begin terra.ctlFrame fraPrefPage 
         Height          =   1090
         Index           =   10
         Left            =   1680
         Top             =   0
         Width           =   6010
         _ExtentX        =   10601
         _ExtentY        =   1923
         Caption         =   "error reporting"
         Backcolour      =   -2147483633
         Begin VB.ComboBox cbbOptErrNoplay 
            Height          =   260
            Left            =   360
            Style           =   2  'Dropdown List
            TabIndex        =   91
            Top             =   480
            Width           =   4570
         End
         Begin VB.Label lblPrefErr 
            AutoSize        =   -1  'True
            Caption         =   "when a file can not be found or is not playable..."
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   90
            Top             =   240
            Width           =   3240
         End
      End
      Begin terra.ctlFrame fraPrefPage 
         Height          =   1090
         Index           =   9
         Left            =   1680
         Top             =   0
         Width           =   6010
         _ExtentX        =   10601
         _ExtentY        =   1923
         Caption         =   "playlists"
         Backcolour      =   -2147483633
         Begin MSComctlLib.Slider sldOptPlAutosave 
            Height          =   250
            Left            =   120
            TabIndex        =   43
            Top             =   480
            Width           =   5770
            _ExtentX        =   10178
            _ExtentY        =   441
            _Version        =   393216
            SmallChange     =   5
            Max             =   240
            TickFrequency   =   5
         End
         Begin VB.Label lblPrefPl 
            AutoSize        =   -1  'True
            Caption         =   "auto-save timer:"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   60
            Top             =   240
            Width           =   1110
         End
         Begin VB.Label lblOptPlAutosave 
            Caption         =   "0"
            Height          =   250
            Left            =   1320
            TabIndex        =   59
            Top             =   240
            Width           =   2410
         End
      End
      Begin terra.ctlFrame fraPrefPage 
         Height          =   2890
         Index           =   4
         Left            =   1680
         Top             =   0
         Width           =   6010
         _ExtentX        =   10601
         _ExtentY        =   5098
         Caption         =   "global hot keys"
         Backcolour      =   -2147483633
         Begin VB.TextBox txtOptHk 
            Alignment       =   2  'Center
            Height          =   270
            Index           =   4
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   2160
            Width           =   2050
         End
         Begin VB.CommandButton cmdOptHk 
            Caption         =   "set"
            Height          =   250
            Index           =   4
            Left            =   3240
            TabIndex        =   33
            Top             =   2160
            Width           =   610
         End
         Begin VB.CommandButton cmdOptHk 
            Caption         =   "set"
            Height          =   250
            Index           =   3
            Left            =   3240
            TabIndex        =   31
            Top             =   1800
            Width           =   610
         End
         Begin VB.TextBox txtOptHk 
            Alignment       =   2  'Center
            Height          =   270
            Index           =   3
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   30
            Top             =   1800
            Width           =   2050
         End
         Begin VB.CommandButton cmdOptHk 
            Caption         =   "set"
            Height          =   250
            Index           =   2
            Left            =   3240
            TabIndex        =   29
            Top             =   1440
            Width           =   610
         End
         Begin VB.TextBox txtOptHk 
            Alignment       =   2  'Center
            Height          =   270
            Index           =   2
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   1440
            Width           =   2050
         End
         Begin VB.CommandButton cmdOptHk 
            Caption         =   "set"
            Height          =   250
            Index           =   1
            Left            =   3240
            TabIndex        =   27
            Top             =   1080
            Width           =   610
         End
         Begin VB.TextBox txtOptHk 
            Alignment       =   2  'Center
            Height          =   270
            Index           =   1
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   1080
            Width           =   2050
         End
         Begin VB.CommandButton cmdOptHk 
            Caption         =   "set"
            Height          =   250
            Index           =   0
            Left            =   3240
            TabIndex        =   25
            Top             =   720
            Width           =   610
         End
         Begin VB.TextBox txtOptHk 
            Alignment       =   2  'Center
            Height          =   270
            Index           =   0
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   720
            Width           =   2050
         End
         Begin VB.Label lblPrefHk 
            AutoSize        =   -1  'True
            Caption         =   "goto:"
            Height          =   180
            Index           =   5
            Left            =   120
            TabIndex        =   101
            Top             =   2160
            Width           =   350
         End
         Begin VB.Label lblOptHk 
            AutoSize        =   -1  'True
            Caption         =   "key not set."
            Height          =   180
            Index           =   4
            Left            =   3960
            TabIndex        =   100
            Top             =   2160
            Width           =   780
         End
         Begin VB.Label lblPrefHk 
            AutoSize        =   -1  'True
            Caption         =   "todo: replace key code numbers with actual key description."
            Height          =   180
            Index           =   6
            Left            =   120
            TabIndex        =   71
            Top             =   2520
            Width           =   4110
         End
         Begin VB.Label lblOptHk 
            AutoSize        =   -1  'True
            Caption         =   "key not set."
            Height          =   180
            Index           =   3
            Left            =   3960
            TabIndex        =   70
            Top             =   1800
            Width           =   780
         End
         Begin VB.Label lblPrefHk 
            AutoSize        =   -1  'True
            Caption         =   "next:"
            Height          =   180
            Index           =   4
            Left            =   120
            TabIndex        =   69
            Top             =   1800
            Width           =   320
         End
         Begin VB.Label lblOptHk 
            AutoSize        =   -1  'True
            Caption         =   "key not set."
            Height          =   180
            Index           =   2
            Left            =   3960
            TabIndex        =   68
            Top             =   1440
            Width           =   780
         End
         Begin VB.Label lblPrefHk 
            AutoSize        =   -1  'True
            Caption         =   "stop:"
            Height          =   180
            Index           =   3
            Left            =   120
            TabIndex        =   67
            Top             =   1440
            Width           =   340
         End
         Begin VB.Label lblOptHk 
            AutoSize        =   -1  'True
            Caption         =   "key not set."
            Height          =   180
            Index           =   1
            Left            =   3960
            TabIndex        =   66
            Top             =   1080
            Width           =   780
         End
         Begin VB.Label lblPrefHk 
            AutoSize        =   -1  'True
            Caption         =   "play / pause:"
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   65
            Top             =   1080
            Width           =   870
         End
         Begin VB.Label lblOptHk 
            AutoSize        =   -1  'True
            Caption         =   "key not set."
            Height          =   180
            Index           =   0
            Left            =   3960
            TabIndex        =   64
            Top             =   720
            Width           =   780
         End
         Begin VB.Label lblPrefHk 
            AutoSize        =   -1  'True
            Caption         =   "show / hide:"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   63
            Top             =   720
            Width           =   820
         End
         Begin VB.Label lblPrefHk 
            AutoSize        =   -1  'True
            Caption         =   $"frmMain.frx":6B90
            Height          =   420
            Index           =   0
            Left            =   120
            TabIndex        =   62
            Top             =   240
            Width           =   5760
            WordWrap        =   -1  'True
         End
      End
      Begin terra.ctlFrame fraPrefPage 
         Height          =   1570
         Index           =   3
         Left            =   1680
         Tag             =   "  > "
         Top             =   0
         Width           =   6010
         _ExtentX        =   10601
         _ExtentY        =   2769
         Caption         =   "gui"
         Backcolour      =   -2147483633
         Begin VB.ComboBox cbbOptColourscheme 
            Height          =   260
            Left            =   360
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   1200
            Width           =   4570
         End
         Begin VB.ComboBox cbbOptDrawmode 
            Height          =   260
            Left            =   360
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   480
            Width           =   4570
         End
         Begin VB.Label lblPrefGui 
            AutoSize        =   -1  'True
            Caption         =   "when using the internal theme, use the following colours..."
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   75
            Top             =   960
            Width           =   3910
         End
         Begin VB.Label lblPrefGui 
            AutoSize        =   -1  'True
            Caption         =   "draw gdi using..."
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   61
            Top             =   240
            Width           =   1130
         End
      End
      Begin terra.ctlFrame fraPrefPage 
         Height          =   3970
         Index           =   8
         Left            =   1680
         Tag             =   "  > "
         Top             =   0
         Width           =   6010
         _ExtentX        =   10601
         _ExtentY        =   7003
         Caption         =   "library scan"
         Backcolour      =   -2147483633
         Begin VB.CheckBox chkOptLibMaintR 
            Caption         =   "clear hash code (i.e. flag for regeneration) on re-find of files."
            Height          =   250
            Index           =   2
            Left            =   360
            TabIndex        =   22
            Top             =   2780
            Value           =   1  'Checked
            Width           =   5530
         End
         Begin VB.CheckBox chkOptLibMaintR 
            Caption         =   "by default, prompt to mark items as missing (instead of removing them)."
            Height          =   250
            Index           =   1
            Left            =   360
            TabIndex        =   21
            Top             =   2530
            Value           =   1  'Checked
            Width           =   5530
         End
         Begin VB.CheckBox chkOptLibMaintR 
            Caption         =   "ignore missing files if they are marked as missing."
            Height          =   250
            Index           =   0
            Left            =   360
            TabIndex        =   20
            Top             =   2280
            Value           =   1  'Checked
            Width           =   5530
         End
         Begin VB.CheckBox chkOptLibMaint 
            Caption         =   "clean up unused database file space."
            Height          =   250
            Index           =   5
            Left            =   360
            TabIndex        =   19
            Top             =   1730
            Value           =   1  'Checked
            Width           =   5530
         End
         Begin VB.CheckBox chkOptLibMaint 
            Caption         =   "re-hash all files. *"
            Height          =   250
            Index           =   2
            Left            =   600
            TabIndex        =   16
            Top             =   960
            Width           =   5170
         End
         Begin VB.CheckBox chkOptLibMaint 
            Caption         =   "read file duration (and check that files are valid). *"
            Height          =   250
            Index           =   4
            Left            =   360
            TabIndex        =   18
            Top             =   1480
            Width           =   5530
         End
         Begin VB.CheckBox chkOptLibMaint 
            Caption         =   "scan watch folders."
            Height          =   250
            Index           =   0
            Left            =   360
            TabIndex        =   14
            Top             =   480
            Value           =   1  'Checked
            Width           =   5530
         End
         Begin VB.CheckBox chkOptLibMaint 
            Caption         =   "scan for missing files and generate hash codes."
            Height          =   250
            Index           =   1
            Left            =   360
            TabIndex        =   15
            Top             =   730
            Value           =   1  'Checked
            Width           =   5530
         End
         Begin VB.CheckBox chkOptLibMaint 
            Caption         =   "scan for moved and duplicate files."
            Height          =   250
            Index           =   3
            Left            =   360
            TabIndex        =   17
            Top             =   1230
            Value           =   1  'Checked
            Width           =   5530
         End
         Begin VB.CommandButton cmdPrefLibMaint 
            Caption         =   "start media library scan"
            Height          =   300
            Left            =   120
            TabIndex        =   23
            Top             =   3120
            Width           =   3490
         End
         Begin VB.Label lblPrefLibScan 
            AutoSize        =   -1  'True
            Caption         =   "* these stages may take a long time."
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   110
            Top             =   3480
            Width           =   2500
         End
         Begin VB.Label lblPrefLibScan 
            AutoSize        =   -1  'True
            Caption         =   "you probably want to leave all these advanced options checked."
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   109
            Top             =   2040
            Width           =   4370
         End
         Begin VB.Label lblPrefLibScan 
            AutoSize        =   -1  'True
            Caption         =   "to run library maintenance, first select required stages and options."
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   58
            Top             =   240
            Width           =   4570
         End
      End
      Begin terra.ctlFrame fraPrefPage 
         Height          =   3970
         Index           =   7
         Left            =   1680
         Tag             =   "  > "
         Top             =   0
         Width           =   6010
         _ExtentX        =   10601
         _ExtentY        =   7003
         Caption         =   "file types"
         Backcolour      =   -2147483633
         Begin VB.CommandButton cmdOptFileext 
            Caption         =   "default..."
            Height          =   250
            Left            =   4920
            TabIndex        =   42
            Top             =   1440
            Width           =   970
         End
         Begin VB.TextBox txtOptFileext 
            Height          =   270
            Left            =   360
            TabIndex        =   41
            Top             =   1080
            Width           =   5530
         End
         Begin VB.Label lblPrefLibFileTyp 
            AutoSize        =   -1  'True
            Caption         =   "media file extensions (include ""."", seperate with ""|""):"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   114
            Top             =   840
            Width           =   3530
         End
         Begin VB.Label lblPrefLibFileTyp 
            AutoSize        =   -1  'True
            Caption         =   $"frmMain.frx":6C25
            Height          =   360
            Index           =   0
            Left            =   120
            TabIndex        =   113
            Top             =   240
            Width           =   5770
            WordWrap        =   -1  'True
         End
      End
      Begin terra.ctlFrame fraPrefPage 
         Height          =   1810
         Index           =   6
         Left            =   1680
         Tag             =   "  > "
         Top             =   0
         Width           =   6010
         _ExtentX        =   10601
         _ExtentY        =   3193
         Caption         =   "media folders"
         Backcolour      =   -2147483633
         Begin VB.CommandButton cmdPrefLibFoldersDone 
            Caption         =   "back to library overview"
            Height          =   300
            Left            =   3720
            TabIndex        =   40
            Top             =   1320
            Width           =   2170
         End
         Begin VB.CommandButton cmdLibWatchRem 
            Caption         =   "remove selected folders"
            Height          =   250
            Left            =   2280
            TabIndex        =   38
            Top             =   240
            Width           =   2050
         End
         Begin VB.CommandButton cmdLibWatchAdd 
            Caption         =   "add a folder"
            Height          =   250
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   2050
         End
         Begin VB.ListBox lstLibWatch 
            Height          =   610
            IntegralHeight  =   0   'False
            Left            =   120
            MultiSelect     =   2  'Extended
            TabIndex        =   39
            Top             =   600
            Width           =   5770
         End
      End
      Begin VB.Timer tmrPrefSave 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   0
         Top             =   3240
      End
      Begin VB.CommandButton cmdOptSave 
         Caption         =   "save now"
         Height          =   250
         Left            =   0
         TabIndex        =   108
         Tag             =   "save now"
         Top             =   3600
         Width           =   1450
      End
      Begin VB.ListBox lstPrefPages 
         Height          =   3490
         IntegralHeight  =   0   'False
         Left            =   60
         TabIndex        =   1
         Top             =   60
         Width           =   1450
      End
   End
   Begin terra.ctlFrame fraLib 
      Height          =   1210
      Left            =   3360
      Top             =   1440
      Width           =   4210
      _ExtentX        =   7426
      _ExtentY        =   2134
      Caption         =   ""
      Backcolour      =   -2147483633
      Begin VB.VScrollBar vsbMdbLst 
         Height          =   370
         Left            =   2400
         Max             =   0
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   360
         Width           =   250
      End
      Begin VB.PictureBox picMdbLst 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   370
         Left            =   0
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   37
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   229
         TabIndex        =   47
         Tag             =   "0"
         Top             =   360
         Width           =   2290
      End
      Begin terra.ctlXpThemeBtn cmdLib 
         Height          =   280
         Left            =   2280
         Top             =   0
         Width           =   1330
         _ExtentX        =   2346
         _ExtentY        =   494
         DrawMode        =   0
         Caption         =   "media library"
         BtnType         =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         BackColour      =   -2147483633
      End
      Begin terra.ctlFrame fraLibSearch 
         Height          =   310
         Left            =   0
         Top             =   840
         Visible         =   0   'False
         Width           =   5410
         _ExtentX        =   9543
         _ExtentY        =   547
         Caption         =   ""
         Backcolour      =   -2147483633
         Begin VB.CommandButton cmdLibSearch 
            Caption         =   "á"
            BeginProperty Font 
               Name            =   "Wingdings"
               Size            =   8
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   250
            Index           =   1
            Left            =   3600
            TabIndex        =   46
            Top             =   30
            Width           =   490
         End
         Begin VB.CommandButton cmdLibSearch 
            Caption         =   "â"
            BeginProperty Font 
               Name            =   "Wingdings"
               Size            =   8
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   250
            Index           =   0
            Left            =   3000
            TabIndex        =   45
            Top             =   30
            Width           =   490
         End
         Begin VB.TextBox txtLibSearch 
            Height          =   270
            Left            =   960
            TabIndex        =   44
            Top             =   30
            Width           =   1930
         End
         Begin terra.ctlXpThemeBtn cmdLibSearchClose 
            Height          =   250
            Left            =   0
            Top             =   0
            Width           =   250
            _ExtentX        =   441
            _ExtentY        =   441
            DrawMode        =   1
            Caption         =   "r"
            BtnType         =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Marlett"
               Size            =   8
               Charset         =   2
               Weight          =   500
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   -1  'True
            BackColour      =   -2147483633
         End
         Begin VB.Label lblLibFind 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "find:"
            Height          =   180
            Index           =   12
            Left            =   360
            TabIndex        =   73
            Top             =   60
            Width           =   290
         End
      End
      Begin terra.ctlXpThemeBtn cmdLibOrder 
         Height          =   280
         Left            =   1440
         Top             =   0
         Width           =   730
         _ExtentX        =   1288
         _ExtentY        =   494
         DrawMode        =   0
         Caption         =   "view"
         BtnType         =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         BackColour      =   -2147483633
      End
      Begin VB.Label lblDbStat 
         AutoSize        =   -1  'True
         Caption         =   "[db stat]"
         Height          =   180
         Left            =   0
         TabIndex        =   52
         Top             =   0
         Width           =   550
      End
   End
   Begin VB.Timer tmrPlAutoSave 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7680
      Top             =   1200
   End
   Begin terra.ctlFrame fraNotice 
      Height          =   250
      Index           =   1
      Left            =   2640
      Top             =   7320
      Visible         =   0   'False
      Width           =   2290
      _ExtentX        =   4039
      _ExtentY        =   441
      Caption         =   ""
      Backcolour      =   -2147483643
      Begin terra.ctlXpThemeBtn cmdLibStop 
         Height          =   250
         Index           =   0
         Left            =   60
         Top             =   0
         Width           =   490
         _ExtentX        =   864
         _ExtentY        =   441
         DrawMode        =   0
         Caption         =   "stop"
         BtnType         =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         BackColour      =   -2147483643
      End
      Begin VB.Label lblLibHash 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "starting maintenance..."
         Height          =   180
         Left            =   600
         TabIndex        =   55
         Top             =   30
         Width           =   1560
      End
   End
   Begin terra.ctlXpThemeSld sldPlay 
      Height          =   240
      Left            =   120
      Top             =   360
      Width           =   3130
      _ExtentX        =   5521
      _ExtentY        =   423
      Enabled         =   -1  'True
      DrawMode        =   0
   End
   Begin terra.ctlFrame fraVid 
      Height          =   490
      Left            =   120
      Top             =   2280
      Width           =   3130
      _ExtentX        =   5521
      _ExtentY        =   864
      Caption         =   ""
      Backcolour      =   -2147483633
      Begin terra.ctlFrame fraVidPlace 
         Height          =   250
         Left            =   0
         Top             =   240
         Width           =   1090
         _ExtentX        =   1923
         _ExtentY        =   441
         Caption         =   ""
         Backcolour      =   -2147483633
      End
      Begin terra.ctlXpThemeBtn cmdVidFull 
         Height          =   280
         Left            =   1800
         Top             =   0
         Width           =   1330
         _ExtentX        =   2346
         _ExtentY        =   494
         DrawMode        =   0
         Caption         =   "full screen"
         BtnType         =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         BackColour      =   -2147483633
      End
      Begin terra.ctlXpThemeBtn cmdVidMinivid 
         Height          =   280
         Left            =   600
         Top             =   0
         Width           =   1330
         _ExtentX        =   2346
         _ExtentY        =   494
         DrawMode        =   0
         Caption         =   "mini-display"
         BtnType         =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         BackColour      =   -2147483633
      End
      Begin VB.Label lblVidStat 
         AutoSize        =   -1  'True
         Caption         =   "no file loaded."
         Height          =   180
         Left            =   0
         TabIndex        =   54
         Top             =   0
         Width           =   970
      End
   End
   Begin VB.Timer tmrUpdateState 
      Interval        =   100
      Left            =   7680
      Top             =   840
   End
   Begin terra.ctlFrame fraPl 
      Height          =   730
      Left            =   1320
      Top             =   1440
      Width           =   1930
      _ExtentX        =   3404
      _ExtentY        =   1288
      Caption         =   ""
      Backcolour      =   -2147483633
      Begin VB.VScrollBar vsbPl 
         Height          =   610
         Left            =   1680
         Max             =   0
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   0
         Width           =   250
      End
      Begin VB.PictureBox picPl 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   250
         Left            =   840
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   85
         TabIndex        =   50
         Tag             =   "0"
         Top             =   360
         Width           =   850
      End
      Begin terra.ctlXpThemeBtn cmdPl 
         Height          =   280
         Left            =   360
         Top             =   0
         Width           =   1330
         _ExtentX        =   2346
         _ExtentY        =   494
         DrawMode        =   0
         Caption         =   "playlist"
         BtnType         =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         BackColour      =   -2147483633
      End
      Begin VB.ListBox lstPlArc 
         Height          =   370
         IntegralHeight  =   0   'False
         Left            =   0
         TabIndex        =   98
         Top             =   480
         Visible         =   0   'False
         Width           =   1330
      End
      Begin VB.Label lblPlStat 
         AutoSize        =   -1  'True
         Caption         =   "[pl stat]"
         Height          =   180
         Left            =   0
         TabIndex        =   57
         Top             =   0
         Width           =   490
      End
   End
   Begin terra.ctlXpThemeBtn cmdMnu 
      Height          =   300
      Index           =   3
      Left            =   4680
      Top             =   0
      Width           =   1330
      _ExtentX        =   2346
      _ExtentY        =   529
      DrawMode        =   0
      Caption         =   "play mode"
      BtnType         =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      BackColour      =   -2147483633
   End
   Begin terra.ctlXpThemeBtn cmdPlayState 
      Height          =   300
      Index           =   1
      Left            =   480
      Top             =   0
      Width           =   370
      _ExtentX        =   653
      _ExtentY        =   529
      DrawMode        =   0
      Caption         =   "4"
      BtnType         =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      BackColour      =   -2147483633
   End
   Begin terra.ctlXpThemeBtn cmdPlayNext 
      Height          =   300
      Left            =   1560
      Top             =   0
      Width           =   490
      _ExtentX        =   864
      _ExtentY        =   529
      DrawMode        =   0
      Caption         =   ":"
      BtnType         =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      BackColour      =   -2147483633
   End
   Begin terra.ctlXpThemeBtn cmdPlayState 
      Height          =   300
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   370
      _ExtentX        =   653
      _ExtentY        =   529
      DrawMode        =   0
      Caption         =   "<"
      BtnType         =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      BackColour      =   -2147483633
   End
   Begin terra.ctlFrame fraNotice 
      Height          =   250
      Index           =   0
      Left            =   120
      Top             =   7320
      Visible         =   0   'False
      Width           =   2410
      _ExtentX        =   4251
      _ExtentY        =   441
      Caption         =   ""
      Backcolour      =   -2147483643
      Begin terra.ctlXpThemeBtn cmdLibStop 
         Height          =   250
         Index           =   1
         Left            =   60
         Top             =   0
         Width           =   490
         _ExtentX        =   864
         _ExtentY        =   441
         DrawMode        =   0
         Caption         =   "stop"
         BtnType         =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         BackColour      =   -2147483643
      End
      Begin VB.Label lblLibAddStat 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "adding media to library..."
         Height          =   180
         Left            =   600
         TabIndex        =   56
         Top             =   30
         Width           =   1710
      End
   End
   Begin terra.ctlXpThemeBtn cmdMnu 
      Height          =   300
      Index           =   0
      Left            =   7560
      Top             =   0
      Width           =   610
      _ExtentX        =   1076
      _ExtentY        =   529
      DrawMode        =   0
      Caption         =   ""
      BtnType         =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      BackColour      =   -2147483633
   End
   Begin terra.ctlFrame fraCue 
      Height          =   730
      Left            =   120
      Top             =   1440
      Width           =   1090
      _ExtentX        =   1923
      _ExtentY        =   1288
      Caption         =   ""
      Backcolour      =   -2147483633
      Begin VB.PictureBox picCue 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   490
         Left            =   0
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   49
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   85
         TabIndex        =   49
         Tag             =   "0"
         Top             =   240
         Width           =   850
      End
      Begin VB.VScrollBar vsbCue 
         Enabled         =   0   'False
         Height          =   610
         Left            =   840
         Max             =   0
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   0
         Width           =   250
      End
      Begin VB.Label lblCueStat 
         AutoSize        =   -1  'True
         Caption         =   "[cue stat]"
         Height          =   180
         Left            =   0
         TabIndex        =   99
         Top             =   0
         Width           =   620
      End
   End
   Begin MSComctlLib.TabStrip tsMain 
      Height          =   370
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   7450
      _ExtentX        =   13141
      _ExtentY        =   653
      TabWidthStyle   =   1
      MultiRow        =   -1  'True
      ShowTips        =   0   'False
      TabMinWidth     =   529
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "display"
            Key             =   "vid"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "library"
            Key             =   "lib"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "archive"
            Key             =   "arc"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "pl(0)"
            Key             =   "pl"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin terra.ctlFrame fraNotice 
      Height          =   250
      Index           =   2
      Left            =   5280
      Top             =   7320
      Visible         =   0   'False
      Width           =   2290
      _ExtentX        =   4039
      _ExtentY        =   441
      Caption         =   ""
      Backcolour      =   -2147483643
      Begin terra.ctlXpThemeBtn cmdLibStop 
         Height          =   250
         Index           =   2
         Left            =   60
         Top             =   0
         Width           =   490
         _ExtentX        =   864
         _ExtentY        =   441
         DrawMode        =   0
         Caption         =   "stop"
         BtnType         =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         BackColour      =   -2147483643
      End
      Begin VB.Label lblFileCopy 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "copying files..."
         Height          =   180
         Left            =   600
         TabIndex        =   77
         Top             =   30
         Width           =   980
      End
   End
   Begin terra.ctlXpThemeBtn cmdMnu 
      Height          =   300
      Index           =   2
      Left            =   6120
      Top             =   0
      Width           =   610
      _ExtentX        =   1076
      _ExtentY        =   529
      DrawMode        =   0
      Caption         =   ""
      BtnType         =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      BackColour      =   -2147483633
   End
   Begin terra.ctlXpThemeBtn cmdPlayHistory 
      Height          =   300
      Left            =   960
      Top             =   0
      Width           =   490
      _ExtentX        =   864
      _ExtentY        =   529
      DrawMode        =   0
      Caption         =   "9"
      BtnType         =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      BackColour      =   -2147483633
   End
   Begin terra.ctlXpThemeBtn cmdMnu 
      Height          =   300
      Index           =   1
      Left            =   6840
      Top             =   0
      Width           =   610
      _ExtentX        =   1076
      _ExtentY        =   529
      DrawMode        =   0
      Caption         =   ""
      BtnType         =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      BackColour      =   -2147483633
   End
   Begin VB.Label lblPlayInfo 
      AutoSize        =   -1  'True
      Caption         =   " "
      Height          =   180
      Left            =   2640
      TabIndex        =   53
      Top             =   120
      Width           =   40
   End
   Begin VB.Menu mnuOpt 
      Caption         =   "options"
      Visible         =   0   'False
      Begin VB.Menu mnuOptTab 
         Caption         =   "show preferences ta&b"
      End
      Begin VB.Menu mnuOptMisc 
         Caption         =   "misc options"
         Begin VB.Menu mnuOptSpeed 
            Caption         =   "set playback &rate"
         End
      End
      Begin VB.Menu mnuOptSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&help"
         Begin VB.Menu mnuHelpAbout 
            Caption         =   "&about..."
         End
         Begin VB.Menu mnuHelpWiki 
            Caption         =   "terra &wiki..."
         End
      End
      Begin VB.Menu mnuDebug 
         Caption         =   "&debug"
         Begin VB.Menu mnuDebugDiv0 
            Caption         =   "this will crash terra with a divide by 0 error."
         End
         Begin VB.Menu mnuDebugRehook 
            Caption         =   "rehook."
         End
         Begin VB.Menu mnuDebugBumpvidwindow 
            Caption         =   "BumpVidWindow"
         End
      End
   End
   Begin VB.Menu mnuMode 
      Caption         =   "mode"
      Visible         =   0   'False
      Begin VB.Menu mnuModeI 
         Caption         =   "sequencial"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuModeI 
         Caption         =   "random"
         Index           =   1
      End
      Begin VB.Menu mnuModeI 
         Caption         =   "by start-count"
         Index           =   2
      End
      Begin VB.Menu mnuModeI 
         Caption         =   "by last-played"
         Index           =   3
      End
   End
   Begin VB.Menu mnuModeLists 
      Caption         =   "mode-lists"
      Visible         =   0   'False
      Begin VB.Menu mnuModeListsI 
         Caption         =   "single list"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuModeListsI 
         Caption         =   "all playlists"
         Index           =   1
      End
      Begin VB.Menu mnuModeListsI 
         Caption         =   "all lists"
         Index           =   2
      End
   End
   Begin VB.Menu mnuJump 
      Caption         =   "jump"
      Visible         =   0   'False
      Begin VB.Menu mnuJumpCurrent 
         Caption         =   "jump to current item"
      End
      Begin VB.Menu mnuJumpAuto 
         Caption         =   "auto-jump at start of track"
      End
   End
   Begin VB.Menu mnuNxsp 
      Caption         =   "nextspecial"
      Visible         =   0   'False
      Begin VB.Menu mnuNxspI 
         Caption         =   "next sequentially|[list mode]"
         Index           =   0
      End
      Begin VB.Menu mnuNxspI 
         Caption         =   "random|[any list]"
         Index           =   1
      End
      Begin VB.Menu mnuNxspI 
         Caption         =   "shuffle by start-count|[any list]"
         Index           =   2
      End
      Begin VB.Menu mnuNxspI 
         Caption         =   "shuffle by last played date|[any list]"
         Index           =   3
      End
      Begin VB.Menu mnuNxspI 
         Caption         =   "un-played|[library]"
         Index           =   4
      End
      Begin VB.Menu mnuNxspI 
         Caption         =   "prefered (finishes / starts > 0.7)|[library]"
         Index           =   5
      End
      Begin VB.Menu mnuNxspI 
         Caption         =   "not played in the last 30 days|[library]"
         Index           =   6
      End
      Begin VB.Menu mnuNxspI 
         Caption         =   "new (within 3 days of most recently added)|[library]"
         Index           =   7
      End
      Begin VB.Menu mnuNxspI 
         Caption         =   "old (1 of 10 played longest ago)|[library]"
         Index           =   8
      End
   End
   Begin VB.Menu mnuHistory 
      Caption         =   "history"
      Visible         =   0   'False
      Begin VB.Menu mnuHistoryI 
         Caption         =   "[]"
         Index           =   0
      End
   End
   Begin VB.Menu mnuLib 
      Caption         =   "lib"
      Visible         =   0   'False
      Begin VB.Menu mnuLibReq 
         Caption         =   "re&query"
      End
      Begin VB.Menu mnuLibOrder 
         Caption         =   "item order"
         Visible         =   0   'False
         Begin VB.Menu mnuLibOrderI 
            Caption         =   "sort by file path"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuLibOrderI 
            Caption         =   "sort by play count - started"
            Index           =   1
         End
         Begin VB.Menu mnuLibOrderI 
            Caption         =   "sort by play count - finished"
            Index           =   2
         End
         Begin VB.Menu mnuLibOrderI 
            Caption         =   "sort by date added to mdb"
            Index           =   3
         End
         Begin VB.Menu mnuLibOrderI 
            Caption         =   "sort by date last played"
            Index           =   4
         End
         Begin VB.Menu mnuLibOrderI 
            Caption         =   "sort by file duration"
            Index           =   5
         End
         Begin VB.Menu mnuLibOrderSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuLibOrderDirection 
            Caption         =   "sort ascending"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuLibOrderDirection 
            Caption         =   "sort descending"
            Index           =   1
         End
         Begin VB.Menu mnuLibOrderSep2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuLibOrderShowmis 
            Caption         =   "show missing files"
         End
      End
      Begin VB.Menu mnuLibFind 
         Caption         =   "&find..."
      End
      Begin VB.Menu mnuLibGoto 
         Caption         =   "&goto..."
      End
      Begin VB.Menu mnuLibReport 
         Caption         =   "report..."
      End
      Begin VB.Menu mnuLibConfig 
         Caption         =   "library configuration..."
      End
      Begin VB.Menu mnuLibSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLibSel 
         Caption         =   "selected items"
         Begin VB.Menu mnuLibAddcue 
            Caption         =   "add to cue"
         End
         Begin VB.Menu mnuLibAddpl 
            Caption         =   "add to playlist"
            Begin VB.Menu mnuLibAddplI 
               Caption         =   "[playlist]"
               Index           =   0
            End
            Begin VB.Menu mnuLibAddplSep1 
               Caption         =   "-"
            End
            Begin VB.Menu mnuLibAddplNew 
               Caption         =   "&new playlist"
            End
         End
         Begin VB.Menu mnuLibSelPathcopy 
            Caption         =   "copy file paths to clipboard"
         End
         Begin VB.Menu mnuLibSelFilecopy 
            Caption         =   "copy files to folder..."
         End
         Begin VB.Menu mnuLibSelSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuLibEnabled 
            Caption         =   "invert enabled"
         End
         Begin VB.Menu mnuLibPlaycnt 
            Caption         =   "play count"
            Begin VB.Menu mnuLibPlaycntI 
               Caption         =   "set..."
               Index           =   0
            End
            Begin VB.Menu mnuLibPlaycntI 
               Caption         =   "reset..."
               Index           =   1
            End
         End
         Begin VB.Menu mnuLibRemcrc 
            Caption         =   "remove crc data..."
         End
         Begin VB.Menu mnuLibRemsel 
            Caption         =   "&remove from library..."
         End
      End
   End
   Begin VB.Menu mnuPl 
      Caption         =   "playlist"
      Visible         =   0   'False
      Begin VB.Menu mnuPlNew 
         Caption         =   "new &list"
      End
      Begin VB.Menu mnuPlSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPlPlay 
         Caption         =   "play list"
      End
      Begin VB.Menu mnuPlSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPlEnab 
         Caption         =   "&enabled"
      End
      Begin VB.Menu mnuPlName 
         Caption         =   "re&name list..."
      End
      Begin VB.Menu mnuPlMoveTabI 
         Caption         =   "move tab left"
         Index           =   0
      End
      Begin VB.Menu mnuPlMoveTabI 
         Caption         =   "move tab right"
         Index           =   1
      End
      Begin VB.Menu mnuPlArc 
         Caption         =   "archive list"
      End
      Begin VB.Menu mnuPlDel 
         Caption         =   "&delete list"
      End
      Begin VB.Menu mnuPlSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPlSelect 
         Caption         =   "&select"
         Begin VB.Menu mnuPlSelAll 
            Caption         =   "&all"
         End
         Begin VB.Menu mnuPlSelMis 
            Caption         =   "missing items"
         End
         Begin VB.Menu mnuPlSelDup 
            Caption         =   "duplicate items"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuPlSort 
         Caption         =   "s&ort"
         Begin VB.Menu mnuPlSortI 
            Caption         =   "&reverse"
            Index           =   0
         End
         Begin VB.Menu mnuPlSortI 
            Caption         =   "by &path"
            Index           =   1
         End
         Begin VB.Menu mnuPlSortI 
            Caption         =   "by play count - &started"
            Index           =   2
         End
         Begin VB.Menu mnuPlSortI 
            Caption         =   "by play count - &finished"
            Index           =   3
         End
         Begin VB.Menu mnuPlSortI 
            Caption         =   "by &date last played"
            Index           =   4
         End
      End
      Begin VB.Menu mnuPlRemdup 
         Caption         =   "remove duplicates..."
      End
      Begin VB.Menu mnuPlRep 
         Caption         =   "repair..."
      End
      Begin VB.Menu mnuPlSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPlSel 
         Caption         =   "selected &items"
         Begin VB.Menu mnuPlAddcue 
            Caption         =   "add to cue"
         End
         Begin VB.Menu mnuPlMove 
            Caption         =   "move &up"
            Index           =   0
         End
         Begin VB.Menu mnuPlMove 
            Caption         =   "move do&wn"
            Index           =   1
         End
         Begin VB.Menu mnuPlRemitm 
            Caption         =   "&remove"
         End
         Begin VB.Menu mnuPlAddpl 
            Caption         =   "add to playlist"
            Begin VB.Menu mnuPlAddplI 
               Caption         =   "[]"
               Index           =   0
            End
            Begin VB.Menu mnuPlAddplSep1 
               Caption         =   "-"
            End
            Begin VB.Menu mnuPlAddplNew 
               Caption         =   "&new playlist"
            End
         End
         Begin VB.Menu mnuPlPathcopy 
            Caption         =   "copy file paths to clipboard"
         End
         Begin VB.Menu mnuPlFilecopy 
            Caption         =   "copy files to folder..."
         End
         Begin VB.Menu mnuPlSelSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPlPlaycnt 
            Caption         =   "play &count"
            Begin VB.Menu mnuPlPlaycntI 
               Caption         =   "&set..."
               Index           =   0
            End
            Begin VB.Menu mnuPlPlaycntI 
               Caption         =   "&reset..."
               Index           =   1
            End
            Begin VB.Menu mnuPlPlaycntI 
               Caption         =   "get from library"
               Index           =   2
            End
         End
         Begin VB.Menu mnuPlHash 
            Caption         =   "hash files"
         End
         Begin VB.Menu mnuPlGethsh 
            Caption         =   "get hash from library"
         End
         Begin VB.Menu mnuPlFindmdb 
            Caption         =   "find in library"
         End
      End
   End
   Begin VB.Menu mnuDisp 
      Caption         =   "display"
      Visible         =   0   'False
      Begin VB.Menu mnuDispFullI 
         Caption         =   "[]"
         Index           =   0
      End
   End
   Begin VB.Menu mnuTray 
      Caption         =   "tray"
      Visible         =   0   'False
      Begin VB.Menu mnuTrayI 
         Caption         =   "stop"
         Index           =   0
      End
      Begin VB.Menu mnuTrayI 
         Caption         =   "play / pause"
         Index           =   1
      End
      Begin VB.Menu mnuTrayI 
         Caption         =   "next"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Compare Binary
Option Base 0

Private m_lCurrentPl As Long, _
    pl_lAutoSaveCntDn As Long, pl_lAutoSaveTime As Long, _
    ps_sMinusOneFilePath As String, _
    pb_Speed As Double, pb_RetryCount As Long

'media scanning:
Private bAbort As Boolean, bAborted As Boolean

Private m_TrackCurrentItem As Boolean
Public m_bInTray As Boolean, m_WsBeforeMin As Long
Public m_bFullScreen As Boolean, m_bMiniVid As Boolean

Private cHk(0 To 4) As clsHotKey
Private SystemTray As New clsSysTray

Sub pref_StartAutoSave()
If pref_CanSave Then
    tmrPrefSave.Enabled = True
End If
End Sub

Sub pref_Read()
On Error GoTo pref_Read_err

Dim e As String, a, i As Long

'settings============================================================
e = "loading config"

m_TrackCurrentItem = False
If GetFromIniEx("pref", "trackcurrentitem", "1", file_INI) = "1" Then
    mnuJumpAuto_Click
End If
mnuModeI_Click (GetFromIniEx("pref", "playmode", "0", file_INI))
mnuModeListsI_Click (GetFromIniEx("pref", "listmode", "2", file_INI))
cmdMnu(2).Redraw

cbbTabPosition.ListIndex = GetFromIniEx("pref", "tspos", "0", file_INI)

'note: defaults to internal style.
a = GetFromIniEx("pref", "gdimode", "i", file_INI)
If a = "b" Then
    cbbOptDrawmode.ListIndex = 1
Else
    cbbOptDrawmode.ListIndex = 0
End If
cbbOptColourscheme.ListIndex = 0

i = Val(GetFromIniEx("lib", "viewpos", "0", file_INI))
    mdbl_SetScroll i

chkOptCuebtm.Value = GetFromIniEx("pref", "cuebtm", "1", file_INI)
chkOptDisplayRepDashNewLine.Value = GetFromIniEx("pref", "displayrepdashnewline", "1", file_INI)
chkOptDisplayMinivid.Value = GetFromIniEx("pref", "displayminivid", "1", file_INI)
chkOptDisplayMintray.Value = GetFromIniEx("pref", "mintotray", "0", file_INI)

chkOptLibNum.Value = GetFromIniEx("lib", "showindex", "1", file_INI)
For i = 0 To chkOptLibcols.count - 1
    chkOptLibcols(i).Value = GetFromIniEx("lib", "viewcol" & i, chkOptLibcols(i).Value, file_INI)
Next i

a = GetFromIniEx("lib", "fnamehide_cnt", 0, file_INI)
If a > 0 Then
    For i = 0 To a - 1
        a = GetFromIniEx("lib", "fnamehide_i" & i, "", file_INI)
        If a <> "" Then lstPrefDirHide.AddItem a
    Next i
End If
mdb_List.lTrunk = Val(GetFromIniEx("lib", "fnametrnc", "1", file_INI))
    sldOptDirLvls.Value = mdb_List.lTrunk
mdb_List.bTrunkReverse = IIf(GetFromIniEx("lib", "fnametrncrev", "1", file_INI) = "1", True, False)
    chkOptDirLvlsReverse.Value = IIf(mdb_List.bTrunkReverse, 1, 0)
chkOptHideExt.Value = IIf(GetFromIniEx("lib", "hidefileext", "0", file_INI) = "1", 1, 0)

txtOptFileext.Text = GetFromIniEx("lib", "fileextlist", file_ext_list_def, file_INI)
a = GetFromIniEx("lib", "watch_cnt", 0, file_INI)
If a > 0 Then
    For i = 0 To a - 1
        a = GetFromIniEx("lib", "watch_i" & i, "", file_INI)
        If a <> "" Then lstLibWatch.AddItem a
    Next i
End If
lblPrefLibFldCnt.Caption = lstLibWatch.ListCount & lblPrefLibFldCnt.Tag

'note: autosave defaults to 30 min.
pl_lAutoSaveTime = Val(GetFromIniEx("pl", "autosavetime", "30", file_INI))
    sldOptPlAutosave.Value = pl_lAutoSaveTime

cbbOptErrNoplay.ListIndex = GetFromIniEx("errorproc", "noplay", "0", file_INI)

Move _
    GetFromIniEx("form", "normalleft", 0, file_INI), _
    GetFromIniEx("form", "normaltop", 0, file_INI), _
    GetFromIniEx("form", "normalwidth", 9000, file_INI), _
    GetFromIniEx("form", "normalheight", 6000, file_INI)
a = GetFromIniEx("form", "showcmd", "1", file_INI)
Select Case a
    Case "1": WindowState = vbNormal
    Case "2": WindowState = vbMinimized
    Case "3": WindowState = vbMaximized
End Select

'hotkeys:
For i = 0 To UBound(cHk)
    txtOptHk(i).Tag = GetFromIniEx("pref", "hk" & Trim$(Str$(i)), "", file_INI)
    cmdOptHk_Click (i)
    If cHk(i).HasData And Not cHk(i).IsSet Then
        AddToLog txtLog, "unable to register hotkey " & cHk(i).GetKeyDes & "."
    End If
Next i

Exit Sub
pref_Read_err:
Main_Err "pref_Read_err."
err.Clear
End Sub

Sub pref_Write()
On Error GoTo pref_Write_err

Dim i As Long

If Not pref_CanSave Then
    MsgBox "settings have not been saved, as they were never loaded in the first place."
    Exit Sub
End If

WriteToIni "pref", "trackcurrentitem", IIf(m_TrackCurrentItem, "1", "0"), file_INI
For i = 0 To mnuModeI.count - 1
    If mnuModeI(i).Checked Then
        WriteToIni "pref", "playmode", (i), file_INI
        Exit For
    End If
Next i
For i = 0 To mnuModeListsI.count - 1
    If mnuModeListsI(i).Checked Then
        WriteToIni "pref", "listmode", (i), file_INI
        Exit For
    End If
Next i

WriteToIni "pref", "tspos", tsMain.Placement, file_INI
WriteToIni "pref", "gdimode", IIf(cbbOptDrawmode.ListIndex = 1, "b", "i"), file_INI

WriteToIni "view", "activetab", tsMain.SelectedItem.Key, file_INI

WriteToIni "pref", "cuebtm", Trim$(Str$(chkOptCuebtm.Value)), file_INI
WriteToIni "pref", "displayrepdashnewline", Trim$(Str$(chkOptDisplayRepDashNewLine.Value)), file_INI
WriteToIni "pref", "displayminivid", Trim$(Str$(chkOptDisplayMinivid.Value)), file_INI
WriteToIni "pref", "mintotray", Trim$(Str$(chkOptDisplayMintray.Value)), file_INI

WriteToIni "lib", "showindex", Trim$(Str$(chkOptLibNum.Value)), file_INI
For i = 0 To chkOptLibcols.count - 1
    WriteToIni "lib", "viewcol" & i, Trim$(Str$(chkOptLibcols(i).Value)), file_INI
Next i

WriteToIni "lib", "fnamehide_cnt", Trim$(Str$(lstPrefDirHide.ListCount)), file_INI
If lstPrefDirHide.ListCount > 0 Then
    For i = 0 To lstPrefDirHide.ListCount - 1
        WriteToIni "lib", "fnamehide_i" & i, lstPrefDirHide.List(i), file_INI
    Next i
End If
WriteToIni "lib", "fnametrnc", Trim$(Str$(mdb_List.lTrunk)), file_INI
WriteToIni "lib", "fnametrncrev", IIf(mdb_List.bTrunkReverse, "1", "0"), file_INI
WriteToIni "lib", "hidefileext", Trim$(Str$(chkOptHideExt.Value)), file_INI

WriteToIni "lib", "viewpos", vsbMdbLst.Value, file_INI

WriteToIni "lib", "fileextlist", file_ext_list, file_INI
WriteToIni "lib", "watch_cnt", Trim$(Str$(lstLibWatch.ListCount)), file_INI
If lstLibWatch.ListCount > 0 Then
    For i = 0 To lstLibWatch.ListCount - 1
        WriteToIni "lib", "watch_i" & i, lstLibWatch.List(i), file_INI
    Next i
End If

WriteToIni "pl", "autosavetime", Trim$(Str$(pl_lAutoSaveTime / 60)), file_INI

WriteToIni "errorproc", "noplay", cbbOptErrNoplay.ListIndex, file_INI

Dim wp As WINDOWPLACEMENT
GetWindowPlacement hWND, wp
WriteToIni "form", "showcmd", Trim$(Str$(wp.showCmd)), file_INI
WriteToIni "form", "normalleft", Trim$(Str$(wp.rcNormalPosition.Left * Screen.TwipsPerPixelX)), file_INI
WriteToIni "form", "normaltop", Trim$(Str$(wp.rcNormalPosition.Top * Screen.TwipsPerPixelY)), file_INI
WriteToIni "form", "normalwidth", Trim$(Str$((wp.rcNormalPosition.right - wp.rcNormalPosition.Left) * Screen.TwipsPerPixelX)), file_INI
WriteToIni "form", "normalheight", Trim$(Str$((wp.rcNormalPosition.bottom - wp.rcNormalPosition.Top) * Screen.TwipsPerPixelY)), file_INI

For i = 0 To UBound(cHk)
    WriteToIni "pref", "hk" & Trim$(Str$(i)), cHk(i).GetKeyDataString, file_INI
Next i

WriteToIni "minivid", "active", IIf(m_bMiniVid, "1", "0"), file_INI

'clean up
tmrPrefSave.Enabled = False
pref_AutoSaveCounter = 0
cmdOptSave.Caption = cmdOptSave.Tag

Exit Sub
pref_Write_err:
Main_Err "pref_Write_err."
err.Clear
End Sub

Sub mdbl_Requery(Optional bRedraw As Boolean = True)
Dim sSql As String, sOrderBy As String, lTimer As Long

lTimer = GetTickCount

'STEP 1 - build query.

'sSql = "select sfile, dadded, lstartcnt, lendcnt, dlastplay, lmd5, lduration, benabled, bmissing from tbl_mediafiles order by "
sSql = "SELECT sfile, dadded, lstartcnt, lendcnt, dlastplay, " & _
    "lmd5, lduration, benabled, bmissing FROM tbl_mediafiles"

'WHERE...

If Not mnuLibOrderShowmis.Checked Then
    sSql = sSql & " WHERE (bmissing=0 OR bmissing is NULL)"
End If

'ORDER BY...

If mnuLibOrderDirection(1).Checked Then
    sOrderBy = " DESC"
Else 'assume index=0
    sOrderBy = " ASC"
End If

sSql = sSql & " ORDER BY "

If mnuLibOrderI(1).Checked Then
    sSql = sSql & "lstartcnt " & sOrderBy & ", lendcnt " & sOrderBy & _
        ", sfile COLLATE NOCASE " & sOrderBy
ElseIf mnuLibOrderI(2).Checked Then
    sSql = sSql & "lendcnt " & sOrderBy & ",lstartcnt " & sOrderBy & _
        ",sfile COLLATE NOCASE " & sOrderBy
ElseIf mnuLibOrderI(3).Checked Then
    sSql = sSql & "dadded " & sOrderBy & ",sfile COLLATE NOCASE " & sOrderBy
ElseIf mnuLibOrderI(4).Checked Then
    sSql = sSql & "dlastplay " & sOrderBy & ",dadded " & sOrderBy & _
        ",sfile COLLATE NOCASE " & sOrderBy
ElseIf mnuLibOrderI(5).Checked Then
    sSql = sSql & "lduration " & sOrderBy & ",sfile COLLATE NOCASE " & sOrderBy
Else 'assume index=0
    sSql = sSql & "sfile COLLATE NOCASE " & sOrderBy
End If

sSql = sSql & ";"

'STEP 2 - run query.

'Debug.Print sSQL
If mdb_QueryToMdb(sSql, mdb_List) Then
    'all is ok
Else
    Debug.Print "mdbl_Requery/mdb_QueryToMdb error."
End If

'STEP 3 - output and clean up.

lblDbStat.Caption = "query returned " & mdb_List.lCnt & " items in " & (GetTickCount - lTimer) / 1000 & " seconds."

mdbl_Rebuild bRedraw
End Sub

Sub mdbl_Rebuild(Optional bRedraw As Boolean = True, Optional bForceRecalc As Boolean = False)
Dim lTotalH As Long

If fraLib.Visible = False And mdbl_ItmH > 0 And bForceRecalc = False Then Exit Sub

mdbl_ItmH = picMdbLst.TextHeight("WA|MJT") * 1.1 'some tall chrs
lTotalH = mdb_List.lCnt * mdbl_ItmH

If lTotalH > picMdbLst.ScaleHeight Then
    vsbMdbLst.Min = 0
    vsbMdbLst.Max = mdb_List.lCnt - (picMdbLst.ScaleHeight / mdbl_ItmH) + 1
    vsbMdbLst.LargeChange = picMdbLst.ScaleHeight / mdbl_ItmH
    vsbMdbLst.SmallChange = 1
    If vsbMdbLst.Enabled <> True Then vsbMdbLst.Enabled = True
Else
    If vsbMdbLst.Value <> 0 Then vsbMdbLst.Value = 0
    If vsbMdbLst.Enabled <> False Then vsbMdbLst.Enabled = False
End If

If bRedraw Then mdbl_Redraw
End Sub

Sub mdbl_Redraw()
Dim i As Long, x As Long, t As Long, b As Boolean, _
    rcMain As RECT, _
    rc As RECT, rc2 As RECT, _
    s As String, _
    lColL(0 To 5) As Long, lFlg As Long, _
    lDrawTo As Long

'Dim lTimer As Long
'lTimer = GetTickCount

'ensure the buffer is the right size
m_ListDc.Width = picMdbLst.ScaleWidth
m_ListDc.Height = picMdbLst.ScaleHeight

'the rc of the draw area
rcMain.Left = 0
rcMain.Top = 0
rcMain.right = picMdbLst.ScaleWidth
rcMain.bottom = picMdbLst.ScaleHeight

'cls the buffer
FillRect m_ListDc.hDC, rcMain, gdi_Main_Brush(0)

'is there anything to draw?====================================================
If mdb_List.lCnt < 1 Then
    s = "the media library is empty." & vbNewLine & vbNewLine & _
        "see 'media library' > 'library configuration' to add media."
    SelectObject m_ListDc.hDC, gdi_Main_hFontNormal
    SetTextColor m_ListDc.hDC, GetSysColor(COLOR_BTNTEXT)
    DrawText m_ListDc.hDC, s, Len(s), rcMain, DT_CENTER
    GoTo NoItems
End If

'check the position of the focus box===========================================
'removed and moved to "got focus" event.
'If mdb_List.lIndex < vsbMdbLst.Value Then
'    mdb_List.lIndex = vsbMdbLst.Value
'ElseIf mdb_List.lIndex > vsbMdbLst.Value + Fix(picMdbLst.ScaleHeight / mdbl_ItmH) - 1 Then
'    mdb_List.lIndex = vsbMdbLst.Value + Fix(picMdbLst.ScaleHeight / mdbl_ItmH) - 1
'End If

'calculate columns=============================================================
SelectObject m_ListDc.hDC, gdi_Main_hFontNormal

'duration
If chkOptLibcols(4).Value = 1 Then
    s = "000:00"
    DrawText m_ListDc.hDC, s, Len(s), rc, dt_left + DT_CALCRECT
    lColL(5) = picMdbLst.ScaleWidth - rc.right * dListCellPaddingH
Else
    lColL(5) = picMdbLst.ScaleWidth
End If

'hash code
If chkOptLibcols(3).Value = 1 Then
    s = "DDDDDDDD"
    DrawText m_ListDc.hDC, s, Len(s), rc, dt_left + DT_CALCRECT
    lColL(4) = lColL(5) - rc.right * dListCellPaddingH
Else
    lColL(4) = lColL(5)
End If

'date last played
If chkOptLibcols(2).Value = 1 Then
    s = Format(Now, sDateFormatString)
    DrawText m_ListDc.hDC, s, Len(s), rc, dt_left + DT_CALCRECT
    lColL(3) = lColL(4) - rc.right * dListCellPaddingH
Else
    lColL(3) = lColL(4)
End If

'date added
If chkOptLibcols(1).Value = 1 Then
    s = Format(Now, sDateFormatString)
    DrawText m_ListDc.hDC, s, Len(s), rc, dt_left + DT_CALCRECT
    lColL(2) = lColL(3) - rc.right * dListCellPaddingH
Else
    lColL(2) = lColL(3)
End If

'play counts
If chkOptLibcols(0).Value = 1 Then
    s = "000/000"
    DrawText m_ListDc.hDC, s, Len(s), rc, dt_left + DT_CALCRECT
    lColL(1) = lColL(2) - rc.right * dListCellPaddingH
Else
    lColL(1) = lColL(2)
End If

lColL(0) = 1 'leave a 1 px gap at the left side.

'Debug.Print lColL(0) & "," & lColL(1) & "," & lColL(2)

'draw the items================================================================
rc.Left = 0
rc.right = picMdbLst.ScaleWidth

SetBkMode m_ListDc.hDC, TRANSPARENT

For i = vsbMdbLst.Value To mdb_List.lCnt - 1
    'item area
    rc.Top = (i - vsbMdbLst.Value) * mdbl_ItmH
    rc.bottom = rc.Top + mdbl_ItmH
    
    If mdb_List.Items(i).bSel Then
        FillRect m_ListDc.hDC, rc, gdi_Main_Brush(1)
        SetTextColor m_ListDc.hDC, GetSysColor(COLOR_HIGHLIGHTTEXT)
    Else
        SetTextColor m_ListDc.hDC, GetSysColor(COLOR_BTNTEXT)
    End If
    
    'active item / item (dis/en)abled?
    If pb_GetPlayState > 0 And cMedia.FileName = mdb_List.Items(i).sFile Then
        If mdb_List.Items(i).bEnabled Then
            SelectObject m_ListDc.hDC, gdi_Main_hFontSel
        Else
            SelectObject m_ListDc.hDC, gdi_Main_hFontSelStrike
        End If
    Else
        If mdb_List.Items(i).bMissing Then
            SelectObject m_ListDc.hDC, gdi_Main_hFontNormalStrikeItalic
        ElseIf mdb_List.Items(i).bEnabled Then
            SelectObject m_ListDc.hDC, gdi_Main_hFontNormal
        Else
            SelectObject m_ListDc.hDC, gdi_Main_hFontNormalStrike
        End If
    End If
    
    'draw text
    rc2 = rc
    For x = 0 To UBound(lColL)
        Select Case x
            Case 1: b = IIf(chkOptLibcols(0).Value = 1, True, False)
            Case 2: b = IIf(chkOptLibcols(1).Value = 1, True, False)
            Case 3: b = IIf(chkOptLibcols(2).Value = 1, True, False)
            Case 4: b = IIf(chkOptLibcols(3).Value = 1, True, False)
            Case 5: b = IIf(chkOptLibcols(4).Value = 1, True, False)
            Case Else: b = True
        End Select
        If Not b Then GoTo NextX
        
        rc2.Left = lColL(x)
        If x = UBound(lColL) Then
            rc2.right = rc.right
        Else
            rc2.right = lColL(x + 1)
        End If
        
        Select Case x
            Case 0
                s = IIf(chkOptLibNum.Value = "1", i, "") & " " & gen_GetShowName(mdb_List.Items(i).sFile)
                lFlg = dt_left
            
            Case 1
                If mdb_List.Items(i).lStartCnt > 0 Then
                    s = mdb_List.Items(i).lStartCnt & "/" & mdb_List.Items(i).lEndCnt
                    lFlg = DT_CENTER
                Else
                    s = ""
                End If
            
            Case 2
                s = IIf(mdb_List.Items(i).dLastPlay < 1, "", Format(mdb_List.Items(i).dLastPlay, sDateFormatString))
                lFlg = dt_left
            
            Case 3
                s = IIf(mdb_List.Items(i).dAdded < 1, "", Format(mdb_List.Items(i).dAdded, sDateFormatString))
                lFlg = dt_left
            
            Case 4 'has code
                If mdb_List.Items(i).lMD5 <> 0 Then
                    s = Hex(mdb_List.Items(i).lMD5)
                    lFlg = DT_CENTER
                Else
                    s = ""
                End If
            
            Case 5 'duration
                If mdb_List.Items(i).lDuration <> 0 Then
                    s = ConvertSecToMin(mdb_List.Items(i).lDuration)
                    lFlg = DT_RIGHT
                Else
                    s = ""
                End If
            
        End Select
        
        If Len(s) > 0 Then
            DrawText m_ListDc.hDC, s, Len(s), rc2, lFlg + DT_NOPREFIX + DT_VCENTER
        End If
NextX:
    Next x
    
    'list index box?
    If mdb_List.lIndex = i And picMdbLst.Tag = "1" Then
        FrameRect m_ListDc.hDC, rc, gdi_Main_Brush(2)
    End If
    
    'at end of display area?
    If rc.bottom > m_ListDc.Height Then Exit For
    
NextI:
Next i

'clean up======================================================================

'Debug.Print "redraw time = " & (GetTickCount - lTimer) / 1000

NoItems:
picMdbLst_Paint
End Sub

Sub mdbl_SetScroll(ByVal v As Long)
If v >= vsbMdbLst.Max Then
    v = vsbMdbLst.Max
ElseIf v <= vsbMdbLst.Min Then
    v = vsbMdbLst.Min
End If
vsbMdbLst.Value = v
End Sub

Function list_GetCurrentIndex(Optional lDef As Long = -1) As Long
Dim b As Boolean, i As Long

Select Case pb_lItemSource
    Case 0
        b = False
        For i = 0 To mdb_List.lCnt - 1
            If cMedia.FileName = mdb_List.Items(i).sFile Then
                b = True
                Exit For
            End If
        Next i
        
        If b = False Then i = lDef
    
    Case Is > 0
        i = mdb_PL(pb_lItemSource - 1).lCurrent
        If i < 0 Then i = lDef
    
    Case Else
        i = -1
    
End Select

list_GetCurrentIndex = i
End Function

Sub list_JumpToCurrent(Optional bForceRedraw As Boolean = False, Optional bSetTab As Boolean = True)
If pb_CheckCurrentIndex = False Then Exit Sub

Dim i As Long

If mdbl_ItmH <= 0 Then mdbl_Rebuild False

i = list_GetCurrentIndex
If i < 0 Then Exit Sub

list_JumpTo pb_lItemSource, i, bSetTab

If bForceRedraw = True Then mdbl_Redraw   'pb_lItemSource
End Sub

Sub list_JumpTo(lSource As Long, lItem As Long, Optional bSetTab As Boolean = True)
Dim i As Long

If bSetTab Then
    Select Case lSource
        Case 0: tsMain.Tabs("lib").Selected = True
        Case Is > 0:
            If mdb_PL(lSource - 1).bArc Then
                tsMain.Tabs("arc").Selected = True
                For i = 0 To lstPlArc.ListCount - 1
                    If lstPlArc.ItemData(i) = lSource - 1 Then
                        lstPlArc.ListIndex = i
                        Exit For
                    End If
                Next i
            Else
                tsMain.Tabs("pl" & lSource - 1).Selected = True
            End If
    End Select
End If

If lSource = 0 Then
    If lItem < vsbMdbLst.Value Then
        mdbl_SetScroll lItem - 1
    ElseIf lItem > vsbMdbLst.Value + Fix(picMdbLst.ScaleHeight / mdbl_ItmH) - 1 Then
        mdbl_SetScroll lItem - Fix(picMdbLst.ScaleHeight / mdbl_ItmH) + 2
    End If
Else
    If lItem < vsbPl.Value Then
        pl_SetScroll vsbPl, lItem - 1
    ElseIf lItem > vsbPl.Value + Fix(picPl.ScaleHeight / mdbl_ItmH) - 1 Then
        pl_SetScroll vsbPl, lItem - Fix(picPl.ScaleHeight / mdbl_ItmH) + 2
    End If
End If
End Sub

Sub mdb_AddDir(f As String, Optional bClearHashOnRefind As Boolean = False)
On Error GoTo ErrorProc

Dim cFlds As Collection, cFiles As Collection, _
    i As Long, n As Long, e As String

e = "checking path exists"
If Not fsoMain.DriveExists(f) And Not fsoMain.FolderExists(f) Then
    AddToLog txtLog, "directory not found: '" & f & "'"
    Exit Sub
End If

e = "set busy"
'set busy
bAbort = False
bAborted = False
fraNotice(0).Visible = True
Form_Resize

'AddToLog txtLog, "processing folder '" & txtAddDir.Text & "'..."
lblLibAddStat.Caption = "adding media to library: processing folder '" & f & "'..."

'AddToLog txtLog, "building sub folder list..."
lblLibAddStat.Caption = "adding media to library: building folder list..."

e = "building dir collection"
Set cFlds = New Collection
BuildDirCollection f, cFlds  ', True,  txtLog
'AddToLog txtLog, " " & cFlds.count & " found.", False

DoEvents
If bAbort = True Then GoTo AbortSub

e = "building file collection"
'AddToLog txtLog, "building file list..."
lblLibAddStat.Caption = "adding media to library: building file list..."
Set cFiles = New Collection
For i = 1 To cFlds.count
    AddFilesToCollection cFlds(i), cFiles
Next i
'AddToLog txtLog, " " & cFiles.count & " found.", False

DoEvents
If bAbort = True Then GoTo AbortSub

'AddToLog txtLog, "adding media files to db..."
lblLibAddStat.Caption = "adding media to library: adding media files to db..."

e = "adding items"
n = 0
For i = 1 To cFiles.count
    e = "IsMediaFile"
    If IsMediaFile(cFiles(i)) Then
        'If IsUnicode(cFiles(i)) Then Debug.Print "unicode file listed: " & cFiles(i)
        
        lblLibAddStat.Caption = "adding media to library: handeling file " & i & " of " & cFiles.count & "."
        
        e = "mdb_AddFile(" & cFiles(i) & ")"
        If mdb_AddFile(cFiles(i), bClearHashOnRefind) Then
            'AddToLog txtLog, "file added: " & cFiles(i)
            n = n + 1
        End If
    End If
    
    DoEvents
    If bAbort = True Then GoTo AbortSub
Next i
'AddToLog txtLog, n & " files added to db."
lblLibAddStat.Caption = "adding media to library: " & n & " files added to db."
'AddToLog txtLog, n_uni & " unicode filenames disguarded."

AbortSub:

'refresh
e = "mdbl_Requery"
If n > 0 Then mdbl_Requery False

CleanUp:

'sub ready
fraNotice(0).Visible = False
Form_Resize

bAborted = True

Exit Sub
ErrorProc:
Main_Err "error adding directory '" & f & "'.  mdb_AddDir/" & e
err.Clear
GoTo CleanUp
End Sub

Function mdbl_FindIndexFromFile(ByVal f As String) As Long
Dim i As Long, b As Boolean

If mdb_List.lCnt < 1 Then Exit Function

f = LCase$(f)
b = False

For i = 0 To mdb_List.lCnt - 1
    If LCase$(mdb_List.Items(i).sFile) = f Then
        b = True
        Exit For
    End If
Next i

If b Then
    mdbl_FindIndexFromFile = i
Else
    mdbl_FindIndexFromFile = -1
End If
End Function

Function mdbl_FindIndexFromHash(ByVal lCRC As Long) As Long
Dim i As Long, b As Boolean

If mdb_List.lCnt < 1 Then Exit Function

b = False

For i = 0 To mdb_List.lCnt - 1
    If mdb_List.Items(i).lMD5 = lCRC Then
        b = True
        Exit For
    End If
Next i

If b Then
    mdbl_FindIndexFromHash = i
Else
    mdbl_FindIndexFromHash = -1
End If
End Function

Sub mdbl_RebuildMenu()
On Error GoTo mdbl_RebuildMenu_err

Dim i As Long, x As Long

If mnuLibAddplI.count > 1 Then
    For i = mnuLibAddplI.count - 1 To 1 Step -1
        Unload mnuLibAddplI(i)
    Next i
End If

For i = 0 To UBound(mdb_PL)
    If Not mdb_PL(i).bArc Then
        If i > 0 Then
            x = mnuLibAddplI.count
            Load mnuLibAddplI(x)
        Else
            x = 0
        End If
        mnuLibAddplI(x).Caption = mdb_PL(i).sName
        mnuLibAddplI(x).Tag = Trim$(Str$(i))
    End If
Next i

Exit Sub
mdbl_RebuildMenu_err:
Main_Err "mdbl_RebuildMenu_err."
err.Clear
End Sub

Sub pll_RebuildTabs(Optional bResize As Boolean = False)
Dim i As Long, tPl As MSComctlLib.Tab

With tsMain.Tabs
    'remove all existing tabs
    For i = .count To 1 Step -1
        If Mid$(.Item(i).Key, 1, 2) = "pl" Then
            .Remove i
        End If
    Next i
    
    'add tabs back
    For i = 0 To UBound(mdb_PL)
        If Not mdb_PL(i).bArc Then
            Set tPl = .Add(.count, "pl" & i, mdb_PL(i).sName) '"pl"
        End If
    Next i
End With

If bResize Then Form_Resize

pll_RebuildEnab
End Sub

Sub pll_RebuildAndPrepArc(Optional lPlSel As Long = -1)
Dim i As Long

'build list
lstPlArc.Clear
For i = 0 To UBound(mdb_PL)
    If mdb_PL(i).bArc Then
        lstPlArc.AddItem mdb_PL(i).sName
        lstPlArc.ItemData(lstPlArc.NewIndex) = i
    End If
Next i

If lPlSel = -1 Then
    'set no playlist visible
    m_lCurrentPl = -1
    pll_RedrawCurrent
Else
    'set pl with matching index
    If lstPlArc.ListCount > 0 Then
        For i = 0 To lstPlArc.ListCount - 1
            If lstPlArc.ItemData(i) = lPlSel Then
                lstPlArc.ListIndex = i
                Exit For
            End If
        Next i
    End If
End If
End Sub

Sub pll_RebuildEnab()
Dim i As Long
For i = 0 To UBound(mdb_PL)
    If Not mdb_PL(i).bArc Then
        If mdb_PL(i).bEnab Then
            tsMain.Tabs("pl" & i).Caption = mdb_PL(i).sName
        Else
            tsMain.Tabs("pl" & i).Caption = "[" & mdb_PL(i).sName & "]"
        End If
    End If
Next i
End Sub

Sub pll_RebuildMenu()
Dim i As Long, x As Long

If mnuPlAddplI.count > 1 Then
    For i = mnuPlAddplI.count - 1 To 1 Step -1
        Unload mnuPlAddplI(i)
    Next i
End If

For i = 0 To UBound(mdb_PL)
    If Not mdb_PL(i).bArc Then
        If i > 0 Then
            x = mnuPlAddplI.count
            Load mnuPlAddplI(x)
        Else
            x = 0
        End If
        mnuPlAddplI(x).Caption = mdb_PL(i).sName
        mnuPlAddplI(x).Tag = Trim$(Str$(i))
    End If
Next i

mnuPlEnab.Checked = mdb_PL(m_lCurrentPl).bEnab
mnuPlArc.Checked = mdb_PL(m_lCurrentPl).bArc
End Sub

Function pll_New(Optional bRebuild As Boolean = True, Optional bResize As Boolean = False) As Long
ReDim Preserve mdb_PL(LBound(mdb_PL) To UBound(mdb_PL) + 1)
pl_SetAsNew mdb_PL(UBound(mdb_PL)), "list" & UBound(mdb_PL)
pll_New = UBound(mdb_PL)
If bRebuild Then pll_RebuildTabs bResize
pll_UpdateStatus True, True 'pl_SetAutoSave True
End Function

Sub pll_UpdateStatus(Optional bSetAutoSave As Boolean = False, Optional bNewAutoSave As Boolean)
Dim sStat As String ', b As Boolean

If m_lCurrentPl < 0 Then
    sStat = ""
Else
    sStat = mdb_PL(m_lCurrentPl).lCnt & " items totaling " & _
        IIf(mdb_PL(m_lCurrentPl).bTotalDurationComplete, "", "more than ") & _
        ConvertSecToHours(mdb_PL(m_lCurrentPl).lTotalDuration) & ".  "
End If

If pl_lAutoSaveTime < 1 Then
    sStat = sStat & "autosave disabled."
    GoTo LeaveSub
ElseIf bSetAutoSave Then
    If tmrPlAutoSave.Enabled <> bNewAutoSave Then
        tmrPlAutoSave.Enabled = bNewAutoSave
        If bNewAutoSave Then
            pl_lAutoSaveCntDn = pl_lAutoSaveTime
            tmrPlAutoSave_Timer
        End If
    End If
End If

If tmrPlAutoSave.Enabled Then
    sStat = sStat & "auto save in " & _
        Fix(pl_lAutoSaveCntDn / 60) & ":" & _
        right$("00" & (pl_lAutoSaveCntDn Mod 60), 2) & _
        ".  click here to save now."
    Else
        sStat = sStat & "no changes to save."
End If

LeaveSub:
lblPlStat.Caption = sStat
lblPlStat.Refresh
End Sub

Sub pll_WriteAllLists()
Dim i As Long, sBuf As cStringBuilder
Set sBuf = New cStringBuilder

For i = 0 To UBound(mdb_PL)
    If Len(mdb_PL(i).sFilePath) <= 0 Then mdb_PL(i).sFilePath = GetFreePlaylistFileName
    pl_Write mdb_PL(i), mdb_PL(i).sFilePath
    
    sBuf.Append FileNameFromPath(mdb_PL(i).sFilePath) & vbNewLine
Next i

UnicodeFile_Write_FSO file_PlayListIndex, sBuf.ToString

pll_UpdateStatus True, False 'pl_SetAutoSave False
End Sub

Sub pll_ReadAllLists()
Dim i As Long, cFlds As Collection, cFiles As Collection, _
     lList As Long, arr() As String

Set cFiles = New Collection

'look for index list for order:
If fsoMain.FileExists(file_PlayListIndex) Then
    arr = Split(UnicodeFile_Read_FSO(file_PlayListIndex), vbNewLine)
    For i = 0 To UBound(arr)
        arr(i) = folder_Playlists & Trim$(arr(i))
        If Len(arr(i)) > 0 Then
            If fsoMain.FileExists(arr(i)) Then
                cFiles.Add arr(i)
            End If
        End If
    Next i
End If

'old method: grab all playlist files in the folder.
'Dim cDir As Collection
'Set cDir = New Collection
'BuildDirCollection folder_Playlists, cDir
'For i = 1 To cDir.count
'    AddFilesToCollection cDir(i), cFiles
'Next i

If cFiles.count < 1 Then Exit Sub

For i = 1 To cFiles.count
    If right$(cFiles(i), Len(file_ext_playlist)) = file_ext_playlist Then
        'If mdb_PL(UBound(mdb_PL)).lCnt < 1 Then
        If UBound(mdb_PL) <= 0 And mdb_PL(0).lCnt < 1 Then
            lList = 0 'UBound(mdb_PL)
        Else
            lList = pll_New(False)
        End If
        pl_Read mdb_PL(lList), cFiles(i)
    End If
Next i

pll_RebuildTabs
pll_UpdateStatus True, False ' pl_SetAutoSave False
End Sub

Private Sub pll_Draw(List As typPl, lFrom As Long, _
    lHdc As Long, lW As Long, lH As Long, _
    Optional bFocused As Boolean = False, _
    Optional sEmptyText As String = "this playlist is empty.")

Dim sCols As String
sCols = ""

If chkOptLibcols(4).Value = 1 Then 'duration
    sCols = sCols & "1"
Else
    sCols = sCols & "0"
End If

If chkOptLibcols(3).Value = 1 Then 'hash
    sCols = sCols & "1"
Else
    sCols = sCols & "0"
End If

If chkOptLibcols(1).Value = 1 Then 'date last played
    sCols = sCols & "1"
Else
    sCols = sCols & "0"
End If

If chkOptLibcols(0).Value = 1 Then 'play count
    sCols = sCols & "1"
Else
    sCols = sCols & "0"
End If

pl_Draw List, lFrom, lHdc, lW, lH, bFocused, sEmptyText, sCols, _
    (chkOptHideExt.Value = 1), mdb_List.lTrunk, mdb_List.bTrunkReverse, _
    IIf(chkOptLibNum.Value = "1", True, False)
End Sub

Sub pll_RedrawCurrent()
On Error GoTo pl_RedrawCurrent_err

If m_lCurrentPl < 0 Then
    pl_ListDc.Width = picPl.ScaleWidth
    pl_ListDc.Height = picPl.ScaleHeight
    pl_Blank pl_ListDc.hDC, picPl.ScaleWidth, picPl.ScaleHeight
    picPl_Paint
    pl_DebuildScroll vsbPl
    Exit Sub
End If

vsbPl.Tag = "1"
pl_RebuildScroll mdb_PL(m_lCurrentPl), picPl.ScaleHeight, vsbPl
If vsbPl.Enabled Then
    If mdb_PL(m_lCurrentPl).lScroll > vsbPl.Max Then
        mdb_PL(m_lCurrentPl).lScroll = vsbPl.Max
        vsbPl.Value = vsbPl.Max
    Else
        vsbPl.Value = mdb_PL(m_lCurrentPl).lScroll
    End If
Else
    mdb_PL(m_lCurrentPl).lScroll = 0
End If
vsbPl.Tag = ""
'Debug.Print Format(Now, "hh:mm:ss") & " drawing pl " & m_lCurrentPl

pl_ListDc.Width = picPl.ScaleWidth
pl_ListDc.Height = picPl.ScaleHeight

pll_Draw mdb_PL(m_lCurrentPl), mdb_PL(m_lCurrentPl).lScroll, pl_ListDc.hDC, _
    picPl.ScaleWidth, picPl.ScaleHeight, (picPl.Tag = "1")

picPl_Paint

Exit Sub
pl_RedrawCurrent_err:
Main_Err "pl_RedrawCurrent_err."
err.Clear
End Sub

Private Sub pll_SwapLists(lFirst As Long, lSecond As Long)
Dim iFirst As typPl, iSecond As typPl

iFirst = mdb_PL(lFirst)
iSecond = mdb_PL(lSecond)

mdb_PL(lFirst) = iSecond
mdb_PL(lSecond) = iFirst

'+1 because cue treats 0=mdb and pl0 as 1
If cue_ListInCue(lFirst + 1) Or cue_ListInCue(lSecond + 1) Then _
    cue_ListsSwapped lFirst + 1, lSecond + 1

If pb_lItemSource = lFirst + 1 Then
    pb_lItemSource = lSecond + 1
ElseIf pb_lItemSource = lSecond + 1 Then
    pb_lItemSource = lFirst + 1
End If
End Sub

Private Sub pll_RemoveList(lList As Long)
On Error GoTo pll_RemoveList_err

Dim f As String, i As Long
f = mdb_PL(lList).sFilePath

'first take care of the in-app list:
If lList < UBound(mdb_PL) Then
    For i = lList To UBound(mdb_PL) - 1
        pll_SwapLists i, i + 1
    Next i
End If

'update the cue.
If cue_ListInCue(UBound(mdb_PL)) Then
    cue_ListsSwapped UBound(mdb_PL), -1, True
End If

If UBound(mdb_PL) > 0 Then
    ReDim Preserve mdb_PL(UBound(mdb_PL) - 1)
Else
    pl_SetAsNew mdb_PL(0), "list0"
End If

'then take care of the saved file:
If Len(f) >= 0 Then
    If fsoMain.FileExists(f) Then
        MoveFileToRecycle f
    End If
End If

'then clean up:
If m_lCurrentPl > UBound(mdb_PL) Then
    'find last non-arc'ed pl
    For i = UBound(mdb_PL) To 0 Step -1
        If Not mdb_PL(i).bArc Then
            tsMain.Tabs("pl" & i).Selected = True
            GoTo LastNonArcFound
        End If
    Next i
    
    'if no pl is not arc'ed, goto lib
    tsMain.Tabs("lib").Selected = True
    
LastNonArcFound:
End If

Exit Sub
pll_RemoveList_err:
Main_Err "pll_RemoveList_err."
err.Clear
End Sub

Sub history_AddTo(lItemSource As Long, sFile As String, _
    Optional lTrackSource As Long = -1, Optional lMD5 As Long = 0)

Dim i As Long

'ensure list does not get too long
If mdb_History.lCnt >= lHistoryLimit Then
    pl_RemoveItem mdb_History, mdb_History.lCnt - 1
End If

'check item is not already in there
If mdb_History.lCnt > 0 Then
    For i = mdb_History.lCnt - 1 To 0 Step -1
        If lItemSource = mdb_History.Items(i).lSource(0) Then
            If IIf(lTrackSource >= 0, _
                lTrackSource = mdb_History.Items(i).lSource(1), _
                sFile = mdb_History.Items(i).sFile) Then
                
                pl_RemoveItem mdb_History, i
            End If
        End If
    Next i
End If

'add new item
pl_AddItem mdb_History, sFile, lMD5, , , , , lItemSource, lTrackSource

'move it to the top
If mdb_History.lCnt > 0 Then
    For i = mdb_History.lCnt - 1 To 1 Step -1
        pl_SwapItems mdb_History, i, i - 1
    Next i
End If
End Sub

Sub history_RebuildMenu()
Dim i As Long, s As String

If mnuHistoryI.count > 1 Then
    For i = mnuHistoryI.count - 1 To 1 Step -1
        Unload mnuHistoryI(i)
    Next i
End If

If mdb_History.lCnt <= 0 Then
    mnuHistoryI(0).Caption = "(empty)"
    mnuHistoryI(0).Enabled = False
Else
    For i = 0 To mdb_History.lCnt - 1
        If i > mnuHistoryI.count - 1 Then Load mnuHistoryI(i)
        
        s = i & " " & gen_GetShowName(mdb_History.Items(i).sFile)
        mnuHistoryI(i).Caption = s
        
        mnuHistoryI(i).Enabled = True
        mnuHistoryI(i).Visible = True
    Next i
End If
End Sub

Sub cue_AddTo(lItemSource As Long, sFile As String, _
    Optional lTrackSource As Long = -1, Optional lMD5 As Long = 0, Optional lDuration As Long = 0)

pl_AddItem mdb_Cue, sFile, lMD5, , , lDuration, , lItemSource, lTrackSource
cue_CheckVis
cue_Redraw
End Sub

Sub cue_RemSel()
Dim i As Long

If mdb_Cue.lCnt < 1 Then Exit Sub

For i = mdb_Cue.lCnt - 1 To 0 Step -1
    If mdb_Cue.Items(i).bSel Then
        pl_RemoveItem mdb_Cue, i
    End If
Next i

cue_CheckVis
cue_Redraw
End Sub

Function cue_TrackWaiting() As Boolean
cue_TrackWaiting = (mdb_Cue.lCnt > 0)
End Function

Function cue_GetNext(lSource As Long, lIndex As Long, Optional bRemove As Boolean = False)
lSource = mdb_Cue.Items(0).lSource(0)

If lSource = 0 Then
    lIndex = mdbl_FindIndexFromFile(mdb_Cue.Items(0).sFile)
ElseIf lSource = -1 Then
    'MsgBox "play file '" & mdb_Cue.Items(0).sFile & "'."
    ps_sMinusOneFilePath = mdb_Cue.Items(0).sFile
ElseIf lSource > 0 Then
    If mdb_Cue.Items(0).lSource(1) >= 0 Then
        lIndex = mdb_Cue.Items(0).lSource(1)
    Else
        lIndex = -1
    End If
Else
    Debug.Print "ooppss..."
End If

If bRemove Then
    pl_RemoveItem mdb_Cue, 0
    cue_CheckVis
    cue_Redraw
End If
End Function

Function cue_ItemInCue(lSource As Long, lIndex As Long) As Boolean
Dim i As Long

cue_ItemInCue = False

With mdb_Cue
    If .lCnt < 1 Then Exit Function
    
    For i = 0 To .lCnt - 1
        If .Items(i).lSource(0) = lSource Then
            If .Items(i).lSource(1) = lIndex Then
                cue_ItemInCue = True
                Exit Function
            End If
        End If
    Next i
End With
End Function

Function cue_ListInCue(lList As Long) As Boolean
Dim i As Long
cue_ListInCue = False

With mdb_Cue
    If .lCnt < 1 Then Exit Function
    
    For i = 0 To .lCnt - 1
        If .Items(i).lSource(0) = lList Then
            cue_ListInCue = True
            Exit Function
        End If
    Next i
End With
End Function

Sub cue_ItemMightHaveChanged(lListI As Long, lFirst As Long, lSecond As Long)
If cue_ItemInCue(lListI, lFirst) Then
    cue_ItemChanged lListI, lFirst, lSecond
End If
If cue_ItemInCue(lListI, lSecond) Then
    cue_ItemChanged lListI, lSecond, lFirst
End If
End Sub

Sub cue_ItemChanged(lSource As Long, lFrom As Long, lTo As Long)
Dim i As Long, x As Long

With mdb_Cue
    If .lCnt >= 1 Then
        For i = .lCnt - 1 To 0 Step -1
            If .Items(i).lSource(0) = lSource Then
                If .Items(i).lSource(1) = lFrom Then
                    .Items(i).lSource(1) = lTo
                End If
            End If
        Next i
    End If
End With

'this is a cheep hack because i can't be bothered to rename the sub and anything
'   that triggers a change in the cue will also affect the history list
With mdb_History
    If .lCnt >= 1 Then
        For i = .lCnt - 1 To 0 Step -1
            If .Items(i).lSource(0) = lSource Then
                If .Items(i).lSource(1) = lFrom Then
                    .Items(i).lSource(1) = lTo
                End If
            End If
        Next i
    End If
End With
End Sub

Sub cue_ListsSwapped(lFirst As Long, lSecond As Long, Optional bRem As Boolean = False)
Dim i As Long

With mdb_Cue
    If .lCnt >= 1 Then
        For i = 0 To .lCnt - 1
            If .Items(i).lSource(0) = lFirst Then
                .Items(i).lSource(0) = lSecond
            ElseIf .Items(i).lSource(0) = lSecond And Not bRem Then
                .Items(i).lSource(0) = lFirst
            End If
        Next i
    End If
End With

'see note in above sub.
With mdb_History
    If .lCnt >= 1 Then
        For i = 0 To .lCnt - 1
            If .Items(i).lSource(0) = lFirst Then
                .Items(i).lSource(0) = lSecond
            ElseIf .Items(i).lSource(0) = lSecond And Not bRem Then
                .Items(i).lSource(0) = lFirst
            End If
        Next i
    End If
End With
End Sub

Sub cue_Redraw()
pl_RebuildScroll mdb_Cue, picCue.ScaleHeight, vsbCue

cue_ListDc.Width = picCue.ScaleWidth
cue_ListDc.Height = picCue.ScaleHeight

'pll_Draw mdb_Cue, vsbCue.Value, cue_ListDc.hDC, _
    picCue.ScaleWidth, picCue.ScaleHeight, (picCue.Tag = "1"), _
    "the cue is empty." & vbNewLine & "tracks will be selected based on play mode."

'drawn directly so that we can overide the columns.
pl_Draw mdb_Cue, vsbCue.Value, cue_ListDc.hDC, _
    picCue.ScaleWidth, picCue.ScaleHeight, _
    (picCue.Tag = "1"), _
    "the cue is empty." & vbNewLine & "tracks will be selected based on play mode.", _
    "1000", _
    (chkOptHideExt.Value = 1), mdb_List.lTrunk, mdb_List.bTrunkReverse

picCue_Paint
End Sub

Sub cue_CheckVis()
Dim i As Long, bCueTabSel As Boolean, k As String, b As Boolean

bCueTabSel = False
For i = 1 To tsMain.Tabs.count
    If tsMain.Tabs(i).Key = "cue" Then bCueTabSel = True: Exit For
Next i

pl_RecountTotalDuration mdb_Cue
lblCueStat.Caption = mdb_Cue.lCnt & " items totaling " & _
    IIf(mdb_Cue.bTotalDurationComplete, "", "more than ") & ConvertSecToHours(mdb_Cue.lTotalDuration) & "."

If chkOptCuebtm.Value = 1 Then
    fraCue.Visible = (mdb_Cue.lCnt > 0)
    Form_Resize
End If

If bCueTabSel = (mdb_Cue.lCnt > 0) Then Exit Sub

If mdb_Cue.lCnt > 0 Then
    k = tsMain.SelectedItem.Key
    tsMain.Tabs.Add 2, "cue", "cue"
    tsMain.Tabs.Item("cue").Image = "cue"
    tsMain.Tabs(k).Selected = True
Else
    b = tsMain.Tabs("cue").Selected
    k = tsMain.SelectedItem.Key
    tsMain.Tabs.Remove "cue"
    If b Then tsMain_Click Else tsMain.Tabs(k).Selected = True
End If

'pll_RebuildEnab
End Sub

Sub pb_StartPlayback(ByVal lSource As Long, Optional ByVal lTrack As Long = -1)
On Error GoTo pb_StartPlayback_err

Dim sFile As String, e As String

e = "stopping current item."
If pb_GetPlayState = stPaused Then pb_SetPlayState stStopped

e = "setting source."
pb_lItemSource = lSource

e = "retreving target item."
If lSource > 0 And lTrack >= 0 Then
    mdb_PL(lSource - 1).lCurrent = lTrack
    pb_llItemIndex = -1
    sFile = mdb_PL(lSource - 1).Items(lTrack).sFile
ElseIf lSource = 0 Then
    pb_llItemIndex = lTrack
    sFile = mdb_List.Items(lTrack).sFile
ElseIf lSource = -1 And Len(ps_sMinusOneFilePath) > 0 Then
    pb_llItemIndex = 0
    sFile = ps_sMinusOneFilePath
Else
    Debug.Print "ooppss..."
End If

e = "setting playback state."
pb_SetPlayState stPlaying

e = "adding previous item to history."
history_AddTo lSource, sFile, lTrack

Exit Sub
pb_StartPlayback_err:
Main_Err "pb_StartPlayback_err/" & e & ".  " & lSource & "," & lTrack & "," & sFile & "."
err.Clear
End Sub

Function pb_CheckCurrentIndex() As Boolean
pb_CheckCurrentIndex = False

If pb_lItemSource = 0 Then
    If pb_llItemIndex >= 0 Then
        pb_CheckCurrentIndex = True
    End If
ElseIf pb_lItemSource > 0 Then
    If mdb_PL(pb_lItemSource - 1).lCurrent >= 0 Then
        pb_CheckCurrentIndex = True
    End If
ElseIf pb_lItemSource = -1 And Len(ps_sMinusOneFilePath) > 0 Then
    pb_CheckCurrentIndex = True
End If
End Function

Function pb_GetPlayState() As Integer
If cMedia Is Nothing Then
    pb_GetPlayState = 0
    Exit Function
End If
If pb_CheckCurrentIndex = False And cMedia.State = stStopped Then
    pb_GetPlayState = 0
    Exit Function
End If

Select Case cMedia.State
    Case stStopped
        pb_GetPlayState = 0
    Case stPlaying
        pb_GetPlayState = 1
    Case stPaused
        pb_GetPlayState = 2
    Case Else
        pb_GetPlayState = 0
End Select
End Function

Sub pb_SetPlayState(iState As Integer)
On Error GoTo pb_SetPlayState_err

If pb_CheckCurrentIndex = False And cMedia.State = stStopped Then Exit Sub

Dim f As String, l As Long, b As Boolean, e As String, a As String

Select Case iState
    Case 0 'stop
        If cMedia.State <> stStopped Then
            cMedia.StopPlaying
            cMedia.FileName = ""
            If pb_lItemSource = 0 Then
                mdbl_Redraw
            ElseIf pb_lItemSource - 1 = m_lCurrentPl Then
                pll_RedrawCurrent
            End If
            
            UpdateFormCaption
            gen_UpdateVidPos
        End If
    
    Case 1 'playing
        If cMedia.State = stPaused Then
            cMedia.Play
        Else
            'get ready
            cMedia.StopPlaying
            
            'find the file to play
            If pb_lItemSource = 0 Then
                f = mdb_List.Items(pb_llItemIndex).sFile
            ElseIf pb_lItemSource > 0 Then
                'an item from a playlist
                f = mdb_PL(pb_lItemSource - 1).Items(mdb_PL(pb_lItemSource - 1).lCurrent).sFile
            ElseIf pb_lItemSource = -1 Then
                f = ps_sMinusOneFilePath
                ps_sMinusOneFilePath = ""
            Else
                Debug.Print "ooppss... this should never happen."
            End If
            
            If Not fsoMain.FileExists(f) Then
                e = "file not found: '" & f & "'."
                GoTo FileCouldNotBePlayed
            End If
            
            'more getting ready
            cMedia.FileName = f
            cMedia.Volume = 0 'max
            cMedia.Balance = 0 'middle
            cMedia.Speed = pb_Speed '1 '1:1
            
            If cMedia.Duration <= 0 Then
                e = "file can not be played: '" & f & "' has a duration <= 0.  system codec missing or file may be damaged."
                GoTo FileCouldNotBePlayed
            End If
            
            cMedia.Play 'play the file
            gen_UpdateVidPrevVis 'gen_UpdateVidPos 'update vid position
            pb_RetryCount = 0 'if we get this far then this file must have worked, so reset retry counter.
            
            tmrVidBumper.Enabled = True 'this is a really really really messy work-around.
            
            'update list play-count data
            If pb_lItemSource = 0 Then
                mdb_List.Items(pb_llItemIndex).lStartCnt = mdb_List.Items(pb_llItemIndex).lStartCnt + 1
                mdb_List.Items(pb_llItemIndex).dLastPlay = Now
            ElseIf pb_lItemSource > 0 Then
                mdb_PL(pb_lItemSource - 1).Items(mdb_PL(pb_lItemSource - 1).lCurrent).lStartCnt = _
                    mdb_PL(pb_lItemSource - 1).Items(mdb_PL(pb_lItemSource - 1).lCurrent).lStartCnt + 1
                mdb_PL(pb_lItemSource - 1).Items(mdb_PL(pb_lItemSource - 1).lCurrent).dLastPlay = Now
                pll_UpdateStatus True, True 'pl_SetAutoSave True
            Else
                'form a non-list source, so no play-count we can increase.
            End If
            
            'update list duration data
            If pb_lItemSource = 0 Then
                mdb_List.Items(pb_llItemIndex).lDuration = cMedia.Duration
            ElseIf pb_lItemSource > 0 Then
                mdb_PL(pb_lItemSource - 1).Items(mdb_PL(pb_lItemSource - 1).lCurrent).lDuration = cMedia.Duration
                pl_RecountTotalDuration mdb_PL(pb_lItemSource - 1)
            Else
                'form a non-list source, so nothing to do.
            End If
            
            'tracking current item?
            If m_TrackCurrentItem Then
                If cMedia.HasVideo Then
                    list_JumpToCurrent False, False
                    tsMain.Tabs("vid").Selected = True
                Else
                    list_JumpToCurrent True
                End If
            Else
                If pb_lItemSource = 0 Then
                    mdbl_Redraw
                ElseIf pb_lItemSource = m_lCurrentPl + 1 Then
                    pll_RedrawCurrent
                Else
                    'must be in a playlist that is not visable.
                End If
            End If
            
            UpdateFormCaption 'update the from
            
            mdb_IncPlyCnt f                     'update the db counts
            mdb_SetDuration f, cMedia.Duration  'and duration data
            
        End If
    
    Case 2 'pause
        cMedia.Pause
    
End Select

EscapeSub:
cmdPlayState(1).Caption = IIf(pb_GetPlayState = 1, ";", "4")
If m_bMiniVid Then frmMiniVid.UpdatedatePlaystate

Exit Sub

FileCouldNotBePlayed:
AddToLog txtLog, e
If cbbOptErrNoplay.ListIndex = 1 Then
    If pb_RetryCount > pb_RetryCount_Max Then
        a = "the last " & pb_RetryCount_Max & " attempts to play a file have failed.  " & _
            "this was because they could not be found or their duration was <= 0.  " & _
            "play back has been halted to prevent an endless loop."
        AddToLog txtLog, a
        Main_Err a, 1
        GoTo EscapeSub
    End If
    pb_NextFile
    pb_RetryCount = pb_RetryCount + 1
Else
    Main_Err e, 1
End If

GoTo EscapeSub

Exit Sub

pb_SetPlayState_err:
Main_Err "pb_SetPlayState_err."
err.Clear
End Sub

Sub pb_NextFile()
On Error GoTo pb_NextFile_err

Dim e As String, _
    lMode As Long, lListMode As Long, i As Long, _
    lList As Long, lItem As Long

e = "0"

'see if there is a track in the cue we should use
If cue_TrackWaiting Then
    cue_GetNext lList, lItem, True
    GoTo TrackChosen
End If

'cue is empty, so find a track based on play mode

e = "1"
'find play modes
For i = 0 To mnuModeI.count - 1
    If mnuModeI(i).Checked = True Then
        lMode = i
        Exit For
    End If
Next i

For i = 0 To mnuModeListsI.count - 1
    If mnuModeListsI(i).Checked = True Then
        lListMode = i
        Exit For
    End If
Next i

e = "2"
'start on the previous list
lList = pb_lItemSource

e = "3"
'see if we need to jump lists
Select Case lMode
    Case 1, 2, 3 'all non-sequencial
        'todo: make something better for when in /play-start or /last-played mode?
        If lListMode = 1 Then
            'note: by excluding the "lSource > 0" term it will "jump out"
            '      of the lib when "use lib" is unchecked.
            lList = pbq_GetListTruerand(False)
        ElseIf lListMode = 2 Then
            lList = pbq_GetListTruerand(True)
        End If
    
End Select

e = "3.5"
If lList < 0 Then
    MsgBox "can not find a list to play tracks from."
    Exit Sub
End If

e = "4"
'now we have the list, find the item
Select Case lMode
    Case 0 'sequencial
        lItem = pbq_GetItemSeq(lList, Not lListMode = 0)
    
    Case 1 'random
        lItem = pbq_GetItemTruerand(lList)
    
    Case 2 'semi-random
        lItem = pbq_GetItemSemirand(lList)
    
    Case 3 '[not played recently]
        lItem = pbq_GetItemLastplayed2(lList)
    
End Select

e = "5"
'all done!  now play the file.
TrackChosen:

If lItem >= 0 Then
    pb_StartPlayback lList, lItem
Else
    MsgBox "unable to find a file to play."
End If

Exit Sub
pb_NextFile_err:
Main_Err "pb_NextFile_err/" & e & ".  " & lListMode & "," & lList & ";" & lMode & "," & lItem & "."
err.Clear
End Sub

Sub pb_SetPlaybackPosition(ByVal l As Long)
'If pb_CheckCurrentIndex = False Then Exit Sub
If pb_GetPlayState <= 0 Then Exit Sub

If l < 0 Then
    l = 0
ElseIf l > cMedia.Duration Then
    l = cMedia.Duration - 1 'just to make sure
End If

cMedia.Position = l
cMedia.Play
cmdPlayState(1).Caption = ";"
End Sub

Public Sub pb_DrawCurrentItem(lHdc As Long, lW As Long, lH As Long, lTextCol As Long)
Dim rcTitle As RECT, rcPlayCount As RECT, rcSource As RECT, rcPath As RECT, _
    rcCalc As RECT, s As String, sPlayCnt As String, _
    i As Long, e As String

On Error GoTo pb_DrawCurrentItem_err

e = "1"
SetTextColor lHdc, lTextCol

'draw title
e = "2"
s = RemExtFromPath(FileNameFromPath(cMedia.FileName))
If chkOptDisplayRepDashNewLine.Value = 1 Then s = Replace(s, " - ", vbNewLine)
SelectObject lHdc, gdi_Main_hFontTripple
rcCalc.Left = 0
rcCalc.right = lW
rcCalc.Top = 0
rcCalc.bottom = 1
DrawText lHdc, s, Len(s), rcCalc, DT_CALCRECT + DT_WORDBREAK + DT_NOPREFIX
rcTitle.Left = 0
rcTitle.right = lW
rcTitle.Top = (lH / 2) - ((rcCalc.bottom - rcCalc.Top) / 2)
rcTitle.bottom = rcTitle.Top + (rcCalc.bottom - rcCalc.Top)
DrawText lHdc, s, Len(s), rcTitle, DT_CENTER + DT_WORDBREAK + DT_NOPREFIX

'draw source and find play count
e = "3"
sPlayCnt = ""
If pb_lItemSource = 0 Then
    s = "(media library)"
    i = mdbl_FindIndexFromFile(cMedia.FileName)
    If i >= 0 Then
        sPlayCnt = mdb_List.Items(i).lStartCnt & _
            " / " & mdb_List.Items(i).lEndCnt & "     " & _
            ConvertSecToMin(mdb_List.Items(i).lDuration)
    End If
ElseIf pb_lItemSource > 0 Then
    s = mdb_PL(pb_lItemSource - 1).sName
    If mdb_PL(pb_lItemSource - 1).lCurrent >= 0 Then
        sPlayCnt = _
            mdb_PL(pb_lItemSource - 1).Items(mdb_PL(pb_lItemSource - 1).lCurrent).lStartCnt _
            & " / " & _
            mdb_PL(pb_lItemSource - 1).Items(mdb_PL(pb_lItemSource - 1).lCurrent).lEndCnt & _
            "     " & _
            ConvertSecToMin(mdb_PL(pb_lItemSource - 1).Items(mdb_PL(pb_lItemSource - 1).lCurrent).lDuration)
    Else
        sPlayCnt = "/"
    End If
Else
    s = "(no list source)"
End If

'draw source
e = "4"
SelectObject lHdc, gdi_Main_hFontDouble
rcCalc.Left = 0
rcCalc.right = lW
rcCalc.Top = 0
rcCalc.bottom = 1
DrawText lHdc, s, Len(s), rcCalc, DT_CALCRECT + DT_WORDBREAK + DT_NOPREFIX
rcSource.Left = 0
rcSource.right = lW
rcSource.Top = rcTitle.Top - (rcCalc.bottom - rcCalc.Top)
rcSource.bottom = rcTitle.Top
DrawText lHdc, s, Len(s), rcSource, DT_CENTER + DT_WORDBREAK + DT_NOPREFIX

'if found, draw play count
e = "5"
If Len(sPlayCnt) > 0 Then
    SelectObject lHdc, gdi_Main_hFontDouble
    rcCalc.Left = 0
    rcCalc.right = lW
    rcCalc.Top = 0
    rcCalc.bottom = 1
    DrawText lHdc, sPlayCnt, Len(sPlayCnt), rcCalc, DT_CALCRECT + DT_WORDBREAK + DT_NOPREFIX
    rcPlayCount.Left = 0
    rcPlayCount.right = lW
    rcPlayCount.Top = rcTitle.bottom
    rcPlayCount.bottom = rcPlayCount.Top + (rcCalc.bottom - rcCalc.Top)
    DrawText lHdc, sPlayCnt, Len(sPlayCnt), rcPlayCount, DT_CENTER + DT_WORDBREAK + DT_NOPREFIX
End If

'draw file path
e = "6"
s = FileFolderFromPath(cMedia.FileName)
SelectObject lHdc, gdi_Main_hFontNormal
rcCalc.Left = 0
rcCalc.right = lW
rcCalc.Top = 0
rcCalc.bottom = 1
DrawText lHdc, s, Len(s), rcCalc, DT_CALCRECT + DT_WORDBREAK + DT_NOPREFIX
rcPath.Left = 0
rcPath.right = lW
If Len(sPlayCnt) > 0 Then
    rcPath.Top = rcPlayCount.bottom
Else
    rcPath.Top = rcTitle.bottom
End If
rcPath.bottom = rcPath.Top + (rcCalc.bottom - rcCalc.Top)
DrawText lHdc, s, Len(s), rcPath, DT_CENTER + DT_WORDBREAK + DT_NOPREFIX

Exit Sub
pb_DrawCurrentItem_err:
Main_Err "pb_DrawCurrentItem_err/" & e
err.Clear
End Sub

Public Sub pb_DrawCurrentItemSmall(lHdc As Long, lW As Long, lH As Long, lTextCol As Long)
Dim s As String, rcCalc As RECT, rcTitle As RECT

SetTextColor lHdc, lTextCol

'draw title
s = RemExtFromPath(FileNameFromPath(cMedia.FileName))
If chkOptDisplayRepDashNewLine.Value = 1 Then s = Replace(s, " - ", vbNewLine)
SelectObject lHdc, gdi_Main_hFontNormal
rcCalc.Left = 0
rcCalc.right = lW
rcCalc.Top = 0
rcCalc.bottom = 1
DrawText lHdc, s, Len(s), rcCalc, DT_CALCRECT + DT_WORDBREAK + DT_NOPREFIX
rcTitle.Left = 0
rcTitle.right = lW
If (rcCalc.bottom - rcCalc.Top) > lH Then
    rcTitle.Top = lH - (rcCalc.bottom - rcCalc.Top)
Else
    rcTitle.Top = (lH / 2) - ((rcCalc.bottom - rcCalc.Top) / 2)
End If
rcTitle.bottom = rcTitle.Top + (rcCalc.bottom - rcCalc.Top)
DrawText lHdc, s, Len(s), rcTitle, DT_CENTER + DT_WORDBREAK + DT_NOPREFIX
End Sub

Function pbq_GetListTruerand(bIncLib As Boolean) As Long
Dim l As Long, i As Long, _
    lMax As Long, lTarget As Long, lOutputSource As Long, _
    lLibTotal As Long

'just in case?!
lOutputSource = -1

'are we using lib?
If bIncLib Then
    'if so, find max, taking into account disabled items.
    lLibTotal = 0
    For i = 0 To mdb_List.lCnt - 1
        If mdb_List.Items(i).bEnabled And Not mdb_List.Items(i).bMissing Then
            lLibTotal = lLibTotal + 1
        End If
    Next i
    
    lMax = lLibTotal - 1
    l = -1
Else
    'if not, set to starting values.
    lMax = -1
    l = 0
End If

'total playlist items, taking into account disalbed lists.
For i = 0 To UBound(mdb_PL)
    If mdb_PL(i).bEnab And Not mdb_PL(i).bArc Then
        lMax = lMax + mdb_PL(i).lCnt
    End If
Next i

'check that there actually are some files to look for.
If lMax < 1 Then GoTo ReturnResult

'all traks accounts for, pick our target number.
Randomize
lTarget = Rnd() * lMax

'find where our target is.
For i = l To UBound(mdb_PL)
    If i = -1 Then
        lTarget = lTarget - lLibTotal
    ElseIf mdb_PL(i).bEnab And Not mdb_PL(i).bArc Then
        lTarget = lTarget - mdb_PL(i).lCnt
    Else
        GoTo SkipDisabledPL
    End If
    If lTarget < 0 Then
        lOutputSource = IIf(i < 0, 0, i + 1)
        Exit For
    End If
SkipDisabledPL:
Next i

'return target.
ReturnResult:
pbq_GetListTruerand = lOutputSource
End Function

Function pbq_GetItemSeq(lSourceList As Long, bUseAllLists As Boolean) As Long
Dim i As Long, _
    lLibTotal As Long, lMax As Long, lOutputItem As Long

'add 1 to prev. item.  simple, unless in lib where items may be disabled.
If lSourceList = 0 Then
    'count the enabled items.
    lLibTotal = 0
    For i = 0 To mdb_List.lCnt - 1
        If mdb_List.Items(i).bEnabled And Not mdb_List.Items(i).bMissing Then
            lLibTotal = lLibTotal + 1
        End If
    Next i
    
    'check that at least one item is still avaible.
    If lLibTotal < 1 Then
        lOutputItem = -1
        GoTo ReturnValue
    End If
    
    'try adding 1, and keep going until an active item is found.
    'list_GetCurrentIndex will default to -1 if no item is found.
    lOutputItem = list_GetCurrentIndex
    Do
        lOutputItem = lOutputItem + 1
        If lOutputItem > lLibTotal - 1 Then Exit Do
        
        If mdb_List.Items(lOutputItem).bEnabled And _
            Not mdb_List.Items(lOutputItem).bMissing Then
            
            Exit Do
        End If
    Loop
Else
    'its a playlist, simple (atm).
    lOutputItem = list_GetCurrentIndex + 1
End If

'see which list we are using, and get the total number of items.
If lSourceList = 0 Then
    lMax = lLibTotal - 1
Else
    lMax = mdb_PL(lSourceList - 1).lCnt - 1
End If

'check to see if we have gone off the btm.
If lOutputItem > lMax Then
    If bUseAllLists Then
        If (lSourceList - 1) >= UBound(mdb_PL) Then
            lSourceList = 0
        Else
            lSourceList = lSourceList + 1
        End If
    End If
    lOutputItem = 0
End If

'return final value.
ReturnValue:
pbq_GetItemSeq = lOutputItem
End Function

Function pbq_GetItemTruerand(lSourceList As Long) As Long
Dim i As Long, lLibTotal As Long, lOutputItem As Long

'set default result.
lOutputItem = -1

'if no items, might as well give up now.
If lSourceList = 0 Then
    If mdb_List.lCnt < 1 Then GoTo ReturnResult
Else
    If mdb_PL(lSourceList - 1).lCnt < 1 Then GoTo ReturnResult
End If

'always a good idea.
Randomize

If lSourceList = 0 Then
    'count the enabled items.
    lLibTotal = 0
    For i = 0 To mdb_List.lCnt - 1
        If mdb_List.Items(i).bEnabled And Not mdb_List.Items(i).bMissing Then
            lLibTotal = lLibTotal + 1
        End If
    Next i
    
    'pick a value
    lOutputItem = Rnd() * lLibTotal
    
    'work out which one it is
    For i = 0 To mdb_List.lCnt - 1
        If mdb_List.Items(i).bEnabled And Not mdb_List.Items(i).bMissing Then
            lOutputItem = lOutputItem - 1
            If lOutputItem < 0 Then
                lOutputItem = i
                Exit For
            End If
        End If
    Next i
Else
    'its a nice simple playlist.
    lOutputItem = Rnd() * mdb_PL(lSourceList - 1).lCnt
End If

'return value.
ReturnResult:
pbq_GetItemTruerand = lOutputItem
End Function

Function pbq_GetItemSemirand(lSourceList As Long) As Long
Dim sSql As String, vRet As Variant, vI As Variant, _
    lSum As Long, lMax As Long, lTarget As Long, _
    i As Long, _
    lOutputItem As Long

'which list?
If lSourceList = 0 Then
    'check that there are actually some items.
    If mdb_List.lCnt < 1 Then
        'abort!
        lOutputItem = -1
        GoTo LeaveSub
    End If
    
    'in the mdb we can use queries to find the most played file.
    'sSQL = "select max(lstartcnt) from tbl_mediafiles;"
    sSql = "select max(lstartcnt) from tbl_mediafiles where benabled=1 or benabled is NULL;"
    vRet = mdb_Query(sSql)
    For Each vI In vRet
        If Mid$(vI, 1, 3) <> "max" Then
            lMax = vI
            Exit For
        End If
    Next vI
    
    'build sum of all selection-indicies.
    lSum = 0
    For i = 0 To mdb_List.lCnt - 1
        If mdb_List.Items(i).bEnabled And Not mdb_List.Items(i).bMissing Then
            lSum = lSum + (lMax - mdb_List.Items(i).lStartCnt)
        End If
    Next i
    
    'generate target selection index.
    Randomize
    lTarget = Rnd() * lSum
    
    'find the target item.
    For i = 0 To mdb_List.lCnt - 1
        If mdb_List.Items(i).bEnabled And Not mdb_List.Items(i).bMissing Then
            lTarget = lTarget - (lMax - mdb_List.Items(i).lStartCnt)
            If lTarget <= 0 Then
                lOutputItem = i
                Exit For
            End If
        End If
    Next i
Else
    'todo: something better here than exit sub?
    '      though this should never come up, as the list selection code
    '      should skip over empty lists.
    If mdb_PL(lSourceList - 1).lCnt < 1 Then
        'abort!
        lOutputItem = -1
        GoTo LeaveSub
    End If
    
    'find the most played file
    lMax = 0
    For i = 0 To mdb_PL(lSourceList - 1).lCnt - 1
        If mdb_PL(lSourceList - 1).Items(i).lStartCnt > lMax Then
            lMax = mdb_PL(lSourceList - 1).Items(i).lStartCnt
        End If
    Next i
    
    'build sum of all selection-indicies
    lSum = 0
    For i = 0 To mdb_PL(lSourceList - 1).lCnt - 1
        lSum = lSum + (lMax - mdb_PL(lSourceList - 1).Items(i).lStartCnt)
    Next i
    
    'generate target selection index
    Randomize
    lTarget = Rnd() * lSum
    
    'find the target item.
    For i = 0 To mdb_PL(lSourceList - 1).lCnt - 1
        lTarget = lTarget - (lMax - mdb_PL(lSourceList - 1).Items(i).lStartCnt)
        If lTarget <= 0 Then
            lOutputItem = i
            Exit For
        End If
    Next i
End If

LeaveSub:
pbq_GetItemSemirand = lOutputItem
End Function

'Function pbq_GetItemLastplayed(lSourceList As Long) As Long
'Dim lOutputItem As Long
'
'If lSourceList = 0 Then
'    lOutputItem = pbq_GetItemLibraryRandFromQueryFull( _
'        "SELECT sfile FROM tbl_mediafiles WHERE benabled=1 OR benabled IS NULL " & _
'        "ORDER BY dlastplay,dadded ASC LIMIT 10;")
'Else
'    'todo: something better here than exit sub?
'    '      though this should never come up, as the list selection code
'    '      should skip over empty lists.
'    If mdb_PL(lSourceList - 1).lCnt < 1 Then
'        'abort!
'        lOutputItem = -1
'        GoTo LeaveSub
'    End If
'
'    Randomize Timer
'    'if there are more than 10 items...
'    If mdb_PL(lSourceList - 1).lCnt > 10 Then
'        'duplicate the playlist
'        Dim lDupList As typPl, i As Long
'        lDupList = mdb_PL(lSourceList - 1)
'
'        'tag each item so we know where it came from before we sorted the list
'        For i = 0 To UBound(lDupList.Items)
'            lDupList.Items(i).lSource(1) = i
'        Next i
'
'        'sort the list by date
'        pl_Sort lDupList, 3, , , False
'
'        'pick an item at random from the top 10
'        i = Int(Rnd() * 10) 'a number between 0 and 9 (int() does no rounding)
'
'        'return item
'        lOutputItem = lDupList.Items(i).lSource(1)
'    Else
'        'if there are less than 10 items, then everything above makes no difference
'        lOutputItem = Int(Rnd() * mdb_PL(lSourceList - 1).lCnt)
'    End If
'
'End If
'
'LeaveSub:
'pbq_GetItemLastplayed = lOutputItem
'End Function

Function pbq_GetItemLastplayed2(lSourceList As Long) As Long
Dim lOutputItem As Long, i As Long, _
    lOldest As Long, lSumDays As Long, lTarget As Long

If lSourceList = 0 Then
    'check that there are actually some items.
    If mdb_List.lCnt < 1 Then
        lOutputItem = -1
        GoTo LeaveSub
    End If
    
    'find oldest date
    i = pbq_GetItemLibraryRandFromQueryFull( _
        "SELECT sfile FROM tbl_mediafiles WHERE (benabled=1 OR benabled IS NULL) AND dlastplay>'2000-01-01' " & _
        "ORDER BY dlastplay ASC, dadded ASC LIMIT 1;")
    lOldest = DateDiff("d", mdb_List.Items(i).dLastPlay, Now)
    Debug.Print "lOldest=" & lOldest & " (" & Format(mdb_List.Items(i).dLastPlay, "YYYY/MM/DD hh:mm:ss") & ")"
    
    'build sum of all selection-indicies in units of days.
    lSumDays = 0
    For i = 0 To mdb_List.lCnt - 1
        If mdb_List.Items(i).bEnabled And Not mdb_List.Items(i).bMissing Then
            'lSumDays = lSumDays + (lOldest - DateDiff("d", mdb_List.Items(i).dLastPlay, Now))
            
            'if an item has not been played, assume it is as old as the oldest item.
            lSumDays = lSumDays + IIf(mdb_List.Items(i).dLastPlay < 1, lOldest, _
                lOldest - DateDiff("d", mdb_List.Items(i).dLastPlay, Now))
        End If
    Next i
    
    'generate target number
    Randomize
    lTarget = Rnd() * lSumDays
    
    'find the target item.
    For i = 0 To mdb_List.lCnt - 1
        If mdb_List.Items(i).bEnabled And Not mdb_List.Items(i).bMissing Then
            'lTarget = lTarget - (lOldest - DateDiff("d", mdb_List.Items(i).dLastPlay, Now))
            lTarget = lTarget - IIf(mdb_List.Items(i).dLastPlay < 1, lOldest, _
                lOldest - DateDiff("d", mdb_List.Items(i).dLastPlay, Now))
            
            If lTarget <= 0 Then
                lOutputItem = i
                Exit For
            End If
        End If
    Next i
Else
    If mdb_PL(lSourceList - 1).lCnt < 1 Then
        lOutputItem = -1
        GoTo LeaveSub
    End If
    
    'find oldest item in list
    lOldest = 0
    For i = 0 To mdb_PL(lSourceList - 1).lCnt - 1
        If DateDiff("d", mdb_PL(lSourceList - 1).Items(i).dLastPlay, Now) > lOldest Then
            lOldest = DateDiff("d", mdb_PL(lSourceList - 1).Items(i).dLastPlay, Now)
        End If
    Next i
    
    'build sum of all selection-indicies
    lSumDays = 0
    For i = 0 To mdb_PL(lSourceList - 1).lCnt - 1
        'lSumDays = lSumDays + (lOldest - DateDiff("d", mdb_PL(lSourceList - 1).Items(i).dLastPlay, Now))
        lSumDays = lSumDays + IIf(mdb_PL(lSourceList - 1).Items(i).dLastPlay < 1, lOldest, _
            lOldest - DateDiff("d", mdb_PL(lSourceList - 1).Items(i).dLastPlay, Now))
    Next i
    
    'generate target selection index
    Randomize
    lTarget = Rnd() * lSumDays
    
    'find the target item.
    For i = 0 To mdb_PL(lSourceList - 1).lCnt - 1
        'lTarget = lTarget - (lOldest - DateDiff("d", mdb_PL(lSourceList - 1).Items(i).dLastPlay, Now))
        lTarget = lTarget - IIf(mdb_PL(lSourceList - 1).Items(i).dLastPlay < 1, lOldest, _
            lOldest - DateDiff("d", mdb_PL(lSourceList - 1).Items(i).dLastPlay, Now))
        If lTarget <= 0 Then
            lOutputItem = i
            Exit For
        End If
    Next i
End If

LeaveSub:
pbq_GetItemLastplayed2 = lOutputItem
End Function

Function pbq_GetItemLibraryRandFromQuery(sSqlWhere As String, _
    Optional sExtraFields As String = "") As Long

Dim lOutputItem As Long, i As Long, x As Long, _
    sSql As String, vRet As Variant, vI As Variant, lQueryRowCnt As Long

sSql = "select sfile" & _
    IIf(Len(sExtraFields) > 0, ", " & sExtraFields, "") & _
    " from tbl_mediafiles where " & sSqlWhere & " and (benabled=1 or benabled is NULL);"

vRet = mdb_Query(sSql)
lQueryRowCnt = number_of_rows_from_last_call

If lQueryRowCnt < 1 Then
    lOutputItem = -1
    GoTo LeaveSub
End If

Randomize
i = Int(Rnd() * lQueryRowCnt)

x = -1
For Each vI In vRet
    If x = i Then
        lOutputItem = mdbl_FindIndexFromFile(us_Decode((vI)))
        Exit For
    End If
    x = x + 1
Next vI

LeaveSub:
pbq_GetItemLibraryRandFromQuery = lOutputItem
End Function

Function pbq_GetItemLibraryRandFromQueryFull(sSql As String) As Long

Dim lOutputItem As Long, i As Long, x As Long, _
    vRet As Variant, vI As Variant, lQueryRowCnt As Long

vRet = mdb_Query(sSql)
lQueryRowCnt = number_of_rows_from_last_call

'Debug.Print "sSQL=" & sSQL
'Debug.Print "lQueryRowCnt=" & lQueryRowCnt

If lQueryRowCnt < 1 Then
    lOutputItem = -1
    GoTo LeaveSub
End If

Randomize
i = Int(Rnd() * lQueryRowCnt)

'Debug.Print "i=" & i

x = -1 'first row is field names
For Each vI In vRet
    If x = i Then
        lOutputItem = mdbl_FindIndexFromFile(us_Decode((vI)))
        Exit For
    End If
    x = x + 1
Next vI

'Debug.Print "lOutputItem=" & lOutputItem

LeaveSub:
pbq_GetItemLibraryRandFromQueryFull = lOutputItem
End Function

Sub UpdateFormCaption()
Dim a As String

If pb_GetPlayState > 0 Then
    a = gen_GetShowName(cMedia.FileName)
Else
    a = ""
End If

a = a & sMainWindowID

If Caption <> a Then
    Caption = a
    If m_bInTray Then SystemTray.ChangeCaption a
End If

If tsMain.SelectedItem.Key = "vid" Then
    fraVidPlace.ForceRedraw
End If
If m_bFullScreen Then frmDisp.UpdateDisplay
If m_bMiniVid Then frmMiniVid.fraVid.ForceRedraw
End Sub

Sub tray_Min()
On Error GoTo tray_Min_err:

SystemTray.PlaceIcon sytMinimize.GetHwnd, Me.Icon, Caption
If WindowState <> vbMinimized Then SystemTray.FlyingWindow Me, 1
m_bInTray = True
Hide

Exit Sub
tray_Min_err:
Main_Err "tray_Min_err."
err.Clear
End Sub

Sub tray_Rest()
If WindowState = vbMinimized Then
    WindowState = m_WsBeforeMin
Else
    SystemTray.FlyingWindow Me, 0
End If
Show
SystemTray.RemoveIcon
m_bInTray = False
End Sub

Sub wMsgProc_Hotkey(lParam As Long)
Dim i As Long

For i = 0 To UBound(cHk)
    If cHk(i).IsKey(lParam) Then
        Select Case i
            Case 0: If Me.Visible Then tray_Min Else tray_Rest
            Case 1: cmdPlayState_Click 1
            Case 2: cmdPlayState_Click 0
            Case 3: cmdPlayNext_Click
            Case 4: gen_ShowGoto
        End Select
    End If
Next i
End Sub

Sub wMsgProc_MouseWheel(wParam As Long)
Dim v As Long
If m_lCurrentPl >= 0 And vsbPl.Enabled Then
    Select Case Sgn(wParam)
        Case Is < 0: v = vsbPl.Value + lScrollWheelMovement
        Case Is > 0: v = vsbPl.Value - lScrollWheelMovement
    End Select
    pl_SetScroll vsbPl, v
ElseIf fraLib.Visible And vsbMdbLst.Enabled Then
    Select Case Sgn(wParam)
        Case Is < 0: v = vsbMdbLst.Value + lScrollWheelMovement
        Case Is > 0: v = vsbMdbLst.Value - lScrollWheelMovement
    End Select
    mdbl_SetScroll v
End If
End Sub

Sub gen_CopyFilesToFolder(cFiles As Collection)
On Error GoTo gen_CopyFilesToFolder_err

Static sPrevFolder As String
If Len(sPrevFolder) < 1 Then sPrevFolder = ""

Dim i As Long, f As String, sTargetFile As String, _
    cFilesFailed As Collection

mnuLibSelFilecopy.Enabled = False
mnuPlFilecopy.Enabled = False

lblFileCopy.Caption = "waiting for user to select target folder..."
fraNotice(2).Visible = True
Form_Resize

f = BrowseForFolder(sPrevFolder, "copy " & cFiles.count & " to a folder...")
If f = "" Then GoTo CancelCopy
If right$(f, 1) <> "\" Then f = f & "\"
If Not fsoMain.FolderExists(f) Then GoTo AbortCopy
sPrevFolder = f

Set cFilesFailed = New Collection
bAbort = False

For i = 1 To cFiles.count
    lblFileCopy.Caption = "copying file " & i & " of " & cFiles.count & " to " & f & "..."
    
    sTargetFile = f & FileNameFromPath(cFiles(i))
    
    If fsoMain.FileExists(cFiles(i)) Then
        If fsoMain.FileExists(sTargetFile) Then
            cFilesFailed.Add "target already exists <" & sTargetFile & ">"
        Else
            fsoMain.CopyFile cFiles(i), sTargetFile, False
        End If
    Else
        cFilesFailed.Add "source not found <" & cFiles(i) & ">"
    End If
    
    DoEvents
    If bAbort Then Exit For
Next i

AbortCopy:

If cFilesFailed.count > 0 Then
    Dim a As String
    a = "errors occred while copying files." & vbNewLine & vbNewLine
    For i = 1 To IIf(cFilesFailed.count > 12, 12, cFilesFailed.count)
        a = a & cFilesFailed(i) & vbNewLine
    Next i
    
    If cFilesFailed.count > 12 Then
        a = a & "... (too many to show)"
    End If
    
    MsgBox a
End If

CancelCopy:

fraNotice(2).Visible = False
Form_Resize

mnuLibSelFilecopy.Enabled = True
mnuPlFilecopy.Enabled = True

Exit Sub
gen_CopyFilesToFolder_err:
Main_Err "gen_CopyFilesToFolder."
err.Clear
GoTo CancelCopy
End Sub

Sub gen_UpdateVidPos()
On Error GoTo gen_UpdateVidPos_err

Dim lhWnd As Long, bShowMouse As Boolean
bShowMouse = False

If m_bFullScreen Then
    lhWnd = frmDisp.hWND
ElseIf m_bMiniVid Then
    lhWnd = frmMiniVid.fraVid.hWND
    bShowMouse = True
Else
    If chkOptDisplayMinivid.Value <> 1 Then
        lhWnd = fraVidPlace.hWND
    Else
        If tsMain.SelectedItem.Key = "vid" Then
            lhWnd = fraVidPlace.hWND
        Else
            lhWnd = fraVidPrev.hWND
            bShowMouse = True
        End If
    End If
End If

If cMedia.Window <> lhWnd Then
    If cMedia.JumphWnd(lhWnd) Then
        '
    Else
        Debug.Print "vid FAILED to jump to " & lhWnd
    End If
End If

gen_UpdateVidPrevVis

cMedia.ShowMouse = bShowMouse

Exit Sub
gen_UpdateVidPos_err:
Main_Err "gen_UpdateVidPos_err."
err.Clear
End Sub

Sub gen_UpdateVidPrevVis()
Dim bVidPrev As Boolean

bVidPrev = False

If Not m_bFullScreen And Not m_bMiniVid Then
    If chkOptDisplayMinivid.Value = 1 Then
        If tsMain.SelectedItem.Key <> "vid" Then
            bVidPrev = cMedia.HasVideo And (cMedia.State <> stStopped)
        End If
    End If
End If

If fraVidPrev.Visible <> bVidPrev Then
    fraVidPrev.Visible = bVidPrev
    Form_Resize
End If
End Sub


Public Sub gen_GoFullscreen(cMon As cMonitor)
If Not m_bFullScreen Then
    Load frmDisp
    frmDisp.SetDisp cMon
    m_bFullScreen = True
    gen_UpdateVidPos
End If
End Sub

Public Sub gen_ShowGoto()
Dim m As cMonitor, pt As POINTAPI, i As Long

If m_bGotoVis Then
    frmGoto.CloseMe
Else
    GetCursorPos pt
    m_cM.Refresh
    For i = 1 To m_cM.MonitorCount
        If m_cM.Monitor(i).Left <= pt.x And m_cM.Monitor(i).Left + m_cM.Monitor(i).Width >= pt.x And _
            m_cM.Monitor(i).Top <= pt.Y And m_cM.Monitor(i).Top + m_cM.Monitor(i).Height >= pt.Y Then
            
            Set m = m_cM.Monitor(i)
        End If
    Next i
    
    frmGoto.Move (m.WorkLeft + (m.WorkWidth / 2)) * Screen.TwipsPerPixelX - (frmGoto.Width / 2), _
           (m.WorkTop + (m.WorkHeight / 2)) * Screen.TwipsPerPixelY - (frmGoto.Height / 2)
    frmGoto.Show
End If
End Sub

Public Function gen_GetShowName(f As String) As String
On Error GoTo gen_GetShowName_err

Dim bRev As Boolean, lLevels As Long
Dim sOut As String

Dim sItem As String

Dim lLen As Long, lCutPont As Long
Dim i As Long, x As Long

If lstPrefDirHide.ListCount > 0 Then
    For i = 0 To lstPrefDirHide.ListCount - 1
        sItem = LCase$(lstPrefDirHide.List(i))
        If sItem = LCase$(Mid$(f, 1, Len(sItem))) Then
            sOut = Mid$(f, Len(sItem) + 1)
            Exit For
        End If
    Next i
End If

If Len(sOut) <= 0 Then
    lLevels = mdb_List.lTrunk
    bRev = mdb_List.bTrunkReverse
    
    If lLevels < 1 Then
        sOut = f
    Else
        lLen = Len(f)
        i = 0
        x = IIf(bRev, lLen, 0)
        
        Do
            x = IIf(bRev, InStrRev(f, "\", x - 1), InStr(x + 1, f, "\"))
            If x = 0 Then Exit Do
            
            lCutPont = x + 1
            i = i + 1
            If i >= lLevels Then Exit Do
        Loop
        
        sOut = Mid$(f, lCutPont)
    End If
End If

If chkOptHideExt.Value = 1 Then
    sOut = RemExtFromPath(sOut)
End If

gen_GetShowName = sOut

Exit Function
gen_GetShowName_err:
Main_Err "gen_GetShowName_err", 1
End Function

Private Sub cbbOptDrawmode_Click()
Dim i As Long, iMode As Integer

iMode = cbbOptDrawmode.ListIndex

sldPlay.DrawMode = iMode

cmdPlayState(0).DrawMode = iMode
cmdPlayState(1).DrawMode = iMode
cmdPlayHistory.DrawMode = iMode
cmdPlayNext.DrawMode = iMode
For i = 0 To cmdMnu.count - 1
    cmdMnu(i).DrawMode = iMode
Next i

cmdVidFull.DrawMode = iMode

cmdLibStop(0).DrawMode = iMode
cmdLibStop(1).DrawMode = iMode

cmdLib.DrawMode = iMode
cmdPl.DrawMode = iMode

cmdLibSearchClose.DrawMode = iMode
End Sub

Private Sub cbbOptErrNoplay_Click()
pref_StartAutoSave
End Sub

Private Sub cbbTabPosition_Click()
Dim k As String
k = tsMain.SelectedItem.Key
tsMain.Placement = cbbTabPosition.ListIndex
tsMain.Tabs(k).Selected = True

Form_Resize

pref_StartAutoSave
End Sub

Private Sub chkOptCuebtm_Click()
pref_StartAutoSave
End Sub

Private Sub chkOptDirLvlsReverse_Click()
mdb_List.bTrunkReverse = IIf(chkOptDirLvlsReverse.Value = 1, True, False)
pref_StartAutoSave
End Sub

Private Sub chkOptDisplayMinivid_Click()
gen_UpdateVidPos
pref_StartAutoSave
End Sub

Private Sub chkOptDisplayMintray_Click()
pref_StartAutoSave
End Sub

Private Sub chkOptDisplayRepDashNewLine_Click()
pref_StartAutoSave
End Sub

Private Sub chkOptHideExt_Click()
pref_StartAutoSave
End Sub

Private Sub chkOptLibcols_Click(Index As Integer)
pref_StartAutoSave
End Sub

Private Sub chkOptLibNum_Click()
pref_StartAutoSave
End Sub

Private Sub cmdLib_Click()
mdbl_RebuildMenu
mnuLibSel.Visible = True
'todo: update "enabled"
PopupMenu mnuLib, vbPopupMenuRightAlign, _
    fraLib.Left + cmdLib.Left + cmdLib.Width, fraLib.Top + cmdLib.Top + cmdLib.Height
End Sub

Private Sub cmdLibOrder_Click()
PopupMenu mnuLibOrder, vbPopupMenuRightAlign, _
    fraLib.Left + cmdLibOrder.Left + cmdLibOrder.Width, fraLib.Top + cmdLibOrder.Top + cmdLibOrder.Height
End Sub

Private Sub cmdLibSearch_Click(Index As Integer)
Dim i As Long, lStart As Long, lEnd As Long, lStep As Long, x As Long, _
    bFound As Boolean, bSecondPass As Boolean

If mdb_List.lCnt < 1 Then Exit Sub
If Len(txtLibSearch.Text) < 1 Then Exit Sub

If Index = 0 Then
    If mdb_List.lIndex >= 0 Then lStart = mdb_List.lIndex + 1 Else lStart = 0
    lEnd = mdb_List.lCnt - 1
    lStep = 1
Else
    If mdb_List.lIndex >= 0 Then lStart = mdb_List.lIndex - 1 Else lStart = mdb_List.lCnt - 1
    lEnd = 0
    lStep = -1
End If

bFound = False
bSecondPass = False

SearchAgain:
For i = lStart To lEnd Step lStep
    If InStr(1, LCase$(mdb_List.Items(i).sFile), LCase$(txtLibSearch.Text)) > 0 Then
        For x = 0 To mdb_List.lCnt - 1
            mdb_List.Items(x).bSel = False
        Next x
        
        mdb_List.Items(i).bSel = True
        mdb_List.lIndex = i
        list_JumpTo 0, i
        bFound = True
        
        Exit For
    End If
Next i

If Not bFound And Not bSecondPass Then
    bSecondPass = True
    If Index = 0 Then
        lStart = 0
    Else
        lStart = mdb_List.lCnt - 1
    End If
    GoTo SearchAgain
End If

If bFound Then mdbl_Redraw
txtLibSearch.BackColor = IIf(bFound, vbWindowBackground, vbRed)
End Sub

Private Sub cmdLibSearchClose_Click()
mnuLibFind_Click
End Sub

Private Sub cmdLibStop_Click(Index As Integer)
bAbort = True
End Sub

Private Sub cmdLibWatchAdd_Click()
On Error GoTo cmdLibWatchAdd_Click_err:

Dim f As String

f = BrowseForFolder("", "select a folder which contains media")
If f = "" Then Exit Sub

lstLibWatch.AddItem f

lblPrefLibFldCnt.Caption = lstLibWatch.ListCount & lblPrefLibFldCnt.Tag

pref_StartAutoSave

Exit Sub
cmdLibWatchAdd_Click_err:
Main_Err "cmdLibWatchAdd_Click_err."
err.Clear
End Sub

Private Sub cmdLibWatchRem_Click()
Dim i As Long

If lstLibWatch.ListCount < 1 Then Exit Sub

For i = lstLibWatch.ListCount - 1 To 0 Step -1
    If lstLibWatch.Selected(i) Then
        lstLibWatch.RemoveItem i
    End If
Next i

lblPrefLibFldCnt.Caption = lstLibWatch.ListCount & lblPrefLibFldCnt.Tag

pref_StartAutoSave
End Sub

Private Sub cmdMnu_AfterRedraw(Index As Integer, lHdc As Long, lW As Long, lH As Long, lForeColour As Long)
Dim i As Long, x As Long, Y As Long

x = (lW / 2) - 8
Y = (lH / 2) - 8

Select Case Index
    Case 0
        i = 5
    
    Case 1 'track current item
        i = IIf(m_TrackCurrentItem, 3, 4)
    
    Case 2 'list mode
        For i = 0 To mnuModeListsI.count - 1
            If mnuModeListsI(i).Checked Then Exit For
        Next i
    
End Select

Select Case Index
    Case 0, 1, 2
        BitBlt lHdc, x, Y, 16, 16, picIconMask(i).hDC, 0, 0, SRCAND
        BitBlt lHdc, x, Y, 16, 16, picIcon(i).hDC, 0, 0, SRCPAINT
    
End Select
End Sub

Private Sub cmdMnu_Click(Index As Integer)
Dim m As Menu

Select Case Index
    Case 0: Set m = mnuOpt
    Case 1: Set m = mnuJump
    Case 2: Set m = mnuModeLists
    Case 3: Set m = mnuMode
End Select

PopupMenu m, vbPopupMenuRightAlign, cmdMnu(Index).Left + cmdMnu(Index).Width, _
    cmdMnu(Index).Top + cmdMnu(Index).Height
End Sub

Private Sub cmdOptFileext_Click()
If MsgBox("reset file extensions list to default?", vbYesNo) <> vbYes Then Exit Sub
txtOptFileext.Text = file_ext_list_def
End Sub

Private Sub cmdOptHk_Click(Index As Integer)
On Error GoTo SetKeyError

Dim lKey As Long, lShift As Long, arrlKeys() As String

If txtOptHk(Index).Tag = "" Then
    If cHk(Index).HasData Then
        If MsgBox("clear this hotkey?", vbYesNo) = vbYes Then
            cHk(Index).ClearKey
            cHk(Index).ClearData
        End If
    End If
Else
    cHk(Index).SetKeyFromData Me.hWND, txtOptHk(Index).Tag
End If

SetStat:
lblOptHk(Index).Caption = IIf(cHk(Index).IsSet, _
    "key set: " & cHk(Index).GetKeyDes & ".", _
    "key not set.")


pref_StartAutoSave

Exit Sub
SetKeyError:
err.Clear
Resume SetStat
End Sub

'this is the big function of doom that processing watch folders, etc. to update the
'  media library.

Private Sub cmdPrefLibMaint_Click()
On Error GoTo cmdOptLibMaint_err

Dim i As Long, e As String

'start by setting some GUI stuff, so the user knows waht is happenning.

lblLibHash.Caption = "running maintenance..."
e = "starting maint."

cmdPrefLibMaint.Enabled = False
cmdPrefLibRunDefScan.Enabled = False
For i = 0 To chkOptLibMaint.count - 1
    chkOptLibMaint(i).Enabled = False
Next i
For i = 0 To chkOptLibMaintR.count - 1
    chkOptLibMaintR(i).Enabled = False
Next i
txtOptFileext.Enabled = False
cmdOptFileext.Enabled = False
cmdLibWatchAdd.Enabled = False
cmdLibWatchRem.Enabled = False
fraNotice(1).Visible = True
Form_Resize
bAbort = False
bAborted = False

DoEvents

Dim x As Long, _
    nRem As Long, nHash As Long, _
    cStream As cBinaryFileStream, cCRC32 As cCRC32, lCRC32 As Long, lOldCRC32 As Long, _
    a As Boolean, b As Boolean

Dim lStart As Long, lEnd As Long, _
    dDateAdded As Date, dDateAddedSet As Boolean, _
    dLastPlayed As Date, dLastPlayedSet As Boolean

Dim mMDB As typMdb 'library object to use for this procedure.

'as a precaution, ensure all playlists are saved.
'  (e.g., if running this crashes terra).
If tmrPlAutoSave.Enabled = True Then pll_WriteAllLists

AddToLog txtLog, "starting library scan."

'step 0: scan watch folders
e = "step 0."
If chkOptLibMaint(0).Value = 1 And lstLibWatch.ListCount > 0 Then
    For i = 0 To lstLibWatch.ListCount - 1
        lblLibHash.Caption = "running maintenance: scanning watch folders... " & i & " of " & lstLibWatch.ListCount - 1 & "."
        lblLibHash.Refresh
        
        mdb_AddDir lstLibWatch.List(i), IIf(chkOptLibMaintR(2).Value = 1, True, False)
    Next i
End If

DoEvents

Dim fRep As New frmScanRes

'step 1: gen hash codes
e = "step 1."
If chkOptLibMaint(1).Value = 1 Then
    mdb_QueryToMdb mdb_DefaultSqlQuery, mMDB 'get temp library object ready.
    
    If mMDB.lCnt < 1 Then GoTo SkipStep1
    
    nRem = 0
    nHash = 0
    
    For i = 0 To mMDB.lCnt - 1
        lblLibHash.Caption = "running maintenance: scanning for un-hashed files... " & i & " of " & mMDB.lCnt - 1 & "."
        lblLibHash.Refresh
        
        If Not fsoMain.FileExists(mMDB.Items(i).sFile) Then
            Debug.Print "item not found: " & mMDB.Items(i).sFile
            
            If chkOptLibMaintR(0).Value = 1 And mMDB.Items(i).bMissing Then
                'item is disabled, and the user is not interested in missing items
                'that are disabled.
                GoTo NextI
            End If
            
            'NOTE:
            'here we can default to "disable item", because if its a duplicate,
            'it will be replaced later, and the new item will default to remove.
            fRep.AddItem mMDB.Items(i).sFile, _
                mMDB.Items(i).lMD5, str_mdb_FileMis, _
                IIf(chkOptLibMaintR(1).Value = 1, trrMarkMissing, trrDelLibRef)
            
            GoTo NextI
        End If
        
        If Not IsMediaFile(mMDB.Items(i).sFile) Then
            'mdb_RemFile mMDB.Items(i).sFile
            
            fRep.AddItem mMDB.Items(i).sFile, mMDB.Items(i).lMD5, _
                str_mdb_FileNotMedia, trrDelLibRef
            
            nRem = nRem + 1
            GoTo NextI
        End If
        
        'todo: make this unicode-compatable
        If (mMDB.Items(i).lMD5 = 0 Or chkOptLibMaint(2).Value = 1) Then
            If IsUnicode(mMDB.Items(i).sFile) Then
                AddToLog txtLog, "unable to hash unicode filename: '" & mMDB.Items(i).sFile & "'"
            Else
                lblLibHash.Caption = "running maintenance: hashing file " & i & " of " & mMDB.lCnt - 1 & "."
                lblLibHash.Refresh
                
                Set cStream = New cBinaryFileStream
                Set cCRC32 = New cCRC32
                
                cStream.File = mMDB.Items(i).sFile
                lCRC32 = cCRC32.GetFileCrc32(cStream)
                
                DoEvents
                
                lOldCRC32 = mMDB.Items(i).lMD5
                
                mMDB.Items(i).lMD5 = lCRC32
                mdb_SetHash mMDB.Items(i).sFile, lCRC32
                
                If lOldCRC32 <> 0 And lOldCRC32 <> lCRC32 Then
                    fRep.AddItem mMDB.Items(i).sFile, lCRC32, _
                        "old crc32: " & Hex$(lOldCRC32), trrDoNothing
                End If
                
                nHash = nHash + 1
            End If
        End If
        
NextI:
        DoEvents
        If bAbort Then GoTo Abort
    Next i
    
    If nRem > 0 Then mdbl_Requery False
    
SkipStep1:
End If

'step 2: check for duplicates and missing files
e = "step 2."
If chkOptLibMaint(3).Value = 1 Then
    mdb_QueryToMdb mdb_DefaultSqlQuery, mMDB 'get temp library object ready.
    
    If mMDB.lCnt < 1 Then GoTo SkipStep2
    
    lblLibHash.Caption = "running maintenance: checking for duplicate hash codes..."
    lblLibHash.Refresh
    
    For i = 0 To mMDB.lCnt - 1
        lblLibHash.Caption = "running maintenance: checking for duplicate hash codes... " & i & " of " & mMDB.lCnt - 1 & "."
        lblLibHash.Refresh
        
        If mMDB.Items(i).lMD5 <> 0 Then
            For x = i + 1 To mMDB.lCnt - 1
                If mMDB.Items(x).lMD5 <> 0 Then
                    If mMDB.Items(i).lMD5 = mMDB.Items(x).lMD5 Then
                        
                        Debug.Print i & " and " & x & " = " & _
                            mMDB.Items(i).lMD5 & "; '" & _
                            mMDB.Items(i).sFile & "' and '" & _
                            mMDB.Items(x).sFile & "'"
                        
                        If mMDB.Items(i).sFile = mMDB.Items(x).sFile Then
                            'a duplication error.  this should not happen,
                            ' but if it does, take care of it selently.
                            mdb_RemFile mMDB.Items(i).sFile
                            Debug.Print "oops: file in lib twice."
                            
                        Else
                            a = fsoMain.FileExists(mMDB.Items(i).sFile)
                            b = fsoMain.FileExists(mMDB.Items(x).sFile)
                            
                            If a And b Then 'both exist
                                fRep.AddItem mMDB.Items(i).sFile, _
                                    mMDB.Items(i).lMD5, _
                                    str_mdb_FileDupFst, trrDoNothing, True
                                
                                fRep.AddItem mMDB.Items(x).sFile, _
                                    mMDB.Items(x).lMD5, _
                                    str_mdb_FileDupSub, trrMoveFile, True
                                
                            ElseIf a <> b Then 'only 1 exists
                                fRep.AddItem mMDB.Items(i).sFile, _
                                    mMDB.Items(i).lMD5, _
                                    IIf(a, str_mdb_FileMisDupExi, str_mdb_FileMisDupMis), _
                                    IIf(a, trrDoNothing, trrDelLibRef), True
                                
                                fRep.AddItem mMDB.Items(x).sFile, _
                                    mMDB.Items(x).lMD5, _
                                    IIf(b, str_mdb_FileMisDupExi, str_mdb_FileMisDupMis), _
                                    IIf(b, trrDoNothing, trrDelLibRef), True
                                
                            Else 'neither exist
                                'NOTE: this is disabled because the option
                                'says "scan for dup's", so if the user wanted
                                'just missing items, they would use the prev. stage.
                                
                                'If there is a need to transfer metadata from a
                                ' 'missing' item to another item that does exist,
                                ' then that will be caught by the previous
                                ' if-statement conditions.
                                
                                'fRep.AddItem mMDB.Items(i).sFile, _
                                    mMDB.Items(i).lMD5, _
                                    str_mdb_FileMis, trrDelLibRef, True
                                
                                'fRep.AddItem mMDB.Items(x).sFile, _
                                    mMDB.Items(x).lMD5, _
                                    str_mdb_FileMis, trrDelLibRef, True
                                
                            End If
                            
                        End If
                        
                    End If
                End If
            Next x
        End If
        
        DoEvents
        If bAbort Then GoTo Abort
    Next i
SkipStep2:
End If

'step 3: check for playability
e = "step 3."
If chkOptLibMaint(4).Value = 1 Then
    mdb_QueryToMdb mdb_DefaultSqlQuery, mMDB 'get temp library object ready.
    
    If mMDB.lCnt < 1 Then GoTo SkipStep3
    
    Dim cPlay As New clsMedia
    
    cPlay.m_RaiseErrors = False
    cPlay.Volume = -10000
    cPlay.Speed = 1
    cPlay.Window = 0
    
    For i = 0 To mMDB.lCnt - 1
        lblLibHash.Caption = "running maintenance: checking for file playability... " & i & " of " & mMDB.lCnt - 1 & "."
        lblLibHash.Refresh
        
        'only test items which have not played before.
        If mMDB.Items(i).lDuration <= 0 Then
            cPlay.FileName = mMDB.Items(i).sFile
            If Not cPlay.Duration > 0 Then
                fRep.AddItem mMDB.Items(i).sFile, mMDB.Items(i).lMD5, "duration <= 0.", trrDoNothing
            Else
                mdb_SetDuration mMDB.Items(i).sFile, cPlay.Duration
            End If
        End If
        
        DoEvents
        If bAbort Then GoTo Abort
    Next i
    
    Set cPlay = Nothing
    
SkipStep3:
End If

AbortHashing:
If bAbort Then GoTo Abort

'step 4: vacume.
e = "step 4."
If chkOptLibMaint(5).Value = 1 Then
    mdb_Query "VACUUM"
End If

If fRep.lvRes.ListItems.count > 0 Then
    lblLibHash.Caption = "running maintenance: querying user..."
    e = "query user."
    fRep.Show 1, Me
    
    If fRep.lRes = 1 Then
        e = "proc user descisions."
        lblLibHash.Caption = "running maintenance: processing users descisions..."
        
        mdb_QueryToMdb mdb_DefaultSqlQuery, mMDB 'get temp library object ready.
        
        'proc chkCountMerge
        'loop through each item we are totalling the counts of.
        For i = 1 To fRep.lvRes.ListItems.count
            'check that this item should be affected
            If fRep.chkCountMerge.Value = 1 And fRep.lvRes.ListItems(i).Tag <> trrDoNothing Then
                
                'we are going to total for all items with a given crc.
                lStart = 0
                lEnd = 0
                lCRC32 = Val("&H" & fRep.lvRes.ListItems(i).SubItems(1) & "&")
                dDateAdded = Empty
                dDateAddedSet = False
                dLastPlayed = Empty
                dLastPlayedSet = False
                
                'this will only work with items that have been crc'ed
                If lCRC32 <> 0 Then
                    'sum cnt.s of items with a given crc.
                    'dDateAdded is set to the oldest date.
                    For x = 0 To mMDB.lCnt - 1
                        If mMDB.Items(x).lMD5 = lCRC32 Then
                            lStart = lStart + mMDB.Items(x).lStartCnt
                            lEnd = lEnd + mMDB.Items(x).lEndCnt
                            
                            If Not dDateAddedSet Then
                                dDateAdded = mMDB.Items(x).dAdded
                                dDateAddedSet = True
                            ElseIf DateDiff("s", dDateAdded, mMDB.Items(x).dAdded) < 0 Then
                                dDateAdded = mMDB.Items(x).dAdded
                            End If
                            
                            If Not dLastPlayedSet Then
                                dLastPlayed = mMDB.Items(x).dLastPlay
                                dLastPlayedSet = True
                            ElseIf DateDiff("s", dLastPlayed, mMDB.Items(x).dLastPlay) > 0 Then
                                dLastPlayed = mMDB.Items(x).dLastPlay
                            End If
                        End If
                    Next x
                    
                    'put the total count and dadded date back for all items
                    'with the crc.
                    '(we will remove dup.s in a moment).
                    For x = 0 To mMDB.lCnt - 1
                        If mMDB.Items(x).lMD5 = lCRC32 Then
                            If fsoMain.FileExists(mMDB.Items(x).sFile) Then
                                mMDB.Items(x).lStartCnt = lStart
                                mMDB.Items(x).lEndCnt = lEnd
                                mdb_SetPlaybackCnt mMDB.Items(x).sFile, lStart, lEnd
                                
                                mMDB.Items(x).dAdded = dDateAdded
                                mdb_SetdAdded mMDB.Items(x).sFile, dDateAdded
                                
                                mMDB.Items(x).dLastPlay = dLastPlayed
                                mdb_SetdLastPlayed mMDB.Items(x).sFile, dLastPlayed
                            End If
                        End If
                    Next x
                End If
            End If
        Next i
        
        'proc the rest...
        For i = 1 To fRep.lvRes.ListItems.count
            Select Case fRep.lvRes.ListItems(i).Tag
                Case trrDoNothing
                    'the name says it all...
                
                Case trrDelLibRef
                    mdb_RemFile fRep.lvRes.ListItems(i).Text
                    nRem = nRem + 1
                    Debug.Print "removed from lib: '" & fRep.lvRes.ListItems(i).Text & "'."
                
                Case trrMoveFile
                    'todo: rem from lib, then from hd.
                    mdb_RemFile fRep.lvRes.ListItems(i).Text
                    AddToLog txtLog, "todo: move '" & fRep.lvRes.ListItems(i).Text & "' to the bin."
                    MsgBox "todo: move '" & fRep.lvRes.ListItems(i).Text & "' to the bin."
                
                Case trrMarkMissing 'find the library item, and flag it.
                    mdb_SetMissing fRep.lvRes.ListItems(i).Text, True
                
            End Select
        Next i
    End If
End If


e = "cleaning up."
Unload fRep

mdbl_Requery False

Abort:
CleanUp:
cmdPrefLibMaint.Enabled = True
cmdPrefLibRunDefScan.Enabled = True
For i = 0 To chkOptLibMaint.count - 1
    chkOptLibMaint(i).Enabled = True
Next i
For i = 0 To chkOptLibMaintR.count - 1
    chkOptLibMaintR(i).Enabled = True
Next i
txtOptFileext.Enabled = True
cmdOptFileext.Enabled = True
cmdLibWatchAdd.Enabled = True
cmdLibWatchRem.Enabled = True
fraNotice(1).Visible = False
Form_Resize

AddToLog txtLog, "library scan result:  " & nHash & " items hashed.  " & nRem & " items removed."

bAborted = True

Exit Sub
cmdOptLibMaint_err:
bAborted = True
Main_Err "cmdOptLibMaint_Click - " & e
err.Clear
GoTo CleanUp
End Sub

Private Sub cmdOptSave_Click()
pref_Write
End Sub

Private Sub cmdPl_Click()
If m_lCurrentPl < 0 Then Exit Sub

pll_RebuildMenu
mnuPlSep2.Visible = True
mnuPlSelect.Visible = True
mnuPlSort.Visible = True
mnuPlSep3.Visible = True
mnuPlRep.Visible = True
mnuPlSel.Visible = True

PopupMenu mnuPl, vbPopupMenuRightAlign, _
    fraPl.Left + cmdPl.Left + cmdPl.Width, _
    fraPl.Top + cmdPl.Top + cmdPl.Height
End Sub

Private Sub cmdPlayHistory_Click()
history_RebuildMenu
PopupMenu mnuHistory, vbPopupMenuLeftAlign, cmdPlayHistory.Left, _
    cmdPlayHistory.Top + cmdPlayHistory.Height
End Sub

Private Sub cmdPlayNext_Click()
pb_NextFile
End Sub

Private Sub cmdPlayNext_ClickSide()
PopupMenu mnuNxsp, vbPopupMenuLeftAlign, cmdPlayNext.Left, _
    cmdPlayNext.Top + cmdPlayNext.Height
End Sub

Public Sub cmdPlayState_Click(Index As Integer)
Select Case Index
    Case 0: pb_SetPlayState 0
    Case 1: If pb_GetPlayState = 1 Then pb_SetPlayState 2 Else pb_SetPlayState 1
End Select
End Sub

Private Sub cmdPrefClose_Click()
mnuOptTab_Click
End Sub

Private Sub cmdPrefDirHideAdd_Click()
On Error GoTo cmdPrefDirHideAdd_Click_err:

Dim a As String

a = InputBox("enter a string to hide from the front of paths:")
If a = "" Then Exit Sub

lstPrefDirHide.AddItem a

pref_StartAutoSave

Exit Sub
cmdPrefDirHideAdd_Click_err:
Main_Err "cmdPrefDirHideAdd_Click_err."
err.Clear
End Sub

Private Sub cmdPrefDirHideRem_Click()
Dim i As Long

If lstPrefDirHide.ListCount < 1 Then Exit Sub

For i = lstPrefDirHide.ListCount - 1 To 0 Step -1
    If lstPrefDirHide.Selected(i) Then
        lstPrefDirHide.RemoveItem i
    End If
Next i

pref_StartAutoSave
End Sub

Private Sub cmdPrefLibFoldersDone_Click()
lstPrefPages.ListIndex = 5
End Sub

Private Sub cmdPrefLibGotoFolders_Click()
lstPrefPages.ListIndex = 6
End Sub

Private Sub cmdPrefLibGotoScan_Click()
lstPrefPages.ListIndex = 8
End Sub

Private Sub cmdPrefLibRunDefScan_Click()
If Not cmdPrefLibMaint.Enabled Then Exit Sub

chkOptLibMaint(0).Value = 1
chkOptLibMaint(1).Value = 1
chkOptLibMaint(2).Value = 0
chkOptLibMaint(3).Value = 1
chkOptLibMaint(4).Value = 0
chkOptLibMaint(5).Value = 1

chkOptLibMaintR(0).Value = 1
chkOptLibMaintR(1).Value = 1
chkOptLibMaintR(2).Value = 1

cmdPrefLibMaint_Click
End Sub

Sub cmdVidFull_Click()
Dim i As Long

If m_bFullScreen Then
    'remove:
    m_bFullScreen = False
    gen_UpdateVidPos
    Unload frmDisp
    Set frmDisp = Nothing
Else
    m_cM.Refresh
    
    If mnuDispFullI.count > 1 Then
        For i = mnuDispFullI.count - 1 To 1 Step -1
            Unload mnuDispFullI(i)
        Next i
    End If
    
    For i = 1 To m_cM.MonitorCount
        If i > mnuDispFullI.count Then Load mnuDispFullI(i - 1)
        
        mnuDispFullI(i - 1).Caption = "on display " & i '& " (" & m_cM.Monitor(i).Name & ")"
        mnuDispFullI(i - 1).Enabled = True
        mnuDispFullI(i - 1).Visible = True
    Next i
    
    PopupMenu mnuDisp, vbPopupMenuRightAlign, _
        fraVid.Left + cmdVidFull.Left + cmdVidFull.Width, _
        fraVid.Top + cmdVidFull.Top + cmdVidFull.Height
End If
End Sub

Public Sub cmdVidMinivid_Click()
If m_bMiniVid Then
    m_bMiniVid = False
    gen_UpdateVidPos
    Unload frmMiniVid
    Set frmMiniVid = Nothing
Else
    Load frmMiniVid
    frmMiniVid.SetPosition
    'If m_bFullScreen Then cmdVidFull_Click 'check for and remove full screen
    m_bMiniVid = True
    gen_UpdateVidPos
End If

pref_StartAutoSave
End Sub

Private Sub Form_Activate()
Form_Resize
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Long

If (Shift And vbAltMask) Then
    Select Case KeyCode
        Case vbKeyHome
            list_JumpToCurrent
        
        Case vbKeyZ 'stop
            cmdPlayState_Click 0
        
        Case vbKeyX 'play / pause
            cmdPlayState_Click 1
        
        Case vbKeyC 'next
            cmdPlayNext_Click
        
        Case vbKeyG 'goto
            gen_ShowGoto
        
    End Select
End If

If KeyCode = vbKeyTab And (Shift And vbCtrlMask) Then
    If (Shift And vbShiftMask) Then
        i = tsMain.SelectedItem.Index - 1
        If i < 1 Then i = tsMain.Tabs.count
        tsMain.Tabs(i).Selected = True
    Else
        i = tsMain.SelectedItem.Index + 1
        If i > tsMain.Tabs.count Then i = 1
        tsMain.Tabs(i).Selected = True
    End If
End If

If (Shift And vbCtrlMask) Then
    Select Case KeyCode
        Case vbKeyHome
            list_JumpToCurrent
        
        Case vbKeyF
            If fraLib.Visible Then
                If mnuLibFind.Checked Then
                    If Screen.ActiveControl = txtLibSearch Then
                        mnuLibFind.Checked = False
                    Else
                        'do nothing.
                    End If
                Else
                    mnuLibFind.Checked = True
                End If
                mnuLibFind.Checked = Not mnuLibFind.Checked
                mnuLibFind_Click
            End If
        
        Case vbKeyC 'ctrl+c
            If Len(cMedia.FileName) > 0 Then
                Clipboard.Clear
                Clipboard.SetText cMedia.FileName, vbCFText
                Clipboard.SetText cMedia.FileName, vbCFRTF
            End If
        
    End Select
End If

If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyF12: cmdMnu(0).SimClick
        Case vbKeyF11: cmdMnu(1).SimClick
        Case vbKeyF10: cmdMnu(2).SimClick
        Case vbKeyF9:  cmdMnu(3).SimClick
        
        Case vbKeyF5:  cmdPlayState(0).SimClick
        Case vbKeyF6:  cmdPlayState(1).SimClick
        Case vbKeyF7:  cmdPlayHistory.SimClick
        Case vbKeyF8:  cmdPlayNext.SimClick
        
        Case vbKeyF1
            Select Case tsMain.SelectedItem.Key
                Case "vid": cmdVidFull.SimClick
                Case "lib": cmdLib.SimClick
                Case "arc": cmdPl.SimClick
                Case Else: If Mid$(tsMain.SelectedItem.Key, 1, 2) = "pl" Then cmdPl.SimClick
            End Select
        
        Case vbKeyF2
            Select Case tsMain.SelectedItem.Key
                Case "vid": cmdVidMinivid.SimClick
                Case "lib": cmdLibOrder.SimClick
            End Select
        
    End Select
End If

If (Shift And vbCtrlMask) Then
    Select Case KeyCode
        Case vbKeyF8:  cmdPlayNext.SimClick True
    End Select
End If
End Sub

Private Sub Form_Load()
On Error GoTo frmMain_Form_Load_err

Dim i As Long, a, _
    e As String, _
    x As Long, Y As Long, c As Long, _
    arr() As String

Const lTrns = &HFF00FF

'Dim lTimer As Long
'lTimer = GetTickCount

e = "0"

'logging
AddToLog txtLog, "terra main window loaded."
If Len(Command) > 0 Then
    AddToLog txtLog, "cmd=" & Command
End If

'start-up messages
If pref_NoHook Then
    AddToLog txtLog, "--nohook switch detected.  system hooks will be disabled."
    MsgBox "--nohook switch detected.  system hooks will be disabled.", vbInformation
End If

If pref_Debug Then
    AddToLog txtLog, "--debug switch detected.  debug menu is now visible."
    MsgBox "--debug switch detected.  debug menu is now visible.", vbInformation
End If

e = "0.1"

'init varables
bAbort = False
bAborted = True

'general caption stuff
mnuOptSpeed.Tag = mnuOptSpeed.Caption

'debug menu
mnuDebug.Visible = IIf(IsCompiled, pref_Debug, True)

e = "init mdb structure"

'next list
For i = 0 To mnuNxspI.count - 1
    mnuNxspI(i).Caption = Replace(mnuNxspI(i).Caption, "|", vbTab)
Next i

'media lib
Set m_ListDc = New pcMemDC
mnuLibFind.Caption = mnuLibFind.Caption & vbTab & "ctrl + f"
mnuLibGoto.Caption = mnuLibGoto.Caption & vbTab & "alt + g"
mnuLibAddcue.Caption = mnuLibAddcue.Caption & vbTab & "q"

e = "init pl stuff"

'play lists
mnuPlNew.Caption = mnuPlNew.Caption & vbTab & "ctrl + n"
mnuPlName.Caption = mnuPlName.Caption & vbTab & "ctrl + e"
mnuPlSelAll.Caption = mnuPlSelAll.Caption & vbTab & "ctrl + a"
mnuPlMove(0).Caption = mnuPlMove(0).Caption & vbTab & "ctrl + pg-up"
mnuPlMove(1).Caption = mnuPlMove(1).Caption & vbTab & "ctrl + pg-dn"
mnuPlRemitm.Caption = mnuPlRemitm.Caption & vbTab & "del"
mnuPlAddcue.Caption = mnuPlAddcue.Caption & vbTab & "q"
mnuPlSelMis.Caption = mnuPlSelMis.Caption & vbTab & "ctrl + m"
mnuPlMoveTabI(0).Caption = mnuPlMoveTabI(0).Caption & vbTab & "ctrl + l"
mnuPlMoveTabI(1).Caption = mnuPlMoveTabI(1).Caption & vbTab & "ctrl + r"

m_lCurrentPl = -1
ReDim mdb_PL(0)
Set pl_ListDc = New pcMemDC
pl_SetAsNew mdb_PL(0), "list0"

e = "init cue"
'cue
Set cue_ListDc = New pcMemDC
pl_SetAsNew mdb_Cue, "cue"

'history
pl_SetAsNew mdb_History, "history"

e = "init pb stuff"
'playback
pb_lItemSource = 0 'default to media library
pb_llItemIndex = -1
pb_Speed = 1
pb_RetryCount = 0
'cMedia.Window = fraVidPlace.hwnd
m_bFullScreen = False
m_bMiniVid = False
gen_UpdateVidPos

e = "init gui"
'gui=============================================
gdi_Main_MakeObjects hDC, Font
If IsCompiled And Not pref_NoHook Then AddMenuItem Me.hWND

'tabs
tsMain.ImageList = imlTabs
With tsMain.Tabs
    .Item("vid").Image = "vid"
    .Item("lib").Image = "lib"
    .Item("arc").Image = "arc"
End With
pll_RebuildTabs

'pref
lstPrefPages.Clear
For i = 0 To fraPrefPage.count - 1
    lstPrefPages.AddItem fraPrefPage(i).Tag & fraPrefPage(i).Caption
Next i
lstPrefPages.ListIndex = 0

cbbTabPosition.Clear
arr = Split("top|bottom|left|right", "|")
For i = 0 To UBound(arr)
    cbbTabPosition.AddItem arr(i)
Next i

cbbOptDrawmode.Clear
arr = Split("internal style.|internal bitmaps.", "|")
For i = 0 To UBound(arr)
    cbbOptDrawmode.AddItem arr(i)
Next i

cbbOptColourscheme.Clear
arr = Split("system colours.", "|")
For i = 0 To UBound(arr)
    cbbOptColourscheme.AddItem arr(i)
Next i

cbbOptErrNoplay.Clear
arr = Split("stop and display an error message.|ignore the problem and try another file.", "|")
For i = 0 To UBound(arr)
    cbbOptErrNoplay.AddItem arr(i)
Next i

'init icons======================================
e = "building icon masks"
For i = 0 To picIcon.count - 1
    picIconMask(i).Move picIconMask(i).Left, picIconMask(i).Top, _
        picIcon(i).Width, picIcon(i).Height
    
    For x = 0 To picIcon(i).ScaleWidth - 1
        For Y = 0 To picIcon(i).ScaleHeight - 1
            If GetPixel(picIcon(i).hDC, x, Y) = lTrns Then
                c = &HFFFFFF
            Else
                c = &H0&
            End If
            SetPixel picIconMask(i).hDC, x, Y, c
        Next Y
    Next x
    picIcon(i).BackColor = &H0
Next i

'settings========================================
e = "init hotkeys"
'hotkeys
For i = 0 To UBound(cHk)
    Set cHk(i) = New clsHotKey
Next i

e = "starting library"
mdbl_Requery False

e = "loading preferences"
pref_Read

e = "loading playlists"
pll_ReadAllLists

'active tab:
e = "setting active tab"
a = GetFromIniEx("view", "activetab", "lib", file_INI)
e = "setting active tab to " & a
If a = "sys" Then
    mnuOptTab_Click
ElseIf a = "cue" Then
    a = "lib"
End If
If Not tsMain.Tabs(a) Is Nothing Then Set tsMain.SelectedItem = tsMain.Tabs(a)

'====================================================================
e = "rebuilding mdb"
mdbl_Rebuild True, True

e = "hooking appcom"
If IsCompiled And Not pref_NoHook Then AppCom_Hook Me.hWND

'Debug.Print "frmMain_Load time = " & GetTickCount - lTimer

'====================================================================
Exit Sub
frmMain_Form_Load_err:
Main_Err "sub frmMain_Load - " & e
err.Clear
End Sub

Private Sub Form_Resize()
On Error GoTo Form_Resize_err

Dim i As Long, Y As Long, _
    fraCue_h As Long, fraVidPrev_w As Long, sldPlay_h As Long, fraNotice_h As Long, _
    l As Long, t As Long, w As Long, h As Long

If WindowState = vbMinimized Then
    If chkOptDisplayMintray.Value = 1 Then tray_Min
    Exit Sub
End If

'save the current window state, as we can't get it once minimize has been called.
m_WsBeforeMin = WindowState

'arange controls
For i = 0 To cmdPlayState.count - 1
    cmdPlayState(i).Move cmdPlayState(i).Width * i + lGp, lGp
Next i
cmdPlayHistory.Move cmdPlayState(cmdPlayState.count - 1).Left + _
    cmdPlayState(cmdPlayState.count - 1).Width, lGp

cmdPlayNext.Move cmdPlayHistory.Left + cmdPlayHistory.Width, lGp

lblPlayInfo.Move cmdPlayNext.Left + cmdPlayNext.Width + lSp, lSp

'note: 27 pixels is a slightly made up number but seems to work well.
sldPlay_h = 27 * Screen.TwipsPerPixelY

fraVidPrev.Move ScaleWidth - fraVidPrev.Width, 0, _
    (cmdPlayState(0).Top + cmdPlayState(0).Height + sldPlay_h) * (4 / 3), _
    cmdPlayState(0).Top + cmdPlayState(0).Height + sldPlay_h
If fraVidPrev.Visible Then
    fraVidPrev_w = fraVidPrev.Width + lGp
Else
    fraVidPrev_w = 0
End If

'slider
sldPlay.Move 0, cmdPlayState(0).Top + cmdPlayState(0).Height, _
    ScaleWidth - fraVidPrev_w, sldPlay_h

'menu buttons
For i = 0 To cmdMnu.count - 1
    If i = 0 Then
        cmdMnu(i).Move ScaleWidth - fraVidPrev_w - cmdMnu(0).Width - lGp, lGp
    Else
        cmdMnu(i).Move cmdMnu(i - 1).Left - cmdMnu(i).Width, lGp
    End If
Next i

'cue
Const lMaxH As Long = 1000
'Const lItemExtH = 60
Dim lItemExtH As Long
lItemExtH = lblCueStat.Height + lGp * 2 + 60
If mdb_Cue.lCnt > 0 And tsMain.SelectedItem.Key <> "cue" And chkOptCuebtm.Value = 1 Then
    fraCue_h = (mdb_Cue.lCnt * mdbl_ItmH) * Screen.TwipsPerPixelY + lItemExtH
    If fraCue_h > lMaxH Then fraCue_h = lMaxH
    fraCue.Move 0, ScaleHeight - (fraCue_h), ScaleWidth, fraCue_h
Else
    fraCue_h = 0
End If

'notice boxes
fraNotice_h = 0
For i = 0 To fraNotice.count - 1
    If fraNotice(i).Visible Then
        fraNotice_h = fraNotice_h + fraNotice(i).Height + lGp
    End If
Next i
Y = ScaleHeight - fraNotice_h - fraCue_h
For i = 0 To fraNotice.count - 1
    If fraNotice(i).Visible Then
        fraNotice(i).Move 0, Y + lGp, ScaleWidth
        Y = Y + fraNotice(i).Height + lGp
    End If
Next i

'form body
t = sldPlay.Top + sldPlay.Height + lGp
tsMain.Move 0, t, ScaleWidth, ScaleHeight - fraNotice_h - fraCue_h - t

l = tsMain.ClientLeft
t = tsMain.ClientTop
w = tsMain.ClientWidth
h = tsMain.ClientHeight

If tsMain.SelectedItem.Key = "vid" Then
    fraVid.Move l, t, w, h
ElseIf tsMain.SelectedItem.Key = "sys" Then
    fraPref.Move l, t, w, h
ElseIf tsMain.SelectedItem.Key = "lib" Then
    fraLib.Move l, t, w, h
ElseIf tsMain.SelectedItem.Key = "cue" Then
    fraCue.Move l, t, w, h
ElseIf tsMain.SelectedItem.Key = "arc" Or Mid$(tsMain.SelectedItem.Key, 1, 2) = "pl" Then
    fraPl.Move l, t, w, h
End If

Form_Resize_err:
err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Form_Unload_err
'check we are ready to end=====================================================
bAbort = True
'Do Until bAborted
    DoEvents
'Loop

'save settings=================================================================
pref_Write
If tmrPlAutoSave.Enabled = True Then pll_WriteAllLists

'clean up======================================================================
AppCom_Unhook

Set m_ListDc = Nothing
Set pl_ListDc = Nothing
Set cue_ListDc = Nothing

'end===========================================================================
main_end

Exit Sub
Form_Unload_err:
Main_Err "Form_Unload_err."
err.Clear
End Sub

Private Sub fraCue_AfterPaint()
lblCueStat.Refresh
End Sub

Private Sub fraCue_AfterResize(lW As Long, lH As Long)
Dim lTopGap As Long
lTopGap = lblCueStat.Height + lGp * 2
lblCueStat.Move lGp, lGp
picCue.Move 0, lTopGap, lW - vsbCue.Width - lGpBtwnLstAndSB, lH - lTopGap
vsbCue.Move lW - vsbCue.Width, picCue.Top, vsbCue.Width, picCue.Height
cue_Redraw
End Sub

Private Sub fraLib_AfterPaint()
lblDbStat.Refresh
End Sub

Private Sub fraLib_AfterResize(lW As Long, lH As Long)
On Error Resume Next

Dim fraLibSearch_h As Long

fraLibSearch_h = IIf(mnuLibFind.Checked, fraLibSearch.Height, 0)

cmdLib.Move lW - lGp - cmdLib.Width, lGp
cmdLibOrder.Move cmdLib.Left - cmdLibOrder.Width - lGp, lGp
lblDbStat.Move lGp, lGp + (cmdLib.Height / 2) - (lblDbStat.Height / 2)

picMdbLst.Move 0, cmdLib.Top + cmdLib.Height + lGp, _
    lW - vsbMdbLst.Width - lGpBtwnLstAndSB, _
    lH - cmdLib.Height - (lGp * 2) - fraLibSearch_h

vsbMdbLst.Move picMdbLst.Left + picMdbLst.Width + lGpBtwnLstAndSB, picMdbLst.Top, _
    vsbMdbLst.Width, picMdbLst.Height

fraLibSearch.Move 0, picMdbLst.Top + picMdbLst.Height, lW

mdbl_Rebuild True, True

On Error GoTo 0
End Sub

Private Sub fraLibSearch_AfterResize(lW As Long, lH As Long)
cmdLibSearchClose.Move cmdLibSearchClose.Left, (lH / 2) - (cmdLibSearchClose.Height / 2)
End Sub

Private Sub fraNotice_AfterPaint(Index As Integer)
lblLibHash.Refresh
End Sub

Private Sub fraPl_AfterPaint()
lblPlStat.Refresh
End Sub

Private Sub fraPl_AfterResize(lW As Long, lH As Long)
Dim lArc As Long

lblPlStat.Move lGp, lGp + (cmdPl.Height / 2) - (lblPlStat.Height / 2)
cmdPl.Move lW - cmdPl.Width - lGp - lArc, lGp

If tsMain.SelectedItem.Key = "arc" Then
    lArc = lstPlArc.Width
    lstPlArc.Move 0, cmdPl.Top + cmdPl.Height + lGp, lstPlArc.Width, lH - cmdPl.Height - lGp * 2
Else
    lArc = 0
End If

picPl.Move lArc, cmdPl.Top + cmdPl.Height + lGp, lW - vsbPl.Width - lArc - lGpBtwnLstAndSB, lH - cmdPl.Height - lGp * 2
vsbPl.Move picPl.Left + picPl.Width + lGpBtwnLstAndSB, picPl.Top, vsbMdbLst.Width, picPl.Height

pll_RedrawCurrent
End Sub

Private Sub fraPref_AfterResize(lW As Long, lH As Long)
On Error GoTo SkipRedraw
Dim i As Long

lstPrefPages.Move 0, 0, lstPrefPages.Width, _
    lH - cmdOptSave.Height - cmdPrefClose.Height - lGp * 2
cmdOptSave.Move lstPrefPages.Left, lstPrefPages.Top + lstPrefPages.Height + lGp, _
    lstPrefPages.Width
cmdPrefClose.Move lstPrefPages.Left, cmdOptSave.Top + cmdOptSave.Height + lGp, _
    lstPrefPages.Width

For i = 0 To fraPrefPage.count - 1
    fraPrefPage(i).Move lstPrefPages.Left + lstPrefPages.Width + lGp, _
        0, lW - lstPrefPages.Width - lGp, lH
Next i

SkipRedraw:
err.Clear
End Sub

Private Sub fraPrefPage_AfterPaint(Index As Integer)
Dim i As Long

Select Case Index
    Case 0 'display
        lblPrefDisp(0).Refresh
        lblPrefDisp(1).Refresh
    
    Case 1 'lists
        lblPrefLists(0).Refresh
    
    Case 2 'names
        lblPrefNames(0).Refresh
        lblPrefNames(1).Refresh
    
    Case 3 'gui
        lblPrefGui(0).Refresh
        lblPrefGui(1).Refresh
    
    Case 4 'hot keys
        For i = 0 To lblPrefHk.count - 1
            lblPrefHk(i).Refresh
        Next i
        For i = 0 To lblOptHk.count - 1
            lblOptHk(i).Refresh
        Next i
    
    Case 5 'lib home
        lblPrefLibPage(0).Refresh
        lblPrefLibPage(1).Refresh
        lblPrefLibFldCnt.Refresh
    
    Case 7 'file type
        lblPrefLibFileTyp(0).Refresh
        lblPrefLibFileTyp(1).Refresh
    
    Case 8
        lblPrefLibScan(0).Refresh
        lblPrefLibScan(1).Refresh
        lblPrefLibScan(2).Refresh
    
    Case 9 'PL
        lblPrefPl(0).Refresh
    
    Case 10 'err
        lblPrefErr(0).Refresh
    
End Select
End Sub

Private Sub fraPrefPage_AfterResize(Index As Integer, lW As Long, lH As Long)
On Error GoTo err:

Select Case Index
    Case 0 'display
        cbbTabPosition.Move cbbTabPosition.Left, cbbTabPosition.Top, _
            lW - cbbTabPosition.Left - lSp
    
    Case 1 'display.lists
        
    
    Case 2 'names
        cmdPrefDirHideAdd.Move lSp, cmdPrefDirHideAdd.Top, (lW - lSp * 3) / 2
        cmdPrefDirHideRem.Move cmdPrefDirHideAdd.Left + cmdPrefDirHideAdd.Width + lSp, _
            cmdPrefDirHideAdd.Top, cmdPrefDirHideAdd.Width
        lstPrefDirHide.Move lstPrefDirHide.Left, lstPrefDirHide.Top, lW - lSp - lstPrefDirHide.Left
        sldOptDirLvls.Move sldOptDirLvls.Left, sldOptDirLvls.Top, lW - sldOptDirLvls.Left - lSp
    
    Case 3 'display.gui
        cbbOptDrawmode.Move cbbOptDrawmode.Left, cbbOptDrawmode.Top, _
            lW - cbbOptDrawmode.Left - lSp
        cbbOptColourscheme.Move cbbOptColourscheme.Left, cbbOptColourscheme.Top, _
            lW - cbbOptColourscheme.Left - lSp
    
    Case 4 'hotkeys
'        lblCap(6).Move lblCap(6).Left, lblCap(6).Top, lW - lblCap(6).Left - lGp
    
    Case 5 'lib home
        lblPrefLibPage(0).Move lSp, lblPrefLibPage(0).Top, lW - lSp * 2
        lblPrefLibPage(0).AutoSize = True
        cmdPrefLibGotoFolders.Move lSp * 2, _
            lblPrefLibPage(0).Top + lblPrefLibPage(0).Height + lSp
        lblPrefLibFldCnt.Move cmdPrefLibGotoFolders.Left + cmdPrefLibGotoFolders.Width + lSp * 3, cmdPrefLibGotoFolders.Top
        lblPrefLibPage(1).Move lSp, _
            cmdPrefLibGotoFolders.Top + cmdPrefLibGotoFolders.Height + lSp, _
            lW - lSp * 2
        lblPrefLibPage(1).AutoSize = True
        cmdPrefLibRunDefScan.Move lSp * 2, _
            lblPrefLibPage(1).Top + lblPrefLibPage(1).Height + lSp
        cmdPrefLibGotoScan.Move lSp * 2, _
            cmdPrefLibRunDefScan.Top + cmdPrefLibRunDefScan.Height + lSp
    
    Case 6 'watch folders
        cmdLibWatchAdd.Move lSp, cmdLibWatchAdd.Top, (lW - lSp * 3) / 2
        cmdLibWatchRem.Move cmdLibWatchAdd.Left + cmdLibWatchAdd.Width + lSp, cmdLibWatchRem.Top, cmdLibWatchAdd.Width
        cmdPrefLibFoldersDone.Move lW - cmdPrefLibFoldersDone.Width - lSp, _
            lH - cmdPrefLibFoldersDone.Height - lGp
        lstLibWatch.Move lSp, lstLibWatch.Top, lW - lSp * 2, _
            cmdPrefLibFoldersDone.Top - lstLibWatch.Top - lSp
    
    Case 7 'file types
        lblPrefLibFileTyp(0).Move lSp, lblPrefLibFileTyp(0).Top, lW - lSp * 2
        lblPrefLibFileTyp(0).AutoSize = True
        
        lblPrefLibFileTyp(1).Move lSp, lblPrefLibFileTyp(0).Top + lblPrefLibFileTyp(0).Height + lSp
        
        txtOptFileext.Move txtOptFileext.Left, lblPrefLibFileTyp(1).Top + lblPrefLibFileTyp(1).Height + lSp, _
            lW - txtOptFileext.Left - lSp
        cmdOptFileext.Move lW - lSp - cmdOptFileext.Width, cmdOptFileext.Top
        'lblCap(15).Move lSp, txtOptFileext.Top - lblCap(15).Height - lSp
        
        'txtOptFileext.Move txtOptFileext.Left, txtOptFileext.Top, lW - txtOptFileext.Left - lSp
        'cmdOptFileext.Move txtOptFileext.Left + txtOptFileext.Width - cmdOptFileext.Width

    Case 8 'lib scan
        
    
    Case 9 'playlists
        sldOptPlAutosave.Move lSp, sldOptPlAutosave.Top, lW - lSp * 2
    
    Case 10 'error reporting
        cbbOptErrNoplay.Move cbbOptErrNoplay.Left, cbbOptErrNoplay.Top, _
            lW - cbbOptErrNoplay.Left - lSp
    
    Case 11 'log
        txtLog.Move lSp, txtLog.Top, lW - lSp * 2, lH - txtLog.Top - lSp
    
End Select

err:
err.Clear
End Sub

Private Sub fraVid_AfterPaint()
lblVidStat.Refresh
End Sub

Private Sub fraVid_AfterResize(lW As Long, lH As Long)
cmdVidFull.Move lW - lGp - cmdVidFull.Width, lGp
cmdVidMinivid.Move cmdVidFull.Left - cmdVidMinivid.Width, lGp
lblVidStat.Move lGp, lGp + (cmdVidFull.Height / 2) - (lblVidStat.Height / 2)
fraVidPlace.Move lGp, cmdVidFull.Top + cmdVidFull.Height + lGp, _
    lW - lGp * 2, lH - cmdVidFull.Top - cmdVidFull.Height - lGp * 2
End Sub

Private Sub fraVidPlace_AfterRedraw(lHdc As Long, lW As Long, lH As Long)
pb_DrawCurrentItem lHdc, lW, lH, GetSysColor(COLOR_BTNTEXT)
End Sub

Private Sub fraVidPrev_DblClick()
frmMain.gen_GoFullscreen m_cM.MonitorForWindow(hWND)
End Sub

Private Sub lblDbStat_Click()
mdbl_Requery
End Sub

Private Sub lblDbStat_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If GetCursor <> IDC_HAND Then SetCursor LoadCursor(0, IDC_HAND)
End Sub

Private Sub lblPlayInfo_Click()
list_JumpToCurrent
End Sub

Private Sub lblPlayInfo_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If GetCursor <> IDC_HAND Then SetCursor LoadCursor(0, IDC_HAND)
End Sub

Private Sub lblPlStat_Click()
pl_lAutoSaveCntDn = 0
tmrPlAutoSave_Timer
End Sub

Private Sub lblPlStat_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If GetCursor <> IDC_HAND Then SetCursor LoadCursor(0, IDC_HAND)
End Sub

Private Sub lstPlArc_Click()
picPl.Enabled = True
m_lCurrentPl = lstPlArc.ItemData(lstPlArc.ListIndex)
pll_RedrawCurrent
End Sub

Private Sub lstPrefPages_Click()
Dim i As Long
For i = 0 To lstPrefPages.ListCount - 1
    If i = lstPrefPages.ListIndex Then
        fraPrefPage(i).Visible = True
    Else
        fraPrefPage(i).Visible = False
    End If
Next i
End Sub

Private Sub mnuDebugBumpvidwindow_Click()
cMedia.BumpVidWindow
End Sub

Private Sub mnuDebugDiv0_Click()
'On Error GoTo mnuDebugDiv0_Click_err

MsgBox (1 / 0)

'Exit Sub
'mnuDebugDiv0_Click_err:
'Main_Err "div by 0!"
End Sub

Private Sub mnuDebugRehook_Click()
AppCom_Unhook
AppCom_Hook Me.hWND
End Sub

Private Sub mnuDispFullI_Click(Index As Integer)
gen_GoFullscreen m_cM.Monitor(Index + 1)
End Sub

Private Sub mnuHelpAbout_Click()
Main_About
End Sub

Private Sub mnuHelpWiki_Click()
ShellExecuteURL "http://terra.aefaradien.net"
End Sub

Private Sub mnuHistoryI_Click(Index As Integer)
If mdb_History.Items(Index).lSource(1) < 0 Then
    ps_sMinusOneFilePath = mdb_History.Items(Index).sFile
    pb_StartPlayback -1, 0
Else
    pb_StartPlayback mdb_History.Items(Index).lSource(0), _
         mdb_History.Items(Index).lSource(1)
End If
End Sub

Private Sub mnuJumpAuto_Click()
m_TrackCurrentItem = Not m_TrackCurrentItem
mnuJumpAuto.Checked = m_TrackCurrentItem

pref_StartAutoSave
End Sub

Private Sub mnuJumpCurrent_Click()
list_JumpToCurrent
End Sub

Private Sub mnuLibAddcue_Click()
If mdb_List.lCnt < 1 Then Exit Sub

Dim i As Long

For i = 0 To mdb_List.lCnt - 1
    If mdb_List.Items(i).bSel Then
        cue_AddTo 0, mdb_List.Items(i).sFile, , mdb_List.Items(i).lMD5, mdb_List.Items(i).lDuration
    End If
Next i
End Sub

Private Sub mnuLibAddplI_Click(Index As Integer)
If mdb_List.lCnt < 1 Then Exit Sub
Dim i As Long, lTargetList As Long

lTargetList = Val(mnuLibAddplI(Index).Tag)

For i = 0 To mdb_List.lCnt - 1
    If mdb_List.Items(i).bSel Then
        pl_AddItem mdb_PL(lTargetList), mdb_List.Items(i).sFile, mdb_List.Items(i).lMD5, , , mdb_List.Items(i).lDuration
    End If
Next i

pll_UpdateStatus True, True
End Sub

Private Sub mnuLibAddplNew_Click()
If mdb_List.lCnt < 1 Then Exit Sub
Dim i As Long, lPl As Long

lPl = pll_New(True, True)

For i = 0 To mdb_List.lCnt - 1
    If mdb_List.Items(i).bSel Then
        pl_AddItem mdb_PL(lPl), mdb_List.Items(i).sFile, mdb_List.Items(i).lMD5, , , mdb_List.Items(i).lDuration
    End If
Next i

pll_UpdateStatus True, True
End Sub

Private Sub mnuLibConfig_Click()
If Not mnuOptTab.Checked Then
    mnuOptTab_Click
Else
    tsMain.Tabs("sys").Selected = True
End If
lstPrefPages.ListIndex = 5
End Sub

Private Sub mnuLibEnabled_Click()
If mdb_List.lCnt < 1 Then Exit Sub

Dim i As Long

For i = 0 To mdb_List.lCnt - 1
    If mdb_List.Items(i).bSel Then
        mdb_List.Items(i).bEnabled = Not mdb_List.Items(i).bEnabled
        mdb_SetEnabled mdb_List.Items(i).sFile, mdb_List.Items(i).bEnabled
    End If
Next i

mdbl_Redraw
End Sub

Private Sub mnuLibFind_Click()
On Error GoTo mnuLibFind_Click_err

mnuLibFind.Checked = Not mnuLibFind.Checked

fraLibSearch.Visible = mnuLibFind.Checked
fraLib.ForceResize

On Error Resume Next
If fraLibSearch.Visible Then txtLibSearch.SetFocus
On Error GoTo mnuLibFind_Click_err

Exit Sub
mnuLibFind_Click_err:
Main_Err "mnuLibFind_Click_err."
err.Clear
End Sub

Private Sub mnuLibGoto_Click()
gen_ShowGoto
End Sub

Private Sub mnuLibOrderDirection_Click(Index As Integer)
Dim i As Long
For i = 0 To mnuLibOrderDirection.count - 1
    mnuLibOrderDirection(i).Checked = (i = Index)
Next i
mdbl_Requery
End Sub

Private Sub mnuLibOrderI_Click(Index As Integer)
Dim i As Long
For i = 0 To mnuLibOrderI.count - 1
    mnuLibOrderI(i).Checked = (i = Index)
Next i
mdbl_Requery
End Sub

Private Sub mnuLibOrderShowmis_Click()
mnuLibOrderShowmis.Checked = Not mnuLibOrderShowmis.Checked
mdbl_Requery
End Sub

Private Sub mnuLibPlaycntI_Click(Index As Integer)
Dim i As Long
If mdb_List.lCnt < 1 Then Exit Sub

Select Case Index
    Case 0 'set
        If mdb_List.lIndex < 0 Then Exit Sub
        
        Dim a As String, lNewS As Long, lNewE As Long
        
        a = _
            InputBox("new start count:  (was " & _
            mdb_List.Items(mdb_List.lIndex).lStartCnt & ")", , _
            mdb_List.Items(mdb_List.lIndex).lStartCnt)
        If Len(a) < 1 Then Exit Sub
        lNewS = Val(a)
        
        a = _
            InputBox("new end count:  (was " & _
            mdb_List.Items(mdb_List.lIndex).lEndCnt & ")", , _
            mdb_List.Items(mdb_List.lIndex).lEndCnt)
        If Len(a) < 1 Then Exit Sub
        lNewE = Val(a)
        
        If MsgBox( _
            "sure to set this playback data? (there is no undo for this)" & vbNewLine & vbNewLine & _
            "file: " & mdb_List.Items(mdb_List.lIndex).sFile & vbNewLine & _
            "start count: " & lNewS & " (was " & mdb_List.Items(mdb_List.lIndex).lStartCnt & ")" & vbNewLine & _
            "end count: " & lNewE & " (was " & mdb_List.Items(mdb_List.lIndex).lEndCnt & ")" _
            , vbYesNo) <> vbYes Then Exit Sub
        
        mdb_SetPlaybackCnt mdb_List.Items(mdb_List.lIndex).sFile, lNewS, lNewE
        mdbl_Requery True
    
    Case 1 'reset
        If MsgBox("sure to reset these item's playback data? (there is no undo for this)", vbYesNo) <> vbYes Then Exit Sub
        
        For i = 0 To mdb_List.lCnt - 1
            If mdb_List.Items(i).bSel Then
                mdb_SetPlaybackCnt mdb_List.Items(i).sFile, 0, 0
            End If
        Next i
        
        mdbl_Requery True

End Select
End Sub

Private Sub mnuLibRemcrc_Click()
Dim i As Long

If mdb_List.lCnt < 1 Then Exit Sub

If MsgBox("sure to reset these item's crc data? (there is no undo for this)", vbYesNo) <> vbYes Then Exit Sub

For i = 0 To mdb_List.lCnt - 1
    If mdb_List.Items(i).bSel Then
        mdb_SetHash mdb_List.Items(i).sFile, 0
    End If
Next i

mdbl_Requery True
End Sub

Private Sub mnuLibRemsel_Click()
Dim i As Long

If mdb_List.lCnt < 1 Then Exit Sub

If MsgBox("sure to remove these items from mdb?", vbYesNo) <> vbYes Then Exit Sub

For i = 0 To mdb_List.lCnt - 1
    If mdb_List.Items(i).bSel Then
        mdb_RemFile mdb_List.Items(i).sFile
    End If
Next i

mdbl_Requery True
End Sub

Private Sub mnuLibReport_Click()
On Error GoTo mnuLibReport_Click_err

Dim fRep As frmOutput, cStrB As cStringBuilder, _
    i As Long, _
    vRet As Variant, vI As Variant, vJ As Variant, sSql As String, lQueryRowCnt As Long

Set fRep = New frmOutput
Set cStrB = New cStringBuilder

With cStrB
    .Append "// terra media library report" & vbNewLine
    .Append "generated on " & Format(Now, "yyyy mm dd") & _
        " at " & Format(Now, "hh:mm:ss") & "." & _
        vbNewLine & vbNewLine
    
    .Append "files in library: "
    .Append mdb_Function("select count(sfile) from tbl_mediafiles;")
    .Append vbNewLine
    
    .Append "total starts: "
    vI = mdb_Function("select total(lstartcnt) from tbl_mediafiles;")
    .Append Trim$(Str$(vI))
    .Append vbNewLine

    .Append "total finishes: "
    vJ = mdb_Function("select total(lendcnt) from tbl_mediafiles;")
    .Append Trim$(Str$(vJ)) & " (" & Round((vJ / vI) * 100, 3) & "% of starts)"
    .Append vbNewLine
    
    .Append "mean start: "
    .Append Trim$(Str$(Round(mdb_Function("select total(lstartcnt)/count(lstartcnt) from tbl_mediafiles;"), 3)))
    .Append vbNewLine
    
    .Append "mean finish: "
    .Append Trim$(Str$(Round(mdb_Function("select total(lendcnt)/count(lendcnt) from tbl_mediafiles;"), 3)))
    .Append vbNewLine
    
    .Append "first added: "
    .Append Trim$(mdb_Function("select min(dadded) from tbl_mediafiles;"))
    .Append " (" & Trim$(Str$(Round(mdb_Function("select julianday('now')-julianday(dadded) from tbl_mediafiles;"), 3))) & _
        " days ago)"
    .Append vbNewLine
    
    .Append "total duration of measured items (h:mm:ss): "
    .Append ConvertSecToHours(Trim$(Str$(mdb_Function("select total(lduration) from tbl_mediafiles;"))))
    .Append vbNewLine
    
    .Append "missing items: "
    .Append mdb_Function("select count(sfile) from tbl_mediafiles WHERE bmissing=1;")
    .Append vbNewLine
    
    .Append "disabled items: "
    .Append mdb_Function("select count(sfile) from tbl_mediafiles WHERE benabled=0;")
    .Append vbNewLine
    
    .Append vbNewLine
    .Append "top 5 (by end count):" & vbNewLine
    vRet = mdb_Query( _
        "select sfile from tbl_mediafiles order by lendcnt desc limit 5;")
        '"select sfile, ' (', lstartcnt, '/', lendcnt, ')' as a from tbl_mediafiles order by lendcnt desc limit 5;"
    If number_of_rows_from_last_call > 0 Then
        i = 0
        For Each vI In vRet
            If i > 5 Then
                Exit Sub
            ElseIf i > 0 Then
                .Append us_Decode((vI)) & vbNewLine
            End If
            i = i + 1
        Next vI
    Else
        .Append "query returned no files."
    End If
    
    
    
    fRep.Caption = "terra media library report"
    fRep.txtOutput.Text = .ToString
End With

fRep.Show

Exit Sub
mnuLibReport_Click_err:
Main_Err "mnuLibReport_Click_err."
err.Clear
End Sub

Private Sub mnuLibReq_Click()
mdbl_Requery
End Sub

Private Sub mnuLibSelFilecopy_Click()
If mdb_List.lCnt < 1 Then Exit Sub

Dim i As Long, cFiles As Collection
Set cFiles = New Collection

For i = 0 To mdb_List.lCnt - 1
    If mdb_List.Items(i).bSel Then
        cFiles.Add mdb_List.Items(i).sFile
    End If
Next i

If cFiles.count < 1 Then Exit Sub

gen_CopyFilesToFolder cFiles
End Sub

Private Sub mnuLibSelPathcopy_Click()
If mdb_List.lCnt < 1 Then Exit Sub

Dim i As Long, cStrB As New cStringBuilder

For i = 0 To mdb_List.lCnt - 1
    If mdb_List.Items(i).bSel Then
        cStrB.Append mdb_List.Items(i).sFile & vbNewLine
    End If
Next i

If cStrB.Length < 1 Then Exit Sub

Clipboard.Clear
Clipboard.SetText cStrB.ToString, vbCFText
Clipboard.SetText cStrB.ToString, vbCFRTF
End Sub

Private Sub mnuModeI_Click(Index As Integer)
Dim i As Long
For i = 0 To mnuModeI.count - 1
    mnuModeI(i).Checked = False
Next i
mnuModeI(Index).Checked = True
cmdMnu(3).Caption = mnuModeI(Index).Caption

pref_StartAutoSave
End Sub

Private Sub mnuModeListsI_Click(Index As Integer)
Dim i As Long
For i = 0 To mnuModeListsI.count - 1
    mnuModeListsI(i).Checked = False
Next i
mnuModeListsI(Index).Checked = True

pref_StartAutoSave
End Sub

Private Sub mnuNxspI_Click(Index As Integer)
On Error GoTo mnuNxspI_Click_err

Dim lNextList As Long, lNextItem As Long, _
    e As String, i As Long, _
    lListMode As Long

lNextList = -1

Select Case Index
    Case 0 'sequential
        e = "0"
        For i = 0 To mnuModeListsI.count - 1
            If mnuModeListsI(i).Checked = True Then
                lListMode = i
                Exit For
            End If
        Next i
        e = "1"
        lNextList = pb_lItemSource
        e = "2"
        lNextItem = pbq_GetItemSeq(lNextList, Not lListMode = 0)
    
    Case 1 'random
        e = "1"
        lNextList = pbq_GetListTruerand(True)
        e = "2"
        If lNextList >= 0 Then lNextItem = pbq_GetItemTruerand(lNextList)
    
    Case 2 'shuffle by start count
        e = "1"
        lNextList = pbq_GetListTruerand(True)
        e = "2"
        If lNextList >= 0 Then lNextItem = pbq_GetItemSemirand(lNextList)
    
    Case 3 'shuffle by last played date
        e = "1"
        lNextList = pbq_GetListTruerand(True)
        e = "2"
        lNextItem = pbq_GetItemLastplayed2(lNextList)
    
    Case 4 'un-played
        e = "1"
        lNextList = 0 'must be in lib
        e = "2"
        lNextItem = pbq_GetItemLibraryRandFromQuery("lStartCnt<=0")
    
    Case 5 'prefered (> 0.7)
        e = "1"
        lNextList = 0 'must be in lib
        e = "2"
        lNextItem = pbq_GetItemLibraryRandFromQuery( _
            "p>=0.7", _
            "cast(lendcnt as real) / cast(lstartcnt as real) AS p")
    
    Case 6 'not played in the last 30 days
        e = "1"
        lNextList = 0 'must be in lib
        e = "2"
        lNextItem = pbq_GetItemLibraryRandFromQuery("date('now', '-30 day') >= dlastplay")
    
    Case 7 'new (last 3 days)
        e = "1"
        lNextList = 0 'must be in lib
        'lNextItem = pbq_GetItemLibraryRandFromQuery("date('now', '-3 day') <= dadded")
        e = "2"
        lNextItem = pbq_GetItemLibraryRandFromQueryFull("SELECT sfile FROM tbl_mediafiles WHERE dadded>=date((SELECT max(dadded) FROM tbl_mediafiles),'-3 day') AND (benabled=1 OR benabled IS NULL);")
    
    Case 8 'old(1 of 10 played longest ago)
        e = "1"
        lNextList = 0 'must be in lib
        'lNextItem = pbq_GetItemLibraryRandFromQuery("date('now', '-30 day') >= dadded")
        e = "2"
        lNextItem = pbq_GetItemLibraryRandFromQueryFull("SELECT sfile FROM tbl_mediafiles WHERE (benabled=1 OR benabled IS NULL) ORDER BY dlastplay ASC LIMIT 10;")
    
End Select

e = "3"
If lNextList >= 0 And lNextItem >= 0 Then
    pb_StartPlayback lNextList, lNextItem
End If

Exit Sub
mnuNxspI_Click_err:
Main_Err "mnuNxspI_Click_err.  e=" & e & _
    ", i=" & Index & _
    ", lNextList=" & lNextList & _
    ", lNextItem=" & lNextItem & "."
err.Clear
End Sub

Private Sub mnuOptSpeed_Click()
Dim a As String, b As Double
a = InputBox("warning: this is an advanced feature.  set playback rate (0.5 to 3):", , Trim$(Str$(pb_Speed)))
If Len(a) < 1 Or Val(a) <= 0 Then Exit Sub

b = Val(a)
If b < 0.5 Then
    b = 0.5
ElseIf b > 3 Then
    b = 3
End If

pb_Speed = b

mnuOptSpeed.Caption = mnuOptSpeed.Tag & " (" & pb_Speed & "×)"
End Sub

Private Sub mnuOptTab_Click()
mnuOptTab.Checked = Not mnuOptTab.Checked

If mnuOptTab.Checked Then
    tsMain.Tabs.Add 2, "sys", "preferences", "sys"
    tsMain.Tabs("sys").Selected = True
Else
    Dim b As Boolean, k As String
    b = tsMain.Tabs("sys").Selected
    k = tsMain.SelectedItem.Key
    tsMain.Tabs.Remove "sys"
    If b Then tsMain_Click Else tsMain.Tabs(k).Selected = True
End If

'pll_RebuildEnab

Form_Resize
End Sub

Private Sub mnuPlAddcue_Click()
If m_lCurrentPl < 0 Then Exit Sub
If mdb_PL(m_lCurrentPl).lCnt < 1 Then Exit Sub

Dim i As Long

For i = 0 To mdb_PL(m_lCurrentPl).lCnt - 1
    If mdb_PL(m_lCurrentPl).Items(i).bSel Then
        cue_AddTo m_lCurrentPl + 1, mdb_PL(m_lCurrentPl).Items(i).sFile, i, _
            mdb_PL(m_lCurrentPl).Items(i).lMD5, mdb_PL(m_lCurrentPl).Items(i).lDuration
    End If
Next i
End Sub

Private Sub mnuPlAddplI_Click(Index As Integer)
If m_lCurrentPl < 0 Then Exit Sub
If mdb_PL(m_lCurrentPl).lCnt < 1 Then Exit Sub

Dim i As Long, lTargetList As Long, _
    sFile As String, lHash As Long, lDuration As Long

lTargetList = Val(mnuPlAddplI(Index).Tag)

For i = 0 To mdb_PL(m_lCurrentPl).lCnt - 1
    If mdb_PL(m_lCurrentPl).Items(i).bSel Then
        sFile = mdb_PL(m_lCurrentPl).Items(i).sFile
        lHash = mdb_PL(m_lCurrentPl).Items(i).lMD5
        lDuration = mdb_PL(m_lCurrentPl).Items(i).lDuration
        pl_AddItem mdb_PL(lTargetList), sFile, lHash, , , lDuration
    End If
Next i

If m_lCurrentPl = lTargetList Then pll_RedrawCurrent
pll_UpdateStatus True, True
End Sub

Private Sub mnuPlAddplNew_Click()
If m_lCurrentPl < 0 Then Exit Sub
If mdb_PL(m_lCurrentPl).lCnt < 1 Then Exit Sub

Dim i As Long, sFile As String, lHash As Long, lDuration As Long, lPl As Long

lPl = pll_New(True, True)

For i = 0 To mdb_PL(m_lCurrentPl).lCnt - 1
    If mdb_PL(m_lCurrentPl).Items(i).bSel Then
        sFile = mdb_PL(m_lCurrentPl).Items(i).sFile
        lHash = mdb_PL(m_lCurrentPl).Items(i).lMD5
        lDuration = mdb_PL(m_lCurrentPl).Items(i).lDuration
        pl_AddItem mdb_PL(lPl), sFile, lHash, , , lDuration
    End If
Next i

pll_UpdateStatus True, True
End Sub

Private Sub mnuPlArc_Click()
If m_lCurrentPl < 0 Then Exit Sub
mnuPlArc.Checked = Not mnuPlArc.Checked
mdb_PL(m_lCurrentPl).bArc = mnuPlArc.Checked
pll_RebuildTabs
If Not mdb_PL(m_lCurrentPl).bArc Then pll_RebuildAndPrepArc
tsMain_Click
pll_UpdateStatus True, True 'pl_SetAutoSave True
End Sub

Private Sub mnuPlDel_Click()
If m_lCurrentPl < 0 Then Exit Sub

If MsgBox("are you want to delete the playlist '" & mdb_PL(m_lCurrentPl).sName & _
    "', there is no undo.", vbYesNo) <> vbYes Then Exit Sub

pll_RemoveList m_lCurrentPl
pll_RebuildTabs
pll_RedrawCurrent
End Sub

Private Sub mnuPlEnab_Click()
If m_lCurrentPl < 0 Then Exit Sub
mnuPlEnab.Checked = Not mnuPlEnab.Checked
mdb_PL(m_lCurrentPl).bEnab = mnuPlEnab.Checked
pll_RebuildEnab
pll_UpdateStatus True, True 'pl_SetAutoSave True
End Sub

Private Sub mnuPlFilecopy_Click()
If m_lCurrentPl < 0 Then Exit Sub
If mdb_PL(m_lCurrentPl).lCnt < 1 Then Exit Sub

Dim i As Long, cFiles As Collection
Set cFiles = New Collection

For i = 0 To mdb_PL(m_lCurrentPl).lCnt - 1
    If mdb_PL(m_lCurrentPl).Items(i).bSel Then
        cFiles.Add mdb_PL(m_lCurrentPl).Items(i).sFile
    End If
Next i

If cFiles.count < 1 Then Exit Sub

gen_CopyFilesToFolder cFiles
End Sub

Private Sub mnuPlFindmdb_Click()
If m_lCurrentPl < 0 Then Exit Sub
If mdb_PL(m_lCurrentPl).lCnt < 1 Then Exit Sub
If mdb_PL(m_lCurrentPl).lIndex < 0 Then Exit Sub

Dim i As Long, x As Long
x = mdbl_FindIndexFromFile(mdb_PL(m_lCurrentPl).Items(mdb_PL(m_lCurrentPl).lIndex).sFile)
If x < 0 Then Exit Sub

For i = 0 To mdb_List.lCnt - 1
    mdb_List.Items(i).bSel = (i = x)
Next i
mdb_List.lIndex = x

list_JumpTo 0, x, True
End Sub

Private Sub mnuPlGethsh_Click()
If m_lCurrentPl < 0 Then Exit Sub
If mdb_PL(m_lCurrentPl).lCnt < 1 Then Exit Sub

Dim i As Long, x As Long

For i = 0 To mdb_PL(m_lCurrentPl).lCnt - 1
    If mdb_PL(m_lCurrentPl).Items(i).bSel Then
        x = mdbl_FindIndexFromFile(mdb_PL(m_lCurrentPl).Items(i).sFile)
        If x < 0 Then GoTo NextI
        mdb_PL(m_lCurrentPl).Items(i).lMD5 = mdb_List.Items(x).lMD5
    End If
NextI:
Next i

pll_RedrawCurrent
pll_UpdateStatus True, True ' pl_SetAutoSave True
End Sub

Private Sub mnuPlHash_Click()
MsgBox "todo: hash selected pl files."
End Sub

Private Sub mnuPlMove_Click(Index As Integer)
'move selected items
If m_lCurrentPl < 0 Then Exit Sub
pl_MoveItems mdb_PL(m_lCurrentPl), Index, True, m_lCurrentPl + 1
pll_UpdateStatus True, True 'pl_SetAutoSave True
pll_RedrawCurrent
End Sub

Private Sub mnuPlMoveTabI_Click(Index As Integer)
Dim lTo As Long

If m_lCurrentPl < 0 Then Exit Sub

Select Case Index
    Case 0 'left
        If m_lCurrentPl = 0 Then Exit Sub
        lTo = m_lCurrentPl - 1
    
    Case 1 'right
        If m_lCurrentPl = UBound(mdb_PL) Then Exit Sub
        lTo = m_lCurrentPl + 1
    
End Select

pll_SwapLists m_lCurrentPl, lTo
pll_RebuildTabs
If mdb_PL(lTo).bArc Then
    pll_RebuildAndPrepArc lTo
Else
    tsMain.Tabs("pl" & lTo).Selected = True
End If

pll_UpdateStatus True, True  'pl_SetAutoSave True
End Sub

Private Sub mnuPlName_Click()
If m_lCurrentPl < 0 Then Exit Sub

Dim a As String
a = InputBox("name playlist.", , mdb_PL(m_lCurrentPl).sName)
If a = "" Then Exit Sub
mdb_PL(m_lCurrentPl).sName = a
pll_RebuildTabs

pll_UpdateStatus True, True ' pl_SetAutoSave True
End Sub

Private Sub mnuPlNew_Click()
pll_New True, True
End Sub

Private Sub mnuPlPathcopy_Click()
If m_lCurrentPl < 0 Then Exit Sub
If mdb_PL(m_lCurrentPl).lCnt < 1 Then Exit Sub

Dim i As Long, cStrB As New cStringBuilder

For i = 0 To mdb_PL(m_lCurrentPl).lCnt - 1
    If mdb_PL(m_lCurrentPl).Items(i).bSel Then
        cStrB.Append mdb_PL(m_lCurrentPl).Items(i).sFile & vbNewLine
    End If
Next i

If cStrB.Length < 1 Then Exit Sub

Clipboard.Clear
Clipboard.SetText cStrB.ToString, vbCFText
Clipboard.SetText cStrB.ToString, vbCFRTF
End Sub

Private Sub mnuPlPlaycntI_Click(Index As Integer)
Dim i As Long, a As String, lSt As Long, lNd As Long, x As Long

If m_lCurrentPl < 0 Then Exit Sub
If mdb_PL(m_lCurrentPl).lCnt < 1 Then Exit Sub

'** query user **
Select Case Index
    Case 0 'set
        If mdb_PL(m_lCurrentPl).lIndex < 0 Then Exit Sub
        
        a = _
            InputBox("new start count:  (was " & _
            mdb_PL(m_lCurrentPl).Items(mdb_PL(m_lCurrentPl).lIndex).lStartCnt & ")", , _
            mdb_PL(m_lCurrentPl).Items(mdb_PL(m_lCurrentPl).lIndex).lStartCnt)
        If Len(a) < 1 Then Exit Sub
        lSt = Val(a)
        
        a = _
            InputBox("new end count:  (was " & _
            mdb_PL(m_lCurrentPl).Items(mdb_PL(m_lCurrentPl).lIndex).lEndCnt & ")", , _
            mdb_PL(m_lCurrentPl).Items(mdb_PL(m_lCurrentPl).lIndex).lEndCnt)
        If Len(a) < 1 Then Exit Sub
        lNd = Val(a)
        
        If MsgBox( _
            "sure to set this playback data? (there is no undo for this)" & vbNewLine & vbNewLine & _
            "file: " & mdb_PL(m_lCurrentPl).Items(mdb_PL(m_lCurrentPl).lIndex).sFile & vbNewLine & _
            "start count: " & lSt & " (was " & mdb_PL(m_lCurrentPl).Items(mdb_PL(m_lCurrentPl).lIndex).lStartCnt & ")" & vbNewLine & _
            "end count: " & lNd & " (was " & mdb_PL(m_lCurrentPl).Items(mdb_PL(m_lCurrentPl).lIndex).lEndCnt & ")" _
            , vbYesNo) <> vbYes Then Exit Sub
    
    Case 1 'reset
        If MsgBox("sure to reset these item's playback data? (there is no undo for this)", _
            vbYesNo) <> vbYes Then Exit Sub
    
    Case 2 'get from library
        If MsgBox("sure to get these item's playback data from the library? (there is no undo for this)", _
            vbYesNo) <> vbYes Then Exit Sub
        
        mdbl_Requery False

End Select

'** do **
Select Case Index
    Case 0 'set
        mdb_PL(m_lCurrentPl).Items(mdb_PL(m_lCurrentPl).lIndex).lStartCnt = lSt
        mdb_PL(m_lCurrentPl).Items(mdb_PL(m_lCurrentPl).lIndex).lEndCnt = lNd
    
    Case 1, 2 'reset,get
        For i = 0 To mdb_PL(m_lCurrentPl).lCnt - 1
            If mdb_PL(m_lCurrentPl).Items(i).bSel Then
                Select Case Index
                    Case 0 'set
                        
                    
                    Case 1 'reset
                        lSt = 0
                        lNd = 0
                    
                    Case 2 'get cnt from mdb
                        x = mdbl_FindIndexFromFile(mdb_PL(m_lCurrentPl).Items(i).sFile)
                        If x < 0 Then GoTo NextI
                        lSt = mdb_List.Items(x).lStartCnt
                        lNd = mdb_List.Items(x).lEndCnt
                    
                End Select
                
                mdb_PL(m_lCurrentPl).Items(i).lStartCnt = lSt
                mdb_PL(m_lCurrentPl).Items(i).lEndCnt = lNd
            End If
NextI:
        Next i
    
End Select

pll_RedrawCurrent
pll_UpdateStatus True, True ' pl_SetAutoSave True
End Sub

Private Sub mnuPlPlay_Click()
Dim i As Long, lMode As Long, lList As Long, lItem As Long

'list is specified by user
lList = m_lCurrentPl + 1

'find play mode
For i = 0 To mnuModeI.count - 1
    If mnuModeI(i).Checked = True Then
        lMode = i
        Exit For
    End If
Next i

'find track based on mode
Select Case lMode
    Case 0 'sequencial
        lItem = pbq_GetItemSeq(lList, False)
    
    Case 1 'semi-random
        lItem = pbq_GetItemSemirand(lList)
    
    Case 2 'random
        lItem = pbq_GetItemTruerand(lList)
    
    Case 3 'by date last played
        lItem = pbq_GetItemLastplayed2(lList)
    
End Select

'play item
pb_StartPlayback lList, lItem
End Sub

Private Sub mnuPlRemdup_Click()
Dim i As Long, j As Long, f As String, _
    lStart As Long, lEnd As Long

If m_lCurrentPl < 0 Then Exit Sub
If mdb_PL(m_lCurrentPl).lCnt < 1 Then Exit Sub

If MsgBox("this will remove all duplicate files, merging play count data to the remaining items.  continue?", vbYesNo) <> vbYes Then Exit Sub

'unsel all
For i = 0 To mdb_PL(m_lCurrentPl).lCnt - 1
    mdb_PL(m_lCurrentPl).Items(i).bSel = False
Next i

'check all items
For i = 0 To mdb_PL(m_lCurrentPl).lCnt - 1
    'selected items have already been processed
    If Not mdb_PL(m_lCurrentPl).Items(i).bSel Then
        'stage 1: merge all playcounts
        lStart = 0
        lEnd = 0
        For j = 0 To mdb_PL(m_lCurrentPl).lCnt - 1
            If LCase$(mdb_PL(m_lCurrentPl).Items(i).sFile) = LCase$(mdb_PL(m_lCurrentPl).Items(j).sFile) Then
                lStart = lStart + mdb_PL(m_lCurrentPl).Items(j).lStartCnt
                lEnd = lEnd + mdb_PL(m_lCurrentPl).Items(j).lEndCnt
            End If
        Next j
        For j = 0 To mdb_PL(m_lCurrentPl).lCnt - 1
            If LCase$(mdb_PL(m_lCurrentPl).Items(i).sFile) = LCase$(mdb_PL(m_lCurrentPl).Items(j).sFile) Then
                mdb_PL(m_lCurrentPl).Items(j).lStartCnt = lStart
                mdb_PL(m_lCurrentPl).Items(j).lEndCnt = lEnd
            End If
        Next j
        
        'stage 2: select subsiquent duplicates
        For j = i + 1 To mdb_PL(m_lCurrentPl).lCnt - 1
            If LCase$(mdb_PL(m_lCurrentPl).Items(i).sFile) = LCase$(mdb_PL(m_lCurrentPl).Items(j).sFile) Then
                mdb_PL(m_lCurrentPl).Items(j).bSel = True
            End If
        Next j
        
    End If
Next i

'remove all selected items
For i = mdb_PL(m_lCurrentPl).lCnt - 1 To 0 Step -1
    If mdb_PL(m_lCurrentPl).Items(i).bSel Then
        pl_RemoveItem mdb_PL(m_lCurrentPl), i, True, m_lCurrentPl + 1
    End If
Next i

pll_UpdateStatus True, True  'pl_SetAutoSave True
pll_RedrawCurrent
End Sub

Private Sub mnuPlRemitm_Click()
Dim i As Long

If m_lCurrentPl < 0 Then Exit Sub
If mdb_PL(m_lCurrentPl).lCnt < 1 Then Exit Sub

For i = mdb_PL(m_lCurrentPl).lCnt - 1 To 0 Step -1
    If mdb_PL(m_lCurrentPl).Items(i).bSel Then
        pl_RemoveItem mdb_PL(m_lCurrentPl), i, True, m_lCurrentPl + 1
    End If
Next i

pll_UpdateStatus True, True 'pl_SetAutoSave True
pll_RedrawCurrent
End Sub

Private Sub mnuPlRep_Click()
On Error GoTo mnuPlRep_Click_err

If m_lCurrentPl < 0 Then Exit Sub

If MsgBox("this will attempt to repair the current pl's missing items by looking up their hash codes in the library.  continue?" & vbNewLine & _
    "(ensure that library contains no duplicate hash codes before continuing.)", _
    vbYesNo) <> vbYes Then Exit Sub

Dim i As Long, x As Long, _
    lFound As Long, lUnfound As Long

mdbl_Requery False

lFound = 0
lUnfound = 0

For i = 0 To mdb_PL(m_lCurrentPl).lCnt - 1
    mdb_PL(m_lCurrentPl).Items(i).bSel = False
    
    If Not fsoMain.FileExists(mdb_PL(m_lCurrentPl).Items(i).sFile) Then
        x = mdbl_FindIndexFromHash(mdb_PL(m_lCurrentPl).Items(i).lMD5)
        
        If x >= 0 Then
            mdb_PL(m_lCurrentPl).Items(i).sFile = mdb_List.Items(x).sFile
            lFound = lFound + 1
        Else
            mdb_PL(m_lCurrentPl).Items(i).bSel = True
            lUnfound = lUnfound + 1
        End If
        
    End If
Next i

MsgBox "repair complete." & vbNewLine & vbNewLine & _
    "items repaired: " & lFound & "." & vbNewLine & _
    "items not repaired: " & lUnfound & " (now selected)."

If lFound > 0 Then pll_UpdateStatus True, True ' pl_SetAutoSave True

pll_RedrawCurrent

Exit Sub
mnuPlRep_Click_err:
Main_Err "mnuPlRep_Click_err."
err.Clear
End Sub

Private Sub mnuPlSelall_Click()
If m_lCurrentPl < 0 Then Exit Sub
pl_SelAll mdb_PL(m_lCurrentPl)
pll_RedrawCurrent
End Sub

Private Sub mnuPlSelDup_Click()
'If m_lCurrentPl < 0 Then Exit Sub
'If mdb_PL(m_lCurrentPl).lCnt < 1 Then Exit Sub
'
'Dim i As Long, j As Long, _
'    plFound As typPl
'
'pl_SetAsNew plFound, "found"
'
''go through each item in the pl
'For i = 0 To mdb_PL(m_lCurrentPl).lCnt - 1
'    'check we don't already have this item
'    If plFound.lCnt > 0 Then
'        For i = 0 To plFound.lCnt - 1
'            If LCase$(mdb_PL(m_lCurrentPl).Items(i).sFile) = LCase$(plFound.Items(i).sFile) Then
'            'if we do already have this item, check which has been played the most
'
'
'        Next i
'    End If
'Next i
End Sub

Private Sub mnuPlSelMis_Click()
On Error GoTo mnuPlSelMis_Click_err

Dim i As Long, lFound As Long

If m_lCurrentPl < 0 Then Exit Sub
If mdb_PL(m_lCurrentPl).lCnt < 1 Then Exit Sub

lFound = 0

For i = 0 To mdb_PL(m_lCurrentPl).lCnt - 1
    If fsoMain.FileExists(mdb_PL(m_lCurrentPl).Items(i).sFile) Then
        mdb_PL(m_lCurrentPl).Items(i).bSel = False
    Else
        mdb_PL(m_lCurrentPl).Items(i).bSel = True
        lFound = lFound + 1
    End If
Next i

If lFound > 0 Then
    pll_RedrawCurrent
    MsgBox lFound & " missing items selected."
Else
    MsgBox "no missing items found :)."
End If

Exit Sub
mnuPlSelMis_Click_err:
Main_Err "mnuPlSelMis_Click_err."
err.Clear
End Sub

Private Sub mnuPlSortI_Click(Index As Integer)
If m_lCurrentPl < 0 Then Exit Sub
If mdb_PL(m_lCurrentPl).lCnt < 1 Then Exit Sub

'note to self: lListI is 1 more becuase it is part of all lists, not just playlists.

Select Case Index
    Case 0 'reverse
        pl_Reverse mdb_PL(m_lCurrentPl), True, m_lCurrentPl + 1
    
    Case 1 'path
        pl_Sort mdb_PL(m_lCurrentPl), 0, , , True, m_lCurrentPl + 1
    
    Case 2 'playcount - started
        pl_Sort mdb_PL(m_lCurrentPl), 1, , , True, m_lCurrentPl + 1
    
    Case 3 'playcount - finished
        pl_Sort mdb_PL(m_lCurrentPl), 2, , , True, m_lCurrentPl + 1
    
    Case 4 'date last played
        pl_Sort mdb_PL(m_lCurrentPl), 3, , , True, m_lCurrentPl + 1
    
End Select

pll_UpdateStatus True, True 'pl_SetAutoSave True
pll_RedrawCurrent
End Sub

Private Sub mnuTrayI_Click(Index As Integer)
Select Case Index
    Case 0: cmdPlayState_Click 0 'stop
    Case 1: cmdPlayState_Click 1 'play / pause
    Case 2: cmdPlayNext_Click 'next
End Select
End Sub

Private Sub picCue_GotFocus()
picCue.Tag = "1"
cue_Redraw
End Sub

Private Sub picCue_KeyDown(KeyCode As Integer, Shift As Integer)
If pl_KeyDown(mdb_Cue, vsbCue, picCue.ScaleHeight, KeyCode, Shift) Then
    cue_Redraw
End If

Select Case KeyCode
    Case vbKeyDelete:  cue_RemSel
    Case vbKeyA: If (Shift And vbCtrlMask) Then pl_SelAll mdb_Cue: cue_Redraw
End Select
End Sub

Private Sub picCue_LostFocus()
picCue.Tag = "0"
cue_Redraw
End Sub

Private Sub picCue_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
pl_MouseDown mdb_Cue, vsbCue, Button, Shift, x, Y
cue_Redraw
End Sub

Private Sub picCue_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim i As Long

If Not Data.GetFormat(vbCFFiles) Then Exit Sub

For i = 1 To Data.Files.count
    If fsoMain.FileExists(Data.Files(i)) Then
        If IsMediaFile(Data.Files(i)) Then
            cue_AddTo -1, Data.Files(i)
        End If
    ElseIf fsoMain.FolderExists(Data.Files(i)) Then
        
        Dim j As Long, cFlds As Collection, cFiles As Collection
    
        Set cFlds = New Collection
        BuildDirCollection Data.Files(i), cFlds
        
        Set cFiles = New Collection
        For j = 1 To cFlds.count
            AddFilesToCollection cFlds(j), cFiles
        Next j
        
        If cFiles.count < 1 Then Exit Sub
        
        For j = 1 To cFiles.count
            If IsMediaFile(cFiles(j)) Then
                frmMain.cue_AddTo -1, cFiles(j)
            End If
        Next j
        
    End If
Next i
End Sub

Private Sub picCue_Paint()
cue_ListDc.Draw picCue.hDC, 0, 0, cue_ListDc.Width, cue_ListDc.Height, 0, 0
End Sub

Private Sub picMdbLst_DblClick()
If mdb_List.lIndex < 0 Then Exit Sub
pb_StartPlayback 0, mdb_List.lIndex
End Sub

Private Sub picMdbLst_GotFocus()
'bring foxus box to visable area
If mdb_List.lIndex < vsbMdbLst.Value Then
    mdb_List.lIndex = vsbMdbLst.Value
ElseIf mdb_List.lIndex > vsbMdbLst.Value + Fix(picMdbLst.ScaleHeight / mdbl_ItmH) - 1 Then
    mdb_List.lIndex = vsbMdbLst.Value + Fix(picMdbLst.ScaleHeight / mdbl_ItmH) - 1
End If

picMdbLst.Tag = "1"
mdbl_Redraw
End Sub

Private Sub picMdbLst_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Long

Select Case KeyCode
    Case vbKeyPageUp, vbKeyPageDown
        mdbl_SetScroll vsbMdbLst.Value + _
            (vsbMdbLst.LargeChange * IIf(KeyCode = vbKeyPageUp, -1, 1))
    
    Case vbKeyUp, vbKeyDown
        i = mdb_List.lIndex + IIf(KeyCode = vbKeyUp, -1, 1)
        If i < 0 Then
            i = 0
        ElseIf i > mdb_List.lCnt - 1 Then
            i = mdb_List.lCnt - 1
        End If
        
        If (Shift And vbShiftMask) Then
            'mdb_List.Items(mdb_List.lIndex).bSel = Not mdb_List.Items(mdb_List.lIndex).bSel
            mdb_List.Items(i).bSel = mdb_List.Items(mdb_List.lIndex).bSel
        End If
        
        mdb_List.lIndex = i
        
        If i < vsbMdbLst.Value Then
            mdbl_SetScroll i
        ElseIf i > vsbMdbLst.Value + Fix(picMdbLst.ScaleHeight / mdbl_ItmH) - 1 Then
            mdbl_SetScroll i - Fix(picMdbLst.ScaleHeight / mdbl_ItmH) + 1
        End If
        
        mdbl_Redraw
    
    Case vbKeySpace
        If (Shift And vbShiftMask) Then
            
        ElseIf (Shift And vbCtrlMask) Then
            
        Else
            For i = 0 To mdb_List.lCnt - 1
                mdb_List.Items(i).bSel = False
            Next i
        End If
        If mdb_List.lIndex >= 0 Then
            mdb_List.Items(mdb_List.lIndex).bSel = Not mdb_List.Items(mdb_List.lIndex).bSel
            mdbl_Redraw
        End If
    
    Case vbKeyReturn
        picMdbLst_DblClick
    
    Case vbKeyQ
        mnuLibAddcue_Click
    
End Select
End Sub

Private Sub picMdbLst_LostFocus()
picMdbLst.Tag = "0"
mdbl_Redraw
End Sub

Private Sub picMdbLst_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim cI As Long, i As Long, d As Long ', sCnt As Long

If Button <> 1 Then Exit Sub
If Y < 0 Then Exit Sub
If mdb_List.lCnt < 1 Then Exit Sub

cI = vsbMdbLst.Value + Fix(Y / mdbl_ItmH)

If cI > mdb_List.lCnt - 1 Then
    cI = -1
    For i = 0 To mdb_List.lCnt - 1
        mdb_List.Items(i).bSel = False
    Next i
Else
    If Shift = 0 Then
        For i = 0 To mdb_List.lCnt - 1
            mdb_List.Items(i).bSel = False
        Next i
        mdb_List.Items(cI).bSel = Not mdb_List.Items(cI).bSel
    ElseIf (Shift And vbCtrlMask) Then
        mdb_List.Items(cI).bSel = Not mdb_List.Items(cI).bSel
    ElseIf (Shift And vbShiftMask) And mdb_List.lIndex >= 0 Then
        d = IIf(mdb_List.lIndex <= cI, 1, -1)
        For i = mdb_List.lIndex + d To cI Step d
            mdb_List.Items(i).bSel = Not mdb_List.Items(i).bSel
        Next i
    End If
End If

If cI >= 0 Then mdb_List.lIndex = cI

mdbl_Redraw

'sCnt = 0
'For i = 0 To mdb_List.lCnt - 1
'    If mdb_List.Items(i).bSel = True Then sCnt = sCnt + 1
'Next i
End Sub

Private Sub picMdbLst_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 2 And x >= 0 And Y >= 0 And x <= picMdbLst.ScaleWidth And Y <= picMdbLst.ScaleHeight Then
    mdbl_RebuildMenu
    'todo: update "enabled"
    PopupMenu mnuLibSel
End If
End Sub

Private Sub picMdbLst_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error GoTo picMdbLst_OLEDragDrop_err:

If Not Data.GetFormat(vbCFFiles) Then Exit Sub

Dim i As Long

For i = 1 To Data.Files.count
    If fsoMain.FileExists(Data.Files(i)) Then
        If IsMediaFile(Data.Files(i)) Then
            mdb_AddFile Data.Files(i)
        End If
    ElseIf fsoMain.FolderExists(Data.Files(i)) Then
        mdb_AddDir Data.Files(i)
    End If
Next i

mdbl_Requery False

Exit Sub
picMdbLst_OLEDragDrop_err:
Main_Err "picMdbLst_OLEDragDrop_err."
err.Clear
End Sub

Private Sub picMdbLst_Paint()
m_ListDc.Draw picMdbLst.hDC, 0, 0, m_ListDc.Width, m_ListDc.Height, 0, 0
End Sub

Private Sub picPl_DblClick()
If m_lCurrentPl < 0 Then Exit Sub
If mdb_PL(m_lCurrentPl).lIndex < 0 Then Exit Sub

'pb_lItemSource = m_lCurrentPl + 1
'pb_llItemIndex = mdb_PL(m_lCurrentPl).lIndex
'pb_SetPlayState 1

'mdb_PL(m_lCurrentPl).lCurrent = mdb_PL(m_lCurrentPl).lIndex
pb_StartPlayback m_lCurrentPl + 1, mdb_PL(m_lCurrentPl).lIndex
End Sub

Private Sub picPl_GotFocus()
If m_lCurrentPl < 0 Then Exit Sub

'check focus box visible
If mdb_PL(m_lCurrentPl).lIndex < mdb_PL(m_lCurrentPl).lScroll Then
    mdb_PL(m_lCurrentPl).lIndex = mdb_PL(m_lCurrentPl).lScroll
ElseIf mdb_PL(m_lCurrentPl).lIndex > mdb_PL(m_lCurrentPl).lScroll + Fix(picPl.ScaleHeight / mdbl_ItmH) - 1 Then
    mdb_PL(m_lCurrentPl).lIndex = mdb_PL(m_lCurrentPl).lScroll + Fix(picPl.ScaleHeight / mdbl_ItmH) - 1
End If

picPl.Tag = "1"
pll_RedrawCurrent
End Sub

Private Sub picPl_KeyDown(KeyCode As Integer, Shift As Integer)
If m_lCurrentPl < 0 Then Exit Sub

If pl_KeyDown(mdb_PL(m_lCurrentPl), vsbPl, picPl.ScaleHeight, KeyCode, Shift, True, m_lCurrentPl + 1) Then
    pll_RedrawCurrent
End If

Select Case KeyCode
    Case vbKeyReturn:  picPl_DblClick
    Case vbKeyDelete:  mnuPlRemitm_Click
    Case vbKeyN: If (Shift And vbCtrlMask) Then mnuPlNew_Click
    Case vbKeyE: If (Shift And vbCtrlMask) Then mnuPlName_Click
    Case vbKeyA: If (Shift And vbCtrlMask) Then mnuPlSelall_Click
    Case vbKeyM: If (Shift And vbCtrlMask) Then mnuPlSelMis_Click
    Case vbKeyQ: mnuPlAddcue_Click
    
    Case vbKeyL: If (Shift And vbCtrlMask) Then mnuPlMoveTabI_Click 0
    Case vbKeyR: If (Shift And vbCtrlMask) Then mnuPlMoveTabI_Click 1
End Select
End Sub

Private Sub picPl_LostFocus()
If m_lCurrentPl < 0 Then Exit Sub
picPl.Tag = "0"
pll_RedrawCurrent
End Sub

Private Sub picPl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If m_lCurrentPl < 0 Then Exit Sub
pl_MouseDown mdb_PL(m_lCurrentPl), vsbPl, Button, Shift, x, Y
pll_RedrawCurrent
End Sub

Private Sub picPl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If m_lCurrentPl < 0 Then Exit Sub
If Button = 2 And x >= 0 And Y >= 0 And x <= picPl.ScaleWidth And Y <= picPl.ScaleHeight Then
    pll_RebuildMenu
    PopupMenu mnuPlSel
End If
End Sub

Private Sub picPl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error GoTo picPl_OLEDragDrop_err

If m_lCurrentPl < 0 Then Exit Sub

Dim i As Long, a As String

If Not Data.GetFormat(vbCFFiles) Then Exit Sub

For i = 1 To Data.Files.count
    
    'a = StringFromAddr(StrPtr(Data.Files(i)), Len(Data.Files(i)))
    'If isUnicode(a) Then Debug.Print "unicode! " & a
    
    If fsoMain.FileExists(Data.Files(i)) Then
        If IsMediaFile(Data.Files(i)) Then
            pl_AddItem mdb_PL(m_lCurrentPl), Data.Files(i)
        End If
    ElseIf fsoMain.FolderExists(Data.Files(i)) Then
        pl_AddDir mdb_PL(m_lCurrentPl), Data.Files(i)
    End If
Next i

pll_RedrawCurrent
pll_UpdateStatus True, True

Exit Sub
picPl_OLEDragDrop_err:
Main_Err "picPl_OLEDragDrop_err."
err.Clear
End Sub

Private Sub picPl_Paint()
On Error GoTo picPl_err
pl_ListDc.Draw picPl.hDC, 0, 0, pl_ListDc.Width, pl_ListDc.Height, 0, 0
picPl_err:
err.Clear
End Sub

Private Sub sldOptDirLvls_Change()
lblOptDirLvls.Caption = IIf(sldOptDirLvls.Value = 0, "disabled", sldOptDirLvls.Value)
mdb_List.lTrunk = sldOptDirLvls.Value

pref_StartAutoSave
End Sub

Private Sub sldOptPlAutosave_Change()
If sldOptPlAutosave.Value <= 0 Then
    lblOptPlAutosave.Caption = "disabled."
Else
    lblOptPlAutosave.Caption = sldOptPlAutosave.Value & ":00 minutes."
End If
pl_lAutoSaveTime = sldOptPlAutosave.Value * 60

pref_StartAutoSave
End Sub

Private Sub sldPlay_ValueChanged(v As Double)
pb_SetPlaybackPosition cMedia.Duration * v
End Sub

Private Sub sytMinimize_MouseMove(x As Single)
Select Case x
    Case WM_LBUTTONDOWN
    
    Case WM_LBUTTONUP
        tray_Rest
    
    Case WM_LBUTTONDBLCLICK
    
    Case WM_RBUTTONUP
        PopupMenu mnuTray
    
    Case WM_MOUSEMOVE
    
End Select
End Sub

Private Sub tmrPlAutoSave_Timer()
If pl_lAutoSaveCntDn <= 0 Then
    pl_lAutoSaveCntDn = pl_lAutoSaveTime
    pll_WriteAllLists
Else
    pl_lAutoSaveCntDn = pl_lAutoSaveCntDn - 1
End If

pll_UpdateStatus
End Sub

Private Sub tmrPrefSave_Timer()
pref_AutoSaveCounter = pref_AutoSaveCounter + 1
If pref_AutoSaveCounter >= pref_AutoSaveTime Then
    pref_Write
Else
    cmdOptSave.Caption = cmdOptSave.Tag & " (" & pref_AutoSaveTime - pref_AutoSaveCounter & ")"
End If
End Sub

Private Sub tmrUpdateState_Timer()
On Error GoTo tmrUpdateState_Timer_err:

Dim i As Long, rc As RECT, a As String, b0 As Boolean, b1 As Boolean, e As String

'check this when in tray
e = "0"
If m_bInTray Then SystemTray.CheckStillInTray

e = "1"
'check playstay indicator is up-to-date
i = pb_GetPlayState
e = "1.a"
a = Array("stoped.", "playing ", "paused ")(i)
e = "1.b"
If i > 0 Then a = a & _
    ConvertSecToMin(cMedia.Position) & " of " & _
    ConvertSecToMin(cMedia.Duration) & _
    " at " & cMedia.Speed & "× speed."
lblPlayInfo.Caption = a

e = "2"
'are we playing?
If pb_GetPlayState <= 0 Then
    If sldPlay.Enabled <> False Then
        sldPlay.Enabled = False
        If m_bMiniVid Then frmMiniVid.sldSeek.Enabled = False
    End If
    Exit Sub
Else
    If sldPlay.Enabled <> True Then
        sldPlay.Enabled = True
        If m_bMiniVid Then frmMiniVid.sldSeek.Enabled = True
    End If
End If

e = "3"
'draw seek bar
If cMedia.Duration > 0 Then
    sldPlay.Value = cMedia.Position / cMedia.Duration
    If m_bMiniVid Then frmMiniVid.UpdatePogbar
End If

e = "4"
'vid stat
If tsMain.Tabs("vid").Selected Then
    If cMedia.HasVideo Then
        a = "size: " & cMedia.Width & "×" & cMedia.Height & " scaled " & _
            Round((fraVidPlace.Width / Screen.TwipsPerPixelX) / cMedia.Width, 3) & " times to " & _
            Int(fraVidPlace.Width / Screen.TwipsPerPixelX) & "×" & _
            Int(fraVidPlace.Height / Screen.TwipsPerPixelY) & "."
        
        b0 = False
    Else
        a = "no video."
        
        b0 = True
    End If
    
    b1 = ss_GetActive
    If b0 <> b1 Then
        ss_SetActive b0
        b1 = ss_GetActive
    End If
    a = a & "  screensaver is " & IIf(b1, "enabled", "disabled") & "."
    
    If lblVidStat.Caption <> a Then lblVidStat.Caption = a
End If

e = "5"
'goto next item?
If cMedia.State <> stStopped And Val(cMedia.Position) >= cMedia.Duration Then
    'end this track
    i = list_GetCurrentIndex
    If i >= 0 Then
        If pb_lItemSource = 0 Then
            mdb_List.Items(i).lEndCnt = mdb_List.Items(i).lEndCnt + 1
        ElseIf pb_lItemSource > 0 Then
            mdb_PL(pb_lItemSource - 1).Items(mdb_PL(pb_lItemSource - 1).lCurrent).lEndCnt = _
                mdb_PL(pb_lItemSource - 1).Items(mdb_PL(pb_lItemSource - 1).lCurrent).lEndCnt + 1
            If pb_lItemSource - 1 = m_lCurrentPl Then pll_RedrawCurrent
            pll_UpdateStatus True, True ' pl_SetAutoSave True
        Else
            'from a non-list source, no playback data to update.
        End If
    End If
    
    mdb_IncPlyCnt cMedia.FileName, 1, True
    
    i = pb_lItemSource
    
    pb_NextFile
    
    If i = 0 Then
        mdbl_Redraw
    Else
        pll_RedrawCurrent
    End If
End If

Exit Sub
tmrUpdateState_Timer_err:
Main_Err "tmrUpdateState_Timer_err/" & e
err.Clear
End Sub

Private Sub tmrVidBumper_Timer()
tmrVidBumper.Enabled = False
cMedia.BumpVidWindow
End Sub

Private Sub tsMain_Click()
Dim b As Boolean, iNewTab As Long

Select Case tsMain.SelectedItem.Key
    Case "vid": iNewTab = 0
    Case "sys": iNewTab = 1
    Case "lib": iNewTab = 2
    Case "cue": iNewTab = 3
    Case "arc": iNewTab = 5
    Case Else: If Mid$(tsMain.SelectedItem.Key, 1, 2) = "pl" Then iNewTab = 4
End Select

If iNewTab = 4 Then
    m_lCurrentPl = Val(Mid$(tsMain.SelectedItem.Key, 3))
    pll_UpdateStatus
Else
    m_lCurrentPl = -1
End If

Form_Resize

'this bit is for the archive function
Select Case iNewTab
    Case 4
        lstPlArc.Visible = False
    Case 5
        lstPlArc.Visible = True
        pll_RebuildAndPrepArc
End Select

'force the current frame to redraw
Select Case iNewTab
    Case 0: fraVid.ForceResize: fraVidPlace.ForceRedraw
    Case 1: fraPref.ForceResize
    Case 2: fraLib.ForceResize
    Case 3: fraCue.ForceResize
    Case 4, 5: fraPl.ForceResize
End Select

fraVid.Visible = (iNewTab = 0)
fraPref.Visible = (iNewTab = 1)
fraLib.Visible = (iNewTab = 2)
fraCue.Visible = (iNewTab = 3) Or (chkOptCuebtm.Value = 1 And mdb_Cue.lCnt > 0)
fraPl.Visible = (iNewTab = 4) Or (iNewTab = 5)

gen_UpdateVidPos
End Sub

Private Sub tsMain_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbRightButton Then
    Dim i As Long
    
    For i = 1 To tsMain.Tabs.count
        With tsMain.Tabs(i)
            If x + tsMain.Left >= .Left And x + tsMain.Left <= .Left + .Width And _
                Y + tsMain.Top >= .Top And Y + tsMain.Top <= .Top + .Height Then
                
                Select Case Mid$(.Key, 1, 2)
                    Case "pl"
                        tsMain.Tabs(.Key).Selected = True
                        
                        mnuPlSep2.Visible = False
                        mnuPlSelect.Visible = False
                        mnuPlSort.Visible = False
                        mnuPlRep.Visible = False
                        mnuPlSep3.Visible = False
                        mnuPlSel.Visible = False
                        
                        mnuPlEnab.Checked = mdb_PL(m_lCurrentPl).bEnab
                        mnuPlArc.Checked = mdb_PL(m_lCurrentPl).bArc
                        PopupMenu mnuPl
                End Select
                
                Exit For
            End If
        End With
    Next i
End If
End Sub

Private Sub txtLibSearch_Change()
If txtLibSearch.BackColor = vbRed Then txtLibSearch.BackColor = vbWindowBackground
End Sub

Private Sub txtLibSearch_GotFocus()
cmdLibSearch(0).Default = True
txtLibSearch.SelStart = 0
txtLibSearch.SelLength = Len(txtLibSearch.Text)
End Sub

Private Sub txtLibSearch_LostFocus()
cmdLibSearch(0).Default = False
End Sub

Private Sub txtOptFileext_Change()
file_ext_list = txtOptFileext.Text
pref_StartAutoSave
End Sub

Private Sub txtOptHk_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim arrlKeys(0 To 3) As Long, i As Long

If KeyCode = vbKeyShift Or KeyCode = vbKeyControl Or KeyCode = 18 Or Shift <= 0 Then
    If Shift = 0 And (KeyCode = vbKeyBack Or KeyCode = vbKeyDelete) Then
        txtOptHk(Index).Tag = ""
        txtOptHk(Index).Text = ""
    End If
    
    Exit Sub
End If

arrlKeys(0) = KeyCode
arrlKeys(1) = Shift And vbShiftMask
arrlKeys(2) = Shift And vbCtrlMask
arrlKeys(3) = Shift And vbAltMask

With txtOptHk(Index)
    .Tag = ""
    For i = 0 To 3
        .Tag = .Tag & Trim$(Str$(arrlKeys(i))) & "|"
    Next i
    
    .Text = _
        IIf(arrlKeys(1), "shift+", "") & _
        IIf(arrlKeys(2), "ctrl+", "") & _
        IIf(arrlKeys(3), "alt+", "") & _
        arrlKeys(0)
    
    .SelStart = 0
    .SelLength = 0
End With
End Sub

Private Sub vsbCue_Change()
cue_Redraw
End Sub

Private Sub vsbCue_Scroll()
cue_Redraw
End Sub

Private Sub vsbMdbLst_Change()
mdbl_Redraw
End Sub

Private Sub vsbMdbLst_GotFocus()
On Error Resume Next
picMdbLst.SetFocus
On Error GoTo 0
End Sub

Private Sub vsbMdbLst_Scroll()
mdbl_Redraw
End Sub

Private Sub vsbPl_Change()
If m_lCurrentPl < 0 Or vsbPl.Tag <> "" Then Exit Sub

mdb_PL(m_lCurrentPl).lScroll = vsbPl.Value
'Debug.Print Format(Now, "hh:mm:ss") & " set scroll for " & m_lCurrentPl
pll_RedrawCurrent
End Sub

Private Sub vsbPl_Scroll()
vsbPl_Change
End Sub
