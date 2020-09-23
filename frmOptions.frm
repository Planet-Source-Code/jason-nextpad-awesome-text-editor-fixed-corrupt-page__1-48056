VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options - Window"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   1500
   ClientWidth     =   8910
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicOptionsMisc 
      BorderStyle     =   0  'None
      Height          =   3540
      Left            =   2070
      ScaleHeight     =   3540
      ScaleWidth      =   6795
      TabIndex        =   66
      Top             =   180
      Visible         =   0   'False
      Width           =   6795
      Begin VB.Frame Frame4 
         Caption         =   "Miscellaneous"
         Height          =   2550
         Left            =   45
         TabIndex        =   17
         Top             =   720
         Width           =   6600
         Begin VB.CommandButton CmdRemoveFavorites 
            Caption         =   "Clear Favorites"
            Height          =   420
            Left            =   3180
            TabIndex        =   71
            Top             =   225
            Width           =   1620
         End
         Begin VB.CommandButton CmdRemoveRecentFiles 
            Caption         =   "Clear Recent Files"
            Height          =   420
            Left            =   1485
            TabIndex        =   70
            Top             =   225
            Width           =   1620
         End
         Begin VB.CheckBox ChckAllowFavorites 
            Caption         =   "Enable Favorites menu."
            Height          =   195
            Left            =   105
            TabIndex        =   69
            Top             =   1185
            Width           =   2550
         End
         Begin VB.CheckBox ChckAllowRecentFiles 
            Caption         =   "Enable Recent Files menu."
            Height          =   225
            Left            =   105
            TabIndex        =   68
            Top             =   1410
            Width           =   2775
         End
         Begin VB.CheckBox ChckSaveFindHistory 
            Caption         =   "Save find history."
            Height          =   270
            Left            =   105
            TabIndex        =   22
            Top             =   1875
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.CheckBox ChckQuickExit 
            Caption         =   "Enable quick exit (Esc key)."
            Height          =   255
            Left            =   105
            TabIndex        =   20
            Top             =   930
            Width           =   2955
         End
         Begin VB.CommandButton cmdResetSettings 
            Caption         =   "Reset Settings"
            Height          =   420
            Left            =   105
            TabIndex        =   18
            Top             =   225
            Width           =   1290
         End
         Begin VB.CheckBox ChckSaveReplaceHistory 
            Caption         =   "Save replace history."
            Height          =   255
            Left            =   105
            TabIndex        =   23
            Top             =   2115
            Visible         =   0   'False
            Width           =   2520
         End
         Begin VB.CheckBox ChckAutoSave 
            Caption         =   "Auto-Save files instead of asking me."
            Height          =   270
            Left            =   105
            TabIndex        =   21
            Top             =   1635
            Width           =   3150
         End
         Begin VB.CheckBox ChckRemoveDeadRecentFiles 
            Caption         =   "Remove dead recent files on startup."
            Height          =   300
            Left            =   105
            TabIndex        =   19
            Top             =   660
            Visible         =   0   'False
            Width           =   3270
         End
         Begin VB.Label Label11 
            BackColor       =   &H000000FF&
            Caption         =   $"frmOptions.frx":000C
            Height          =   1575
            Left            =   3405
            TabIndex        =   72
            Top             =   765
            Visible         =   0   'False
            Width           =   3015
         End
      End
      Begin VB.Label Label6 
         Caption         =   "Miscellaneous options."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   930
         TabIndex        =   16
         Top             =   180
         Width           =   5670
      End
      Begin VB.Image PicOptions 
         Height          =   480
         Left            =   60
         Top             =   90
         Width           =   720
      End
   End
   Begin MSComDlg.CommonDialog CFontDialog 
      Left            =   6555
      Top             =   810
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.TreeView TvOptions 
      Height          =   3525
      Left            =   75
      TabIndex        =   5
      Top             =   255
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   6218
      _Version        =   393217
      Indentation     =   1058
      LabelEdit       =   1
      Style           =   3
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      SingleSel       =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   720
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":0146
            Key             =   "window"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":046A
            Key             =   "fileassociations"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":0F36
            Key             =   "externaleditor"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":2C72
            Key             =   "printer"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":5426
            Key             =   "fileopening"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":62AA
            Key             =   "textboxoptions"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":6B42
            Key             =   "misc"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame6 
      Caption         =   "Help"
      Height          =   1215
      Left            =   90
      TabIndex        =   44
      Top             =   3780
      Width           =   8730
      Begin VB.TextBox TxtHelp 
         Height          =   855
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   240
         Width           =   8520
      End
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   195
      Top             =   4995
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.PictureBox Picoptionsd 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3780
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1785
         Left            =   2100
         TabIndex        =   35
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox nonusable 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3780
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1785
         Left            =   1545
         TabIndex        =   34
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox ippy 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3780
      Index           =   19
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1785
         Left            =   645
         TabIndex        =   33
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   7725
      TabIndex        =   8
      ToolTipText     =   "Saves any changes you have made without closing this dialog box."
      Top             =   5070
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6525
      TabIndex        =   7
      ToolTipText     =   "Cancels any changes you have made and closes this dialog box."
      Top             =   5070
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5325
      TabIndex        =   6
      ToolTipText     =   "Saves any changes you have made and closes this dialog box."
      Top             =   5070
      Width           =   1095
   End
   Begin VB.PictureBox PicOptionsWindow 
      BorderStyle     =   0  'None
      Height          =   3405
      Left            =   2055
      ScaleHeight     =   3405
      ScaleWidth      =   7125
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   225
      Width           =   7125
      Begin VB.Frame Frame12 
         Caption         =   "Browsebar"
         Height          =   855
         Left            =   15
         TabIndex        =   63
         Top             =   2505
         Width           =   6765
         Begin VB.CommandButton CmdBrowseBarAdvanced 
            Caption         =   "Advanced..."
            Height          =   375
            Left            =   2370
            TabIndex        =   65
            Top             =   300
            Width           =   1095
         End
         Begin VB.CheckBox ChckBrowsebar 
            Caption         =   "Always show &Browsebar."
            Height          =   210
            Left            =   120
            TabIndex        =   64
            Top             =   375
            Width           =   2565
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "View"
         Height          =   855
         Left            =   0
         TabIndex        =   48
         Top             =   600
         Width           =   6780
         Begin VB.CheckBox ChckLastWinPos 
            Caption         =   "&Remember main windows previous size and position."
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   360
            Width           =   6375
         End
      End
      Begin VB.Frame fraSample1 
         Caption         =   "Toolbar"
         Height          =   855
         Left            =   0
         TabIndex        =   29
         Top             =   1575
         Width           =   6780
         Begin VB.CheckBox Check2 
            Caption         =   "&Always show toolbar  (default)."
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   360
            Width           =   3735
         End
      End
      Begin VB.Label Label10 
         Caption         =   "Options for altering NextPad's view."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   765
         TabIndex        =   47
         Top             =   120
         Width           =   4695
      End
      Begin VB.Image PicWindow 
         Height          =   480
         Left            =   135
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.PictureBox PicOptionsExternalEditor 
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   2055
      ScaleHeight     =   2655
      ScaleWidth      =   6960
      TabIndex        =   38
      Top             =   195
      Visible         =   0   'False
      Width           =   6960
      Begin VB.Frame Frame2 
         Caption         =   "External Editor"
         Height          =   2055
         Left            =   135
         TabIndex        =   39
         Top             =   480
         Width           =   6585
         Begin VB.CheckBox Chckaskiftoobig 
            Caption         =   "Auto-launch external editor when file is to large for NextPad to open."
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   720
            Width           =   6000
         End
         Begin VB.Frame Frame3 
            Caption         =   "Current External Editor"
            Height          =   735
            Left            =   120
            TabIndex        =   42
            Top             =   1080
            Width           =   5055
            Begin VB.CommandButton cmdChooseexternaleditor 
               Caption         =   "Select &External Editor ......"
               Height          =   375
               Left            =   120
               TabIndex        =   43
               Top             =   240
               Width           =   4815
            End
         End
         Begin VB.CheckBox ChckExternalEditor 
            Caption         =   "&Use external editor to open files To large for NextPad to open."
            Height          =   375
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Width           =   5175
         End
      End
      Begin VB.Label Label8 
         Caption         =   "Options for configuring the external editor."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   45
         Top             =   120
         Width           =   4695
      End
      Begin VB.Image PicExternalEditor 
         Height          =   480
         Left            =   0
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.PictureBox PicOptionsFileAssociations 
      BorderStyle     =   0  'None
      Height          =   2700
      Left            =   2070
      ScaleHeight     =   2700
      ScaleMode       =   0  'User
      ScaleWidth      =   7005
      TabIndex        =   36
      Top             =   180
      Visible         =   0   'False
      Width           =   7005
      Begin VB.Frame Frame1 
         Caption         =   "File associations"
         Height          =   2010
         Left            =   120
         TabIndex        =   37
         Top             =   615
         Width           =   6645
         Begin VB.CheckBox Chckassociations 
            Caption         =   "&Notify me when NextPad does not currently open text documents on this computer."
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   720
            Width           =   6335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "&Allow NextPad to open text documents on this computer."
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Width           =   5215
         End
      End
      Begin VB.Label Label9 
         Caption         =   "Options for NextPad's file Association."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   46
         Top             =   120
         Width           =   4815
      End
      Begin VB.Image PicFileAssociations 
         Height          =   480
         Left            =   0
         Top             =   0
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   375
         Left            =   1560
         Top             =   1440
         Width           =   615
      End
   End
   Begin VB.PictureBox PicOptionsPrinter 
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   2070
      ScaleHeight     =   2655
      ScaleWidth      =   7050
      TabIndex        =   50
      Top             =   195
      Visible         =   0   'False
      Width           =   7050
      Begin VB.Frame Frame7 
         Caption         =   "Printer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   120
         TabIndex        =   51
         Top             =   600
         Width           =   6705
         Begin VB.CommandButton CmdPrintSetup 
            Caption         =   "&Print Setup..."
            Height          =   495
            Left            =   3300
            TabIndex        =   53
            Top             =   810
            Width           =   2895
         End
         Begin VB.CommandButton CmdPageSetup 
            Caption         =   "Page &Setup..."
            Height          =   495
            Left            =   150
            TabIndex        =   52
            Top             =   810
            Width           =   2895
         End
      End
      Begin VB.Label Label2 
         Caption         =   "Options for configuring and setting up how NextPad prints Documents."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   61
         Top             =   120
         Width           =   6935
      End
      Begin VB.Image PicPrinter 
         Height          =   600
         Left            =   120
         Top             =   0
         Width           =   675
      End
   End
   Begin VB.PictureBox picOptionsTextBox 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3435
      Left            =   2040
      ScaleHeight     =   3435
      ScaleWidth      =   6780
      TabIndex        =   62
      Top             =   150
      Visible         =   0   'False
      Width           =   6780
      Begin VB.Frame Frame9 
         Caption         =   "TextBox"
         Height          =   1860
         Left            =   30
         TabIndex        =   2
         Top             =   585
         Width           =   6435
         Begin VB.PictureBox PicTextColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   1500
            ScaleHeight     =   285
            ScaleWidth      =   1710
            TabIndex        =   10
            TabStop         =   0   'False
            ToolTipText     =   "Click to see more colors which you can choose..."
            Top             =   645
            Width           =   1740
         End
         Begin VB.PictureBox PicBackColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   1500
            ScaleHeight     =   285
            ScaleWidth      =   1710
            TabIndex        =   9
            TabStop         =   0   'False
            ToolTipText     =   "Click to see more colors which you can choose..."
            Top             =   300
            Width           =   1740
         End
         Begin VB.CheckBox ChckWordWrap 
            Caption         =   "&Word Wrap"
            Height          =   255
            Left            =   1665
            TabIndex        =   14
            Top             =   1485
            Width           =   1455
         End
         Begin VB.CommandButton CmdFont 
            Caption         =   "&Font..."
            Height          =   360
            Left            =   105
            TabIndex        =   11
            Top             =   1005
            Width           =   1230
         End
         Begin VB.CheckBox ChckItalic 
            Caption         =   "&Italic"
            Height          =   225
            Left            =   855
            TabIndex        =   13
            Top             =   1500
            Width           =   705
         End
         Begin VB.CheckBox ChckBold 
            Caption         =   "&Bold"
            Height          =   195
            Left            =   105
            TabIndex        =   12
            Top             =   1515
            Width           =   660
         End
         Begin VB.Frame Frame10 
            Caption         =   "Sample"
            Height          =   1575
            Left            =   3330
            TabIndex        =   24
            Top             =   120
            Width           =   2940
            Begin VB.TextBox TxtSample 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1110
               Left            =   165
               TabIndex        =   15
               Text            =   "Sample Text"
               Top             =   285
               Width           =   2595
            End
         End
         Begin VB.Label Label7 
            Caption         =   "&Text Color:"
            Height          =   210
            Left            =   120
            TabIndex        =   0
            Top             =   675
            Width           =   1290
         End
         Begin VB.Label Label5 
            Caption         =   "&Background Color:"
            Height          =   255
            Left            =   120
            TabIndex        =   1
            Top             =   330
            Width           =   1380
         End
      End
      Begin VB.Label Label4 
         Caption         =   "These settings allow you to change specific properties of the textbox."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   735
         TabIndex        =   67
         Top             =   135
         Width           =   6000
      End
      Begin VB.Image PicTextBox 
         Height          =   345
         Left            =   60
         Top             =   135
         Width           =   450
      End
   End
   Begin VB.PictureBox PicOptionsFileOpening 
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   2040
      ScaleHeight     =   2655
      ScaleWidth      =   6930
      TabIndex        =   54
      Top             =   210
      Visible         =   0   'False
      Width           =   6930
      Begin VB.Frame Frame11 
         Caption         =   "File Opening"
         Height          =   1935
         Left            =   120
         TabIndex        =   55
         Top             =   600
         Width           =   6645
         Begin VB.OptionButton optOpenMethod 
            Caption         =   "Use binary open method (Slower), most file and character compatibility."
            Height          =   375
            Index           =   1
            Left            =   135
            TabIndex        =   58
            Top             =   1245
            Width           =   6060
         End
         Begin VB.OptionButton optOpenMethod 
            Caption         =   "Use input open method (Fastest), least file and character compatibility."
            Height          =   375
            Index           =   0
            Left            =   135
            TabIndex        =   57
            Top             =   810
            Width           =   5520
         End
         Begin VB.CheckBox ChckSmartFileOpening 
            Caption         =   "Use Smart File Opening "
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   240
            Width           =   2055
         End
         Begin VB.Frame Frame8 
            Caption         =   "File Open Method"
            Height          =   1200
            Left            =   105
            TabIndex        =   59
            Top             =   570
            Width           =   6180
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Options for altering the way NextPad opens file's."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   60
         Top             =   105
         Width           =   6015
      End
      Begin VB.Image PicFileOpening 
         Height          =   480
         Left            =   120
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.Label Label3 
      Caption         =   "&Options :"
      Height          =   195
      Left            =   75
      TabIndex        =   4
      Top             =   30
      Width           =   1920
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private OpenMethod As Integer


Private Sub ChckAllowFavorites_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  TxtHelp.Text = "This option allows the Favorites menu to be visible."
End Sub


Private Sub ChckAllowRecentFiles_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  TxtHelp.Text = "This option allows the Recent Files menu to be visible."
End Sub

Private Sub Chckaskiftoobig_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  TxtHelp.Text = "If checked then the next time a file is to large for NextPad to open, it will open it with the external editor automatically."
End Sub
Private Sub Chckassociations_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  TxtHelp.Text = "If enabled, this option would allow NextPad to notify the user when it does not currently open text documents on the user's computer."
End Sub

Private Sub ChckAutoSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  TxtHelp.Text = "If you choose this option NextPad will not notify you if you have made any changes to a file because it will automatically make the changes for you."
End Sub

Private Sub ChckBold_Click()
  TxtSample.FontBold = CBool(ChckBold.Value)
End Sub
Private Sub ChckBold_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  TxtHelp.Text = "Toggles Bold property of TextBox"
End Sub
Private Sub ChckBrowsebar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  TxtHelp.Text = "Enables the Browsebar to always be shown which allows you to browse through text files with ease."
End Sub
Private Sub ChckExternalEditor_Click()
    Select Case ChckExternalEditor.Value
     Case 0
       Frame3.Enabled = False
       cmdChooseexternaleditor.Enabled = False
       Chckaskiftoobig.Enabled = False
     Case 1
       Frame3.Enabled = True
       cmdChooseexternaleditor.Enabled = True
       Chckaskiftoobig.Enabled = True
    End Select
End Sub
Private Sub ChckExternalEditor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  TxtHelp.Text = "If a file is too large for NextPad too open , It will prompt the user too open the file with an Editor that can."
End Sub
Private Sub ChckItalic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  TxtHelp.Text = "Toggles Italic property of TextBox"
End Sub


Private Sub ChckQuickExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  TxtHelp.Text = "This option will allow you to quickly exit NextPad from the main window by pressing the Escape (ESC) key. Please note however that you will not be asked to save any document you are working on."
End Sub

Private Sub ChckSmartFileOpening_Click()
Dim bVal As Boolean
  bVal = CBool(ChckSmartFileOpening.Value) ' Retreive the boolean value and append it to variable
  bVal = Not bVal ' Make sure the opposite value is returned
  optOpenMethod(0).Enabled = bVal ' Append value
  optOpenMethod(1).Enabled = bVal ' Append value
End Sub
Private Sub Chckwordwrap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  TxtHelp.Text = "Enable this setting to wrap text in NextPads Window to the window."
End Sub
Private Sub ChckLastWinPos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  TxtHelp.Text = "If enabled NextPad will remember it's  Previous Window State and Position, Else it starts always in the center of the screen."
End Sub
Private Sub ChckSmartFileOpening_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  TxtHelp.Text = "Allows NextPad to decide the best open file method when opening a file (highly recommended)."
   '"This option is the most smartest option available , when enabled it allows the use of what is called " _
   '& " Smart File Opening , what this does is this , it enables input open method ( no binary ) which is relatively the fastest open " _
   '& " file method available too NextPad. Although the input method does not have all the character ability of the binary open method " _
   '& " it gives speed and lowers the risk of memory loss (bytes in memory).Now Heres is where smart file opening comes in. " _
   '& " Usually if you are opening a file using the input open method you may get an error stating that the file is too large too be opened " _
   '& " but thats not the case thats just the error control picking it up , why ? because the input open method does not have the same " _
   '& " amount of character and file compatitbility as the binary open method , but when using smart file opening ; if the file is under " _
   '& " the maximum file size (65000 bytes) then nextpad opens the file in the binary open method , else if everything goes well youll see a big " _
   '& " improvement in file opening speed thanks too a new feature called smart file opening "
End Sub
Private Sub Check1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  TxtHelp.Text = "This option will allow NextPad to open Text documents on this computer."
End Sub
Private Sub Check2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  TxtHelp.Text = "This option If enabled always shows the Toolbar (You can toggle the Toolbar ON and OFF in the View Menu)."
End Sub
Private Sub ChckItalic_Click()
   TxtSample.FontItalic = CBool(ChckItalic.Value)
End Sub
Private Sub cmdApply_Click()
   Call SaveMainSettings
   ResizeNoteWithToolbar
End Sub
Private Sub CmdBrowseBarAdvanced_Click()
   FrmBrowseBarOptions.Show vbModal
End Sub
Private Sub CmdBrowseBarAdvanced_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   TxtHelp.Text = "Displays advanced options to configure the browse bar."
End Sub
Private Sub CmdCancel_Click()
   Unload Me
   ResizeNoteWithToolbar
End Sub
Private Sub cmdChooseexternaleditor_Click()

   On Error GoTo CdlcCancelErr:
      CDialog.Flags = cdlOFNHideReadOnly + cdlOFNFileMustExist
      CDialog.Filter = "Executable Files(*.EXE) |*.EXE"
      CDialog.DialogTitle = "Choose An External Viewer IE ;  Wordpad.EXE "
      CDialog.Cancelerror = True
    CDialog.ShowOpen
 
   If CDialog.Filename <> "" Then
      SaveSetting "ExternalEditor", "Path", CDialog.Filename
      cmdChooseexternaleditor.Caption = CDialog.Filename
   End If
 
CdlcCancelErr:
     If Err.Number = 32755 Then
        Exit Sub
     End If
End Sub
Private Sub cmdChooseexternaleditor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  TxtHelp.Text = "Allows you too choose an External editor other than the default that is already being used."
  cmdChooseexternaleditor.ToolTipText = cmdChooseexternaleditor.Caption
End Sub
Private Sub CmdFont_Click()
  On Error GoTo Cancelerror: ' if an error occurs goto that line
  
  CFontDialog.Flags = cdlCFBoth  ' set the Font Dialogs flags
  CFontDialog.ShowFont ' show the Select Font Dialog
  ' set the forms active control's font too the selected font form the dialog
  TxtSample.Font = CFontDialog.fontname
  ' set the forms active control's Bold too the Selected font dialog bold
  TxtSample.FontBold = CFontDialog.FontBold
  ' set the form's active control's Font size too the selected font size from the dialog
  TxtSample.fontsize = CFontDialog.fontsize
  ' save in the registry ; the forms active control's italic (BOOLEAN) too the selected italic form the dialog
  TxtSample.FontItalic = CFontDialog.FontItalic
  ' save in the registry ; The Font For the Form's Text Box
  ' For later Retrieval From the registry
  'SaveSetting "Font", "font", FrmMain.mTextBox.fontname
  ' save in the registry ; the controls current Font size for later
  ' retrieval from the registry
  'SaveSetting "Font", "Fontsize", FrmMain.mTextBox.fontsize
  ' Save in the Registry ; the controls current Bold Attribute
  ' Either TRUE (-1) or FALSE (0)
  'SaveSetting "Font", "FontBold", Abs(FrmMain.mTextBox.FontBold)
  ' Save in the registry ; the controls current italic Attribute
  ' Either TRUE (-1) or FALSE (0)
  'SaveSetting "Font", "FontItalic", Abs(FrmMain.mTextBox.FontItalic)
  
Cancelerror: ' if an error occurs Vb will Jump here
  Exit Sub ' Exit the sub immediately

End Sub
Private Sub CmdFont_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  TxtHelp.Text = "Allows you to change specific properties of the TextBox such as font,bold,italic,size,etc..."
End Sub
Private Sub CmdOk_Click()
  Call SaveMainSettings
  Unload Me
  ResizeNoteWithToolbar
End Sub
Private Sub SaveMainSettings()

   
   Dim StrPathAndExe As String 'String Value to Hold the Correct path and filename of this execuatable
   
   Dim Retval As String ' Declare String variable
    
   Dim Strbackup As String ' Declare String variable
   
   Dim BfState As Boolean ' Declare Boolean variables
   
   'Set the StrPathAndExe String value to hold the Applications current path and Executable name
   StrPathAndExe = App.Path & "\" & App.EXEName & ".EXE"
   ' If the file is in the root Directory "C:\" Then We remove the appended
   ' Backslash \ to fit the current path and executable filename
   If bFileExists(StrPathAndExe) = False Then StrPathAndExe = App.Path & App.EXEName & ".EXE"
   

  If Check1.Value = vbChecked Then
          'Set NextPad as the default TXT Viewer
             'Save the current TXT Viewer Path  for later retrieval
           SaveSetting "Associations", "TXT" _
           , GetSettingString(HKEY_CLASSES_ROOT, "txtfile\shell\open\command", "", "")
           'Save the New one (NextPad)
           SaveSettingString HKEY_CLASSES_ROOT, "txtfile\shell\open\command" _
           , "", StrPathAndExe & " %1"
           SaveSetting "associations", "isassociated", "1"
  Else 'The CheckBox Was Deselected
            'Restore the previous TXT viewer in the registry
  SaveSetting "associations", "isassociated", "0"

  If GetSetting("NextPad", "Associations", "TXT", "") <> "" Then
       SaveSettingString HKEY_CLASSES_ROOT, "txtfile\shell\open\command" _
       , "", GetSetting("NextPad", "Associations", "TXT", "")
  End If
  End If



SaveSetting "OpenMethod", "OpenMethod", Abs(OpenMethod)

'*****************************************************************
' the Subs Here are in ModOptions , they work depending on the
' Checks value then save the setting requested associated with the sub.
      
      SaveSetting_Toolbar CBool(Check2.Value)

      SaveSetting_chckassociations CBool(Chckassociations.Value)

      SaveSetting_UseExternalEditor CBool(ChckExternalEditor.Value)

        
      Select Case ChckSmartFileOpening.Value
        Case 1
          SaveSetting "OpenMethod", "UseSmartFileOpening", 1
          SaveSetting "OpenMethod", "OpenMethod", 0
        Case 0
          SaveSetting "OpenMethod", "UseSmartFileOpening", 0
      End Select
        
      SaveSetting_AutoLaunchExtEditor CBool(Chckaskiftoobig.Value)
      
      SaveSetting_RememberLastWinPos CBool(ChckLastWinPos.Value)
      
     ' Textbox
     '------------------------------------------------------------
      'BackColor
      SaveSetting "Font", "BackColor", CStr(CLng(TxtSample.BackColor))
      'ForeColor
      SaveSetting "Font", "ForeColor", CStr(CLng(TxtSample.ForeColor))
      'FontSize
      SaveSetting "Font", "Fontsize", CInt(TxtSample.fontsize)
      'FontBold
      SaveSetting "Font", "FontBold", Abs(CInt(CBool(TxtSample.FontBold)))
      'FontItalic
      SaveSetting "Font", "FontItalic", Abs(CInt(CBool(TxtSample.FontItalic)))
      'Font
      SaveSetting "Font", "font", CStr(TxtSample.fontname)
      'WordWrap
      SaveSetting "WordWrap", "Wordwrap", CInt(ChckWordWrap.Value)
      'Reload the TextBoxes
      'Backup Current text to prevent loss
      Strbackup = FrmMain.mTextbox.Text 'Backup Current text to prevent loss
      'Save previous Fstate (File State)
      BfState = CBool(Fstate.dirty)
      'Toggle Wordwrap
      ToggleWordWrap CBool(GetSetting("NextPad", "WordWrap", "Wordwrap", 0))
      'Place the text back inside
      FrmMain.mTextbox.Text = Strbackup ' Place the text back inside
      'Return Previous (File State)
      Fstate.dirty = Abs(CInt(BfState))
      '------------------------------------------------------------

   ' BrowseBar
   '------------------------------------------------------------
   ResizeNoteWithToolbar 'Resize the Form To Match
   ' The Current state
   FrmMain.PicBrowseBar.Visible = Abs(CInt(ChckBrowsebar.Value)) 'Show the Browser Bar
   ' Using a Logical Expression Not (The Opposite if each , eg; 1 = 0 0 = 1)
   FrmMain.MnuBrowseBar.Checked = Abs(CInt(ChckBrowsebar.Value))
   ResizeNoteWithToolbar  'Resize the Form To Match the
   ' current state
   SaveSetting "BrowseBar", "Visible", Abs(CInt(ChckBrowsebar.Value))
   ' Current state
   '------------------------------------------------------------
   
   ' Miscellaneous
   '------------------------------------------------------------
   ' Dead recent files
   SaveSetting "RecentFiles", "RemoveDead", Abs(CInt(ChckRemoveDeadRecentFiles.Value))
   SaveSetting "RecentFiles", "Enable", Abs(CInt(ChckAllowRecentFiles.Value))
   ' If the user no longer wants recent files, eradicate them
   If Abs(CInt(ChckAllowRecentFiles.Value)) = 0 Then
   'On Error Resume Next
   DeleteSetting "NextPad", "RecentFiles"
   SaveSetting "RecentFiles", "Enable", Abs(CInt(ChckAllowRecentFiles.Value))
   With FrmMain
     .MnuRecentFiles(0).Caption = "": .MnuRecentFiles(0).Visible = False
     On Error Resume Next
   
     Dim i As Integer, X As Integer
      'Get the menu count
      i = .MnuRecentFiles.Count
      For X = 1 To i Step 1 ' Go through them one by one
        If X = i Then Exit For ' If the variable is equal to the count then stop
        Unload .MnuRecentFiles(X) ' Unload the specified menu
      Next X ' Next X value
   End With
   
   End If
   ' Quick exit
   SaveSetting "QuickExit", "QuickExit", Abs(CInt(ChckQuickExit.Value))
   ' Auto-save
   SaveSetting "AutoSave", "AutoSave", Abs(CInt(ChckAutoSave.Value))
   ' Find history
   SaveSetting "FindHistory", "Save", Abs(CInt(ChckSaveFindHistory.Value))
   ' Replace History
   SaveSetting "ReplaceHistory", "Save", Abs(CInt(ChckSaveReplaceHistory.Value))
   ' Favorites
   SaveSetting "Favorites", "Enable", Abs(CInt(ChckAllowFavorites.Value))
   FrmMain.MnuFavoritesItem.Visible = Abs(CInt(ChckAllowFavorites.Value))
   '******************************************************************

End Sub
Private Sub CmdPageSetup_Click()
 On Error GoTo ObjectUnloaded:

   FrmPageSetup.Show (vbModal)

ObjectUnloaded:
  If Err.Number <> 0 Then
     Exit Sub
  End If
End Sub
Private Sub CmdPageSetup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  TxtHelp.Text = "This when pressed loads the Page Setup dialog enabling you too create Custom Page Settings."
End Sub
Private Sub CmdPrintSetup_Click()
  On Error GoTo CdlcCancelError: 'If an error Occurs VB WIll Jump too that line

   'Set the Common Dialogs Flags too display the Print Setup DIalog
   ' Rather than the Print DIalog
   FrmMain.CPrintSetupDialog.Flags = cdlPDPrintSetup
     ' Show The Dialog
     FrmMain.CPrintSetupDialog.ShowPrinter
 
 Exit Sub ' Exit Sub Immediatelly

CdlcCancelError:
   If Err.Number = 32755 Then 'if cancel was Pressed (CldlcCancelError) then ......
     Exit Sub
   Else ' Else another error Has occured Other than the Cancel Error
     MsgBox "Could not display the Print Setup window, " & Err.Description, vbCritical, "NextPad - Print"
   End If

End Sub
Private Sub CmdPrintSetup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  TxtHelp.Text = "This when pressed opens the Print Setup dialog enabling you too setup any Printer(s) that you have installed."
End Sub

Private Sub CmdRemoveFavorites_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  TxtHelp.Text = "This button will remove all Favorites from the registry"
End Sub

Private Sub CmdRemoveRecentFiles_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  TxtHelp.Text = "This button will remove all Recent Files from the registry"
End Sub

Private Sub cmdResetSettings_Click()
 On Error Resume Next
  If MsgBox("Are you sure that you would like to reset NextPads settings?" & _
     vbNewLine & "(You will lose all of the shortcuts to your recent files and favorites)", vbExclamation + vbYesNo, "NextPad") = vbYes Then
     DeleteSetting "NextPad"
     MsgBox "NextPad's settings in the registry have been removed.", vbInformation, "NextPad"
  Else
     Exit Sub
  End If
  
End Sub
Private Sub cmdResetSettings_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  TxtHelp.Text = "Once this button is clicked, any settings that were previously in the registry would be removed. This is helpful for trouble shooting purposes."
End Sub

Private Sub CmdRemoveRecentFiles_Click()
On Error Resume Next
DeleteSetting "NextPad", "RecentFiles"

With FrmMain
     .MnuRecentFiles(0).Caption = "": .MnuRecentFiles(0).Visible = False
     On Error Resume Next
   
     Dim i As Integer, X As Integer
      'Get the menu count
      i = .MnuRecentFiles.Count
      For X = 1 To i Step 1 ' Go through them one by one
        If X = i Then Exit For ' If the variable is equal to the count then stop
        Unload .MnuRecentFiles(X) ' Unload the specified menu
      Next X ' Next X value
End With

MsgBox "The list of Recent Files has been succesfully removed from the registry", vbInformation, "NextPad"

End Sub
Private Sub CmdRemoveFavorites_Click()
On Error Resume Next
DeleteSetting "NextPad", "Favorites"
FrmMain.MnuFavoritesItem.Visible = False
MsgBox "The list of Favorites has been succesfully removed from the registry", vbInformation, "NextPad"

End Sub

Private Sub Form_Load()
    
 With ImageList1
     PicWindow.Picture = .ListImages("window").Picture
     PicFileOpening.Picture = .ListImages("fileopening").Picture
     PicPrinter.Picture = .ListImages("printer").Picture
     PicExternalEditor.Picture = .ListImages("externaleditor").Picture
     PicFileAssociations.Picture = .ListImages("fileassociations").Picture
     PicTextBox.Picture = .ListImages("textboxoptions").Picture
     PicOptions.Picture = .ListImages("misc").Picture
 End With
      
 With TvOptions
    .Nodes.Add , , "Window", "Window", "window"
    .Nodes.Add , , "FileAssociations", "File Associations", "fileassociations"
    .Nodes.Add , , "FileOpening", "File Opening", "fileopening"
    .Nodes.Add , , "Printer", "Printer", "printer"
    .Nodes.Add , , "ExternalEditor", "External Editor", "externaleditor"
    .Nodes.Add , , "TextBoxOptions", "TextBox", "textboxoptions"
    .Nodes.Add , , "misc", "Miscellaneous", "misc"
 End With
    
  
  PicBackColor.BackColor = GetSetting("NextPad", "Font", "BackColor", vbWhite)
  PicTextColor.BackColor = GetSetting("NextPad", "Font", "ForeColor", vbBlack)
  TxtSample.BackColor = GetSetting("NextPad", "Font", "BackColor", vbWhite)
  TxtSample.ForeColor = GetSetting("NextPad", "Font", "ForeColor", vbBlack)
 On Error Resume Next
  TxtSample.Font.Name = GetSetting("NextPad", "Font", "Font", "Ms Sans Serif")
  TxtSample.Font.Size = GetSetting("NextPad", "Font", "Fontsize", "8")
  ChckBold.Value = CInt(GetSetting("NextPad", "Font", "FontBold", "0"))
  ChckItalic.Value = CInt(GetSetting("NextPad", "Font", "FontItalic", "0"))
  ChckWordWrap.Value = CInt(GetSetting("NextPad", "WordWrap", "Wordwrap", 1))
  TxtSample.FontBold = CBool(GetSetting("NextPad", "Font", "FontBold", "0"))
  TxtSample.FontItalic = CBool(GetSetting("NextPad", "Font", "FontItalic", "0"))
  
 Dim Retval As String, OpenMethod As Integer
 Dim r As String, Response
  
    r = GetSettingString(HKEY_CLASSES_ROOT, _
        "txtfile\Shell\open\command", _
        "", "")

    If r = App.Path & "\" & App.EXEName & ".EXE" & " %1" Or r = App.Path & App.EXEName & ".EXE" & " %1" Then
       Check1.Value = vbChecked
    Else
       Check1.Value = vbUnchecked
    End If
    ' Associations
       Chckassociations.Value = Abs(CInt(Check_Associations_At_Startup))
    ' Toolbar
       Check2.Value = CInt(GetSetting("NextPad", "Toolbar", "Visible", 1))
    ' External editor
       ChckExternalEditor.Value = Abs(CInt(UseExternalEditor))
    ' More
       Chckaskiftoobig.Value = Abs(CInt(AutoLaunchExtEditor))

      Retval = GetSetting("NextPad", "ExternalEditor", "Path", "")
      If Retval <> "" Then
         cmdChooseexternaleditor.Caption = Retval
      Else
         cmdChooseexternaleditor.Caption = DetectExternalEditor
      End If
      
      Select Case Abs(CInt(UseExternalEditor))
       Case 0
         Frame3.Enabled = False
         cmdChooseexternaleditor.Enabled = False
         Chckaskiftoobig.Enabled = False
       Case 1
         Frame3.Enabled = True
         cmdChooseexternaleditor.Enabled = True
         Chckaskiftoobig.Enabled = True
      End Select
    '...........................................
     ' Win position
      ChckLastWinPos.Value = Abs(CInt(RememberLastWinPos))
     ' File opening
      OpenMethod = GetSetting("NextPad", "OpenMethod", "OpenMethod", 0)
     
      optOpenMethod.Item(OpenMethod).Value = True
     
     If Abs(CInt(UseSmartFileOpening)) Then
        ChckSmartFileOpening.Value = vbChecked
        optOpenMethod.Item(0).Value = True
        optOpenMethod.Item(0).Enabled = False
        optOpenMethod.Item(1).Enabled = False
     Else
        ChckSmartFileOpening.Value = vbUnchecked
     End If
        ' BrowseBar
        ChckBrowsebar.Value = Abs(CInt(BrowseBar_Show))
        ' Recent Files
        ChckRemoveDeadRecentFiles.Value = Abs(CInt(RemoveDeadRecentFiles))
        ChckAllowRecentFiles.Value = Abs(CInt(AllowRecentFiles))
        ' Favorites
        ChckAllowFavorites.Value = Abs(CInt(AllowFavorites))
        ' Quick exit
        ChckQuickExit.Value = Abs(CInt(QuickExit))
        ' Auto-save
        ChckAutoSave.Value = Abs(CInt(AutoSave))
        ' Find history
        ChckSaveFindHistory.Value = Abs(CInt(RememberFindHistory))
        ' Replace history
        ChckSaveReplaceHistory.Value = Abs(CInt(RememberReplaceHistory))
' Make these borders high speed looking
On Error Resume Next
MakeBorderStatic (PicTextColor.hWnd)
MakeBorderStatic (PicBackColor.hWnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  ' Release form from  memory.
  Set frmOptions = Nothing
End Sub
Private Sub OptOpenMethod_Click(Index As Integer)
  OpenMethod = Index
End Sub
Private Sub optOpenMethod_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Select Case Index
    Case 0
      TxtHelp.Text = "Input open statement (fastest) least file and character compatibility , also some text files may not open , note: if you experience problems when opening files under the maximum of 65000 bytes (65k) then use the binary open method.It is recommended that you allow NextPad to decide the best open file method (enabled by default)."
    Case 1
      TxtHelp.Text = "Binary open statement (slower) most file and character compatibility (default) , note: this is the slowest method but it is also the most reliable and compatible with all text files ,unlike the input open method it handles all and most text files under 65000 bytes ( 65k ) (recommended).It is recommended that you allow NextPad to decide the best open file method (enabled by default)."
  End Select
End Sub
Private Sub PicBackColor_Click()
  Dim Retval
       Retval = ShowColorDlg
       If Retval = Null Then 'If Cancel Was Pressed Then NULL Will be returned
           Exit Sub 'Exit the sub
       Else   'A Value other Than NULL Exists
           On Error Resume Next
           TxtSample.BackColor = Retval 'Set the Forecolor
           PicBackColor.BackColor = Retval
           Exit Sub
       End If
End Sub
Private Sub PicBackColor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   TxtHelp.Text = "Sets the background color of NextPad's textbox ; Click to see more colors which you can choose..."
End Sub
Private Sub PicTextColor_Click()
  Dim Retval
       Retval = ShowColorDlg
       If Retval = Null Then 'If Cancel Was Pressed Then NULL Will be returned
           Exit Sub 'Exit the sub
       Else   'A Value other Than NULL Exists
           On Error Resume Next
           TxtSample.ForeColor = Retval 'Set the Forecolor
           PicTextColor.BackColor = Retval
           Exit Sub
       End If

End Sub
Private Sub PicTextColor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   TxtHelp.Text = "Sets the color  of text displayed in NextPad's textbox ; Click to see more colors which you can choose..."
End Sub
Private Sub TvOptions_NodeClick(ByVal Node As MSComctlLib.Node)
 
  TxtHelp.Text = ""
  Select Case Node.Key
      Case "misc"
         PicOptionsMisc.Visible = True
         PicOptionsFileAssociations.Visible = False
         PicOptionsFileOpening.Visible = False
         PicOptionsWindow.Visible = False
         PicOptionsPrinter.Visible = False
         PicOptionsExternalEditor.Visible = False
         picOptionsTextBox.Visible = False
         Me.Caption = "Options - Miscellaneous"
      Case "FileAssociations"
         PicOptionsFileAssociations.Visible = True
         PicOptionsFileOpening.Visible = False
         PicOptionsWindow.Visible = False
         PicOptionsPrinter.Visible = False
         PicOptionsExternalEditor.Visible = False
         picOptionsTextBox.Visible = False
         PicOptionsMisc.Visible = False
         Me.Caption = "Options - File Associations"
      Case "FileOpening"
         PicOptionsFileAssociations.Visible = False
         PicOptionsFileOpening.Visible = True
         PicOptionsWindow.Visible = False
         PicOptionsPrinter.Visible = False
         PicOptionsExternalEditor.Visible = False
         picOptionsTextBox.Visible = False
         PicOptionsMisc.Visible = False
         Me.Caption = "Options - File Opening"
      Case "Printer"
         PicOptionsFileAssociations.Visible = False
         PicOptionsFileOpening.Visible = False
         PicOptionsWindow.Visible = False
         PicOptionsPrinter.Visible = True
         PicOptionsExternalEditor.Visible = False
         picOptionsTextBox.Visible = False
         PicOptionsMisc.Visible = False
         Me.Caption = "Options - Printer"
      Case "ExternalEditor"
         PicOptionsFileAssociations.Visible = False
         PicOptionsFileOpening.Visible = False
         PicOptionsWindow.Visible = False
         PicOptionsPrinter.Visible = False
         PicOptionsExternalEditor.Visible = True
         PicOptionsMisc.Visible = False
         picOptionsTextBox.Visible = False
         Me.Caption = "Options - External Editor"
      Case "Window"
         PicOptionsFileAssociations.Visible = False
         picOptionsTextBox.Visible = False
         PicOptionsFileOpening.Visible = False
         PicOptionsWindow.Visible = True
         PicOptionsPrinter.Visible = False
         PicOptionsMisc.Visible = False
         PicOptionsExternalEditor.Visible = False
         Me.Caption = "Options - Window"
      Case "TextBoxOptions"
         picOptionsTextBox.Visible = True
         PicOptionsFileAssociations.Visible = False
         PicOptionsFileOpening.Visible = False
         PicOptionsWindow.Visible = False
         PicOptionsMisc.Visible = False
         PicOptionsPrinter.Visible = False
         PicOptionsExternalEditor.Visible = False
         Me.Caption = "Options - TextBox"
   End Select
End Sub
Private Sub TxtHelp_KeyPress(KeyAscii As Integer)
   Beep
End Sub
Private Sub TxtSample_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   TxtHelp.Text = "Shows how your TextBox will look as you change its properties"
End Sub
