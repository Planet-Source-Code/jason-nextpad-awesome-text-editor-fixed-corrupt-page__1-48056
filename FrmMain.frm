VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmMain 
   Caption         =   "Untitled - NextPad"
   ClientHeight    =   5700
   ClientLeft      =   1470
   ClientTop       =   2220
   ClientWidth     =   8775
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "FrmMain"
   ScaleHeight     =   5700
   ScaleWidth      =   8775
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   7335
      TabIndex        =   7
      Top             =   4950
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   7335
      TabIndex        =   6
      Top             =   5265
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ImgListDrives 
      Left            =   6750
      Top             =   4935
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":27A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":4F56
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":770A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":9EBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":C672
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox PicBrowseBar 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   5340
      Left            =   0
      ScaleHeight     =   5340
      ScaleWidth      =   2820
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   2820
      Begin VB.FileListBox FileBrowseBar 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2040
         Left            =   -15
         Pattern         =   "*.TXT;*.INI"
         TabIndex        =   5
         Top             =   2880
         Width           =   2775
      End
      Begin MSComctlLib.TreeView tvwFolders 
         Height          =   2430
         Left            =   0
         TabIndex        =   8
         Top             =   225
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   4286
         _Version        =   393217
         Style           =   7
         ImageList       =   "ImgListDrives"
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
      Begin VB.Label LblClose 
         BackColor       =   &H00808080&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   2805
         TabIndex        =   11
         ToolTipText     =   "Close BrowseBar"
         Top             =   -15
         Width           =   210
      End
      Begin VB.Label LblSelect 
         BackColor       =   &H00808080&
         Caption         =   "Please Select a File..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   0
         Left            =   0
         TabIndex        =   9
         Top             =   2655
         Width           =   2805
      End
      Begin VB.Label LblSelect 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "Please Select a Directory..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   1
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   2805
      End
   End
   Begin MSComDlg.CommonDialog CPrintSetupDialog 
      Left            =   3120
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6120
      Top             =   4950
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":EE26
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":EF3A
            Key             =   "open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":F04E
            Key             =   "save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":F162
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":F276
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":F38A
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":F49E
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":F5B2
            Key             =   "find"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":F6C6
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":F7DA
            Key             =   "properties"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":F8EE
            Key             =   "options"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":FA02
            Key             =   "about"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":FB16
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":FC36
            Key             =   "printsetup"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":FD42
            Key             =   "time&date"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":FE9E
            Key             =   "fullscreen"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":FFF8
            Key             =   "browsebar"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":100F2
            Key             =   "replace"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1024C
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":103A6
            Key             =   "print"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":10500
            Key             =   "uppercase"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1065A
            Key             =   "lowercase"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   27
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "new"
            Object.ToolTipText     =   "New File"
            ImageKey        =   "new"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "open"
            Object.ToolTipText     =   "Open File"
            ImageKey        =   "open"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "recentfiles"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "save"
            Object.ToolTipText     =   "Save File"
            ImageKey        =   "save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "refresh"
            Object.ToolTipText     =   "Refresh/Revert"
            ImageKey        =   "refresh"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "properties"
            Object.ToolTipText     =   "File Properties"
            ImageKey        =   "properties"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "print"
            Object.ToolTipText     =   "Print"
            ImageKey        =   "print"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "printsetup"
            Object.ToolTipText     =   "Print Setup"
            ImageKey        =   "printsetup"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "undo"
            Object.ToolTipText     =   "Undo"
            ImageKey        =   "undo"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cut"
            Object.ToolTipText     =   "Cut"
            ImageKey        =   "cut"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "copy"
            Object.ToolTipText     =   "Copy"
            ImageKey        =   "copy"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "paste"
            Object.ToolTipText     =   "Paste"
            ImageKey        =   "paste"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "time&date"
            Object.ToolTipText     =   "Time/Date"
            ImageKey        =   "time&date"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "uppercase"
            Object.ToolTipText     =   "Make Selection Uppercase"
            ImageKey        =   "uppercase"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "lowercase"
            Object.ToolTipText     =   "Make Selection Lowercase"
            ImageKey        =   "lowercase"
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "find"
            Object.ToolTipText     =   "Find"
            ImageKey        =   "find"
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "replace"
            Object.ToolTipText     =   "Replace"
            ImageKey        =   "replace"
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "browsebar"
            Object.ToolTipText     =   "Toggle BrowseBar"
            ImageKey        =   "browsebar"
            Style           =   1
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "fullscreen"
            Object.ToolTipText     =   "Full Screen"
            ImageKey        =   "fullscreen"
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "options"
            Object.ToolTipText     =   "Options"
            ImageKey        =   "options"
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "About"
            Object.ToolTipText     =   "About"
            ImageKey        =   "about"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txt 
      Height          =   4215
      Index           =   1
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   6735
   End
   Begin MSComDlg.CommonDialog CfontDialog 
      Left            =   2160
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2640
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Timer MainTimer 
      Interval        =   1
      Left            =   120
      Top             =   4920
   End
   Begin VB.TextBox txt 
      Height          =   4245
      Index           =   0
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   360
      Width           =   6735
   End
   Begin VB.Label lblfilename 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   5280
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Menu MnuFile 
      Caption         =   "&File  "
      Begin VB.Menu MnuNewFile 
         Caption         =   "&New "
         Shortcut        =   ^N
      End
      Begin VB.Menu MnuOpenFile 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu MnuSave 
         Caption         =   "S&ave"
         Shortcut        =   ^S
      End
      Begin VB.Menu MnuSaveAs 
         Caption         =   "&Save As..."
      End
      Begin VB.Menu Line17 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPageSetup 
         Caption         =   "Pa&ge Setup..."
      End
      Begin VB.Menu MnuPrintSetup 
         Caption         =   "Print Set&up..."
      End
      Begin VB.Menu MnuPrint 
         Caption         =   "Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu line12 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCalculateSize 
         Caption         =   "&Calculate Size..."
         Shortcut        =   {F3}
      End
      Begin VB.Menu MnuDeleteFile 
         Caption         =   "&Delete..."
      End
      Begin VB.Menu MnuRename 
         Caption         =   "Re&name..."
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu MnuRefresh 
         Caption         =   "&Refresh/Revert"
         Shortcut        =   {F4}
      End
      Begin VB.Menu MnuProperties 
         Caption         =   "&Properties..."
      End
      Begin VB.Menu line22 
         Caption         =   "-"
      End
      Begin VB.Menu MnuRecentFilesMenu 
         Caption         =   "Recent &Files"
         Begin VB.Menu MnuRecentFilesAdvanced 
            Caption         =   "Advanced..."
         End
         Begin VB.Menu MnuReset 
            Caption         =   "&Clear"
         End
         Begin VB.Menu line19 
            Caption         =   "-"
         End
         Begin VB.Menu MnuRecentFiles 
            Caption         =   ""
            Index           =   0
            Visible         =   0   'False
         End
      End
      Begin VB.Menu line5 
         Caption         =   "-"
      End
      Begin VB.Menu MnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuedititem 
      Caption         =   " &Edit  "
      Begin VB.Menu MnuUndo 
         Caption         =   "Und&o"
         Shortcut        =   ^Z
      End
      Begin VB.Menu line14 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCut 
         Caption         =   "C&ut"
         Shortcut        =   ^X
      End
      Begin VB.Menu MnuCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu MnuPaste 
         Caption         =   "&Paste"
         Enabled         =   0   'False
         Shortcut        =   ^V
      End
      Begin VB.Menu MnuDelete 
         Caption         =   "De&lete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu line11 
         Caption         =   "-"
      End
      Begin VB.Menu MnuFind 
         Caption         =   "Fin&d..."
         Shortcut        =   ^F
      End
      Begin VB.Menu MnuFindNext 
         Caption         =   "&Find Next"
         Shortcut        =   {F2}
      End
      Begin VB.Menu MnuReplace 
         Caption         =   "&Replace..."
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu line6 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSelectAll 
         Caption         =   "&Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu line7 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSelection 
         Caption         =   "Make Selection"
         Begin VB.Menu MnuReverse 
            Caption         =   "&Reverse"
            Shortcut        =   ^R
         End
         Begin VB.Menu MnuUpperCase 
            Caption         =   "&Uppercase"
            Shortcut        =   ^U
         End
         Begin VB.Menu MnuLowerCase 
            Caption         =   "&Lowercase"
            Shortcut        =   ^L
         End
         Begin VB.Menu MnuInvertCase 
            Caption         =   "&Invert Case"
            Shortcut        =   ^I
         End
         Begin VB.Menu MnuMakeTitle 
            Caption         =   "&Title Case"
            Shortcut        =   ^T
         End
         Begin VB.Menu MnuLine2 
            Caption         =   "-"
         End
         Begin VB.Menu MnuMakeTabified 
            Caption         =   "&Tabified"
         End
      End
      Begin VB.Menu line20 
         Caption         =   "-"
      End
      Begin VB.Menu MnuInsert 
         Caption         =   "Insert"
         Begin VB.Menu MnuInsertTime 
            Caption         =   "&Time"
            Shortcut        =   +{F7}
         End
         Begin VB.Menu MnuInsertDate 
            Caption         =   "&Date"
            Begin VB.Menu MnuDateLongFormat 
               Caption         =   "Long Format"
               Shortcut        =   +{F5}
            End
            Begin VB.Menu MnuDateShortFormat 
               Caption         =   "Short Format"
               Shortcut        =   ^{F6}
            End
         End
         Begin VB.Menu MnuInsertTimeAndDate 
            Caption         =   "Time&/Date"
         End
         Begin VB.Menu MnuInsertFileName 
            Caption         =   "File&name"
            Shortcut        =   +{F4}
         End
         Begin VB.Menu MnuLine0 
            Caption         =   "-"
         End
         Begin VB.Menu MnuInsertBeginning 
            Caption         =   "Text &at Beginning of document..."
            Shortcut        =   +{F2}
         End
         Begin VB.Menu MnuInsertEnd 
            Caption         =   "Text at &End of document..."
            Shortcut        =   +{F3}
         End
         Begin VB.Menu MnuLine1 
            Caption         =   "-"
         End
         Begin VB.Menu MnuInsertTab 
            Caption         =   "&Tab"
         End
         Begin VB.Menu MnuInsertNewLine 
            Caption         =   "New &Line"
            Shortcut        =   +{F1}
         End
         Begin VB.Menu MnuInsertTextFromFile 
            Caption         =   "Text From &File..."
            Shortcut        =   ^{INSERT}
            Visible         =   0   'False
         End
      End
   End
   Begin VB.Menu MnuView 
      Caption         =   " &View  "
      Begin VB.Menu MnuToolbar 
         Caption         =   "&Toolbar"
      End
      Begin VB.Menu MnuBrowseBar 
         Caption         =   "&BrowseBar"
      End
      Begin VB.Menu line4 
         Caption         =   "-"
      End
      Begin VB.Menu MnuFullScreen 
         Caption         =   "Full &Screen..."
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu MnuFavoritesItem 
      Caption         =   "&Favorites"
      Begin VB.Menu MnuAddToFavorites 
         Caption         =   "Add To Favorites..."
      End
      Begin VB.Menu MnuOrganizeFavorites 
         Caption         =   "Organize Favorites..."
      End
      Begin VB.Menu MnuSepFavorites 
         Caption         =   "-"
      End
      Begin VB.Menu MnuFavorites 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
   End
   Begin VB.Menu MnuFormat 
      Caption         =   "&Format"
      Begin VB.Menu MnuWordWrap 
         Caption         =   "&Word Wrap"
      End
      Begin VB.Menu Line16 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSetBkcolor 
         Caption         =   "&Background Color..."
      End
      Begin VB.Menu MnuSetForecolor 
         Caption         =   "&Text Color..."
      End
      Begin VB.Menu MnuSetFont 
         Caption         =   "&Font..."
      End
   End
   Begin VB.Menu mnutools 
      Caption         =   "&Tools"
      Begin VB.Menu MnuClipboard 
         Caption         =   "Clipboard"
         Begin VB.Menu MnuClearClipboardText 
            Caption         =   "Cl&ear..."
            Shortcut        =   ^G
         End
         Begin VB.Menu MnuMangeClipboard 
            Caption         =   "&Manage..."
            Shortcut        =   ^M
         End
      End
      Begin VB.Menu MnuSeperator02 
         Caption         =   "-"
      End
      Begin VB.Menu MnuLaunchNewInstance 
         Caption         =   "&Launch New Instance"
         Shortcut        =   {F8}
      End
      Begin VB.Menu line02 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCopyFile 
         Caption         =   "&Copy File..."
         Shortcut        =   {F6}
      End
      Begin VB.Menu MnuMoveFile 
         Caption         =   "&Move File..."
         Shortcut        =   {F7}
      End
      Begin VB.Menu line3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuOptions 
         Caption         =   "&Options... "
         Shortcut        =   {F1}
      End
   End
   Begin VB.Menu MnuHelp 
      Caption         =   "  &Help    "
      Begin VB.Menu MnuTipOfTheDay 
         Caption         =   "&Tip of The Day..."
      End
      Begin VB.Menu Line15 
         Caption         =   "-"
      End
      Begin VB.Menu MnuVersionHistory 
         Caption         =   "&Version History..."
      End
      Begin VB.Menu line9 
         Caption         =   "-"
      End
      Begin VB.Menu MnuAbout 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu MnuBrowseBarPopup 
      Caption         =   "MnuBrowseBarPopup"
      Visible         =   0   'False
      Begin VB.Menu MnuBBVisible 
         Caption         =   "&Visible"
      End
      Begin VB.Menu MnuSeperator1 
         Caption         =   "-"
      End
      Begin VB.Menu MnubbOptions 
         Caption         =   "&Options..."
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FileBrowseBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Screen.MousePointer = vbNormal
End Sub
Private Sub Form_Load()
 GetDrives ' Get the drives and add them to the tree view.
End Sub
Private Sub LblClose_Click()
  Call MnuBrowseBar_Click
End Sub
Private Sub LblClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  LblClose.ForeColor = vbWhite
  Screen.MousePointer = vbNormal
End Sub
Private Sub LblSelect_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  LblClose.ForeColor = vbBlack
  Screen.MousePointer = vbNormal
End Sub
Private Sub MnuRecentFilesMenu_Click()
'MsgBox "Add code that as soon as this menu is activated, the recent files will load making NextPad startup faster.", vbInformation, "Suggestion"
End Sub

Private Sub MnuReverse_Click()
    ' Make the selected text go in reverse
    InsertAtSel mTextbox, StrReverse(mTextbox.SelText)
End Sub

Private Sub PicBrowseBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
   ' Change the mouse pointer to a resize icon.
   Screen.MousePointer = vbSizeWE
   
   If Button = vbLeftButton Then
      ' Initialize resizing.
      ResizeNoteWithToolbar
      ' Change the mouse pointer to a resize icon.
      Screen.MousePointer = vbSizeWE
      ' Make the width of the browsebar the current X position of the mouse cursor.
      PicBrowseBar.Width = X + 50 '2800
      ' Move the close button also.
      LblClose.Move X - 150 '+ 50
      ' Resize code is all the rest...
      FrmMain.FileBrowseBar.Width = FrmMain.PicBrowseBar.Width - 50
      FrmMain.tvwFolders.Width = FrmMain.PicBrowseBar.Width - 50
      FrmMain.LblSelect.Item(0).Width = FrmMain.PicBrowseBar.Width - 50
      FrmMain.LblSelect.Item(1).Width = FrmMain.PicBrowseBar.Width - 50
      ResizeNoteWithToolbar
      SaveSetting "BrowseBar", "Width", X + 50
   End If
End Sub
Private Sub Toolbar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LblClose.ForeColor = vbBlack
End Sub
Private Sub tvwFolders_KeyDown(KeyCode As Integer, Shift As Integer)
  If QuickExit And KeyCode = vbKeyEscape Then
     End
  End If
End Sub
Private Sub tvwFolders_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = vbRightButton Then
    PopupMenu MnuBrowseBarPopup, vbPopupMenuRightAlign, , , MnuBBVisible
 End If
End Sub
Private Sub FileBrowseBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = vbRightButton Then
    PopupMenu MnuBrowseBarPopup, vbPopupMenuRightAlign, , , MnuBBVisible
 End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  
' If the window state is other than minimized then...
If Not Me.WindowState = vbMinimized Then SaveSetting_RememberLastWinPos (RememberLastWinPos)
 
Dim Response  ' declare variables
   ' if the filestate is dirty then go to unload:
   If Fstate.dirty = True Then GoTo Unload:
   ' else if its not END
   If Fstate.dirty = False Then End
   


Unload: ' Vb will jump here if the File Sate is dirty
   
   ' If the user wants to Auto-Save files then just save it and quit
   If AutoSave = True Then
      SaveFile sOpenFileName, Me.mTextbox, False
      End
   End If
   
   
   ' Load the savefile function and recieve the value of the button the user has pressed
   Select Case SaveFile(sOpenFileName, Me.mTextbox, True)
   ' Select a response
    Case vbCancel ' User chose the Cancel Button
       Cancel = True ' Escape the Unload mode
       Exit Sub ' Exit the sub before running any more code in this sub
    Case Else ' the user chose the No Button
       End ' Stop and quit
   End Select
End Sub
Private Sub Form_Resize()
   On Error GoTo Nxterr: ' if An error occurs go to that line
  ' A form CANNOT be Resized while it is minimized so if it is
  ' Exit This Sub Immediatley
  If FrmMain.WindowState = vbMinimized Then Exit Sub
   Call ResizeNoteWithToolbar ' call resize procedure

Nxterr:
 If Err.Number <> 0 Then ' if the errors number equlas anything other than 0 then
   Exit Sub ' exit this sub immediately
 End If
End Sub
Private Sub MnuAddToFavorites_Click()
  FrmAddFavorites.Show vbModal
End Sub
Private Sub MnubbOptions_Click()
  FrmBrowseBarOptions.Show vbModal
End Sub
Private Sub MnuBBVisible_Click()
  Call MnuBrowseBar_Click
End Sub
Private Sub MnuCalculateSize_Click()
Dim Msg As String
  
  ' If the formatted size comes in KB's then display the size correctly.
  If Right(FormatKB(Len(mTextbox.Text)), 2) = "KB" Then
     '                  Show size in Kilobytes then show size in bytes (formatted with comma EX: 56,000)
     Msg = "Text is " & FormatKB(Len(mTextbox.Text)) & " (" & Format(Len(mTextbox.Text), "##,##") & " bytes)" & " long."
  Else ' Its just bytes.
     '                  Make integer with comma EX: 56,000
     Msg = "Text is " & Format(Len(mTextbox.Text), "##,##") & " bytes long."
  End If
  ' If the length (in bytes) of the document is 0 then just tell the user.
  If Len(mTextbox.Text) = 0 Then Msg = "Text is 0 bytes long."
  ' Display the size in bytes and or KB of the document to the user.
  MsgBox Msg, vbInformation, "NextPad - Calculate Size"
End Sub
Private Sub MnuCopyFile_Click()
  FrmFileOperation.Caption = "Copy File"
  FrmFileOperation.CmdOK.Caption = "Copy File"
  FrmFileOperation.Show vbModal
End Sub
Private Sub MnuDateLongFormat_Click()
      On Error Resume Next
     
    ' FrmMain.mTextBox.SelText = Time ' Set the Form's Active Control's
    InsertAtSel Me.mTextbox, CStr(Format(Date, "dddd, mmmm dd, yyyy"))
    ' selection point ( where the system Caret is )
    ' Text to the systems current Time

End Sub
Private Sub MnuDateShortFormat_Click()
      On Error Resume Next
     
    ' FrmMain.mTextBox.SelText = Time ' Set the Form's Active Control's
    InsertAtSel Me.mTextbox, CStr(Date)
    ' selection point ( where the system Caret is )
    ' Text to the systems current Time

End Sub
Private Sub MnuFavorites_Click(Index As Integer)
  Dim sFileName As String
  
  If Fstate.dirty Then ' If the state of the file is dirty then
  ' Save the file, But if the response is cancel then exit this sub (canceling the opening of the recent file)
  If SaveFile(sOpenFileName, mTextbox, True) = vbCancel Then Exit Sub
  End If
 
 sFileName = MnuFavorites(Index).Tag
  'Call the OpenFile sub in modmain
  If bFileExists(sFileName) = False Then: NotifyFileNonExistent (sFileName): Exit Sub
  OpenFile sFileName
  ' Open the Favorite From the registry using
  ' the Procedure in ModMain ,

End Sub
Private Sub MnuInsertBeginning_Click()
   ' Make system caret go to the beginning.
   mTextbox.SelStart = 0
   mText = InputBox$("Enter text here that you would like placed at the beginning of this document.", "NextPad")
   If mText <> "" Then
      InsertAtSel mTextbox, CStr(mText & " ")
   End If
End Sub
Private Sub MnuInsertEnd_Click()
   ' Make system caret go to the beginning.
   mTextbox.SelStart = Len(mTextbox.Text)
   mText = InputBox$("Enter the text here that you would like placed at the end of this document.", "NextPad")
   If mText <> "" Then
      InsertAtSel mTextbox, CStr(mText & " ")
   End If

End Sub
Private Sub MnuInsertFileName_Click()
   ' Insert the filename of the file currently open at the selection point.
   InsertAtSel mTextbox, sOpenFileName
End Sub
Private Sub MnuInsertNewLine_Click()
   ' Insert a new line at the selection point.
   InsertAtSel mTextbox, vbNewLine
End Sub
Private Sub MnuInsertTab_Click()
   ' Insert a tab at the selection point.
   InsertAtSel mTextbox, vbTab
End Sub
Private Sub MnuInsertTextFromFile_Click()
   Dim sFileName As String
     MsgBox "UNDER CONSTRUCTION!"
End Sub
Private Sub MnuInsertTime_Click()
      On Error Resume Next
     
    InsertAtSel Me.mTextbox, CStr(Time)
    ' selection point ( where the system Caret is )
    ' Text to the systems current Time
End Sub
Private Sub MnuInsertTimeAndDate_Click()
   On Error Resume Next
     
    ' FrmMain.mTextBox.SelText = Time ' Set the Form's Active Control's
    InsertAtSel Me.mTextbox, CStr(Time & " " & Date)
    ' selection point ( where the system Caret is )
    ' Text to the systems current Time
End Sub
Private Sub MnuInvertCase_Click()
   InsertAtSel mTextbox, StrInvertCase(mTextbox.SelText)
End Sub
Private Sub MnuLaunchNewInstance_Click()
   ShellNewNextPad vbNormalFocus
End Sub
Private Sub MnuLowerCase_Click()
   InsertAtSel mTextbox, LCase$(mTextbox.SelText)
End Sub
Private Sub MnuMakeTabified_Click()
   ' Make the line tabified.
   InsertAtSel mTextbox, vbTab & mTextbox.SelText
End Sub
Private Sub MnuMakeTitle_Click()
   InsertAtSel mTextbox, StrConv(mTextbox.SelText, vbProperCase)
End Sub
Private Sub MnuMangeClipboard_Click()
   frmClipboard.Show vbModal
End Sub
Private Sub MnuOrganizeFavorites_Click()
   FrmOrganizeFavorites.Show 0, Me
End Sub
Private Sub MnuRecentFilesAdvanced_Click()
   FrmRecentFiles.Show 0, Me
End Sub
Private Sub MnuAbout_Click()
   Load frmAbout ' load the form
   frmAbout.Show (vbModal) ' show the form in Vbmodal mode
End Sub
Private Sub MnuBrowseBar_Click()
   ResizeNoteWithToolbar 'Resize the Form To Match
   ' The Current state
   PicBrowseBar.Visible = Not PicBrowseBar.Visible 'Show the Browser Bar
   ' Using a Logical Expression Not (The Opposite if each , eg; 1 = 0 0 = 1)
   MnuBrowseBar.Checked = PicBrowseBar.Visible
   ResizeNoteWithToolbar  'Resize the Form To Match the
   ' current state
   SaveSetting "BrowseBar", "Visible", Abs(CInt(PicBrowseBar.Visible))
End Sub
Private Sub MnuDeleteFile_Click()
   'Send file to Recycle bin
   DeleteFile Me.hWnd, sOpenFileName, True
  'If the file was infact deleted then close this file because it no longer exists
  If bFileExists(sOpenFileName) = False Then ClearNextPad
  ' Refresh File list control
  FileBrowseBar.Refresh
End Sub
Private Sub MnuRefresh_Click()
   ' Reopen the file, Reverting it to its last saved state.
   If Not bFileExists(sOpenFileName) Then MsgBox "Cannot refresh file currently open." & vbNewLine & "The file may no longer exist or, may have been moved since the last time it was open." & vbNewLine & Err.Description, vbCritical, "NextPad - Refresh/Revert": Exit Sub
   OpenFile sOpenFileName, False
End Sub
Private Sub MnuRename_Click()
On Error GoTo ErrorHandler:
   Dim mText As String, sPath As String, sFinalPath As String
     ' If no file is currently open or the path is empty then set it to the currently open one.
     If sOpenFileName = "" Then
        ' Notify user about the problem giving the user a choice.
        If MsgBox("You cannot rename this file until you save it. Do you want to save it?", vbInformation + vbYesNo, "Rename - NextPad") = vbYes Then
           ' Save the file prompting the user with the save file CMD.
           SaveFile "", mTextbox, False
        Else
           Exit Sub
        End If
     End If
        
     If sOpenFileName = "" Then Exit Sub
     mText = InputBox("New File Name:", "Rename File", Dir(sOpenFileName))
      If mText <> "" Then ' If there is a string...
         sPath = Trim(GetPathFromFileName(sOpenFileName))
         ' Commence renaming of the file, appending the path to the filename.
         sFinalPath = sPath + "\" + mText
      Else
         ' No file title expressed stop.
         Exit Sub
      End If
  ' Rename the file
  Name sOpenFileName As sFinalPath
  ' Tell NextPad what the current open file name is.
  sOpenFileName = sFinalPath
  ' Set the main forms caption.
  FrmMain.Caption = GetFileTitleStr(sOpenFileName) & " - NextPad"
  Exit Sub ' Were done
  
ErrorHandler:
  ' The user renamed the file to the same name...
  If Err.Number = 58 Then
     Exit Sub
  Else ' An unknown error has occured...
     MsgBox "Could not rename file, " & Err.Description, vbCritical, "Rename File Error - NextPad"
     Exit Sub
  End If
    
End Sub
Private Sub MnuUpperCase_Click()
   InsertAtSel mTextbox, UCase$(mTextbox.SelText)
End Sub
Private Sub MnuVersionHistory_Click()
   Load Frmbugz ' load the form
   Frmbugz.Show (vbModal) ' show the form
End Sub
Private Sub MnuClearClipboardText_Click()
Dim Response 'Variable to hold response from user
  'Display the Warning to the user
  Response = MsgBox("Clear all contents currently on the Clipboard ?", vbYesNo + vbInformation, "Clear Clipboard - NextPad")
   
   Select Case Response
       Case vbYes 'User pressed Yes
         Clipboard.Clear 'Clear the clipboard
       Case vbNo 'User Pressed No
         Exit Sub 'Cancel ; Exit this sub
   End Select
End Sub
Private Sub MnuCopy_Click()
' We Now Send Messages Because The Undo Feature Didnt Work
' Properly ,

' Send A Message Too The Forms Active Control , WM_COPY
' Too Simulate a Copy Procedure

   SendMessage FrmMain.mTextbox.hWnd, WM_COPY, 0&, 0&

End Sub
Private Sub MnuCut_Click()
  ' Now We Send Messages Because Undo Wasnt Working Like It Was Supposed Too
  ' Send A Message Too Cut
  SendMessage FrmMain.mTextbox.hWnd, WM_CUT, 0&, 0&
End Sub
Private Sub MnuDelete_Click()
    ' delete the selected text in the form's active control's text
    ' property
    InsertAtSel mTextbox, ""
End Sub
Private Sub MnuExit_Click()
   Call Form_QueryUnload(1, 0)
End Sub
Private Sub MnuFind_Click()
    If FrmMain.mTextbox.SelText <> "" Then
        frmFind.CmbFind.Text = FrmMain.mTextbox.SelText
    Else
        frmFind.CmbFind.Text = gFindString
    End If
    
    gFirstTime = True
    
    If (gFindCase) Then
        frmFind.chkCase = 1
    End If
    
    frmFind.Show (vbModeless), Me ' show the form Telling VB
    ' that the owner form is Me (FrmMain)
End Sub
Private Sub MnuFindNext_Click()
    If Len(gFindString) > 0 Then 'if the length of the
        FindStr  ' search string is over 0 then FindStr
    Else ' if there is nothing too find then
        MnuFind_Click 'load the dialog
    End If
End Sub
Private Sub MnuFullScreen_Click()
 ' ** here we Select a Case of BOOLEAN
 ' ** Frmfullscreen.visible = [BOOLEAN]

  Select Case frmfullscreen.Visible ' select a case of true or false
   Case False ' Frmfullscreen IS NOT visible
      MnuFullScreen.Checked = True 'check the menu
      frmfullscreen.Txtfullscreen.BackColor = Me.mTextbox.BackColor
      frmfullscreen.Txtfullscreen.ForeColor = Me.mTextbox.ForeColor
      frmfullscreen.Txtfullscreen.Font = Me.mTextbox.Font
      frmfullscreen.Txtfullscreen.fontsize = Me.mTextbox.fontsize
      frmfullscreen.Txtfullscreen.FontBold = Me.mTextbox.FontBold
      frmfullscreen.Txtfullscreen.FontItalic = Me.mTextbox.FontItalic
      frmfullscreen.Show 0, Me ' show the form
   Case True ' frmfullscreen IS visible
      MnuFullScreen.Checked = False ' Uncheck the Menu
      Unload frmfullscreen ' Unload the Form
      Unload Frmleavefullscreen ' Unload the form
  End Select

End Sub
Private Sub MnuToolbar_Click()
    Toolbar.Visible = Not Toolbar.Visible
    ' Change the check to match the current state
    MnuToolbar.Checked = Toolbar.Visible
    ' Call the resize procedure
    ResizeNoteWithToolbar ' resize the form
    
    SaveSetting "Toolbar", "Visible", Abs(CInt(CBool(Toolbar.Visible)))
End Sub
Private Sub MnuMoveFile_Click()
  ' Show The Form
  FrmFileOperation.Caption = "Move File"
  FrmFileOperation.CmdOK.Caption = "Move File"
  FrmFileOperation.Show (vbModal)
End Sub
Private Sub MnuNewFile_Click()
  ' If the file has not been saved then save it.
 Dim Response As VbMsgBoxResult
 
  If Fstate.dirty Then
     ' Notify user to save the file.
     Select Case SaveFile(sOpenFileName, mTextbox, True)
       Case vbCancel
            Exit Sub
       Case Else
            ' Were done clear NextPad.
            ClearNextPad
     End Select
  Else ' The file is not dirty.
     ClearNextPad
  End If
End Sub
Private Sub MnuOptions_Click()
  On Error Resume Next
   frmOptions.Show (vbModal) ' show the form In Vbmodal Mode
End Sub
Private Sub MnuPageSetup_Click()
  On Error Resume Next
   FrmPageSetup.Show (vbModal) 'show the form in vbmodal mode
End Sub
Private Sub MnuPrint_Click()
  PrintDocument mTextbox 'Print The Document USing the PrintDocument() Procedure in ModMain
  ' By Referencing MyTextBox As The Real Ones on FrmMain
End Sub
Private Sub MnuPrintSetup_Click()
  On Error GoTo CdlcCancelError: 'If an error Occurs VB WIll Jump too that line

   'Set the Common Dialogs Flags too display the Print Setup DIalog
   ' Rather than the Print DIalog
   CPrintSetupDialog.Flags = cdlPDPrintSetup
     ' Show The Dialog
     CPrintSetupDialog.ShowPrinter
 
 Exit Sub ' Exit Sub Immediatelly

CdlcCancelError:
   If Err.Number = 32755 Then 'if cancel was Pressed (CldlcCancelError) then ......
     Exit Sub
   Else ' Else another error Has occured Other than the Cancel Error
   If Err.Number = 28663 Then Exit Sub ' No Default Printer exists
  
     MsgBox "Could not display Print Setup window, " & Err.Description, vbCritical, "NextPad"
   End If

End Sub
Private Sub MnuReplace_Click()
    If FrmMain.mTextbox.SelText <> "" Then
        frmReplace.CmbFind.Text = FrmMain.mTextbox.SelText
    Else
        frmReplace.CmbFind.Text = gFindString
    End If
    
    gFirstTime = True
    
    If (gFindCase) Then
        frmReplace.chkCase.Value = 1
    End If
    
    frmReplace.Show vbModeless, Me
End Sub
Private Sub MnuReset_Click()
  'If the user presses yes then continue...
  If MsgBox("Are you sure you wish to remove the list of recent files from the registry?", vbYesNo + vbExclamation, "Clear Recent Files") = vbYes Then
     ' Make sure we take out this menu item also.
     MnuRecentFiles(0).Caption = "": MnuRecentFiles(0).Visible = False
     On Error Resume Next
     ' Delete the Setting's Specified , "RecentFiles"
     DeleteSetting "NextPad", "RecentFiles"
   
     Dim i As Integer, X As Integer
      'Get the menu count
      i = MnuRecentFiles.Count
      For X = 1 To i Step 1 ' Go through them one by one
        If X = i Then Exit For ' If the variable is equal to the count then stop
        Unload MnuRecentFiles(X) ' Unload the specified menu
      Next X ' Next X value
  End If
End Sub
Private Sub MnuSetBkcolor_Click()
 On Error GoTo CdlcCancelError:
  ' Set Common DIalog Flags
  ' THis One Shows the Entire Common Dialog Box
  ' including the custom colors Section
  CfontDialog.Flags = CDLCFullopen
  ' Show the color dialog
  CfontDialog.ShowColor
  ' set FrmMain's active controls' BackColor too the one
  ' the user had chosen from the Common Dialog
  FrmMain.mTextbox.BackColor = CfontDialog.Color
  
  SaveSetting "Font", "BackColor", CLng(FrmMain.mTextbox.BackColor)

CdlcCancelError:
 
  If Err.Number = 32755 Then  ' if cancel was Pressed
     Exit Sub
  End If
End Sub
Private Sub MnuSetForecolor_Click()
  On Error GoTo CdlcCancelError:
  
   With CfontDialog
    ' set the Common Dialogs Flags
    ' This Flag Allows The Whole Dialog Too be Displayed
    ' Including the Define Custom Colors
    .Flags = CDLCFullopen
    ' Show the Dialog
    .ShowColor
    ' Set FrmMain's mTextBox's Forecolor ( Font Color )
    ' too the One the user has selected in the Common Dialog
    FrmMain.mTextbox.ForeColor = .Color
  End With
  
  'Save in the Registry the Forms Active Control's ForeColor ( Font Color )
  ' And We Convert the Color Chosen From The Color Dialog Too a Long Using
  ' Clng([Expression])
  SaveSetting "Font", "ForeColor", CLng(FrmMain.mTextbox.ForeColor)
  
CdlcCancelError:
    If Err.Number = 32755 Then ' if the user Pressed Cancel
       Exit Sub
    End If
End Sub
Private Sub MnuTipOfTheDay_Click()
   On Error Resume Next
    ' Show The Tip Of The Day Dialog
    frmTip.Show (vbModal)
End Sub
Private Sub MnuUndo_Click()
'***********************************************************************
'* Here we perform A Simple Undo Procedure Fist We Send A Message Too  *
'* The active form's Text Control's HWND Property its Called EM_UNDO   *
'* And thats it !!!!                                                   *
'***********************************************************************
  On Error Resume Next
  SendMessage FrmMain.mTextbox.hWnd, EM_UNDO, 0&, 0&
End Sub
Private Sub MnuOpenFile_Click()
'*********************************************************************
' This menu Item When clicked Will check if the file Open ( If any)  *
' Needs to be saved if so , Display the Message Too the user ,       *
' And Calculate a response .                                         *
'*********************************************************************
Close #1 'Close the file just in case it is open
Dim Response ' declare Variables
Dim regval As String ' Declare variables
  
   If Fstate.dirty = False Then GoTo OpenFileProc:
   If Fstate.dirty = True Then
    
' Msg Variable

End If
       ' set response variable too display msgbox with the Defined buttons and title
       Response = SaveFile(sOpenFileName, Me.mTextbox, True)

   Select Case Response ' detect Which Button was pressed by the user.....
    
     Case vbYes ' user clicked the yes Button
            SaveFile sOpenFileName, Me.mTextbox, False   ' call procedure in modmain
            CommonDialog1.Filename = ("")
            sOpenFileName = ""
     Case vbNo ' user clicked the No Button
            GoTo OpenFileProc:
     Case vbCancel ' User clicked The Cancel Button
            Exit Sub
    End Select

OpenFileProc:
      Reset ' Reset all open Disks
       Close #1 ' close before using it again
        FreeFile (1) 'Free the File
         On Error GoTo Cdlogerror:
     
     With CommonDialog1
        .Flags = Normal_Cdlogflags    ' use the constant in modmain
        .Cancelerror = True ' set Cancel Error Too true When User clicks cancel an error will generate
        ' set the filter
        .Filter = "Text Files (*.TXT) |*.TXT|Ini Files (*.INI) |*.INI|Log Files (*.LOG) |*.LOG|All Files (*.*) |*.*"
            ' set the commondialogs title
        .DialogTitle = "Open File"
        .ShowOpen ' Show the open dialog
     End With
     
     OpenFile CommonDialog1.Filename, True ' call OpenFile proc in modmain

Cdlogerror: ' error that is triggered when the cancel button on the Commondialog is pressed
    If Err.Number = 32755 Then
       Exit Sub
    End If
End Sub
Private Sub MnuPaste_Click()
     On Error GoTo Outofmemoryerr:
       SendMessage FrmMain.mTextbox.hWnd, WM_PASTE, 0&, 0&
      

Outofmemoryerr:
     If Err.Number <> 0 Then
      ' Notify user.
      MsgBox "Could not paste any more text, " & Err.Description, vbCritical, "NextPad"
      ' Exit this sub
      Exit Sub
     End If
End Sub
Private Sub MnuProperties_Click()
Dim r As Long 'declare variables
     
     r = ShowFileProperties(Me.hWnd, sOpenFileName)
     ' show fileinfo  the hwnd Of this window ; and the sOpenFileName property
End Sub
Private Sub MnuRecentFiles_Click(Index As Integer)
Dim sFileName As String
  
  If Fstate.dirty Then ' If the state of the file is dirty then
  ' Save the file, But if the response is cancel then exit this sub (canceling the opening of the recent file)
  If SaveFile(sOpenFileName, mTextbox, True) = vbCancel Then Exit Sub
  End If
 
 sFileName = MnuRecentFiles(Index).Caption
  'Call the OpenFile sub in modmain
  'OpenFile (sFileName)
  If bFileExists(sFileName) = False Then: NotifyFileNonExistent (sFileName): Exit Sub
  OpenFile sFileName
End Sub
Private Sub MnuSaveAs_Click()
  ' Use the new savefile function without passing a filename to it
  ' in order for it to just show the save as dialog
  SaveFile "", Me.mTextbox, False
End Sub
Private Sub MnuSave_Click()
  'Save the file without prompting the user
  SaveFile sOpenFileName, Me.mTextbox, False
  ' Change the fileState now
  Fstate.dirty = False
End Sub
Private Sub MnuSelectAll_Click()
   ' set the forms active controls
   ' selection start point to the beginning
   FrmMain.mTextbox.SetFocus
   FrmMain.mTextbox.SelStart = 0
   '  set the forms active control's selection length  too
   '  the lenth of The Form's active control's Text
   FrmMain.mTextbox.SelLength = Len(FrmMain.mTextbox.Text)
End Sub
Private Sub MnuSetFont_Click()
  On Error GoTo Cancelerror: ' if an error occurs goto that line
  
  CfontDialog.Flags = cdlCFBoth  ' set the Font Dialogs flags
  CfontDialog.ShowFont ' show the Select Font Dialog
  ' set the forms active control's font too the selected font form the dialog
  FrmMain.mTextbox.Font = CfontDialog.fontname
  ' set the forms active control's Bold too the Selected font dialog bold
  FrmMain.mTextbox.FontBold = CfontDialog.FontBold
  ' set the form's active control's Font size too the selected font size from the dialog
  FrmMain.mTextbox.fontsize = CfontDialog.fontsize
  ' save in the registry ; the forms active control's italic (BOOLEAN) too the selected italic form the dialog
  FrmMain.mTextbox.FontItalic = CfontDialog.FontItalic
  ' save in the registry ; The Font For the Form's Text Box
  ' For later Retrieval From the registry
  SaveSetting "Font", "font", FrmMain.mTextbox.fontname
  ' save in the registry ; the controls current Font size for later
  ' retrieval from the registry
  SaveSetting "Font", "Fontsize", FrmMain.mTextbox.fontsize
  ' Save in the Registry ; the controls current Bold Attribute
  ' Either TRUE (-1) or FALSE (0)
  SaveSetting "Font", "FontBold", Abs(FrmMain.mTextbox.FontBold)
  ' Save in the registry ; the controls current italic Attribute
  ' Either TRUE (-1) or FALSE (0)
  SaveSetting "Font", "FontItalic", Abs(FrmMain.mTextbox.FontItalic)
  
Cancelerror: ' if an error occurs Vb will Jump here
  Exit Sub ' Exit the sub immediately
End Sub
Private Sub MnuWordWrap_Click()
  Dim Retval As Boolean
     
    Retval = Usewordwrap ' set variable too usewordwrap's value
    ' Use logical exclusion.
    Retval = Not Retval
    ' Save the setting in the registry.
    SaveSetting "WordWrap", "WordWrap", CInt(Abs(Retval))
    ' Toggle wordwrap
    ToggleWordWrap (Retval)
End Sub
Private Sub MainTimer_Timer()
'***************************************************************
' MainTimer Events ; makes NextPad update menu choices properly.
' Interval 1 millisecond
'***************************************************************
Dim Retval As Long, bSelText  As Boolean, bVal As Boolean
Dim bFileName As Boolean, bClipBoard As Boolean
    
    On Error Resume Next ' if there is an error
    ' we dont want to exit the sub because the rest of the code wont be executed
    ' so just go too the next line
    Select Case PicBrowseBar.Visible ' Check if the browsebar is visible
       Case True 'It is
           Toolbar.Buttons("browsebar").Value = 1 'Match the browsebar's current state with the buttons value
           MnuBBVisible.Checked = True
       Case False 'It is not
           Toolbar.Buttons("browsebar").Value = 0 'Match the browsebar's current state with the buttons value
           MnuBBVisible.Checked = False
    End Select
    
    On Error Resume Next ' if there is an error
    ' we dont want too exit the sub because the rest of the code wont be executed
    ' so just go too the next line
    Select Case sOpenFileName ' if the Lblfilename's Caption property tells us
        Case Is <> ""
           bFileName = True
        Case Else
           bFileName = False
    End Select
     
     MnuProperties.Enabled = bFileName
     MnuDeleteFile.Enabled = bFileName
     MnuInsertFileName.Enabled = bFileName
     MnuRefresh.Enabled = bFileName
     Toolbar.Buttons("properties").Enabled = bFileName
     Toolbar.Buttons("refresh").Enabled = bFileName

   
   
    On Error Resume Next ' if there is an error
    ' we dont want too exit the sub because the rest of the code wont be executed
    ' so just go too the next line
    Select Case Clipboard.GetText  ' if the clipboard Has text Then ....
        Case Is <> ""
           bClipBoard = True
        Case Else
           bClipBoard = False
    End Select
     MnuClearClipboardText.Enabled = bClipBoard ' enable this menu Because now there is Clipboard text too clear
     MnuPaste.Enabled = bClipBoard ' enable this menu because now there is text too paste
     Toolbar.Buttons("paste").Enabled = bClipBoard
     MnuViewClipboardText.Enabled = bClipBoard ' enable menu this now because now there is text too view
     
  On Error Resume Next
     
     With mTextbox
        Select Case .SelText   'Check if any text is selected
            Case Is <> ""
               bSelText = True
            Case Else
               bSelText = False
        End Select
        MnuCopy.Enabled = bSelText
        MnuCut.Enabled = bSelText
        MnuDelete.Enabled = bSelText
        MnuMakeTitle.Enabled = bSelText
        MnuUpperCase.Enabled = bSelText
        MnuLowerCase.Enabled = bSelText
        MnuInvertCase.Enabled = bSelText
        MnuMakeTabified.Enabled = bSelText
        MnuReverse.Enabled = bSelText
        Toolbar.Buttons("cut").Enabled = bSelText
        Toolbar.Buttons("copy").Enabled = bSelText
        Toolbar.Buttons("uppercase").Enabled = bSelText
        Toolbar.Buttons("lowercase").Enabled = bSelText
     End With
  
  On Error Resume Next
  
  ' Check If We Can Undo Or Not
  ' Send the message too the Active Control of the form
  ' too see if we can undo or not
  Retval = SendMessage(FrmMain.mTextbox.hWnd, EM_CANUNDO, ByVal 0&, ByVal 0&)
  bVal = (Retval <> 0) 'Boolean Value
  
  Select Case bVal
    Case True 'We Can UNDO
       MnuUndo.Enabled = True 'Enable The Menu
       Toolbar.Buttons("undo").Enabled = True
    Case False ' We CANT UNDO
       MnuUndo.Enabled = False 'Disable The Menu
       Toolbar.Buttons("undo").Enabled = False
  End Select
End Sub
Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
      On Error GoTo ErrorHandler: ' if an error occurs Vb will jump too that line

        Select Case Button.Key ' Select The button Being pressed by
        ' using the select case statement also on the Buttons Key
     Case "open" ' user Clicked the Open Button
         Call MnuOpenFile_Click
     Case "save" ' user Clicked the save Button
         Call MnuSave_Click
     Case "refresh" ' user Clicked the Refresh Button
         Call MnuRefresh_Click
     Case "print" 'user Clicked Print Button
         Call MnuPrint_Click
     Case "printsetup" 'User Clicked Page Setup Button
         Call MnuPrintSetup_Click
     Case "copy" ' user Clicked the copy Button
         Call MnuCopy_Click
     Case "cut" ' user Clicked the cut Button
         Call MnuCut_Click
     Case "paste" ' user Clicked the Paste Button
         Call MnuPaste_Click
     Case "find" ' user Clicked the Find Button
         Call MnuFind_Click
     Case "replace" ' user clicked the replace button
         Call MnuReplace_Click
     Case "undo" ' User Clicked Undo Button
         Call MnuUndo_Click
     Case "time&date" ' user Clicked the time&date Button
         Call MnuInsertTimeAndDate_Click
     Case "uppercase"
         Call MnuUpperCase_Click
     Case "lowercase"
         Call MnuLowerCase_Click
     Case "options" ' user Clicked the options Button
         Call MnuOptions_Click
     Case "About" ' user Clicked the about Button
         Call MnuAbout_Click
     Case "new" ' user Clicked the new Button
         Call MnuNewFile_Click
     Case "properties" ' user Clicked the properties Button
         Call MnuProperties_Click
     Case "browsebar"
         Call MnuBrowseBar_Click
     Case "fullscreen"
         Call MnuFullScreen_Click
        End Select
ErrorHandler:
    If Err.Number <> 0 Then ' if an error is equal too anything other or above zero then ......
      ' Display the error message too the user
      MsgBox "Cannot access the toolbar, " & Err.Description, vbCritical, "Toolbar - NextPad"
      Exit Sub ' exit the sub immediately
    End If

End Sub
Private Sub Toolbar_ButtonDropDown(ByVal Button As MSComctlLib.Button)
On Error Resume Next
  ' If the user has disabled recent files, dont display whats left.
  If Not AllowRecentFiles Then Exit Sub
  ' If the open file button menu was clicked then, display the recent files menu.
  If Button.ButtonMenus.Item(1).Key = "recentfiles" Then
      PopupMenu MnuRecentFilesMenu, vbPopupMenuLeftAlign
  End If

End Sub
Private Sub Toolbar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then ' if the user
     ' Right Clicked on the toolbar then...
           
     PopupMenu MnuView, vbPopupMenuLeftAlign, , , MnuToolbar
    ' Display the popup menu too the user
    ' Display the popup menu Left Aligned At X Y POS
    End If
 End Sub
Private Sub tvwFolders_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  LblClose.ForeColor = vbBlack
  Screen.MousePointer = vbNormal
End Sub
Private Sub txt_Change(Index As Integer)
  TextChangecontrol
End Sub
Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If QuickExit And KeyCode = vbKeyEscape Then
     End
  End If
End Sub
Private Sub FileBrowseBar_Click()
   Dim SelectedFile As String
   Dim Response
     
     If Fstate.dirty Then ' If the file's state is changed...
        ' Display the message to the user.
        Response = SaveFile(sOpenFileName, mTextbox, True)
       Select Case Response ' Examine variable.
          Case vbYes ' Yes
              ' Continue
          Case vbNo  ' No
              ' Continue
          Case vbCancel ' Cancel
              Exit Sub ' Dont continue opening the file.
       End Select
     End If
     ' First we append this string with a slash like its in a folder
     SelectedFile = FileBrowseBar.Path & "\" & FileBrowseBar.Filename
     ' If the file is not in a folder its filename may look like this : "C:\\MyFile.txt" ,Repair it removing the slash
     If bFileExists(SelectedFile) = False Then SelectedFile = FileBrowseBar.Path & FileBrowseBar.Filename
     OpenFile SelectedFile ' If all goes well we can now open the file
End Sub
Private Sub FileBrowseBar_KeyDown(KeyCode As Integer, Shift As Integer)
  If QuickExit And KeyCode = vbKeyEscape Then
     End
  End If
End Sub
Private Function SetMainTextBox() As TextBox
'******************************************************************************************
' Please Remember that this Function is a Reference to an object                          *
' that already exists , By the time i had incoroporated the BrowseBar                     *
' The ActiveControl Property no longer worked when the Browse Bar was Visible             *
' Because, one of its objects would take the activecontrol property so we had             *
' to improvise by referencing this fucntion towards the Real TextBox's                    *
'******************************************************************************************
   If txt(0).Visible = True Then Set SetMainTextBox = txt(0) 'Set this Fake Function to the real one
   If txt(1).Visible = True Then Set SetMainTextBox = txt(1) 'Set this Fake Function to the real one
 
End Function
Public Function mTextbox() As TextBox
' another Recursion of a Recursion :)
' All we do is Reference them back to back
' And we can now use this function universally in NextPad

  SetMainTextBox 'Call the Original to set the Property up
  Set mTextbox = SetMainTextBox 'Set this To the Real Function

End Function
Private Sub GetDrives()
    Dim i As Integer
    Dim strPath As String
    Dim IconN As Integer
    tvwFolders.Nodes.Clear
    
    For i = 0 To Drive1.ListCount - 1 'Get all drives
        ' Get drive
        strPath = UCase(Left(Drive1.List(i), 1)) & ":\"

        Select Case GetDriveType(strPath) 'Check to the type of the drive
            Case 2 'Diskette drive
                IconN = 1
            Case 3 'Hard Disk
                IconN = 2
            Case 5 'CDROM drive
                IconN = 3
            Case Else
                IconN = 2
        End Select
        
        ' Add drive
        tvwFolders.Nodes.Add , , strPath, UCase(Drive1.List(i)), IconN
        tvwFolders.Nodes.Add strPath, tvwChild, ""
    Next
End Sub
Private Sub tvwFolders_Expand(ByVal Node As MSComctlLib.Node)
    On Error GoTo TVError
    Dim i As Integer
    Dim strRelative As String
    Dim strFolderName As String
    Dim intFolderPos As Integer
    Dim strNewPath As String

    MousePointer = vbHourglass 'Change mouse pointer to hourglass

    ' Add folders
    If Node.Child.Text = "" Then
             
        tvwFolders.Nodes.Remove Node.Child.Index
        strRelative = Node.Key
        Dir1.Path = strRelative
        intFolderPos = Len(strRelative) + 1

        For i = 0 To Dir1.ListCount - 1
            strFolderName = Mid(Dir1.List(i), intFolderPos)
            strNewPath = strRelative & strFolderName & "\"
            tvwFolders.Nodes.Add strRelative, tvwChild, strNewPath, strFolderName, 4
            Dir1.Path = strNewPath

            If Dir1.ListCount > 0 Then
                tvwFolders.Nodes.Add strNewPath, tvwChild, , ""
                tvwFolders.Nodes(strNewPath).ExpandedImage = 5
            End If

            Dir1.Path = strRelative
        Next

    End If
    MousePointer = vbDefault 'Change mouse pointer to default
    Exit Sub
TVError:
    If Err.Number = 68 Then MsgBox "Device not avaible!", vbCritical, "NextPad - Folder View"
    MousePointer = vbDefault 'Change mouse pointer to hourglass
    Exit Sub
End Sub
Private Sub tvwFolders_NodeClick(ByVal Node As MSComctlLib.Node)
  On Error GoTo ErrorHandler:
  
    FileBrowseBar.Path = Node.Key 'Set path

ErrorHandler:
    If Err.Number <> 0 Then 'If An Error Has Occured ......
       MsgBox "Cannot access the directory you have selected, " & Err.Description, vbCritical, "Folder View - NextPad"
       Exit Sub
    End If
End Sub
Private Sub txt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  LblClose.ForeColor = vbBlack
  Screen.MousePointer = vbNormal
End Sub
