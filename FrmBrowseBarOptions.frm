VERSION 5.00
Begin VB.Form FrmBrowseBarOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BrowseBar"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6480
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmBrowseBarOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "File List"
      Height          =   1545
      Left            =   75
      TabIndex        =   11
      Top             =   90
      Width           =   6255
      Begin VB.CommandButton CmdResetPattern 
         Caption         =   "Default Pattern"
         Height          =   330
         Left            =   4680
         TabIndex        =   0
         Top             =   735
         Width           =   1395
      End
      Begin VB.TextBox TxtPattern 
         Height          =   285
         Left            =   795
         TabIndex        =   2
         Top             =   1125
         Width           =   5295
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   ";*.MYEXT;*.MYEXT2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   135
         TabIndex        =   12
         Top             =   885
         Width           =   1665
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Pa&ttern :"
         Height          =   255
         Left            =   75
         TabIndex        =   1
         Top             =   1125
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   $"FrmBrowseBarOptions.frx":000C
         Height          =   1095
         Left            =   165
         TabIndex        =   13
         Top             =   210
         Width           =   6015
      End
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5250
      TabIndex        =   9
      ToolTipText     =   "Cancels Any Changes you have Made and Closes This Window ..."
      Top             =   2685
      Width           =   1095
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4050
      TabIndex        =   8
      ToolTipText     =   "Save Settings Close Window ...."
      Top             =   2685
      Width           =   1095
   End
   Begin VB.Frame Frame4 
      Caption         =   "Show Files With These Attributes:"
      Height          =   855
      Left            =   75
      TabIndex        =   10
      Top             =   1725
      Width           =   6255
      Begin VB.CheckBox ChckShowHidden 
         Caption         =   "Show Hidden Files"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Shows Files With Hidden Attributes"
         Top             =   240
         Width           =   1695
      End
      Begin VB.CheckBox ChckShowArchive 
         Caption         =   "Show Archive Files"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Shows Files With Archive Attributes"
         Top             =   480
         Width           =   1695
      End
      Begin VB.CheckBox ChckShowNormal 
         Caption         =   "Show Normal Files"
         Height          =   255
         Left            =   2160
         TabIndex        =   4
         ToolTipText     =   "Shows Files With Normal Attributes"
         Top             =   240
         Width           =   1695
      End
      Begin VB.CheckBox ChckShowReadonly 
         Caption         =   "Show Read Only Files"
         Height          =   255
         Left            =   2160
         TabIndex        =   7
         ToolTipText     =   "Shows Files With Read Only Attributes"
         Top             =   480
         Width           =   1935
      End
      Begin VB.CheckBox ChckShowSystem 
         Caption         =   "Show System Files"
         Height          =   255
         Left            =   4080
         TabIndex        =   5
         ToolTipText     =   "Shows Files With System Attributes"
         Top             =   240
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FrmBrowseBarOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancel_Click()
  Unload Me
End Sub

Private Sub CmdOk_Click()
  
  ' update the browsebar of the types of files to show
  '-------------------------------------------------------------------
  With FrmMain.FileBrowseBar
      .ReadOnly = ChckShowReadonly.Value
      .Archive = ChckShowArchive.Value
      .Hidden = ChckShowHidden.Value
      .Normal = ChckShowNormal.Value
      .System = ChckShowSystem.Value
  End With
  '--------------------------------------------------------------------
 
  ' save the settings of the types of files to show
  '--------------------------------------------------------------------
  SaveSetting "BrowseBar", "ReadOnly", ChckShowReadonly.Value
  SaveSetting "BrowseBar", "Archive", ChckShowArchive.Value
  SaveSetting "BrowseBar", "Hidden", ChckShowHidden.Value
  SaveSetting "BrowseBar", "Normal", ChckShowNormal.Value
  SaveSetting "BrowseBar", "System", ChckShowSystem.Value
  '--------------------------------------------------------------------
  
  ' Update the file pattern
  '--------------------------------------------------------------------
  SaveSetting "BrowseBar", "Pattern", TxtPattern.Text
  FrmMain.FileBrowseBar.Pattern = TxtPattern.Text
  '--------------------------------------------------------------------
  Unload Me
End Sub

Private Sub CmdResetPattern_Click()
  TxtPattern.Text = "*.TXT;*.INI"
End Sub

Private Sub Form_Load()
With FrmMain.FileBrowseBar
  '----------------------------------------------------------------------
  ChckShowReadonly.Value = Abs(CInt(.ReadOnly))
  ChckShowArchive.Value = Abs(CInt(.Archive))
  ChckShowHidden.Value = Abs(CInt(.Hidden))
  ChckShowNormal.Value = Abs(CInt(.Normal))
  ChckShowSystem.Value = Abs(CInt(.System))
  '----------------------------------------------------------------------
  
  'Pattern
  TxtPattern.Text = .Pattern
End With
    
  
End Sub
