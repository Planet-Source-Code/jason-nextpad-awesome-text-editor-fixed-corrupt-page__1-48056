VERSION 5.00
Begin VB.Form FrmNotifyAss 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NextPad"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   Icon            =   "FrmNotifyAss.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdYes 
      Caption         =   "&Yes"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2910
      TabIndex        =   1
      Top             =   2220
      Width           =   1155
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   5535
      TabIndex        =   4
      Top             =   0
      Width           =   5535
   End
   Begin VB.CommandButton CmdNo 
      Caption         =   "&No"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4155
      TabIndex        =   2
      Top             =   2220
      Width           =   1155
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Don't show me this in the future."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   75
      TabIndex        =   3
      Top             =   2805
      Width           =   4350
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   15
      X2              =   5600
      Y1              =   750
      Y2              =   750
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   0
      X2              =   5585
      Y1              =   735
      Y2              =   735
   End
   Begin VB.Label Label1 
      Caption         =   $"FrmNotifyAss.frx":000C
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   105
      TabIndex        =   0
      Top             =   930
      Width           =   5220
   End
End
Attribute VB_Name = "FrmNotifyAss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
   SaveSetting "Associations", "Notify", Check1.Value
End Sub

Private Sub CmdNo_Click()
   Unload Me
End Sub

Private Sub CmdYes_Click()
   Dim StrPathAndExe As String 'String Value to Hold the Correct path and filename of this execuatable
   
   Dim Retval As String ' Declare String variable
    
   'Set the StrPathAndExe String value to hold the Applications current path and Executable name
   StrPathAndExe = App.Path & "\" & App.EXEName & ".EXE"
   ' If the file is in the root Directory "C:\" Then We remove the appended
   ' Backslash \ to fit the current path and executable filename
   If bFileExists(StrPathAndExe) = False Then StrPathAndExe = App.Path & App.EXEName & ".EXE"
   

          'Set NextPad as the default TXT Viewer
             'Save the current TXT Viewer Path  for later retrieval
           SaveSetting "Associations", "TXT" _
           , GetSettingString(HKEY_CLASSES_ROOT, "txtfile\shell\open\command", "", "")
           'Save the New one (NextPad)
           SaveSettingString HKEY_CLASSES_ROOT, "txtfile\shell\open\command" _
           , "", StrPathAndExe & " %1"
           SaveSetting "associations", "isassociated", "0"
     
       Unload Me
End Sub



Private Sub Form_Load()
   Picture1.Picture = frmAbout.Picabout.Picture
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set FrmNotifyAss = Nothing 'Release memory taken
End Sub
