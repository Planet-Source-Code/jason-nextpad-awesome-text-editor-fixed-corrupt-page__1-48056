VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   1680
   ClientLeft      =   1500
   ClientTop       =   1875
   ClientWidth     =   5265
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Find.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CmbFind 
      Height          =   315
      Left            =   1155
      TabIndex        =   1
      Top             =   150
      Width           =   2790
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Replace"
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Direction"
      Height          =   735
      Left            =   2280
      TabIndex        =   6
      Top             =   840
      Width           =   1575
      Begin VB.OptionButton optDirection 
         Caption         =   "&Down"
         Height          =   252
         Index           =   1
         Left            =   720
         TabIndex        =   8
         ToolTipText     =   "Search to End of Document"
         Top             =   360
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optDirection 
         Caption         =   "&Up"
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Search to Beginning of Document"
         Top             =   360
         Width           =   612
      End
   End
   Begin VB.CheckBox chkCase 
      Caption         =   "Match &Case"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmdcancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   372
      Left            =   4080
      TabIndex        =   3
      Top             =   600
      Width           =   1092
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find &Next"
      Default         =   -1  'True
      Height          =   372
      Left            =   4080
      TabIndex        =   2
      Top             =   120
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "&Find What:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   195
      Width           =   975
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub chkCase_Click()
    gFindCase = chkCase.Value
End Sub

Private Sub CmdCancel_Click()
    gFindString = CmbFind.Text
    gFindCase = chkCase.Value
    Unload frmFind
End Sub

Private Sub CmdFind_Click()
    FindStrings.Add CmbFind.Text
    CmbFind.AddItem CmbFind.Text
    gFindString = CmbFind.Text
    FindStr
End Sub

Private Sub Command1_Click()
  frmReplace.Show vbModeless, FrmMain
  Unload Me
End Sub

Private Sub Form_Load()
    cmdFind.Enabled = False
    optDirection(gFindDirection).Value = 1
    
    For i = 1 To FindStrings.Count
        CmbFind.AddItem CStr(FindStrings.Item(i))
    Next i

End Sub

Private Sub optDirection_Click(Index As Integer)
    gFindDirection = Index
End Sub

Private Sub CmbFind_Change()
    gFirstTime = True
    If CmbFind.Text = "" Then
        cmdFind.Enabled = False
    Else
        cmdFind.Enabled = True
    End If
End Sub

