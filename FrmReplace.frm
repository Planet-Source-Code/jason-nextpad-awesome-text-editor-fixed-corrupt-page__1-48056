VERSION 5.00
Begin VB.Form frmReplace 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Replace"
   ClientHeight    =   3210
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
   Icon            =   "FrmReplace.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkCase 
      Caption         =   "Match &Case"
      Height          =   330
      Left            =   210
      TabIndex        =   9
      ToolTipText     =   "Case Sensitivity"
      Top             =   1395
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Options"
      Height          =   1995
      Left            =   90
      TabIndex        =   8
      Top             =   1155
      Width           =   3780
      Begin VB.OptionButton OptReplacement 
         Caption         =   "&Don't Change Replacement"
         Height          =   255
         Index           =   0
         Left            =   105
         TabIndex        =   14
         Top             =   1665
         Value           =   -1  'True
         Width           =   2625
      End
      Begin VB.OptionButton OptReplacement 
         Caption         =   "Make Replacement &Invert Case"
         Height          =   255
         Index           =   3
         Left            =   105
         TabIndex        =   13
         Top             =   1410
         Width           =   2670
      End
      Begin VB.OptionButton OptReplacement 
         Caption         =   "Make Replacement &Lowercase"
         Height          =   255
         Index           =   2
         Left            =   105
         TabIndex        =   12
         Top             =   1140
         Width           =   2535
      End
      Begin VB.OptionButton OptReplacement 
         Caption         =   "Make Replacement &Uppercase"
         Height          =   375
         Index           =   1
         Left            =   105
         TabIndex        =   11
         Top             =   810
         Width           =   2790
      End
      Begin VB.OptionButton OptReplacement 
         Caption         =   "Make Replacement &Proper Case"
         Height          =   330
         Index           =   4
         Left            =   105
         TabIndex        =   10
         Top             =   555
         Width           =   2730
      End
   End
   Begin VB.CommandButton CmdReplaceAll 
      Cancel          =   -1  'True
      Caption         =   "Replace &All"
      Height          =   372
      Left            =   4080
      TabIndex        =   6
      Top             =   1065
      Width           =   1092
   End
   Begin VB.CommandButton CmdReplace 
      Caption         =   "&Replace"
      Height          =   372
      Left            =   4080
      TabIndex        =   5
      Top             =   585
      Width           =   1092
   End
   Begin VB.ComboBox CmbFind 
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   165
      Width           =   2670
   End
   Begin VB.ComboBox CmbReplace 
      Height          =   315
      Left            =   1200
      TabIndex        =   3
      Top             =   585
      Width           =   2670
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4845
      Top             =   2775
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Top             =   1545
      Width           =   1095
   End
   Begin VB.CommandButton CmdFind 
      Caption         =   "Find &Next"
      Default         =   -1  'True
      Height          =   372
      Left            =   4080
      TabIndex        =   4
      Top             =   120
      Width           =   1092
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2160
      Left            =   45
      ScaleHeight     =   2160
      ScaleWidth      =   3960
      TabIndex        =   15
      Top             =   1080
      Width           =   3960
   End
   Begin VB.Label Label2 
      Caption         =   "Replace &With:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   615
      Width           =   1095
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
Attribute VB_Name = "frmReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FirstTime As Boolean
Dim rFirstTime As Boolean
Dim OptReplacementVal As Integer

Private Sub chkCase_Click()
    gFindCase = chkCase.Value
End Sub

Private Sub CmdCancel_Click()
    gFindCase = chkCase.Value
    Unload frmReplace
End Sub

Private Sub CmdFind_Click()
  ' Add the new replace string to the collection.
 
 FindStrings.Add CmbFind.Text
 CmbFind.AddItem CmbFind.Text

' If this is the first time then reset the selection start point.
 If FirstTime Then FrmMain.mTextbox.SelStart = 0
    gFindDirection = 1 ' Set the direction to go from the top down.
    ' Add each value to itself + 1.
 If Not FirstTime Then FrmMain.mTextbox.SelStart = FrmMain.mTextbox.SelStart + 1
    gFindString = CmbFind.Text ' Set search string
    FindStr  ' Find the string without replacing any thing.
    FirstTime = False ' Its not the first time any more
End Sub

Private Sub Cmdreplace_Click()
Dim ReplaceStr As String
 
   ReplaceStr = CmbReplace.Text ' Append text to variable
  
  ' Add the new replace string to the collection.
   ReplaceStrings.Add ReplaceStr
   CmbReplace.AddItem ReplaceStr
 
 Select Case OptReplacementVal ' Check the value
     Case 0 ' Propercase
       ' Convert to propercase
       ReplaceStr = StrConv(ReplaceStr, vbProperCase)
     Case 1 ' Uppercase
       ' Convert to uppercase
       ReplaceStr = StrConv(ReplaceStr, vbUpperCase)
     Case 2 ' Lowercase
       ' Convert to lowercase
       ReplaceStr = StrConv(ReplaceStr, vbLowerCase)
     Case 3
       ' Convert to inverted case
       ReplaceStr = StrInvertCase(ReplaceStr)
 End Select
 
' If this is the first time then reset the selection start point.
 If rFirstTime Then FrmMain.mTextbox.SelStart = 0
    gFindDirection = 1 ' Set the direction to go from the top down.
    ' Add each value to itself + 1.
 If Not rFirstTime Then FrmMain.mTextbox.SelStart = FrmMain.mTextbox.SelStart + 1
    gFindString = CmbFind.Text ' Set search string
    ' Find the string but, if the string is not found then stop the replacement.
    If Not FindStr Then Exit Sub
    rFirstTime = False ' Its not the first time any more
    FrmMain.mTextbox.SelText = ReplaceStr
    FrmMain.SetFocus
End Sub

Private Sub Cmdreplaceall_Click()
Dim ReplaceStr As String
Dim ReplacedTimes As Long
   ReplaceStr = CmbReplace.Text ' Append text to variable
  
  ' Add the new replace string to the collection.
   ReplaceStrings.Add ReplaceStr
   CmbReplace.AddItem ReplaceStr
 
 Select Case OptReplacementVal ' Check the value
     Case 4 ' Propercase
       ' Convert to propercase
       ReplaceStr = StrConv(ReplaceStr, vbProperCase)
     Case 1 ' Uppercase
       ' Convert to uppercase
       ReplaceStr = StrConv(ReplaceStr, vbUpperCase)
     Case 2 ' Lowercase
       ' Convert to lowercase
       ReplaceStr = StrConv(ReplaceStr, vbLowerCase)
     Case 3
       ' Convert to inverted case
       ReplaceStr = StrInvertCase(ReplaceStr)
 End Select
 
' If this is the first time then reset the selection start point.
 If rFirstTime Then FrmMain.mTextbox.SelStart = 0
    gFindDirection = 1 ' Set the direction to go from the top down.
    ' Add each value to itself + 1.
 If Not rFirstTime Then FrmMain.mTextbox.SelStart = FrmMain.mTextbox.SelStart + 1
    gFindString = CmbFind.Text ' Set search string
    ' Find the string but, if the string is not found then stop the replacement.
    ' Keep looping, if we find the string replace it...
    Do Until FindStr = False
      Screen.MousePointer = vbHourglass
      ReplacedTimes = ReplacedTimes + 1
      rFirstTime = False ' Its not the first time any more
      FrmMain.mTextbox.SelText = ReplaceStr
      FrmMain.SetFocus
    Loop
    MsgBox "The specified region has been searched." & vbNewLine & _
           ReplacedTimes & " replacements have been made.", vbInformation, "NextPad"
    Screen.MousePointer = vbDefault
       
    
End Sub

Private Sub Form_Load()
    FirstTime = True ' This is the first time we assume.
    rFirstTime = True ' This is the first time we assume.
    CmdFind.Enabled = False ' Disable it
    CmbFind.Text = FrmMain.mTextbox.SelText ' Set the text to the text currently selected
    ' Add the following items to the combo box to allow the user to use.
    With CmbReplace
        .AddItem Format(Date, "dddd, mmmm dd, yyyy")
        .AddItem Date
        .AddItem Format(Time, "hh:mm")
        .AddItem Time
        .AddItem String(70, "-")
      For i = 1 To ReplaceStrings.Count
         CmbReplace.AddItem CStr(ReplaceStrings.Item(i))
      Next i
    End With
    
      For i = 1 To FindStrings.Count
         CmbFind.AddItem CStr(FindStrings.Item(i))
      Next i

End Sub


Private Sub OptReplacement_Click(Index As Integer)
  OptReplacementVal = Index
End Sub

Private Sub Timer1_Timer()
   ' If there is no text in the text boxes then disable the replace button.
   If CmbFind.Text = "" And CmbReplace.Text = "" Or CmbFind.Text = "" Then
      CmdReplace.Enabled = False
      CmdReplaceAll.Enabled = False
   Else
      CmdReplace.Enabled = True
      CmdReplaceAll.Enabled = True
   End If
End Sub

Private Sub CmbFind_Change()
    gFirstTime = True
    If CmbFind.Text = "" Then
       CmdFind.Enabled = False
       CmdReplace.Enabled = False
       CmdReplaceAll.Enabled = False
    Else
       CmdFind.Enabled = True
       CmdReplace.Enabled = True
       CmdReplaceAll.Enabled = True
    End If
End Sub

