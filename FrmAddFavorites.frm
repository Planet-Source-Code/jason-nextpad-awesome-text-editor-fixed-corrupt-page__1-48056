VERSION 5.00
Begin VB.Form FrmAddFavorites 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add To Favorites"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4860
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmAddFavorites.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox FileList 
      Height          =   285
      Left            =   0
      Pattern         =   "*.TXT;*.INI"
      TabIndex        =   7
      Top             =   1515
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Directory..."
      Height          =   360
      Left            =   3345
      TabIndex        =   6
      Top             =   120
      Width           =   1365
   End
   Begin VB.CommandButton CmdAddOther 
      Caption         =   "Browse..."
      Height          =   360
      Left            =   2100
      TabIndex        =   1
      Top             =   120
      Width           =   1125
   End
   Begin VB.CommandButton CmdAddOpenFile 
      Caption         =   "Currently Open File"
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1890
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   360
      Left            =   3600
      TabIndex        =   5
      Top             =   1320
      Width           =   1125
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "Add To Favorites"
      Default         =   -1  'True
      Height          =   360
      Left            =   1845
      TabIndex        =   4
      Top             =   1320
      Width           =   1665
   End
   Begin VB.TextBox TxtFile 
      Height          =   330
      Left            =   120
      TabIndex        =   3
      Top             =   780
      Width           =   4590
   End
   Begin VB.Label Label1 
      Caption         =   "&Directory or File Name:"
      Height          =   225
      Left            =   150
      TabIndex        =   2
      Top             =   555
      Width           =   3525
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   135
      X2              =   4710
      Y1              =   1215
      Y2              =   1215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   150
      X2              =   4710
      Y1              =   1200
      Y2              =   1200
   End
End
Attribute VB_Name = "FrmAddFavorites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancel_Click()
 Unload Me
End Sub

Private Sub CmdOk_Click()
 
 ' If the user clicked the Add Directory button...
 If CBool(PathIsDirectory(TxtFile.Text)) Then
   On Error GoTo ErrorHandler:
    ' If the complete path does not contain the \ then add it we need it
    If Not Right(TxtFile.Text, 1) = "\" Then TxtFile.Text = TxtFile.Text & "\"
    ' Let the file list know what path we are sorting through
    FileList.Path = TxtFile.Text
    ' If this directory does not contain text files or INI files then notify the user.
    If FileList.ListCount = 0 Then
       MsgBox "The directory you wish to use does not contain any text files." & vbNewLine & "Please select a directory that contains text files.", vbInformation, "NextPad"
       TxtFile.Text = ""
       Exit Sub
    End If
    ' Let the user know that NextPad is busy
    Screen.MousePointer = vbHourglass
    ' Now we will get to work by getting the list count then sorting through them
    For i = 0 To FileList.ListCount - 1
        ' Add current file to list of Favorites
        AddFavorite (TxtFile.Text & FileList.List(i))
    Next i
    ' Reset mouse pointer
    Screen.MousePointer = vbNormal
    ' Close this window
    Unload Me
    Exit Sub
  End If
        
 ' If the user decided to add one file
 If bFileExists(TxtFile.Text) = False Then
    MsgBox "Please select a file that exists to add to your Favorites.", vbInformation, "Add Favorites - NextPad"
    Exit Sub
 End If
 
 AddFavorite (TxtFile.Text)
 Unload Me

ErrorHandler:
  If Err.Number <> 0 Then
     MsgBox "The directory you are attempting to use does not exist, Please check the name and path and try again." & vbNewLine & Err.Description, vbCritical, "NextPad"
     Exit Sub
  End If
  
End Sub

Private Sub Command1_Click()
Dim PreviousString As String, NewString As String
  ' Just in case the user clicks cancel
  PreviousString = TxtFile.Text
   NewString = ChooseDir
   If NewString <> "" Then
      TxtFile.Text = NewString
   Else
      TxtFile.Text = PreviousString
   End If
End Sub

Private Sub Form_Load()
    Select Case sOpenFileName
       Case ""
          CmdAddOpenFile.Enabled = False
          cmdOK.Enabled = False
       Case Else
          TxtFile.Text = sOpenFileName
    End Select
End Sub

Private Sub CmdAddOpenFile_Click()
        TxtFile.Text = sOpenFileName
End Sub

Private Sub CmdAddOther_Click()
On Error GoTo ErrorHandler:

         With FrmMain
            With .CommonDialog1
                 .Flags = Normal_Cdlogflags    ' use the constant in modmain
                 .Cancelerror = True ' set Cancel Error Too true When User clicks cancel an error will generate
                 ' set the filter
                 .Filter = "Text Files (*.TXT) |*.TXT|Ini Files (*.INI) |*.INI|Log Files (*.LOG) |*.LOG|All Files (*.*) |*.*"
                 ' set the commondialogs title
                 .DialogTitle = "Please Select A File To Add To Your Favorites..."
                 .ShowOpen ' Show the open dialog
                 TxtFile.Text = .Filename
                 Exit Sub
            End With
        End With

ErrorHandler:
  If Err.Number = 32755 Then ' Cancel was pressed Exit Sub
     Exit Sub
  Else
     MsgBox "An error occured while performing a requested function" & vbNewLine & "Reason :" & Err.Description, vbCritical, "NextPad - Add To Favorites"
     Exit Sub
  End If
End Sub

Private Sub TxtFile_Change()
  If TxtFile.Text = "" Then
     cmdOK.Enabled = False
  Else
     cmdOK.Enabled = True
  End If
End Sub
