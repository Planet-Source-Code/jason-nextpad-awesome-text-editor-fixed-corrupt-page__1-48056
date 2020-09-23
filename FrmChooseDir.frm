VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmChooseDir 
   Caption         =   "Choose Folder"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5160
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmChooseDir.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   5160
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picContainer1 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   4830
      TabIndex        =   7
      Top             =   3090
      Width           =   4830
      Begin VB.TextBox TxtDir 
         Height          =   285
         Left            =   870
         TabIndex        =   8
         Top             =   0
         Width           =   3915
      End
      Begin VB.Label Label1 
         Caption         =   "Folder:"
         Height          =   255
         Left            =   60
         TabIndex        =   9
         Top             =   45
         Width           =   585
      End
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   360
      Left            =   2580
      TabIndex        =   6
      Top             =   3600
      Width           =   1125
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   360
      Left            =   3840
      TabIndex        =   5
      Top             =   3600
      Width           =   1125
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   4
      Top             =   3915
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwFolders 
      Height          =   2280
      Left            =   195
      TabIndex        =   0
      Top             =   660
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   4022
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
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   705
      TabIndex        =   2
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   675
      TabIndex        =   1
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ImgListDrives 
      Left            =   75
      Top             =   3960
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
            Picture         =   "FrmChooseDir.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmChooseDir.frx":27C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmChooseDir.frx":4F74
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmChooseDir.frx":7728
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmChooseDir.frx":9EDC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Please choose a folder, or type one yourself."
      Height          =   240
      Left            =   225
      TabIndex        =   3
      Top             =   195
      Width           =   4035
   End
End
Attribute VB_Name = "FrmChooseDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdCancel_Click()
  Unload Me
End Sub

Private Sub CmdOk_Click()
  If CBool(CStr(PathIsDirectory(CStr(TxtDir.Text)))) = False Then ' If the destination does not exist then
     ' Notify user
     MsgBox "The directory you would like to use does not exist." + vbNewLine + "Please check to make sure that the path given is correct.", vbCritical, "Select A Directory - Invalid Directory"
     Exit Sub
  End If
  
  sDir = TxtDir.Text
  Unload Me
End Sub

Private Sub Form_Load()
  GetDrives
End Sub

Private Sub Form_Resize()
On Error Resume Next
' Use some insane way of resizing the controls on this form, very time consuming.
' Fortunatelly it works fairly well. The controls must be moved and resized at the same time.
' The key is to maintain some order and keep the controls together and have a standard size.
  TxtDir.Width = Me.ScaleWidth - 1200
  tvwFolders.Width = Me.ScaleWidth - 400
  tvwFolders.Height = Me.ScaleHeight - 1800
  CmdCancel.Left = Me.ScaleWidth - 1300
  CmdOK.Left = Me.ScaleWidth - 2600
  CmdCancel.Top = Me.ScaleHeight - 450
  CmdOK.Top = Me.ScaleHeight - 450
  picContainer1.Top = Me.ScaleHeight - 800
  picContainer1.Width = Me.ScaleWidth
  
End Sub

Private Sub txtDir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  With TxtDir
    .ToolTipText = .Text
  End With
End Sub

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
    If Err.Number = 68 Then MsgBox "Cannot access drive you have selected, " & Err.Description, vbCritical, "Folder View - NextPad"
    MousePointer = vbDefault 'Change mouse pointer to hourglass
    Exit Sub
End Sub
Private Sub tvwFolders_NodeClick(ByVal Node As MSComctlLib.Node)
  On Error GoTo ErrorHandler:
  
    TxtDir.Text = Node.Key  'Set path

ErrorHandler:
    If Err.Number <> 0 Then 'If An Error Has Occured ......
      MsgBox "Cannot access drive you have selected, " & Err.Description, vbCritical, "Folder View - NextPad"
      Exit Sub
    End If
End Sub

