VERSION 5.00
Begin VB.Form FrmRecentFiles 
   Caption         =   "Recent Files"
   ClientHeight    =   2955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4980
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmRecentFiles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   4980
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicBottom 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   645
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   4980
      TabIndex        =   5
      Top             =   2310
      Width           =   4980
      Begin VB.CommandButton CmdRemove 
         Caption         =   "Remove"
         Height          =   360
         Left            =   660
         TabIndex        =   2
         Top             =   240
         Width           =   1125
      End
      Begin VB.CommandButton CmdOpen 
         Caption         =   "Open"
         Height          =   360
         Left            =   1875
         TabIndex        =   3
         Top             =   240
         Width           =   1125
      End
      Begin VB.CommandButton CmdDelete 
         Caption         =   "Delete..."
         Height          =   360
         Left            =   3105
         TabIndex        =   4
         Top             =   240
         Width           =   1125
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   90
         X2              =   4890
         Y1              =   150
         Y2              =   150
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderStyle     =   6  'Inside Solid
         Index           =   1
         X1              =   90
         X2              =   4890
         Y1              =   135
         Y2              =   135
      End
   End
   Begin VB.ListBox ListRecentFiles 
      Height          =   2010
      Left            =   60
      TabIndex        =   1
      Top             =   270
      Width           =   4800
   End
   Begin VB.Label Label1 
      Caption         =   "&Recent Files :"
      Height          =   210
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   4545
   End
End
Attribute VB_Name = "FrmRecentFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdDelete_Click()
   DeleteFile Me.hWnd, ListRecentFiles.Text, True
End Sub

Private Sub CmdOpen_Click()
  Dim sFileName As String
 If Fstate.dirty Then SaveFile sOpenFileName, FrmMain.mTextbox, True
 sFileName = ListRecentFiles.Text
  'Call the OpenFile sub in modmain
  If bFileExists(sFileName) = False Then: NotifyFileNonExistent (sFileName): Exit Sub
  OpenFile sFileName
  ' Open the Favorite From the registry using
  ' the Procedure in ModMain ,
  Unload Me
End Sub

Private Sub CmdRemove_Click()
Dim MyInt As Integer
   
   MyInt = ListRecentFiles.ListIndex ' Append selected value to variable.
   
   If MyInt = -1 Then ' No item is currently selected.
      MsgBox "Please select an item to remove.", vbInformation, "Recent Files"
      Exit Sub
   Else
      ListRecentFiles.RemoveItem ListRecentFiles.ListIndex ' Remove the item
   End If

On Error Resume Next ' An error might occur
   DeleteSetting "NextPad", "RecentFiles", MyInt '+ 1 ' Delete the setting
   Form_Unload (1) ' Execute code used when form is finished executing but we dont close.
 
End Sub

Private Sub Form_Load()
On Error GoTo ErrorHandler:
  
  ListRecentFiles.Clear
  
  Dim i As Integer, X As Integer
   i = GetSetting("NextPad", "RecentFiles", "Count", 0)
   If GetSetting("NextPad", "RecentFiles", 0, "") = "" Then Exit Sub
     
     With ListRecentFiles
        For X = 0 To i
          .AddItem GetSetting("NextPad", "RecentFiles", X)
        Next X
     End With

ErrorHandler:
   If Err.Number <> 0 Then
      MsgBox "Could not get Recent Files from registry. If problem persist's please go to the Options window and reset NextPad's settings.", vbCritical, "NextPad - Recent Files"
      Exit Sub
   End If

End Sub

Private Sub Form_Resize()
On Error Resume Next
  ListRecentFiles.Height = Me.ScaleHeight - (PicBottom.Height + Label1.Height + 10)
  ListRecentFiles.Width = Me.ScaleWidth - 100
  Line1(0).X2 = Me.ScaleWidth - 50
  Line1(1).X2 = Me.ScaleWidth - 50
End Sub

Private Sub Form_Unload(Cancel As Integer)
   ' Delete the Setting's Specified , "RecentFiles"
 If FrmMain.MnuRecentFiles.Count > 0 Then
   On Error Resume Next
   DeleteSetting "NextPad", "RecentFiles"
   
   Dim i As Integer, X As Integer
   ' Get the menu count
   ' Make sure we take out this menu item also.
   FrmMain.MnuRecentFiles(0).Caption = "": FrmMain.MnuRecentFiles(0).Visible = False
   i = FrmMain.MnuRecentFiles.Count
    For X = 1 To i Step 1 ' Go through them one by one
      If X = i Then Exit For ' If the variable is equal to the count then stop
      Unload FrmMain.MnuRecentFiles(X) ' Unload the specified menu
    Next X ' Next X value
 End If
 
  If ListRecentFiles.List(0) <> "" Then ' If there is a list then...

   i = 0 ' Reset variable
   X = 0 ' Reset variable
   
   i = ListRecentFiles.ListCount  ' Get the list count
   SaveSetting "RecentFiles", "Count", i - 1
    ' IF there is only one item in the list then...
    If ListRecentFiles.List(0) <> "" And ListRecentFiles.List(1) = "" And ListRecentFiles.List(2) = "" Then
       SaveSetting "RecentFiles", 0, ListRecentFiles.List(0)
       SaveSetting "RecentFiles", "Count", 0
       GetRecentFiles
       Exit Sub
    End If
     
    i = ListRecentFiles.ListCount  ' Get the list count
    For X = 0 To i
      If ListRecentFiles.List(X) = "" Then Exit For
      SaveSetting "RecentFiles", X, ListRecentFiles.List(X)
    Next X
    
    GetRecentFiles

  End If

 
End Sub

