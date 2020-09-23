VERSION 5.00
Begin VB.Form FrmFileOperation 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4875
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmFileOperation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimEnable 
      Interval        =   1
      Left            =   0
      Top             =   1800
   End
   Begin VB.CommandButton CmdDir 
      Caption         =   "..."
      Height          =   390
      Left            =   4485
      TabIndex        =   5
      ToolTipText     =   "Select a directory."
      Top             =   1185
      Width           =   360
   End
   Begin VB.CommandButton CmdOK 
      Default         =   -1  'True
      Height          =   360
      Left            =   2040
      TabIndex        =   6
      Top             =   1770
      Width           =   1125
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   360
      Left            =   3345
      TabIndex        =   7
      Top             =   1770
      Width           =   1125
   End
   Begin VB.CommandButton CmdFile 
      Caption         =   "..."
      Height          =   390
      Left            =   4470
      TabIndex        =   2
      ToolTipText     =   "Select a source file."
      Top             =   360
      Width           =   360
   End
   Begin VB.TextBox TxtDestination 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1185
      Width           =   4335
   End
   Begin VB.TextBox TxtFile 
      Height          =   375
      Left            =   105
      TabIndex        =   1
      Top             =   360
      Width           =   4335
   End
   Begin VB.Label Label2 
      Caption         =   "&Destination:"
      Height          =   300
      Left            =   120
      TabIndex        =   3
      Top             =   915
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "&Source File:"
      Height          =   225
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   945
   End
End
Attribute VB_Name = "FrmFileOperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancel_Click()
  Unload Me 'Unload This form
End Sub

Private Sub CmdDir_Click()
  TxtDestination.Text = ChooseDir
End Sub

Private Sub CmdFile_Click()
On Error GoTo ErrorHandler:
 
 With FrmMain.CommonDialog1
    .Cancelerror = True
    .DialogTitle = "Please select the file you wish to " + Left$(Me.Caption, 4) + "."
    .Filter = "All Files |*.*|"
    .Flags = cdlOFNHideReadOnly
    .ShowOpen
    TxtFile.Text = .Filename
 End With
 
ErrorHandler:
 If Err.Number = 32755 Then ' If cancel was pressed...
    Exit Sub
 Else ' A different error occured...
    Exit Sub
 End If
End Sub

Private Sub CmdOk_Click()

If bFileExists(TxtFile.Text) = False Then ' If the file does not exist then
   ' Notify user
   MsgBox "The file you are attempting to " + Left$(Me.Caption, 4) + " does not exist. Please verify that path and name given are correct.", vbCritical, Me.Caption
   Exit Sub
End If

If CBool(CStr(PathIsDirectory(CStr(TxtDestination.Text)))) = False Then ' If the destination does not exist then
   ' Notify user
   MsgBox "The destination you are attempting to " + Left$(Me.Caption, 4) + " the file to does not exist. Please verify that path and name given are correct.", vbCritical, Me.Caption
   Exit Sub
End If

  Select Case Me.Caption
      Case "Copy File"
         ' Copy the file
         FileOperation FO_COPY, TxtFile.Text, TxtDestination.Text
         Unload Me
      Case "Move File"
         ' Move the file
         FileOperation FO_MOVE, TxtFile.Text, TxtDestination.Text
         Unload Me
  End Select
End Sub

Private Sub Form_Load()
  CmdOK.Caption = Me.Caption
End Sub

Private Sub TimEnable_Timer()
   ' If the text boxes have no text we can not perform a copy or a move operation. Disable the move/copy button.
      If TxtFile.Text = "" And TxtDestination = "" Then
         CmdOK.Enabled = False
      Else ' However, if the user has added some text enable it.
         CmdOK.Enabled = True
      End If
      
End Sub
