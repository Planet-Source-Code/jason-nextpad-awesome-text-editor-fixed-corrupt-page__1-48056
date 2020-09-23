VERSION 5.00
Begin VB.Form frmfullscreen 
   BorderStyle     =   0  'None
   Caption         =   "Full Screen Mode"
   ClientHeight    =   7410
   ClientLeft      =   720
   ClientTop       =   705
   ClientWidth     =   10155
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   7410
   ScaleWidth      =   10155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   720
      Top             =   5160
   End
   Begin VB.TextBox Txtfullscreen 
      Height          =   3495
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "frmfullscreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Me.Width = Screen.Width
Me.Height = Screen.Height
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
Txtfullscreen.Text = FrmMain.mTextbox.Text
Txtfullscreen.Height = Me.ScaleHeight '- 200
Txtfullscreen.Width = Me.ScaleWidth '- 150
Load Frmleavefullscreen
Frmleavefullscreen.Show , Me
End Sub

Private Sub Form_Resize()
Txtfullscreen.Height = Me.ScaleHeight ' - 200
Txtfullscreen.Width = Me.ScaleWidth '- 150

End Sub

Private Sub Timer1_Timer()
If FrmMain.Visible = False Then
Unload Me
Unload Frmleavefullscreen
End If

End Sub

Private Sub Txtfullscreen_Change()
  On Error GoTo outofmemory:
   
   FrmMain.mTextbox.Text = Txtfullscreen.Text
 
outofmemory:
   If Err.Number <> 0 Then
      MsgBox "NextPad Has Encountered The Following Error(s) While Performing The operation you requested : " _
      & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "If The Error Is " & Chr(34) & "Out of memory" & Chr(34) & " Then NextPad Cannot Place Anymore Text Into The Text Box Because It Has Run Out of memory", vbCritical, "NextPad "
      Exit Sub
   End If

End Sub


Private Sub Txtfullscreen_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
       Unload Frmleavefullscreen
       Unload frmfullscreen
       FrmMain.MnuFullScreen.Checked = False
       Set frmfullscreen = Nothing
       Set Frmleavefullscreen = Nothing
       FrmMain.WindowState = vbNormal
       FrmMain.SetFocus
    End If
    
End Sub
