VERSION 5.00
Begin VB.Form Frmbugz 
   Caption         =   "Version History "
   ClientHeight    =   5130
   ClientLeft      =   2235
   ClientTop       =   2985
   ClientWidth     =   8805
   Icon            =   "Frmbugz.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   8805
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicLogoHolder 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   8745
      TabIndex        =   1
      Top             =   0
      Width           =   8805
      Begin VB.PictureBox PicLogo 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   0
         ScaleHeight     =   735
         ScaleWidth      =   6495
         TabIndex        =   2
         Top             =   0
         Width           =   6495
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4365
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "Frmbugz.frx":27A2
      Top             =   720
      Width           =   8760
   End
End
Attribute VB_Name = "Frmbugz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
   On Error GoTo ErrorHandler:
   
     PicLogo.Picture = frmAbout.Picabout.Picture
     
ErrorHandler:
  If Err.Number <> 0 Then
      Exit Sub
  End If
End Sub

Private Sub Form_Resize()

On Error GoTo Resizeerr: ' if an error occurs Vb will Jump too that line

    Text1.Height = Frmbugz.ScaleHeight - PicLogoHolder.Height  'Set the Text1's Height too the height of FrmMain's scale height property
    Text1.Width = Frmbugz.ScaleWidth ' set the Text1's Height too the width of  FrmMain's scale width property

Resizeerr:
    Exit Sub ' Exit the sub immediately

End Sub

Private Sub Form_Unload(Cancel As Integer)

  Set Frmbugz = Nothing ' Release the memory that This form had Held


End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  
  ' Let the user Know that This control cannot Be edited
  Beep

End Sub
