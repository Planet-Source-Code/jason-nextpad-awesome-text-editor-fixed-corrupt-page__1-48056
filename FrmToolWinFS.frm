VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form Frmleavefullscreen 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Full "
   ClientHeight    =   435
   ClientLeft      =   8970
   ClientTop       =   390
   ClientWidth     =   450
   ControlBox      =   0   'False
   Icon            =   "FrmToolWinFS.frx":0000
   LinkTopic       =   "FrmToolWinFS"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   435
   ScaleWidth      =   450
   ShowInTaskbar   =   0   'False
   Begin VB.Timer TimWinState 
      Interval        =   1
      Left            =   0
      Top             =   480
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   360
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   22
      ImageHeight     =   21
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmToolWinFS.frx":000C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TbarFullScreen 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   767
      ButtonWidth     =   767
      ButtonHeight    =   714
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "fullscreen"
            Object.ToolTipText     =   "Click here to return to the normal view. (Esc)"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Frmleavefullscreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
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

Private Sub TimWinState_Timer()
   If FrmMain.WindowState = vbMinimized Then
      Me.Visible = False
   Else
      Me.Visible = True
   End If
End Sub

Private Sub TbarFullScreen_ButtonClick(ByVal Button As MSComctlLib.Button)

  Select Case Button.Key
    Case "fullscreen"
     Unload Me
     Unload frmfullscreen
     FrmMain.MnuFullScreen.Checked = False
     Set frmfullscreen = Nothing
     Set Frmleavefullscreen = Nothing
     FrmMain.WindowState = vbNormal
     FrmMain.SetFocus
  End Select
  

End Sub
