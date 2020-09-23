VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClipboard 
   Caption         =   "Clipboard"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7380
   Icon            =   "FrmClipboard.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   7380
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1980
      Top             =   1770
   End
   Begin VB.TextBox TxtClip 
      Height          =   4200
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   0
      Width           =   7335
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2040
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClipboard.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClipboard.frx":0556
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClipboard.frx":066A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClipboard.frx":077E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   4215
      Width           =   7380
      _ExtentX        =   13018
      _ExtentY        =   1005
      ButtonWidth     =   3201
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Write Text To Clipboard"
            Key             =   "savetoclipboard"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Update Window"
            Key             =   "updatewindow"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Clear Clipboard"
            Key             =   "clearclipboard"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save Text To File..."
            Key             =   "SaveFile"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmClipboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()

     On Error GoTo Outofmemoryerr:
      
   ' Instead of Pasting the Text From the Clipboard we Now
   ' Send A Message.
      SendMessage Me.TxtClip.hWnd, WM_PASTE, 0&, 0&
      Form_Resize
      
Outofmemoryerr:
     If Err.Number <> 0 Then
    ' display the error Message too the user
      MsgBox "NextPad Has Encountered The Following Error(s) While Pasting Your Selection : " _
      & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "If The Error Is " & Chr(34) & "Out of memory" & Chr(34) & " Then NextPad Cannot Paste Anymore Into The Text Box Because It Has Run Out of memory", vbCritical, "NextPad "
      'exit the sub immediately
      Exit Sub
     End If
End Sub

Private Sub Form_Resize()
  
   On Error GoTo Errresize: ' if an error occurs while resizing the form Jump too that line
      
      TxtClip.Width = Me.ScaleWidth - (TxtClip.Left * 2) 'Set Text1's Width too the FOrms Width Multiplying that by the Text1's Left Property
       TxtClip.Height = Me.ScaleHeight - Toolbar1.Height  ' Set Text1's Height too the Forms Height Minus The Frames Hieght

Errresize:
       Exit Sub ' Exit the sub Immediately
End Sub


Private Sub Timer1_Timer()
On Error Resume Next
  Me.Caption = "Clipboard" & ", Size in bytes: " & Len(Clipboard.GetText)
  If Clipboard.GetText <> "" Then 'Detect if clipboard contains text....
      Toolbar1.Buttons("clearclipboard").Enabled = True 'Enable button
  Else
      Toolbar1.Buttons("clearclipboard").Enabled = False 'Disable button
  End If
On Error Resume Next
  If TxtClip.Text <> "" Then
     Toolbar1.Buttons("savetoclipboard").Enabled = True ' Enable button
  Else
     Toolbar1.Buttons("savetoclipboard").Enabled = False ' Disable button
  End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim Response 'Variable to hold response from user
On Error GoTo clipboarderror: ' if an error occurs vb will jump too that line
   
    Select Case Button.Key
         Case "clearclipboard" 'User Pressed ...
          'Display the Warning to the user
            Response = MsgBox("Clear all contents currently on the Clipboard ?", vbYesNo + vbInformation, "Clear Clipboard - NextPad")
                 Select Case Response
                    Case vbYes 'User pressed Yes
                         Clipboard.Clear 'Clear the clipboard
                         TxtClip.Text = "" ' Clear the text box
                    Case vbNo 'User Pressed No
                         Exit Sub 'Cancel ; Exit this sub
                 End Select
         Case "updatewindow" 'User Pressed ...
              ' Instead of Pasting the Text From the Clipboard we Now
              ' Send A Message.
             TxtClip.Text = ""
             SendMessage Me.TxtClip.hWnd, WM_PASTE, 0&, 0&
      
         Case "savetoclipboard" 'User Pressed ...
              If TxtClip.Text = "" Then: Clipboard.Clear: Exit Sub
              TxtClip.SelStart = 0
              TxtClip.SelLength = Len(TxtClip.Text)
              SendMessage TxtClip.hWnd, WM_COPY, 0&, 0&
              TxtClip.SelLength = 0
         Case "SaveFile"
              SaveFile "", TxtClip, False
    End Select
     

clipboarderror:               ' Error Control Starts Here

    If Err.Number <> 0 Then ' IF an Error Number is anything Above or at zero then .. . ...
      ' display error message too user
      MsgBox "Cannot access information in Clipboard" _
      & Chr(13) & "Reason : " & Err.Description, vbCritical, "NextPad - Clipboard"
      Exit Sub ' exit the sub immediately
    End If
            
         
End Sub

