VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About NextPad"
   ClientHeight    =   5070
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5385
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3499.405
   ScaleMode       =   0  'User
   ScaleWidth      =   5056.793
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picabout 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      Picture         =   "frmAbout.frx":000C
      ScaleHeight     =   675
      ScaleWidth      =   5325
      TabIndex        =   9
      Top             =   0
      Width           =   5385
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4080
      TabIndex        =   0
      Top             =   4680
      Width           =   1260
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Licensing and legal information ..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      MouseIcon       =   "frmAbout.frx":0855
      TabIndex        =   13
      ToolTipText     =   "Click for more info..."
      Top             =   4680
      Width           =   2475
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "This release is the 49th  release since Sunday, September 24, 2001"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   2760
      Width           =   5055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Special thanks to Matt for his assistance in testing and suggestions."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      MouseIcon       =   "frmAbout.frx":09A7
      TabIndex        =   11
      ToolTipText     =   "Click for more info..."
      Top             =   4200
      Width           =   3135
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Written by : Jason - Simeone"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   3720
      Width           =   3495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Special thanks to ElitePad for its registry module."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      MouseIcon       =   "frmAbout.frx":0AF9
      TabIndex        =   8
      ToolTipText     =   "Click for more info..."
      Top             =   3960
      Width           =   3735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Please  email me about any bugs or, if you have any suggestions."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   4875
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Compiled on : Wednesday, August 27, 2003"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3000
      Width           =   3825
   End
   Begin VB.Label lbllicense 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "This software is licensed to:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   225.372
      X2              =   4958.192
      Y1              =   2267.366
      Y2              =   2267.366
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "A small application for creating, editing, saving, printing and managing  text documents.  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   4605
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   225.372
      X2              =   4958.192
      Y1              =   2277.719
      Y2              =   2277.719
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 4.121 Beta 30 Revision 4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   4125
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   "Email address : Cyberarea@hotmail.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   240
      TabIndex        =   2
      Top             =   3480
      Width           =   3990
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Sub CmdOk_Click()
   
   Unload Me ' unload the form

End Sub
Private Sub Form_Load()

On Error GoTo Picerr: ' if an error occurs vb will jump too that line
  
    Dim UserName As String
    
    UserName = String(164, 0) ' Set up buffer
    GetUserName UserName, 164 ' Pass it to the function
    StripTerminator UserName ' Strip all null characters
    lbllicense.Caption = UserName ' Set caption


   Beep ' Beep From the Computer Speaker Or The Default WIndows Beep
   ' Too grab the users attention

Picerr:
    If Err.Number <> 0 Then ' if the error's number is equal too anything above or other than zero then.....
     Exit Sub ' exit this sub immediately
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Label4.ForeColor = vbBlack
  Label5.ForeColor = vbBlack
  Label8.ForeColor = vbBlack
End Sub

Private Sub Form_Unload(Cancel As Integer)
'**************************************************
'This event is triggered when the form is unloaded*
'**************************************************
      Set frmAbout.Picabout = Nothing 'set this picture to nothing too release memory it held
      Set frmAbout = Nothing ' set this form too nothing too release memory it held

End Sub


Private Sub Label4_Click()
     MsgBox "Special thanks goes out to Matt my great friend, Matt has spent countless hours testing and making suggestions." & vbNewLine & "Thanks", vbInformation, "Thanks"
     
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Label5.ForeColor = vbBlack
  Label4.ForeColor = &HFF0000
  Label8.ForeColor = vbBlack

End Sub

Private Sub Label5_Click()
     MsgBox "Special thanks goes out too the creator and developer of ElitePad for its great registry module." + vbNewLine + vbNewLine + "Thank you", vbInformation, "Thanks"
   
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Label5.ForeColor = &HFF0000
  Label4.ForeColor = vbBlack
  Label8.ForeColor = vbBlack
  
End Sub

Private Sub Label8_Click()
  MsgBox "By downloading this program and running it on your computer you accept that the author takes no responsibility for any direct or indirect damage caused to you or your computer. " _
  & vbNewLine & vbNewLine & "This program is intended " & _
                            "solely to provide general usage to you the user" & _
                            ", whom accepts full responsibility for its use. " & _
                            "It is provided as is, with no guarantee of " & _
                            "completeness or accuracy and without warranty " & _
                            "of any kind, express or implied." & _
  vbNewLine & vbNewLine & "You may distribute this program freely, infact i encourage you to pass it along to your friends. Although you can distribute it freely you still are not allowed to sell it." & _
  vbNewLine & vbNewLine & "Thank you and enjoy your user experience." & _
  vbNewLine & vbNewLine & "If you need help, or there is a problem with this software that you would like to address in the next release, please feel free to email me at:" & _
  vbNewLine & vbNewLine & "Cyberarea@hotmail.com" _
  , vbInformation, "NextPad License Info"
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Label5.ForeColor = vbBlack
  Label4.ForeColor = vbBlack
  Label8.ForeColor = &HFF0000

End Sub
