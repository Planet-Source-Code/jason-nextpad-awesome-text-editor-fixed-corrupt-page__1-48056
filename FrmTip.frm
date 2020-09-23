VERSION 5.00
Begin VB.Form frmTip 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tip of the Day"
   ClientHeight    =   3090
   ClientLeft      =   2355
   ClientTop       =   2385
   ClientWidth     =   5130
   Icon            =   "FrmTip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton CmdNextTip 
      Caption         =   "&Next Tip"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3945
      TabIndex        =   5
      Top             =   585
      Width           =   1125
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
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
      Height          =   360
      Left            =   3945
      TabIndex        =   4
      Top             =   120
      Width           =   1125
   End
   Begin VB.TextBox TxtTips 
      Height          =   1935
      Left            =   4080
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Text            =   "FrmTip.frx":08CA
      Top             =   2010
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2715
      Left            =   120
      Picture         =   "FrmTip.frx":0F6B
      ScaleHeight     =   2655
      ScaleWidth      =   3675
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Did you know..."
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
         Left            =   540
         TabIndex        =   2
         Top             =   180
         Width           =   2655
      End
      Begin VB.Label lblTipText 
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
         Height          =   1875
         Left            =   60
         TabIndex        =   1
         Top             =   720
         Width           =   3495
      End
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' The in-memory database of tips.
Dim Tips As New Collection

' Name of tips file
Const TIP_FILE = "TIPS.TIPS"

' Index in collection of tip currently being displayed.
Dim CurrentTip As Long
Private Sub DoNextTip()

    ' Select a tip at random.
     CurrentTip = Int((Tips.Count * Rnd) + 1)
    
    ' Or, you could cycle through the Tips in order

  ' CurrentTip = CurrentTip + 1
   ' If Tips.Count < CurrentTip Then
   '     CurrentTip = 1
   ' End If
    
    ' Show it.
    frmTip.DisplayCurrentTip
    
End Sub

Function LoadTips(sFile As String) As Boolean
    Dim NextTip As String   ' Each tip read in from file.
    Dim InFile As Integer   ' Descriptor for file.
    
    ' Obtain the next free file descriptor.
    InFile = FreeFile
    
    ' Make sure a file is specified.
    If sFile = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' Make sure the file exists before trying to open it.
    If Dir(sFile) = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' Read the collection from a text file.
    Open sFile For Input As InFile
    While Not EOF(InFile)
        Line Input #InFile, NextTip
        Tips.Add NextTip
    Wend
    Close InFile

    ' Display a tip at random.
    DoNextTip
    
    LoadTips = True
    
End Function


Private Sub cmdNextTip_Click()
    DoNextTip
End Sub

Private Sub CmdOk_Click()
    Unload Me
End Sub

Private Sub Form_Load()
   On Error GoTo ErrorHandler:
   
    Dim TipsPath As String
    Dim sTipsPath As String
    ' See if we should be shown at startup
            
    
    TipsPath = App.Path & "\" & TIP_FILE
    If bFileExists(TipsPath) = False Then TipsPath = App.Path & TIP_FILE
    
    sTipsPath = App.Path & "\"
    If bFileExists(sTipsPath) = False Then sTipsPath = App.Path
    
    ' Read in the tips file and display a tip at random.
    
    ' Instead Of Telling The User That The Tips File Is Not
    ' Found We Will Just Save A New One , If The Origanal Isnt Found ,
    ' This Way If The User Accidentally Deletes or Misplaces The File
    ' We Have The Whole Document Writted In The TxtTips TextBox So We
    ' Can Just Get The Applications Path And Resave It There
    If LoadTips(TipsPath) = False Then
       Close #1 'Close The Open File
       ' Open The sTipsPath & TIP_FILE For Output
       Open sTipsPath & TIP_FILE For Output As #1
       ' Write Too The File
       Print #1, TxtTips.Text
       ' Close The File
       Close #1
       ' Recall Form_load
       Form_Load
    End If

ErrorHandler:
   If Err.Number <> 0 Then
      Exit Sub
   End If
End Sub

Public Sub DisplayCurrentTip()
    If Tips.Count > 0 Then
        lblTipText.Caption = Tips.Item(CurrentTip)
    End If
End Sub

