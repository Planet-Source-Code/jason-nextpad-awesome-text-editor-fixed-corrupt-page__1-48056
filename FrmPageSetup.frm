VERSION 5.00
Begin VB.Form FrmPageSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Page Setup"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6570
   Icon            =   "FrmPageSetup.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "&Fonts"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   3960
      TabIndex        =   9
      Top             =   120
      Width           =   2535
      Begin VB.CheckBox ChckPfont 
         Caption         =   "Print Out &Fonts"
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
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   2175
      End
      Begin VB.CheckBox ChckPIFont 
         Caption         =   "Print Out &Italic Fonts"
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
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   2175
      End
      Begin VB.CheckBox ChckPfBold 
         Caption         =   "Print Out &Bold Fonts"
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
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   2295
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "&Colors"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   3735
      Begin VB.CheckBox ChckpFontColor 
         Caption         =   "Print Out &Colored Fonts (If Any)"
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
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   3495
      End
      Begin VB.ComboBox CmboColor 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "&Mode :"
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
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CommandButton CmdOK 
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
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Ca&ncel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   7
      Top             =   1920
      Width           =   1215
   End
End
Attribute VB_Name = "FrmPageSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmboColor_Click()

    ChckpFontColor.Enabled = CmboColor.ListIndex

End Sub

Private Sub CmdCancel_Click()
Unload Me
End Sub

Private Sub CmdOk_Click()
  
  
  Dim ColorMethod As Integer 'Declare Variables
  
  ColorMethod = CmboColor.ListIndex
  ColorMethod = ColorMethod + 1
        'Save the REAL Color Method NOT the LIST Reference
        ' Why ? Because the COnstant is 1 and 2 not 0 & 1 so we have too
        ' add the List Index too the List Index + 1 too get the Correct
        ' Constant Number
         SaveSetting "PrinterSettings", "ColorMethod", ColorMethod
          ' Here we Save The Reference for the COMBOBOX Why ?
          ' Because if we used the REAL Value then the COMBOBOX would Generate AN
          ' Error Because the ComboBox's List INdex gos from 0 and up NOT 1 and up
          ' so we use the Reference
           SaveSetting "PrinterSettings", "ColorMethodRef", CmboColor.ListIndex
        
             SaveSetting "PrinterSettings", "PfColors", ChckpFontColor.Value
              SaveSetting "PrinterSettings", "PfBold", ChckPfBold.Value
               SaveSetting "PrinterSettings", "PiFont", ChckPIFont.Value
                SaveSetting "PrinterSettings", "PfFont", ChckPfont.Value
             
            Unload Me
  
End Sub

Private Sub Form_Load()
  On Error GoTo ErrorInvalidRegistryEntrys:
    Dim ColorMethodRef As Integer 'Declare Variables
    Dim PfColors As Integer 'Declare Variables
    Dim PfBold As Integer 'Declare Variables
    Dim PFont As Integer 'Declare Variables
    Dim PiFont As Integer 'Declare Variables
    
'Retrieve the REFERENCE of Method for the ComboBox Of Color  which is "Monochrome" by Default ,if there is none stored
' Then use the Default Which is 1 "Monochrome"
ColorMethodRef = GetSetting("NextPad", "PrinterSettings", "ColorMethodRef", 0)

PfColors = GetSetting("NextPad", "PrinterSettings", "PfColors", 0)

PfBold = GetSetting("NextPad", "PrinterSettings", "PfBold", 0)

PfFont = GetSetting("NextPad", "PrinterSettings", "PfFont", 0)

PiFont = GetSetting("NextPad", "PrinterSettings", "PiFont", 0)



With CmboColor
    .AddItem "Monochrome Black & White", 0
    .AddItem "Color", 1
End With


CmboColor.ListIndex = ColorMethodRef ' Set the Selected item of
' the ComboBox from the Registry


ChckPfBold.Value = PfBold

ChckPfont.Value = PfFont

ChckPIFont.Value = PiFont

ChckpFontColor.Enabled = ColorMethodRef


ErrorInvalidRegistryEntrys:
    If Err.Number <> 0 Then
        ' Here we Have too Repair ANY and ALL Registry Entrys for The Page Setup Dialog
        ' If any Are Corrupt !!!!
        MsgBox "The settings in the registry for the page setup dialog are corrupted. Please click OK to repair them.", vbCritical, "Page Setup Dialog"
         SaveSetting "PrinterSettings", "ColorMethodRef", 0
        SaveSetting "PrinterSettings", "PfColors", 0
       SaveSetting "PrinterSettings", "PfBold", 0
      SaveSetting "PrinterSettings", "PiFont", 0
     SaveSetting "PrinterSettings", "PfFont", 0
     Form_Load
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set FrmPageSetup = Nothing
  
End Sub

