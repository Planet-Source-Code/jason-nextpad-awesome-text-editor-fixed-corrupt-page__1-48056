Attribute VB_Name = "Modoptions"

'*******************************************************************************************************************
' This module makes it easier to load settings without bogging down NextPad.                                       *
' You can also modify the default behavior of NextPad with this module also.                                       *
'*******************************************************************************************************************

Sub SaveSetting_Toolbar(bVal As Boolean)
   FrmMain.Toolbar.Visible = bVal
   FrmMain.MnuToolbar.Checked = FrmMain.Toolbar.Visible
   ResizeNoteWithToolbar ' resize the form
   SaveSetting "Toolbar", "Visible", Abs(CInt(CBool(FrmMain.Toolbar.Visible)))
End Sub

Sub SaveSetting_UseExternalEditor(bVal As Boolean)
   Dim Retval As Long 'declare Variables
    Retval = IIf(bVal, True, False)
    Select Case Retval
      Case True
       SaveSetting "ExternalEditor", "Use", 1
     Case False
       SaveSetting "ExternalEditor", "Use", 0
    End Select
End Sub
 

Sub SaveSetting_Wordwrap(bVal As Boolean)
  ToggleWordWrap (bVal) ' togglewordwrap (Retval as BOOLEAN)
  SaveSetting "Wordwrap", "Wordwrap", Abs(CInt(CBool(bVal)))
End Sub
Sub SaveSetting_chckassociations(bVal As Boolean)
  Dim Retval As Long 'declare Variables
    Retval = IIf(bVal, True, False)
      Select Case Retval
        Case True
          SaveSetting "Associations", "Notify", 1
        Case False
          SaveSetting "Associations", "Notify", 0
      End Select
End Sub
Sub SaveSetting_AutoLaunchExtEditor(bVal As Boolean)
  Dim Retval As Long 'declare Variables
    Retval = IIf(bVal, True, False)
      
      Select Case Retval
       Case True
          SaveSetting "ExternalEditor", "AutoLaunchExtEditor", 1
       Case False
          SaveSetting "ExternalEditor", "AutoLaunchExtEditor", 0
      End Select
End Sub
Sub SaveSetting_RememberLastWinPos(bVal As Boolean)
   Dim Retval As Long 'Declare Variables
     Retval = IIf(bVal, True, False)
     
      Select Case Retval
        Case True
         SaveSetting "LastWinPos", "Remember", 1
        Case False
         SaveSetting "LastWinPos", "Remember", 0
      End Select
      
    '*********************************************************************************************************
    ' We always want too record the current window state and position even if the user has Chosen No
    ' This way we keep a record and NextPad wont Give errors If NextPad was Abruptly Terminated before the
    ' Query_Unload Sub Can be met , Remember this is a New Feature it hasnt been fully Inmplemented
    If FrmMain.WindowState = vbMinimized Then
       SaveSetting "LastWinPos", "WindowState", vbNormal
       SaveSetting "LastWinPos", "Width", 8895
       SaveSetting "LastWinPos", "Height", 6510
       Exit Sub
    Else
       SaveSetting "LastWinPos", "WindowState", CInt(FrmMain.WindowState)
    End If
    SaveSetting "LastWinPos", "Left", FrmMain.Left
    SaveSetting "LastWinPos", "Top", FrmMain.Top
    SaveSetting "LastWinPos", "Width", FrmMain.Width
    SaveSetting "LastWinPos", "Height", FrmMain.Height
    '*********************************************************************************************************
End Sub
Function UseExternalEditor() As Boolean
    UseExternalEditor = CBool(GetSetting("NextPad", "ExternalEditor", "Use", 1))
End Function
Function Usewordwrap() As Boolean
    Usewordwrap = CBool(GetSetting("NextPad", "Wordwrap", "Wordwrap", 1))
End Function
Function Check_Associations_At_Startup() As Boolean
    Check_Associations_At_Startup = CBool(GetSetting("NextPad", "Associations", "Notify", 1))
End Function
Function AutoLaunchExtEditor() As Boolean
    AutoLaunchExtEditor = CBool(GetSetting("NextPad", "ExternalEditor", "AutoLaunchExtEditor", 1))
End Function
Function RememberLastWinPos() As Boolean
    RememberLastWinPos = CBool(GetSetting("NextPad", "LastWinPos", "Remember", 0))
End Function
Function Print_PfBold() As Boolean
    Print_PfBold = CBool(GetSetting("NextPad", "PrinterSettings", "PfBold", 0))
End Function
Function Print_PfItalic() As Boolean
    Print_PfItalic = CBool(GetSetting("NextPad", "PrinterSettings", "Pifont", 0))
End Function
Function Print_PfColor() As Boolean
    Print_PfColor = CBool(GetSetting("NextPad", "PrinterSettings", "PfColors", 0))
End Function
Function Print_Pfont() As Boolean
    Print_Pfont = CBool(GetSetting("NextPad", "PrinterSettings", "PfFont", 0))
End Function
Function Print_ColorMethod()
    Print_ColorMethod = GetSetting("NextPad", "PrinterSettings", "ColorMethod", 1)
End Function
Function Open_Method() As Integer
    Open_Method = GetSetting("NextPad", "OpenMethod", "OpenMethod", 0)
End Function
Function UseSmartFileOpening() As Boolean
    UseSmartFileOpening = CBool(GetSetting("NextPad", "OpenMethod", "UseSmartFileOpening", True))
End Function
Function BrowseBar_Show() As Boolean
   BrowseBar_Show = CBool(GetSetting("NextPad", "BrowseBar", "Visible", 1))
End Function
Function QuickExit() As Boolean
   QuickExit = CBool(GetSetting("NextPad", "QuickExit", "QuickExit", 0))
End Function
Function AutoSave() As Boolean
   AutoSave = CBool(GetSetting("NextPad", "AutoSave", "AutoSave", 0))
End Function
Function RememberFindHistory() As Boolean
   RememberFindHistory = CBool(GetSetting("NextPad", "FindHistory", "Save", 1))
End Function
Function RememberReplaceHistory() As Boolean
   RememberReplaceHistory = CBool(GetSetting("NextPad", "ReplaceHistory", "Save", 1))
End Function
Function RemoveDeadRecentFiles() As Boolean
   RemoveDeadRecentFiles = CBool(GetSetting("NextPad", "RecentFiles", "RemoveDead", 1))
End Function
Function AllowRecentFiles() As Boolean
   AllowRecentFiles = CBool(GetSetting("NextPad", "RecentFiles", "Enable", 1))
End Function
Function AllowFavorites() As Boolean
   AllowFavorites = CBool(GetSetting("NextPad", "Favorites", "Enable", 1))
End Function
