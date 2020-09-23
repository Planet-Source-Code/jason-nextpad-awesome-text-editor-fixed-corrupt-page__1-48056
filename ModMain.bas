Attribute VB_Name = "Modmain"
Private Declare Function ShellExecuteEx Lib "shell32.dll" (sei As SHELLEXECUTEINFO) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function PathRemoveFileSpec Lib "shlwapi.dll" Alias "PathRemoveFileSpecA" (ByVal pszPath As String) As Long
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Declare Function GetFileTitle Lib "comdlg32.dll" Alias "GetFileTitleA" (ByVal lpszFile As String, ByVal lpszTitle As String, ByVal cbBuf As Integer) As Integer
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long
Private Declare Function IsCharLower Lib "user32" Alias "IsCharLowerA" (ByVal cChar As Byte) As Long
Private Declare Function IsCharUpper Lib "user32" Alias "IsCharUpperA" (ByVal cChar As Byte) As Long
Private Declare Function StrFormatByteSize Lib "shlwapi" Alias "StrFormatByteSizeA" (ByVal dw As Long, ByVal pszBuf As String, ByRef cchBuf As Long) As String
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long

Private Type SHELLEXECUTEINFO ' Type Of SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hWnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Enum FO_CONSTANTS
  ' Constants For File Operations
  FO_COPY = &H2 'Copy Operation
  FO_DELETE = &H3 'Delete Operation
  FO_MOVE = &H1 ' Move Operation
  FO_RENAME = &H4 ' Rename Operation
End Enum

Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings As Long
    lpszProgressTitle As String ' only used if FOF_SIMPLEPROGRESS
End Type


Const MAX_FILENAME_LEN = 260
' Constants for SetWindowPos
Public Const HWND_TOPMOST = -1 ' Keep the Window On top of all other Low Level Windows
Public Const HWND_NOTOPMOST = -2 ' Dont Keep the Window on top of all other low level windows
Public Const SWP_NOSIZE = &H1 ' The Window Retains its Size ignoring the Cx and y and x Parameters
Public Const SWP_NOMOVE = &H2 ' ''''''''''''''''''''''' Position on Screen '''''''''''''''''''''''
Public Const SWP_NOACTIVATE = &H10 ' The Window Cannot Be Activated
Public Const SWP_SHOWWINDOW = &H40 ' Show the Window

Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_CLIENTEDGE = &H200
Public Const WS_EX_STATICEDGE = &H20000
Public Const WS_EX_TRANSPARENT = &H20&

Public Const SW_SHOWNORMAL = 1
Public Const EM_UNDO = &HC7
Public Const WM_USER = &H400
Public Const EM_CANUNDO = &HC6 'Constant that evaluates if You Can Undo
Public Const WM_PASTE = &H302 ' Message Too Paste
Public Const WM_CUT = &H300 'Message Too Cut
Public Const WM_COPY = &H301 'Message Too Copy
Public Const EM_REPLACESEL = &HC2 'Message To replace selection or insert at selection point

Public Const Normal_Cdlogflags = cdlOFNHideReadOnly + cdlOFNFileMustExist + cdlOFNLongNames

Type filestring ' The heart of NextPads BOOLEAN Memory
dirty As Integer ' Without This Then NextPad wouldnt
End Type ' Know if a file was changed or not
Public Fstate As filestring

Public Const FOF_ALLOWUNDO = &H40 'Allow The Undo Of An Operation (FO)
Public Const FOF_WANTNUKEWARNING = &H4000 'Allow Warning With Dialog

' Constants for shell execute info
Public Const SEE_MASK_INVOKEIDLIST = &HC
Public Const SEE_MASK_NOCLOSEPROCESS = &H40
Public Const SEE_MASK_FLAG_NO_UI = &H400

Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOOWNERZORDER = &H200
Public Const SWP_NOZORDER = &H4

Public StrFind As String

Public gFindString As String
Public gFindCase As Integer
Public gFindDirection As Integer
Public gCurPos As Integer
Public gFirstTime As Integer

Public ReplaceStrings As New Collection
Public FindStrings As New Collection

Public sDir As String


'****************************************************************************
Function FindStr() As Boolean
       
   On Error GoTo FindSelErr:
    
    Dim intStart As Integer 'declare variables
    Dim intPos As Integer 'declare variables
    Dim strFindString As String 'declare variables
    Dim strSourceString As String 'declare variables
    Dim StrMsg As String 'declare variables
    Dim intResponse As Integer 'declare variables
    Dim intOffset As Integer 'declare variables
    
    FrmMain.mTextbox.SetFocus
    
    If (gCurPos = FrmMain.mTextbox.SelStart) Then
        intOffset = 1
    Else
        intOffset = 0
    End If

    If gFirstTime Then intOffset = 0
    intStart = FrmMain.mTextbox.SelStart + intOffset
        
    If gFindCase Then
        strFindString = gFindString
        strSourceString = FrmMain.mTextbox.Text
    Else
        strFindString = UCase(gFindString)
        strSourceString = UCase(FrmMain.mTextbox.Text)
    End If
            
    If gFindDirection = 1 Then
        intPos = InStr(intStart + 1, strSourceString, strFindString)
    Else
        For intPos = intStart - 1 To 0 Step -1
            If intPos = 0 Then Exit For
            If Mid(strSourceString, intPos, Len(strFindString)) = strFindString Then Exit For
        Next
    End If

    If intPos Then
        FrmMain.mTextbox.SelStart = intPos - 1
        FrmMain.mTextbox.SelLength = Len(strFindString)
        FindStr = True ' We found it.
    Else
        StrMsg = "Cannot find " & Chr(34) & gFindString & Chr(34)
        intResponse = MsgBox(StrMsg, vbInformation, "NextPad")
        FindStr = False ' We could not find it.
    End If
    
    gCurPos = FrmMain.mTextbox.SelStart
    gFirstTime = False
    
FindSelErr:
      
      If Err.Number <> 0 Then
        StrMsg = "Cannot find " & Chr(34) & gFindString & Chr(34)
        intResponse = MsgBox(StrMsg, vbInformation, "NextPad")
        Exit Function
      End If
  End Function
Sub ResizeNoteWithToolbar()
 '**********************************************************************
 ' This Sub is Also Pretty Important it Handles Resizing of the Form   *
 ' Although the code hasnt been updated or Attempted too been Improved *
 ' But it Still Works Fairly Well :)                                   *
 ' Sorry that i didnt Comment This Sub Very Well But,                  *
 ' i Didnt Have Time Too Comment This Sub For 10-15 minutes at a time  *
 ' Sorry :(                                                            *
 '**********************************************************************
On Error GoTo Errresize:
   If FrmMain.WindowState = vbMinimized Then Exit Sub 'We Cant Resize A Form
       ' While it is Minimized
    ' if the toolbar is visible and the 2nd text box is also visible....
    If FrmMain.Toolbar.Visible And FrmMain.txt(1).Visible = True Then
        FrmMain.txt(1).Height = FrmMain.ScaleHeight - FrmMain.Toolbar.Height
        FrmMain.txt(1).Top = FrmMain.Toolbar.Height
    If FrmMain.PicBrowseBar.Visible = True Then
        FrmMain.txt(1).Left = FrmMain.PicBrowseBar.Width
        FrmMain.txt(1).Width = FrmMain.ScaleWidth - FrmMain.PicBrowseBar.Width
        FrmMain.FileBrowseBar.Height = FrmMain.Height - FrmMain.tvwFolders.Height - 1550
    Else
        FrmMain.txt(1).Width = FrmMain.ScaleWidth
        FrmMain.txt(1).Left = 0
    End If
    
    Else
        If FrmMain.Toolbar.Visible And FrmMain.txt(0).Visible = True Then
        FrmMain.txt(0).Height = FrmMain.ScaleHeight - FrmMain.Toolbar.Height
        FrmMain.txt(0).Top = FrmMain.Toolbar.Height
    If FrmMain.PicBrowseBar.Visible = True Then
        FrmMain.txt(0).Left = FrmMain.PicBrowseBar.Width
        FrmMain.txt(0).Width = FrmMain.ScaleWidth - FrmMain.PicBrowseBar.Width
        FrmMain.FileBrowseBar.Height = FrmMain.Height - FrmMain.tvwFolders.Height - 1550
    Else
        FrmMain.txt(0).Width = FrmMain.ScaleWidth
        FrmMain.txt(0).Left = 0
    End If

    Else
        If FrmMain.txt(1).Visible = True Then
        FrmMain.txt(1).Height = FrmMain.ScaleHeight
        FrmMain.txt(1).Top = 0
    If FrmMain.PicBrowseBar.Visible = True Then
        FrmMain.txt(1).Left = FrmMain.PicBrowseBar.Width
        FrmMain.txt(1).Width = FrmMain.ScaleWidth - FrmMain.PicBrowseBar.Width
        FrmMain.FileBrowseBar.Height = FrmMain.Height - FrmMain.tvwFolders.Height - 1400
    Else
        FrmMain.txt(1).Width = FrmMain.ScaleWidth
        FrmMain.txt(1).Left = 0
    End If

        Else
        If FrmMain.txt(0).Visible = True Then
        FrmMain.txt(0).Height = FrmMain.ScaleHeight
        FrmMain.txt(0).Top = 0
    If FrmMain.PicBrowseBar.Visible = True Then
        FrmMain.txt(0).Left = FrmMain.PicBrowseBar.Width
        FrmMain.txt(0).Width = FrmMain.ScaleWidth - FrmMain.PicBrowseBar.Width
        FrmMain.FileBrowseBar.Height = FrmMain.Height - FrmMain.tvwFolders.Height - 1400
    Else
        FrmMain.txt(0).Width = FrmMain.ScaleWidth
        FrmMain.txt(0).Left = 0
    End If

       End If
      End If
    End If

  
Errresize:
    If Err.Number <> 0 Then
       Exit Sub
    End If
   End If
End Sub
Sub IsAssociated()
' If we are supposed to check associations then...
If Check_Associations_At_Startup Then
    ' Examine information in the registry.
    Select Case GetSettingString(HKEY_CLASSES_ROOT, "Txtfile\shell\open\command", "", "")
      Case App.Path & "\" & App.EXEName & ".EXE" & " %1"
             Exit Sub ' Were associated...
      Case App.Path & App.EXEName & ".EXE" & " %1"
             Exit Sub ' Were associated...
      Case Else ' We are not associated...
       ' If the user does not want to be notified then exit...
       If GetSetting("NextPad", "Associations", "Notify", 1) = 0 Then Exit Sub
             ' The user wants to be, display the associations dialog.
             FrmNotifyAss.Show , FrmMain
             Exit Sub
     End Select
Else ' We are not supposed to check...
   Exit Sub
End If

End Sub
Sub FileError(sFileName As String, mTextbox As TextBox)
'************************************************************************
' This function now handles file errors directly instead of indirectly  *
' if it is found that the file is read only the user now has the choice *
' whether or not to change that attribute and save the file.            *
'************************************************************************

 Dim FileAttr As Integer, Response
 ' Check what the file attributes are on this file
 FileAttr = GetAttr(sFileName) And vbReadOnly
   If FileAttr = 1 Then ' File is read only it cannot be saved
       ' Show the message
       Response = MsgBox(sOpenFileName + " is read-only which means you cannot save any changes you have made to it." & _
                  vbNewLine + vbNewLine + "Would you like to remove its read-only attributes so you can save it?", vbYesNo + vbExclamation, "NextPad")
       Select Case Response ' Check which buttons were pressed
            Case vbYes ' Yes was pressed
                SetAttr sFileName, vbNormal ' Set attributes to normal
                SaveFile sFileName, mTextbox, False ' Now we can hopefully save the file
                Exit Sub ' We are now done
            Case Else ' Another button was pressed.
                Exit Sub
       End Select
   End If
     ' Notify user that this file cannot be saved
     MsgBox GetFileTitleStr(sFileName) + " Cannot be saved." _
     + vbNewLine + Err.Description, vbCritical, "NextPad"
     Exit Sub ' We are now done
   
End Sub
Sub Main() ' main startup  Procedures and Misc. for The Project
'*****************************************************************************
' This Procedure is the main  Loading Of NextPad , When NextPad First Starts *
' This Gets Loaded first so we can decide or Change what wed like            *
'*****************************************************************************

   Dim OnOffWordWrap As Boolean, Bwidth As Long

   Dim X As Long, CmdFile As String
   
    On Error Resume Next
    X = InitCommonControls ' Initialize common controls
    On Error Resume Next
    If GetSetting("NextPad", "ExternalEditor", "Path", "") = "" Then: SaveSetting "ExternalEditor", "Path", DetectExternalEditor
    ' togglewordwrap from usewordwraps value from modoptions BOOLEAN
    ToggleWordWrap (Usewordwrap)
    'Declare Public Variables Boolean
    Fstate.dirty = False
  '----------------------------------------------------------------------
  ' Saved Window State Settings - Loading and Retrieving
  On Error GoTo ErrorHandlerWinPos:
   If RememberLastWinPos = True Then ' check if the user wanted too save & Load the Last Window POSition
   ' Check if the Form Was Minimized when its Position was saved
   ' Why ? Because  a form cant be Resized or Moved while it is Minimized !!!
   If GetSetting("NextPad", "LastWinPos", "WindowState", 0) = vbMinimized Then
      ' Set the forms WindowState Too the Previous One in the Registry
      SaveSetting "LastWinPos", "WindowState", vbNormal
      ' Show the form
      FrmMain.WindowState = vbNormal
      ' Center it now
      FrmMain.Move (Screen.Width - FrmMain.Width) / 2, (Screen.Height - FrmMain.Height) / 2
      ' Save Settings
      SaveSetting_RememberLastWinPos True
   Else
       ' If one of the values in the registry is invalid then center the screen and save the new settings
       On Error Resume Next ': FrmMain.Move (Screen.Width - FrmMain.Width) / 2, (Screen.Height - FrmMain.Height) / 2:      SaveSetting_RememberLastWinPos True
      FrmMain.Top = CSng(GetSetting("NextPad", "LastWinPos", "Top", 1335)) ' Convert the ones in the Registry too single
      FrmMain.Left = CSng(GetSetting("NextPad", "LastWinPos", "Left", 1710)) ' Convert the ones in the Reigstry too single
      FrmMain.Width = CSng(GetSetting("NextPad", "LastWinPos", "Width", 8895))
      FrmMain.Height = CSng(GetSetting("NextPad", "LastWinPos", "Height", 6510))
      FrmMain.WindowState = CInt(GetSetting("NextPad", "LastWinPos", "WindowState", vbNormal))
   End If
 
   If RememberLastWinPos = False Then ' If the User Doesent Want the Form Too remember its window Postition
      ' Center the Form on the Users Screen
      FrmMain.Move (Screen.Width - FrmMain.Width) / 2, (Screen.Height - FrmMain.Height) / 2
   End If
   End If
  ' BrowseBar
  '----------------------------------------------------------------------
          ' Resize the browseBar no matter what the result is (Visible or Not)
           With FrmMain
             Bwidth = GetSetting("NextPad", "BrowseBar", "Width", 2850)
             .PicBrowseBar.Width = Bwidth ' Set width
             .LblSelect(0).Width = Bwidth - 50
             .LblSelect(1).Width = Bwidth - 50
             .LblClose.Move Bwidth - 150
             .tvwFolders.Width = Bwidth - 50 ' Set width
             .FileBrowseBar.Width = Bwidth - 50 ' Set width
              ' Set pattern
             .FileBrowseBar.Pattern = GetSetting("NextPad", "BrowseBar", "Pattern", "*.TXT;*.INI")
              ' Set attributes that cvan be shown
             .FileBrowseBar.Normal = GetSetting("NextPad", "BrowseBar", "Normal", True)
             .FileBrowseBar.ReadOnly = GetSetting("NextPad", "BrowseBar", "ReadOnly", False)
             .FileBrowseBar.Archive = GetSetting("NextPad", "BrowseBar", "Archive", False)
             .FileBrowseBar.Hidden = GetSetting("NextPad", "BrowseBar", "Hidden", False)
             .FileBrowseBar.System = GetSetting("NextPad", "BrowseBar", "System", False)
             'Resize the form to match the current state
          End With
     If BrowseBar_Show = True Then
        ResizeNoteWithToolbar 'Resize the form to match the current state
        FrmMain.PicBrowseBar.Visible = True ' show The BrowseBar
        FrmMain.MnuBrowseBar.Checked = True 'Check the Menu to match the current state
     End If
  '----------------------------------------------------------------------
       
     
  ' Toolbar
  '----------------------------------------------------------------------
     With FrmMain
       Dim tVal As Boolean
        tVal = CBool(GetSetting("NextPad", "ToolBar", "Visible", 1))
        .Toolbar.Visible = tVal
        .MnuToolbar.Checked = tVal
        ResizeNoteWithToolbar
     End With
  '----------------------------------------------------------------------
        
  On Error GoTo ErrorHandlerFrmMain:
        FrmMain.Show '  finally show the form
        
  ' Load all this stuff after we show the form so NextPad does not load up slowly.
  On Error Resume Next
        GetRecentFiles ' Call the GetRecentFiles Procedure
        GetFavorites   ' Call the GetFavorites Procedure
        IsAssociated   ' Check if NextPad is associated...
    
  On Error GoTo ErrorHandlerCommandLine:
        If Command$ <> "" Then '  Check if there are any command line args.
           ' If command line args are "quoted"...
           If Left(Command$, 1) = Chr(34) And Right(Command$, 1) = Chr(34) Then
              ' Remove quotes
              CmdFile = Mid$(Command$, 2, Len(Command$) - 2)
           Else ' It is not quoted...
              CmdFile = Command$
           End If
              ' Open the file...
              OpenFile GetLongFileName(CmdFile), True
              Exit Sub
        End If

    
ErrorHandlerWinPos:
     If Err.Number <> 0 Then
        MsgBox "Partial information in the registry is corrupt. Click OK to repair.", vbCritical, "Registry - NextPad"
        DeleteSetting "NextPad", "LastWinPos"
        Resume Next
     End If
ErrorHandlerCommandLine:
     If Err.Number <> 0 Then
        MsgBox Command$ & " Could not be opened, " & Err.Description, vbCritical, "Command Line - NextPad"
        Resume Next
     End If
ErrorHandlerFrmMain:
     If Err.Number <> 0 Then
       MsgBox "An error has occured while NextPad was starting up." & _
       vbNewLine & "NextPad cannot start as a result of this." & _
       vbNewLine & "Please contact the author by emailing him at :" & _
       vbNewLine & "Cyberarea@hotmail.com" & _
       vbNewLine & "We are sorry for any inconvenience this may have caused you." & _
       vbNewLine & vbNewLine & "Reason : " & Err.Description & _
       vbNewLine & "Error Number : " & Err.Number, vbCritical, "NextPad"
       Exit Sub
       End
     End If
End Sub
Sub OpenFile(sFileName As String, Optional AddTooRecentFiles As Boolean = False)
'********************************************************************************************
'* This Sub opens the File Requested and Performs the Required Operation to Open the File   *
'* it looks Pointless at some Points, but it Works :)                                       *
'********************************************************************************************
On Error GoTo FileNotFound:
  
  Wrap$ = Chr$(13) + Chr$(10)  'create wrap character
  
  If bFileExists(sFileName) = False Then GoTo FileNotFound:
  If sFileName <> "" Then
     
     If FileLen(sFileName) > 65000 Then FrmMain.Hide: Query_TooBig (GetShortPath(sFileName)): Close #1: ClearNextPad: Exit Sub
     
     Close #1 'Close the file

      On Error GoTo outofmemory: 'If An Error Occurs goto That line
         'FrmMain.Show  'Show the form
           FrmMain.mTextbox.Text = "" 'Remove any Text Currently there
           FrmMain.Caption = GetFileTitleStr(sFileName) & " - NextPad"  'set the forms caption
       Select Case Open_Method 'check the Open_Method
        Case 1 'Binary Open_Method
           Close #1 'Close the file
           Open sFileName For Input As #1 'Open the file
             Do Until EOF(1)          'then read lines from file
               Line Input #1, LineOfText$
               AllText$ = AllText$ & LineOfText$ & Wrap$
             Loop
           FrmMain.mTextbox.Text = AllText$
            Close #1 'Close the file
        Case 0 'Input Open_Method
           Close #1 'Close the file
           Open sFileName For Input As #1
           FrmMain.mTextbox.Text = Input(LOF(1), 1) 'Open the file
           Close #1 'Close the file
        Case Else 'Else The Value Is Different
           Close #1 'Close the file
           Open sFileName For Input As #1 'Open the file
             Do Until EOF(1)          'then read lines from file
               Line Input #1, LineOfText$
               AllText$ = AllText$ & LineOfText$ & Wrap$
             Loop
           FrmMain.mTextbox.Text = AllText$
            Close #1 'Close the file
     End Select
         
           Fstate.dirty = False 'Set the fstate too not dirty false
        
        If AddTooRecentFiles = True Then 'if we want too addtorecentfiles then
              AddRecentFile (sFileName) 'Add too the recent files menu
        End If
       
          sOpenFileName = sFileName  'set the current file name
          Fstate.dirty = False 'The fstate is NOT dirty Now
          Close #1 'Close the file
          Exit Sub 'Exit the Sub Before executing anymore code
      
  End If

     
outofmemory:      ' error That occurs when NextPad runs out of memory
 If Err.Number <> 0 Then
   
   
 If UseSmartFileOpening = True And FileLen(sFileName) < 65000 And Open_Method = 0 Then
    '********
    'Please Note This May Look Totally Unecessary But it helps
    '(Not The Comment), The Code May Look Dumb But Thats how it
    'Opens files Now
    '********
    'MsgBox "Using Smart©" 'debugging Purposes ONLY
    Close #1 'close the file that is currently open
    SaveSetting "OpenMethod", "OpenMethod", 1 'set the open-Method Back To Binary
    OpenFile sFileName, AddTooRecentFiles 'Open the file Again
    SaveSetting "OpenMethod", "OpenMethod", 0 'set the Open_Method Back Too Input
    Exit Sub 'Exit The Sub Before Running Any New Code
 End If
 
 If Not FileLen(sFileName) < 65000 Then
   FrmMain.Hide 'Hide the Form
   Query_TooBig (GetShortPath(sFileName)) 'Tell the user about the file being too large
   Close #1 'close the file
   ClearNextPad ' Clear NextPad
   Fstate.dirty = False ' set the filestate too Not Dirty FALSE (0)
 End If
 
 If FileLen(sFileName) < 65000 And UseSmartFileOpening = False And OpenMethod = 0 Then
  Dim Response ' Variable used to hold MsgBox() return value
   
   ' Display Message to user
   Response = MsgBox("Smart File Opening is currently disabled." & _
          vbNewLine & "NextPad has detected that this file may not " & _
          "be displayed correctly or at all as a result." & _
          vbNewLine & "In order to have this file displayed correctly you must " & _
          "enable Smart File Opening." & _
          vbNewLine & vbNewLine & "Would you like to enable it now?", vbYesNo + vbCritical, "NextPad")
   
   Select Case Response ' Examine variable
       Case vbYes ' User selected Yes
          ' Enable Smart File Opening
          SaveSetting "OpenMethod", "UseSmartFileOpening", 1
          ' Reset OpenMethod to 0
          SaveSetting "OpenMethod", "OpenMethod", 0
          ' Reopen the file with changes made
          OpenFile sFileName, AddTooRecentFiles
          ' We are now finished exit this sub
          Exit Sub
       Case vbNo ' User selected No
          ClearNextPad ' Make sure we clear everything
          ' We are now finished exit this sub
          Exit Sub
   End Select
 End If
 End If
FileNotFound:
  If Err.Number = 53 Or Err.Number = 75 Or Err.Number = 76 Then  'If The Error Is 'File Not Found'
     NotifyFileNonExistent (sFileName)
  '------------------------------------------------------------------------------------------------------------------------------
  ' The following code displays errors to the user when we can just ignore them because
  ' we know that switching from certain files to another can cause errors...
  'Else
     'MsgBox GetFileTitleStr(sFileName) + " Cannot be opened." + vbNewLine + Err.Description, vbCritical, "Open File - NextPad"
     'Exit Sub
  '------------------------------------------------------------------------------------------------------------------------------
  End If
        
 End Sub
Public Property Get sOpenFileName() As String
   sOpenFileName = FrmMain.lblfilename.Caption
End Property
Public Property Let sOpenFileName(s As String)
   FrmMain.lblfilename.Caption = s
End Property


Function ShellNewNextPad(Thewindowstyle As VbAppWinStyle, Optional CommandLineParams As String = "") As Long
'****************************************************************************
' This function Shells A New NextPad By using the VbAppWinStyle Constants   *
' It Works Rather Well And Can Shell A New NextPad Very Easily With Command *
' Line Params                                                               *
'****************************************************************************
On Error GoTo FileNotFound:

Dim StrApp As String 'Variable too hold the Apps Path And FileName
  
  StrApp = App.Path & "\" & App.EXEName & ".EXE" 'Set the Strapp Variable too the Apps Path and ExeName
  ' if The App is In The root Directory NextPad Will Detect This By Seeing The Extra Slash and Fixing the variable
  If bFileExists(StrApp) = False Then: StrApp = App.Path & App.EXEName & ".EXE"
  ' Shell the New App Including any Command Line Params Nad the TheWinStyle
  Shell StrApp & Space(1) & CommandLineParams, Thewindowstyle

FileNotFound:
   If Err.Number <> 0 Then
     ' Notify User
     MsgBox "NextPad Cannot find its own executable file.", vbCritical, "New Instance - NextPad"
     Exit Function 'exit the Function
   End If

End Function
Sub TextChangecontrol()
'*****************************************************
' This Procedure was Put Here Just too handle the    *
' Seperate Controls FileState Flag , Unfortunately   *
' We Dont Need It Anymore But i kept it Anyway       *
'*****************************************************
   On Error GoTo outofmemoryerror:
     
    ' If the FileState Flag is Not Dirty then Make it Dirty
    If Fstate.dirty = False Then
       Fstate.dirty = True 'Set the Fstate Flag too Dirty
    End If

outofmemoryerror:
    If Err.Number <> 0 Then
       MsgBox "NextPad Has Encountered The Following Error(s) While Performing The operation you requested : " _
       & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "If The Error Is " & Chr(34) & "Out of memory" & Chr(34) & " Then NextPad Cannot Place Anymore Text Into The Text Box Because It Has Run Out of memory", vbCritical, "NextPad "
       Exit Sub
    End If
End Sub
Sub SaveSetting(Section As String, Key As Variant _
, Setting As Variant)
VBA.SaveSetting "NextPad", Section, Key, Setting
End Sub
Sub ExecuteExternalEditor(Currentfilename As String)
'**********************************************************************
' This Function Executes the External Editor with a Filename Provided *
' It works Well :)                                                    *
'**********************************************************************
  Dim ThePathName As String 'variable too hold the pathname of the External Editor
  Dim Success As Long 'Success Variable too hold the Shell Fucntion
  Dim Response 'Variable too hold the MsgBox() Function
 
  On Error GoTo CannotFindExternalEditor:
   
    If UseExternalEditor = False Then
       ' Notify the user about the problem.
       MsgBox sOpenFileName & " " & "Cannot be opened because it is to large for NextPad to open." & vbNewLine & _
       "Please enable the External Editor in the options window. If an external editor cannot be found you will not be able to view this file.", vbCritical, "NextPad"
       ' Exit this sub before we make another mistake.
       Exit Sub
    End If
       
       
   ' Get the path Name of the external editor
   ThePathName = GetSetting("NextPad", "ExternalEditor", "Path", "")
   ' Shell the External Editor With Any Following Command Line Params
   If bFileExists(ThePathName) = False Then: Err.Raise 456798, "NextPad", "Cannot find external editor": GoTo CannotFindExternalEditor:
 
  
     
    ' Shell the External Editor With Any Following Command Line Params
    Success = Shell(ThePathName & Space(1) & GetShortPath(Currentfilename), vbNormalFocus)
    Exit Sub

CannotFindExternalEditor:


 If Err.Number <> 0 Then
   ' Notify User
   Response = MsgBox("NextPad cannot locate the external editor that is used to open files that are to large to open." & vbNewLine & vbNewLine & "Would you like NextPad to try and find it for you?", vbYesNo + vbCritical, "NextPad")
     Select Case Response 'Decide which button the user pressed
        Case vbYes 'the user pressed the yes Button
           ' If NextPad still cannot detect an external editor then we must notify the user.
           If bFileExists(DetectExternalEditor) = False Then
               ' Notify the user about the unfortunate event.
               MsgBox "NextPad is unable to find a program on your computer that can open files that are extremely large." & _
               vbNewLine & "Because of the reason stated above you cannot view files that are extemely large with NextPad." & _
               vbNewLine & "NextPad apologizes for the inconvenience.", vbCritical, "NextPad"
               ' Disable the external editor because none can be found.
               SaveSetting_UseExternalEditor (False)
               ' Exit this sub immediatelly
               Exit Sub
            End If
           'Save the External Editors Path
           SaveSetting "ExternalEditor", "Path", DetectExternalEditor
           ' Notify of Successful Save
           MsgBox "NextPad has successfully found the external editor. Please click OK to open the file you have requested.", vbInformation, "NextPad"
           ExecuteExternalEditor (Currentfilename)
           'ShellNewNextPad vbNormalFocus ' Shell a New NextPad
           End ' Close this instance
        Case vbNo 'User Pressed the NO Button
           'End 'End Execution Stop All Code And Quit Evrything
     End Select
    'Exit Sub
  End If

End Sub
Function DetectExternalEditor() As String
' *********************************************************
' This function was rewritten because of registry access  *
' problems with Windows XP but now it works better        *
' *********************************************************
Dim sFile As String

  ' Name this file to anything
  sFile = App.Path & "\" & "NP02143425642334TMP.RTF"
  'Open the file for saving
  Open sFile For Output As #1
  Print #1, 0 'Print anything to the file
  If bFileExists(sFile) = False Then sFile = App.Path & "NP02143425642334TMP.RTF"
  Close #1
  'Open the file for saving
  Open sFile For Output As #1
  Print #1, 0 'Print anything to the file
  Close #1
 If FindAssociatedExe(sFile) <> "" Then
    SaveSetting "ExternalEditor", "Path", FindAssociatedExe(sFile)
    DetectExternalEditor = FindAssociatedExe(sFile)
    Kill sFile
    Exit Function
 Else
    DetectExternalEditor = ""
    Kill sFile
    Exit Function
 End If
 
End Function
Sub Query_TooBig(sFileName As String)
  '**********************************************************************************
  ' This Procedure Asks the User too lauch the External editor Because the File is  *
  ' Too large too open , or it will just tell the user that the file is too large   *
  ' Too open , if the external editor has chosen too be used there is also an option*
  ' Just Too Launch the External ediotr , Not Prompt too ........                   *
  '**********************************************************************************
    Dim regval As String 'String Variable Too Hold the value form the registry
     Dim Msg, Response ' Response Variable to hold the MsgBox () Fucntion,Msg Too Hold the String in the Message Box

  'set regval variables
    regval = GetSetting("NextPad", "ExternalEditor", "Path", "")

    If UseExternalEditor = True And regval = "" Then
       regval = DetectExternalEditor
    End If

      Select Case UseExternalEditor 'Check if the external editor is going to be used
        Case True ' Case is True *******
           GoTo Query: ' it is so Jump too the Query Line
           Exit Sub 'exit the sub before executing any new code ,
        Case False ' Case is False ******
           FrmMain.Show 'show the form
           ' notify user that the file is too large too be opened
           MsgBox GetFileTitleStr(GetLongFileName(sFileName)) _
           & vbNewLine & "Is to large For NextPad to open." _
           & vbNewLine & vbNewLine & "Be sure to have " & vbNewLine & Chr$(34) & "Use external editor to open files to large for NextPad to open." & Chr$(34) & _
           vbNewLine & " Enabled in the options window.", vbExclamation, "NextPad"
           ' set the caption too Ntohing ("") 0
           ClearNextPad
           Fstate.dirty = False
           ' Exit the sub Immediately
           Exit Sub
        End Select
        
'**************************************************************
'Nothing Suspicous Was Detected ( If Code Gets Below This Line )
'**************************************************************
Query:
  ' check if user Wants too just Launch the External Editor Instead Of Notifying the user ....
  If AutoLaunchExtEditor = True Then: ExecuteExternalEditor (sFileName): End: Exit Sub
  
  ' the user wants too be notified , set up msg variable and display notification
  ' too user
   Msg = "This file is to large for NextPad to open." _
   & vbNewLine & vbNewLine & "Would you like " & StrConv(GetFileTitleStr(GetSetting("NextPad", "ExternalEditor", "Path", "")), vbProperCase) & " to open it ?"
    Response = MsgBox(Msg, vbDefaultButton1 + vbQuestion + vbYesNo, "NextPad")

      Select Case Response 'Check the Response Variable
        Case vbYes 'User Pressed The Yes Button
           ExecuteExternalEditor (sFileName) 'Execute the External Editor With The FileName Specified
           End 'Stop Execution of code and Close ALL of NextPad's Windows Then Quit
        Case vbNo 'User pressed the Cancel Button
           FrmMain.Show 'Show the Form
           FrmMain.Caption = "Untitled - NextPad" ' Reset the Forms Caption too fit The Current State
           Exit Sub 'Exit the sub Before Executing Any New Code
      End Select
End Sub
Sub ToggleWordWrap(Optional ONOrOFF As Boolean)
'****************************************************************************
'* This code is Ugly at Best ,But it gets the job done with a mess          *
'*  of Code too do it, unfortunately attempts too improve this have failed  *
'* Well you know what they say , Dont fix it if its not Broken ,            *
'* And thats exactly what im going too do until i figure out an easier way  *
'****************************************************************************
On Error GoTo Err:
Dim fontname, fontsize, FBold, Fitalic As Integer
Dim FontColor, fBackColor As Long
'onoff = IIf(us, True, False)
 fontname = GetSetting("NextPad", "Font", "font")
 fontsize = GetSetting("NextPad", "Font", "Fontsize", 8)
 fBackColor = GetSetting("NextPad", "Font", "BackColor", vbWhite)
 FontColor = GetSetting("NextPad", "Font", "ForeColor", vbBlack)
 FBold = GetSetting("NextPad", "Font", "FontBold", 0)
 Fitalic = GetSetting("NextPad", "Font", "FontItalic", 0)
 
Select Case ONOrOFF
Case False
ResizeNoteWithToolbar
FrmMain.txt(0).Visible = False
ResizeNoteWithToolbar

FrmMain.txt(1).fontname = FrmMain.txt(0).fontname
FrmMain.txt(1).fontsize = FrmMain.txt(0).fontsize
FrmMain.txt(1).BackColor = FrmMain.txt(0).BackColor
FrmMain.txt(1).ForeColor = FrmMain.txt(0).ForeColor
FrmMain.txt(1).FontBold = FrmMain.txt(0).FontBold
FrmMain.txt(1).FontItalic = FrmMain.txt(0).FontItalic
FrmMain.txt(1).fontname = fontname
FrmMain.txt(1).fontsize = fontsize
FrmMain.txt(1).BackColor = fBackColor
FrmMain.txt(1).ForeColor = FontColor
FrmMain.txt(1).FontBold = FBold
FrmMain.txt(1).FontItalic = Fitalic

FrmMain.txt(1).Visible = True

If Fstate.dirty = True Then
Fstate.dirty = True
Else
If Fstate.dirty = False Then
Fstate.dirty = False
End If
End If

ResizeNoteWithToolbar
FrmMain.txt(1).Text = FrmMain.txt(0).Text
If Fstate.dirty = True Then
Fstate.dirty = True
Else
If Fstate.dirty = False Then
Fstate.dirty = False
End If
End If

'Check off Menu too match current state
FrmMain.MnuWordWrap.Checked = False


Case True

'resize form too match current state
ResizeNoteWithToolbar
FrmMain.txt(1).Visible = False
ResizeNoteWithToolbar
'set font name and size from registry
FrmMain.txt(0).fontname = FrmMain.txt(1).fontname
FrmMain.txt(0).fontsize = FrmMain.txt(1).fontsize
FrmMain.txt(0).BackColor = FrmMain.txt(1).BackColor
FrmMain.txt(0).ForeColor = FrmMain.txt(1).ForeColor
FrmMain.txt(0).FontBold = FrmMain.txt(1).FontBold
FrmMain.txt(0).FontItalic = FrmMain.txt(1).FontItalic

FrmMain.txt(0).Visible = True
FrmMain.txt(0).fontname = fontname
FrmMain.txt(0).fontsize = fontsize
FrmMain.txt(0).BackColor = fBackColor
FrmMain.txt(0).ForeColor = FontColor
FrmMain.txt(0).FontBold = FBold
FrmMain.txt(0).FontItalic = Fitalic

If Fstate.dirty = True Then
Fstate.dirty = True
Else
If Fstate.dirty = False Then
Fstate.dirty = False
End If
End If
ResizeNoteWithToolbar
FrmMain.txt(0).Text = FrmMain.txt(1).Text
If Fstate.dirty = True Then
Fstate.dirty = True
Else
If Fstate.dirty = False Then
Fstate.dirty = False
End If
End If
FrmMain.MnuWordWrap.Checked = True



'error Handler
Err:
Resume Next
End Select

End Sub
Public Function GetShortPath(StrFileName As String) As String
'***************************************************************************************
' This Code Grabs the Short Path of a File Name , Why Do we Need this ? Because When   *
' Launching the external editor , The Long Path gets Cut off Through the Code somewhere*
' and the External Editor Reports the File Is NON Existent                             *
'***************************************************************************************
    Dim lngRes As Long, strPath As String
    'Create a small buffer
    strPath = String$(165, 0)
    'retrieve the short pathname
    lngRes = GetShortPathName(StrFileName, strPath, 164)
    'remove all unnecessary chr$(0)'s
    GetShortPath = Left$(strPath, lngRes)
End Function

Public Function ShowFileProperties(OwnerHwnd As Long, Filename As String)
'***********************************************************************************
' This Function Displays the Windows File Properties Dialog by Retrieving the File *
' Name and then Passing it too The SEI Shell Exectute Info Structure , the Code is *
' Pretty Useful and it works !! :)                                                 *
'***********************************************************************************
 On Error GoTo ErrorHandler:
 
   Dim sei As SHELLEXECUTEINFO
   Dim r As Long
    With sei
         'Set the structure's size
         .cbSize = Len(sei)
         'Set the mask
         .fMask = SEE_MASK_NOCLOSEPROCESS Or _
         SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
         'Set the owner window
         .hWnd = OwnerHwnd
         'Show the properties
         .lpVerb = "properties"
         'Set the Filename
         .lpFile = Filename
         .lpParameters = vbNullChar
         .lpDirectory = vbNullChar
         .nShow = 0
         .hInstApp = 0
         .lpIDList = 0
    End With
   r = ShellExecuteEx(sei)

ErrorHandler:
   If Err.Number <> 0 Then
      MsgBox "Could not display the file properties dialog, " + Err.Description, vbCritical, "NextPad"
      Exit Function
   End If

End Function
Function Msga() As String
   '***************************************************************************************************************************
   ' IN a last Ditch effort too make the final executable smaller i have included this method in a function since it is used  *
   ' Twice in the same way so i created a functioin for it :)                                                                 *
   '***************************************************************************************************************************
    
 If sOpenFileName <> "" Then ' if the filename or anything in that label is equal to anything except "" then
    Msga = "The text in the " _
    & sOpenFileName & " file has changed." _
    & Chr(13) & Chr(13) & "Do you want to save the changes?"
 Else ' * If there isnt then Display The One Below
    Msga = "The text in the Untitled file has changed." _
    & Chr(13) & Chr(13) & "Do you want to save the changes?"
 End If

End Function

Function bFileExists(sFileName As String) As Boolean
 '****************************************************************************
 ' This Fucntion is Used too Check if a File Actually Exists                 *
 ' and Returns A BOOLEAN Value of Either True (-1) or Flase (0) if the File  *
 ' Doesent Exist, Fairly Simple but we need this Function for the Text File  *                                                           *
 ' Manager .                                                                 *
 '****************************************************************************
 ' Also This Fucntion is SO Important Because it Helps With a Lot Of Troubled
 ' Functions , For Example : If You Attempted too Launch a New Instance Of
 ' NextPad in the Boot Drive Eg ;' C:\ ' NextPad would Report an Error Occured ,
 ' Why ? Because When your shelling a File your blind, heres how we Used too
 ' Shell ' StrApp = App.path & "\" & App.exename ' But the Problem is that
 ' It Looks for the Applications Execuatable Thinking its in a Directory
 ' like ' C:\Program Files\NextPad.exe ' But if the App's Executable is in
 ' the Boot drive ' C:\ ' Then It Will Attempt too Shell it Like this
 '  '  C:\\NextPad.exe ' Which is incorrect Because of the Second Slash
 ' But Now That We Can Detect if A File Truly Exists or not We Use This Syntax Now
 '
 ' Dim StrApp As String
 '     StrApp = App.Path & "\" & App.exename
 '    If bFileExists(StrApp) = False then :StrApp = App.Path & Exename
 '   Shell (Strapp,VbnormalFocus)
 '
       
  ' Lets first check if we we just recieved a filename such as this one
  ' C:\\MYFILE
  ' If so let NextPad know that the file does NOT exist.
  If Left$(Mid(sFileName, 4), 1) = "\" Then bFileExists = False: Exit Function
    On Error GoTo FExistsError
    Dim F As String
    F = FreeFile
    Open sFileName For Input As #F 'Open file
    Close #F
FExistsError:
    If Err.Number = 53 Then 'If doesn't exists
        bFileExists = False 'Set FileExists to False
    ElseIf Err.Number = 0 Then 'else if exists
        bFileExists = True 'Set FileExists to True
    End If
 
 End Function
 

Sub NotifyFileNonExistent(sFileName As String)
'*******************************************************************************************************
' This Sub Was Created Because of A stupid Way of NextPad Opening Files That Didnt Exist , For Example *
' If you tried to open '9087(*&0-' Surely that wouldnt Exist Right ? But NextPad Thought That it Would *
' And Laucnhed the External Editor , Which is A Dumb Move Dont you Think ?............................ *
'*******************************************************************************************************
     MsgBox "The file you are attempting to open " & vbNewLine _
     & "'" & sFileName & "'" & " Does not exist." _
     & vbNewLine & "Please verify that the name and path given is correct. " _
     , vbCritical, "NextPad"
       
       
       ClearNextPad
       Fstate.dirty = False
       
End Sub

Sub AddRecentFile(sFileName As String)
'******************************************************
' This Function Adds a File too The Recent Files Menu *
' This Function is Simple But It Does Its Job :)      *
'******************************************************
    ' If the user disabled the recent files menu, then we cannot continue.
    If Not AllowRecentFiles Or sFileName = "" Then Exit Sub
    
    Dim iCount As Integer
    ' Get number of recent files in registry
    iCount = GetSetting("NextPad", "RecentFiles", "Count", 0)
    If GetSetting("NextPad", "RecentFiles", 0, "") <> "" Then iCount = iCount + 1
    
    With FrmMain.MnuRecentFiles
       ' If the default "Disabled" Menu is there use it
       Dim i As Integer ' Variable to hold menu count
       i = .Count ' Append value to integer
       ' If there is nothing there make sure that we show one
       If i = 0 Then .Item(0).Caption = sFileName: .Item(0).Visible = True: GoTo Save:
       i = i + 1 ' Since there is one there already add one
       Load .Item(i) ' Load the new menu
       .Item(i).Visible = True
       .Item(i).Caption = sFileName  ' Put a caption on it
    End With
    
Save:
    ' Save Favorite in registry
   If iCount = 0 Then
     SaveSetting "RecentFiles", 0, sFileName
     ' Update number of RecentFiles
     SaveSetting "RecentFiles", "Count", 0
   Else
     SaveSetting "RecentFiles", (iCount), sFileName
     ' Update number of RecentFiles
     SaveSetting "RecentFiles", "Count", (iCount)
   End If



End Sub

Sub GetRecentFiles()
'*********************************************************
' This Procedure Gets the Recent Files from the Registry *
' And Places there Filename in the Menu ,                *
' Enabeling the user too click them Opening them.        *
' Therefore This Function Works Wonders :)               *
'*********************************************************
   Dim i As Integer, X As Integer
   
    ' If the user disabled the Recent files menu, then we cannot continue.
    If Not AllowRecentFiles Then FrmMain.MnuRecentFilesMenu.Enabled = False:   Exit Sub
    
On Error GoTo ErrorHandler:
   
   i = GetSetting("NextPad", "RecentFiles", "Count", 0)
   If GetSetting("NextPad", "RecentFiles", 0, "") = "" Then Exit Sub ' If there are NO RecentFiles Stop
     
     With FrmMain.MnuRecentFiles
        .Item(0).Visible = True
        .Item(0).Caption = GetSetting("NextPad", "RecentFiles", 0)
        If i = 0 Then Exit Sub
        For X = 1 To i
          Load .Item(X) ' Load the new menu
          ' Give it a caption
          .Item(X).Caption = GetSetting("NextPad", "RecentFiles", X)
          ' Allow user to see it
          .Item(X).Visible = True
        Next X
     End With

ErrorHandler:
   If Err.Number <> 0 Then
      MsgBox "Could not get Recent Files from registry. If problem persist's please go to the Options window and reset NextPad's settings.", vbCritical, "NextPad - Recent Files"
      Exit Sub
   End If
End Sub

Sub PrintDocument(cTextBox As TextBox)
' ***************************************************************
' This Procedure is How We Print Documents It Is Fairly Simple  *
' There are several Boolean                                     *
' Options That Are Loaded When NextPad First Startsup , These   *
' Options Are Public and Are used througout the whole project   *
' So it makes accessing options in realtime and enables Certain *
' Speed Improvements                                            *
'****************************************************************
 Dim NumCopies, EndPage, BeginPage 'Declare Variables
 
     
      On Error Resume Next 'if An Error Occurs We dont want too exit the sub
      ' because then the rest of our code will not be executed
      
 Select Case Print_PfBold 'Check if we want to print out bold fonts
      Case True 'we do
         Printer.FontBold = cTextBox.FontBold 'Set the printer Object's FontBold Property too the TextBox's FontBOld Property
      Case False
         Printer.FontBold = False 'else dont print oput bold fonts
 End Select
       
         On Error Resume Next 'if An Error Occurs We dont want too exit the sub
      ' because then the rest of our code will not be executed

 Select Case Print_PfItalic 'check if we want too print out italic fonts
      Case True 'we do
         Printer.FontItalic = cTextBox.FontItalic 'set the printer objects fontitalic property too the textbox's
      Case False 'we dont
         Printer.FontItalic = False 'so dont print out italic fonts
 End Select
          
          On Error Resume Next 'if An Error Occurs We dont want too exit the sub
      ' because then the rest of our code will not be executed

 Select Case Print_ColorMethod 'check what color method wed like too use , Monochrome , Color
      Case 1 'Monochrome
         Printer.ColorMode = 1 'Set the printer object's Colormode Propertyr too Monochrome
      Case 2 'Color
         Printer.ColorMode = 2 'set the printer object's colormode property too Color
 Select Case Print_PfColor ' if the user wanted too print colors then decide whether the user wanted too print ForeColor's
      Case True 'The user Did
         Printer.ForeColor = cTextBox.ForeColor 'Set the Printer Objects Forecolor Property too the TextBox's forecolor
      Case False
         Printer.ForeColor = vbBlack ' Else set it too Default Black
 End Select
 End Select
             
             On Error Resume Next 'if An Error Occurs We dont want too exit the sub
      ' because then the rest of our code will not be executed

 Select Case Print_Pfont 'Check if the user wanted too print out fonts
      Case True 'The user Did
         Printer.Font = cTextBox.Font 'Set the printer objects font too the TextBox's Font Property
      Case False 'The user didnt
         Printer.Font = "MS Sans Serif" ' Set the font too the default
 End Select
  
  On Error GoTo CdlcCancelError: 'If an error Occurs VB WIll Jump too that line

 FrmMain.CommonDialog1.ShowPrinter ' Show The Print Dialog
  NumCopies = FrmMain.CommonDialog1.Copies ' Grab the Number of copies
   EndPage = FrmMain.CommonDialog1.ToPage ' Grab the EndPage
    BeginPage = FrmMain.CommonDialog1.FromPage ' Grab The Beginning Page

 
 For i = 1 To NumCopies ' GO From Page 1 to the Last Page
   Printer.Print cTextBox.Text ' Print the Text in the Control
   Printer.EndDoc ' End Printing on this Page
 Next
 
 Exit Sub

CdlcCancelError:
   If Err.Number = 32755 Then 'if cancel was Pressed (CldlcCancelError) then ......
     Exit Sub
   Else ' Else another error Has occured Other than the Cancel Error
   
   If Err.Number = 28663 Then Exit Sub ' No Default Printer exists
      
      MsgBox "Cannot print the document you requested, " & Err.Description, vbCritical, "NextPad - Print"
   End If
End Sub
Function FileOperation(fOperation As FO_CONSTANTS, sFileExisting, SFileTo)

'****************************************************************************************************
' This function performs the fle operation requested by using the value passed from  FO_CONSTANTS   *
'****************************************************************************************************
 
 On Error GoTo ErrorHandler:
 
 Dim r As Long
 Dim SHFO As SHFILEOPSTRUCT ' We Need This too Hold The Structure
 
  With SHFO
    .wFunc = fOperation ' Use enum constants
    .pFrom = sFileExisting ' Set The Existsing File
    .pTo = SFileTo ' Set Where We Are going Too Move The File
  End With
  
 r = SHFileOperation(SHFO)   'Perform The File Operation
ErrorHandler:
   If Err.Number <> 0 Then
      MsgBox "Could not perform the file operation requested '" + sFileExisting + "' To '" + SFileTo + "' Reason : " + Err.Description, vbCritical, "Error , NextPad API"
      Exit Function
   End If

End Function


Function DeleteFile(OwnerHwnd As Long, sFileName As String, Optional bSendToRecyclingBin As Boolean = True)
' ***********************************************************
' This Function Deletes Or Can Wipe A File From Disk ,      *
' Similare Too Move File But With Just Different Flags And  *
' A Different File Operation , Deleting                     *
' ***********************************************************
 On Error GoTo ErrorHandler:
 
  Dim NewFflags As Long  ' We Need This Too Store The Flags
  Dim r As Long 'Declare Variables
  Dim SHFO As SHFILEOPSTRUCT ' We Need This Too Hold The Structure
   
     Select Case bSendToRecyclingBin 'see if it is too be wiped from disk or sent too the recycling bin
        Case True 'It Is Too Be Sent too the Recycling bin
          NewFflags = FOF_WANTNUKEWARNING + FOF_ALLOWUNDO 'Send The File Too The Recycling Bin
        Case False 'It is To Be Wiped From Disk
          NewFflags = FOF_WANTNUKEWARNING 'Wipe The File From Disk
     End Select
   
    With SHFO
       .wFunc = FO_DELETE 'The Function We Are Doing is Deleting
       .pFrom = sFileName ' Set What Too Delete
       .fFlags = NewFflags 'Set The Flags that we made
       .hWnd = OwnerHwnd
    End With
  
  r = SHFileOperation(SHFO) 'Perform the file operation
  
ErrorHandler:
   If Err.Number <> 0 Then
      MsgBox sFileName + " Cannot be deleted." & vbNewLine & "The file may no longer exist or may have been moved.", vbCritical, "Error, NextPad API"
      Exit Function
   End If
End Function

Public Function GetFileTitleStr(sFileName As String) As String

' Buffer to hold Data
Dim sBuffer As String
 ' Fill buffer with Nulls
 sBuffer = String(255, Chr(0))
 ' Set Buffer
 GetFileTitle sFileName, sBuffer, Len(sBuffer)
 ' Remove Null Chars
 sBuffer = Left$(sBuffer, InStr(1, sBuffer, Chr$(0)) - 1)
 ' Return FileTitle
 If sBuffer <> "" Then 'If String has Data ....
    GetFileTitleStr = sBuffer
 Else 'Else Return the FileName given
    GetFileTitleStr = sFileName
 End If
End Function


Public Function ShowColorDlg() As Variant
'***************************************************************************************
' This function was built because i wanted to show the color CMD and have it return a  *
' Value of the color selected without rewriting the concept over and over              *
' NOTE : If Cancel is Pressed NULL is Returned.                                        *
'***************************************************************************************

 On Error GoTo CdlcCancelError: 'If an Error Occurs jump to that line
 
    With FrmMain.CFontDialog
         .Cancelerror = True 'Make sure Pressing Cancel Generates an Error
         .Flags = cdlCCFullOpen 'Make sure the user can create a custom color
         .ShowColor 'Show the Dialog
         ShowColorDlg = .Color 'Return the color selected
         Exit Function
    End With

CdlcCancelError:
    If Err.Number = 32755 Then 'Cancel Was Pressed
       ShowColorDlg = Null 'Return NULL Becuase Cancel was Pressed
       Exit Function
    End If
End Function


Public Function MakeBorderStatic(ByVal hWnd As Long)
    
    Dim LngRetVal As Long
    
    'Retrieve the current border style
    LngRetVal = GetWindowLong(hWnd, GWL_EXSTYLE)
    
    'Calculate border style to use
    LngRetVal = LngRetVal Or WS_EX_STATICEDGE And Not WS_EX_CLIENTEDGE
    
    'Apply the changes
    SetWindowLong hWnd, GWL_EXSTYLE, LngRetVal
    SetWindowPos hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or _
                 SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
    
End Function

Public Function SaveFile(sFileName As String, mTextbox As TextBox, PromptUI As Boolean) As VbMsgBoxResult
'************************************************************************************************************************
' This function was created to decrease the constant redundancy of saving a file and it can be used in different ways   *
' it can be used to save a file and if the promptUI argument is true it notifys the user that the file has changed and  *
' then the function returns what button was pressed.                                                                    *
'************************************************************************************************************************
 
 Dim Response
 
 Select Case PromptUI
     Case True ' Prompt user
         ' show the MsgBox with the msg , Buttons ; Style , title
          Response = MsgBox(Msga, vbYesNoCancel + vbExclamation + vbDefaultButton3, "NextPad")

           Select Case Response ' select a response
                Case vbYes     ' User chose the Yes button .
                  SaveFile sFileName, mTextbox, False  ' Call Procedure But now Bypassing MsgBox
                  SaveFile = Response ' Return Response given from user
                  Exit Function ' Stop and quit
                Case vbCancel ' User chose the Cancel Button
                  SaveFile = Response ' Return Response given from user
                  Exit Function ' Exit the sub before running any more code in this sub
                Case vbNo ' the user chose the No Button
                  SaveFile = Response ' Return Response given from user
                  Exit Function ' Stop and quit
           End Select
           
     Case False ' DO NOT prompt user
  On Error GoTo FileError:
   Select Case sFileName ' Check the filename
       Case Is <> "" ' There is a string
          Close #1 ' Close all open files
          Open sFileName For Output As #1 ' Open the file for output
          Print #1, mTextbox.Text ' Save it
          Close #1 ' Close this file range number
          AddRecentFile sFileName ' Add this file to the recent files menu.
          ' If the text box is the main one then...
          If mTextbox = FrmMain.mTextbox Then
             sOpenFileName = sFileName ' Tell NextPad the current filename
             FrmMain.Caption = GetFileTitleStr(sFileName) & " - NextPad" ' Set the caption of the main form
          End If
       Case Is = "" ' There is no string
           With FrmMain
               On Error GoTo Cancelerror: ' If cancel gets pressed an error will occure jump to that line
                .CommonDialog1.Cancelerror = True ' Set the cancel error to true to generate an error if cancel is pressed
                .CommonDialog1.Flags = cdlOFNOverwritePrompt ' Set the normal flags
                .CommonDialog1.DialogTitle = "Save" ' Set the dialogs caption
                .CommonDialog1.Filter = "Text documents (*.TXT)|*.TXT|INI Configuration Files (*.INI)|*.INI|Log Files (*.LOG)|*.LOG|All Files (*.*) |*.*" ' Set which files can be save
                .CommonDialog1.ShowSave ' Show the save dialog box
                    If .CommonDialog1.Filename <> "" Then ' If there is a filename...
                        Close #1 ' Reset all open disks
                        Open .CommonDialog1.Filename For Output As #1 ' Open it for output
                        Print #1, mTextbox.Text ' Save the file
                        Close #1 ' Close the file range number
                        AddRecentFile sFileName ' Add this file to the recent files menu.
                        ' If the text box is the main one then...
                        If mTextbox = FrmMain.mTextbox Then
                           sOpenFileName = .CommonDialog1.Filename ' Tell NextPad the current filename
                           FrmMain.Caption = GetFileTitleStr(sOpenFileName) & " - NextPad" ' Set the caption of the main form
                           Fstate.dirty = False ' Tell NextPad that the file is clean now
                        End If
                    End If
           End With
     End Select ' Second select case statement
  End Select ' First select case statement
Cancelerror:
' If cancel was pressed then
If Err.Number = 32755 Then
   Exit Function 'Exit this sub
End If

FileError:
' A file error has occured
If Err.Number <> 0 Then
   FileError sFileName, mTextbox
   Exit Function 'Exit this Function
End If
End Function

Sub InsertAtSel(mTextbox As TextBox, sInsertion As String)
   On Error GoTo ErrorHandler:
  
    SendMessage mTextbox.hWnd, EM_REPLACESEL, 1, ByVal sInsertion ' Now we paste it and now are unable to undo.

ErrorHandler:
    If Err.Number <> 0 Then
       MsgBox "Cannot insert text at selection point, NextPad might be out of memory.", vbCritical, "NextPad"
       Exit Sub
    End If
End Sub

Public Function FindAssociatedExe(sFileName As String) As String
'************************************************************************************
' This function was created because the old way of retrieving the external editor   *
' did not work any longer and was to cumbersome so now we create an empty file and  *
' look for its associated program.                                                  *
'************************************************************************************
   
   Dim i As Integer, s2 As String
   
   'Check if the file exists
   If Dir(sFileName) = "" Or sFileName = "" Then
        FindAssociatedExe = ""
        Exit Function
   End If
   'Create a buffer
   s2 = String(MAX_FILENAME_LEN, 32)
   'Retrieve the name and handle of the executable, associated with this file
   i = FindExecutable(sFileName, vbNullString, s2)
   If i > 32 Then
      FindAssociatedExe = Left$(s2, InStr(s2, Chr$(0)) - 1)
   Else
      FindAssociatedExe = ""
   End If


End Function

Sub ClearNextPad()
      FrmMain.txt(0).Text = ("") ' Set the forms Active Control's Text Too nothing ("")
      FrmMain.txt(1).Text = ("") ' Set the forms Active Control's Text Too nothing ("")
      sOpenFileName = "" ' Set the current file name
      FrmMain.Caption = "Untitled - NextPad" ' Set the Forms Caption using the Caption Property
      Fstate.dirty = False ' No file has been created or edited yet.
End Sub

Sub AddFavorite(sFileName As String)
        
    ' If the user disabled the favorites menu, then we cannot continue.
    If Not AllowFavorites Then Exit Sub
    
    Dim iCount As Integer
    ' Get number of favorites in registry
    iCount = GetSetting("NextPad", "Favorites", "Count", 0)
    If GetSetting("NextPad", "Favorites", 0, "") <> "" Then iCount = iCount + 1
    
    With FrmMain.MnuFavorites
       ' If the default "Disabled" Menu is there use it
       Dim i As Integer ' Variable to hold menu count
       i = .Count ' Append value to integer
       ' If there is nothing there make sure that we show one
       If i = 0 Then .Item(0).Caption = GetFileTitleStr(sFileName): .Item(0).Tag = sFileName: .Item(0).Visible = True: GoTo Save:
       i = i + 1 ' Since there is one there already add one
       Load .Item(i) ' Load the new menu
       .Item(i).Visible = True
       .Item(i).Tag = sFileName
       .Item(i).Caption = GetFileTitleStr(sFileName)  ' Put a caption on it
    End With
    
Save:
    ' Save Favorite in registry
   If iCount = 0 Then
     SaveSetting "Favorites", 0, sFileName
     ' Update number of favorites
     SaveSetting "Favorites", "Count", 0
   Else
     SaveSetting "Favorites", (iCount), sFileName
     ' Update number of favorites
     SaveSetting "Favorites", "Count", (iCount)
   End If
   
End Sub

Sub GetFavorites()
On Error GoTo ErrorHandler:
    
    ' If the user disabled the favorites menu, then we cannot continue.
    If Not AllowFavorites Then FrmMain.MnuFavoritesItem.Visible = False
     
  Dim i As Integer, X As Integer
   i = GetSetting("NextPad", "Favorites", "Count", 0)
   If GetSetting("NextPad", "Favorites", 0, "") = "" Then Exit Sub ' If there are NO favorites Stop
     
     With FrmMain.MnuFavorites
        .Item(0).Visible = True
        .Item(0).Caption = GetFileTitleStr(GetSetting("NextPad", "Favorites", 0))
        .Item(0).Tag = GetSetting("NextPad", "Favorites", 0)
        If i = 0 Then Exit Sub
        For X = 1 To i
          Load .Item(X) ' Load the new menu
          ' Hide the real filename internally
          .Item(X).Tag = GetSetting("NextPad", "Favorites", X)
          ' Give it a caption
          .Item(X).Caption = GetFileTitleStr(GetSetting("NextPad", "Favorites", X))
          ' Allow user to see it
          .Item(X).Visible = True
        Next X
     End With

ErrorHandler:
   If Err.Number <> 0 Then
      MsgBox "Could not get Favorites from registry. If problem persist's please go to the Options window and reset NextPad's settings.", vbCritical, "NextPad - Favorites"
      Exit Sub
   End If
End Sub

Function StrInvertCase(Text As String) As String
 '***************************************************************************
 ' This function inverts a string by going through it char by char.         *
 ' If it finds an uppercase it lowers it, and so on.                        *
 '***************************************************************************
  
  Dim MyStr, MyStr2
  ' Go through each letter one by one
  FrmMain.MousePointer = vbHourglass ' Let user know NextPad is working
  For i = 1 To Len(Text) Step 1
    ' Get the next letter
    MyStr = Mid$(Text, i, 1)
     
     ' If any lines are a new line character then make a new line and reset the variable
     If Mid$(Text, i, Len(vbNewLine)) = vbNewLine Then MyStr2 = MyStr2 & vbNewLine: MyStr = ""
       ' Add these invalid characters because they have no case.
     Select Case MyStr
       Case " " ' Space
          MyStr2 = MyStr2 & " "
       Case "." ' Other invalid character
          MyStr2 = MyStr2 & MyStr
       Case "," ' Other invalid character
          MyStr2 = MyStr2 & MyStr
       Case ";" ' Other invalid character
          MyStr2 = MyStr2 & MyStr
       Case ":" ' Other invalid character
          MyStr2 = MyStr2 & MyStr
       Case "/" ' Other invalid character
          MyStr2 = MyStr2 & MyStr
       Case "[" ' Other invalid character
          MyStr2 = MyStr2 & MyStr
       Case "]" ' Other invalid character
          MyStr2 = MyStr2 & MyStr
       Case "(" ' Other invalid character
          MyStr2 = MyStr2 & MyStr
       Case ")" ' Other invalid character
          MyStr2 = MyStr2 & MyStr
       Case "'" ' Other invalid character
          MyStr2 = MyStr2 & MyStr
       Case "{" ' Other invalid character
          MyStr2 = MyStr2 & MyStr
       Case "}" ' Other invalid character
          MyStr2 = MyStr2 & MyStr
       Case "!" ' Other invalid character
          MyStr2 = MyStr2 & MyStr
       Case "@" ' Other invalid character
          MyStr2 = MyStr2 & MyStr
       Case "#" ' Other invalid character
          MyStr2 = MyStr2 & MyStr
       Case "$" ' Other invalid character
          MyStr2 = MyStr2 & MyStr
       Case "%" ' Other invalid character
          MyStr2 = MyStr2 & MyStr
       Case "%" ' Other invalid character
          MyStr2 = MyStr2 & MyStr
       Case "^" ' Other invalid character
          MyStr2 = MyStr2 & MyStr
       Case "&" ' Other invalid character
          MyStr2 = MyStr2 & MyStr
       Case "*" ' Other invalid character
          MyStr2 = MyStr2 & MyStr
       Case "-" ' Other invalid character
          MyStr2 = MyStr2 & MyStr
       Case "_" ' Other invalid character
          MyStr2 = MyStr2 & MyStr
       Case "+" ' Other invalid character
          MyStr2 = MyStr2 & MyStr
       Case "=" ' Other invalid character
          MyStr2 = MyStr2 & MyStr
       Case "|" ' Other invalid character
          MyStr2 = MyStr2 & MyStr
       Case "\" ' Other invalid character
          MyStr2 = MyStr2 & MyStr
       Case "<" ' Other invalid character
          MyStr2 = MyStr2 & MyStr
       Case ">" ' Other invalid character
          MyStr2 = MyStr2 & MyStr
       Case Chr(34) ' Other invalid character
          MyStr2 = MyStr2 & MyStr
       Case "~" ' Other invalid character
          MyStr2 = MyStr2 & MyStr
       Case "`" ' Other invalid character
          MyStr2 = MyStr2 & MyStr
       Case "?" ' Other invalid character
          MyStr2 = MyStr2 & MyStr
     End Select
                
     On Error Resume Next ' If an error occurs just jump to the next line...
     If IsNumeric(MyStr) Then
        MyStr2 = MyStr2 & MyStr
     End If
     
     On Error Resume Next ' If an error occurs just jump to the next line...
     If IsCharUpper(AscB(MyStr)) Then
        ' If the character is uppercase then make it lowercase.
        MyStr2 = MyStr2 & LCase(MyStr)
     End If
     
     On Error Resume Next ' If an error occurs just jump to the next line...
     If IsCharLower(AscB(MyStr)) Then
        ' If the character is lowercase then make it uppercase.
        MyStr2 = MyStr2 & UCase(MyStr)
     End If
  Next i
  ' Return the now inverted string
  StrInvertCase = MyStr2
  FrmMain.MousePointer = 0 ' Reset mouse pointer
End Function

Public Function FormatKB(ByVal Amount As Long) _
    As String
    Dim Buffer As String
    Dim Result As String
    Buffer = Space$(255)
    Result = StrFormatByteSize(Amount, Buffer, _
    Len(Buffer))


    If InStr(Result, vbNullChar) > 1 Then


        FormatKB = Left$(Result, InStr(Result, _
            vbNullChar) - 1)
        End If
    End Function

Public Function GetLongFileName(ByVal ShortFileName As String) As String

    Dim intPos As Integer
    Dim strLongFileName As String
    Dim strDirName As String
    
    'Format the filename for later processing
    ShortFileName = ShortFileName & "\"
    
    'Grab the position of the first real slash
    intPos = InStr(4, ShortFileName, "\")
    
    'Loop round all the directories and files
    'in ShortFileName, grabbing the full names
    'of everything within it.
    
    While intPos
    
        strDirName = Dir(Left(ShortFileName, intPos - 1), _
            vbNormal + vbHidden + vbSystem + vbDirectory)
        
        If strDirName = "" Then
            GetLongFileName = ""
            Exit Function
        End If
        
        strLongFileName = strLongFileName & "\" & strDirName
        intPos = InStr(intPos + 1, ShortFileName, "\")
        
    Wend

    'Return the completed long file name
    GetLongFileName = Left(ShortFileName, 2) & strLongFileName
  
End Function
Function GetPathFromFileName(sFullPath As String) As String
On Error GoTo ErrorHandler:
  If sFullPath = "" Then Exit Function
  ' Strip the path from the full path and filename EX: C:\WINDOWS\NOTEPAD.EXE [becomes:] C:\WINDOWS
  PathRemoveFileSpec sFullPath
  ' Return path
  GetPathFromFileName = StripTerminator(sFullPath)

ErrorHandler:
  If Err.Number <> 0 Then
     ' Give back the given filename.
     GetPathFromFileName = sFullPath
     Exit Function ' Exit this function now.
  End If
End Function

'This function is used to strip all the unnecessary chr$(0)'s
Function StripTerminator(sInput As String) As String
    Dim ZeroPos As Integer
    'Search the position of the first chr$(0)
    ZeroPos = InStr(1, sInput, vbNullChar)
    If ZeroPos > 0 Then
        StripTerminator = Left$(sInput, ZeroPos - 1)
    Else
        StripTerminator = sInput
    End If
End Function

Function ChooseDir() As String
 ' Show the 'Choose Directory' form
 FrmChooseDir.Show vbModal
   ' If the variable returned from the form has data...
   If sDir <> "" And PathIsDirectory(sDir) Then
      ' Return path
      ChooseDir = sDir
   Else ' There is no data
      ChooseDir = ""
   End If
   ' Reset the variable
   sDir = ""
End Function
