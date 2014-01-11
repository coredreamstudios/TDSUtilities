Attribute VB_Name = "basSHFileOp"
Option Explicit

' ---------------------------------------------------------
' Constants and variables
' ---------------------------------------------------------
  Private Const NORMAL_PRIORITY_CLASS As Long = 0   '&H20&
  Private Const INFINITE As Long = -1&
  Private Const MAX_PATH As Long = 260
  Private Const BUFFER_SIZE As Long = 32766    ' 2 bytes short
  Private Const SWP_NOACTIVATE As Long = &H10
  Private Const SWP_SHOWWINDOW As Long = &H40
  Private Const SWP_NOMOVE As Long = 2
  Private Const SWP_NOSIZE As Long = 1
  Private Const HWND_FLAGS As Long = SWP_NOMOVE Or SWP_NOSIZE
  Private Const HWND_TOPMOST As Long = -1
  Private Const HWND_NOTOPMOST As Long = -2
  Public sDestPath As String
  Public sTmpFileName As String
  Public bInProgress As Boolean
  
' ------------------------------------------------------------------------
' Declares required for SHFileOperation API call
' ------------------------------------------------------------------------
  Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" _
                 (ByVal lpszPath As String, ByVal lpPrefixString As String, _
                 ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
                 
  Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" _
                 (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

' ------------------------------------------------------------------------
' Declares required for SHFileExecute API call
' ------------------------------------------------------------------------
  Public Const SEE_MASK_INVOKEIDLIST As Long = &HC
  Public Const SEE_MASK_NOCLOSEPROCESS As Long = &H40
  Public Const SEE_MASK_FLAG_NO_UI As Long = &H400

  Public Type SHELLEXECUTEINFO
         cbSize As Long
         fMask As Long
         hwnd As Long
         lpVerb As String
         lpFile As String
         lpParameters As String
         lpDirectory As String
         nShow As Long
         hInstApp As Long
         lpIDList As Long     'Optional parameter
         lpClass As String    'Optional parameter
         hkeyClass As Long    'Optional parameter
         dwHotKey As Long     'Optional parameter
         hIcon As Long        'Optional parameter
         hProcess As Long     'Optional parameter
  End Type

  Public Declare Function ShellExecuteEX Lib "shell32.dll" _
           Alias "ShellExecuteEx" (SEI As SHELLEXECUTEINFO) As Long

' ------------------------------------------------------------------------
' Required for SHBrowseForFolder API call
' ------------------------------------------------------------------------
  Public Type BrowseInfo
         hWndOwner As Long
         pIDLRoot As Long
         pszDisplayName As Long
         lpszTitle As Long
         ulFlags As Long
         lpfnCallback As Long
         lParam As Long
         iImage As Long
  End Type

  Public Const BIF_RETURNONLYFSDIRS = 1

  Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
  Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
  
  Public Declare Function SHGetPathFromIDList Lib "shell32" _
                 (ByVal pidList As Long, ByVal lpBuffer As String) As Long

  Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
                 (ByVal lpString1 As String, ByVal lpString2 As String) As Long

' ------------------------------------------------------------------------
' Operation Functions
' ------------------------------------------------------------------------
  Public Enum sfos_wfunc
       FO_MOVE = &H1
       FO_COPY = &H2
       FO_DELETE = &H3
       FO_RENAME = &H4
  End Enum
     
' ------------------------------------------------------------------------
' Flags that control the file operation. This member can be a
' combination of the following values:
'
' FOF_ALLOWUNDO           Preserves undo information, if possible.
' FOF_CONFIRMMOUSE        Not implemented.
' FOF_FILESONLY           Performs the operation only on files if
'                         a wildcard filename (*.*) is specified.
' FOF_MULTIDESTFILES      Indicates that the pTo member specifies
'                         multiple destination files (one for each
'                         source file) rather than one directory
'                         where all source files are to be deposited.
' FOF_NOCONFIRMATION      Responds with "yes to all" for any dialog
'                         box that is displayed.
' FOF_NOCONFIRMMKDIR      Does not confirm the creation of a new
'                         directory if the operation requires one to
'                         be created.
' FOF_RENAMEONCOLLISION   Gives the file being operated on a new name
'                         (such as "Copy #1 of...") in a move, copy,
'                         or rename operation if a file of the target
'                         name already exists.
' FOF_SILENT              Does not display a progress dialog box.
' FOF_SIMPLEPROGRESS      Displays a progress dialog box, but does
'                         not show the filenames.
' FOF_WANTMAPPINGHANDLE   Fills in the hNameMappings member.
' ------------------------------------------------------------------------
  Public Enum sfos_fflags
       FOF_CREATEPROGRESSDLG = &H0
       FOF_MULTIDESTFILES = &H1
       FOF_CONFIRMMOUSE = &H2
       FOF_SILENT = &H4
       FOF_RENAMEONCOLLISION = &H8
       FOF_NOCONFIRMATION = &H10
       FOF_WANTMAPPINGHANDLE = &H20
       FOF_ALLOWUNDO = &H40
       FOF_FILESONLY = &H80
       FOF_SIMPLEPROGRESS = &H100
       FOF_NOCONFIRMMKDIR = &H200
  End Enum
     
  Type SHFILEOPSTRUCT
        hwnd As Long
        wFunc As Long
        pFrom As String
        pTo As String
        fFlags As Integer
        fAnyOperationsAborted As Long
        hNameMappings As Long
        lpszProgressTitle As String '  only used if FOF_SIMPLEPROGRESS
  End Type

' ------------------------------------------------------------------------
' Declares required
' ------------------------------------------------------------------------
  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
                  (xdest As Any, xsource As Any, ByVal xsize As Long)
  
  Private Declare Function SHFileOperationA Lib "shell32.dll" (lpFileOp As Byte) As Long
  
  Private Declare Function SHFileOperation Lib "shell32.dll" Alias " SHFileOperationA" _
                  (lpFileOp As SHFILEOPSTRUCT) As Long
  
  Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
                  (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
  
  Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
                   ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, _
                   ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
  
  
Public Function ShellFileOp(hwnd As Long, TitleText As String, FromPath As String, _
                            ToPath As String, OpFunction As sfos_wfunc, _
                            OpFlags As sfos_fflags) As Boolean
     
' -------------------------------------------------------------
' Author unknown.  Great code, though.
'
' This function builds a byte array to pass to the
' SHFileOperation API call. We can not use the SHFILEOPSTRUCT
' definition because it has an Integer embedded in it and VB
' pads extra bytes around that throwing the following
' parameters out of whack
'
' -------------------------------------------------------------
'
' Notes:
'     1. The user pressing CANCEL is not apparently considered
'        an error
'     2. On fast copies the animated display may not appear
'     3. Make sure the destination directory exist.  The subs
'        will be created automatically.
'
' To use this function, do something like this for a
' simple copy operation:
'
'  Dim strTitle As String
'  Dim strFrom As String
'  Dim strTo As String
'  Dim lngFlags As Long
'
'  strTitle = "Backing up documents..."
'  strFrom = "C:\My Documents"
'  strTo = "C:\Backup"
'  lngFlags = FOF_SIMPLEPROGRESS Or FOF_RENAMEONCOLLISION
'
' Use Me.hWnd or 0& for a form handle
'
'  If ShellFileOp(Me.hWnd, strTitle, strFrom, strTo, FO_COPY, lngFlags) Then
'      MsgBox "Backup did not report error"
'  Else
'      MsgBox "Backup failed"
'  End If
' -------------------------------------------------------------

  
' -------------------------------------------------------------
' Define local variables
' -------------------------------------------------------------
  Dim sos(30) As Byte          ' our "dummy" SHFILEOPSTRUCT
  Dim bytFrom() As Byte        ' Source array
  Dim bytTo() As Byte          ' Destination array
  Dim bytTitle() As Byte       ' Title array
  Dim intFlags As Integer      ' flags field must be short int
  Dim lngTemp As Long          ' temp variable to hold pointers/etc
  Dim hDummyHandle As Long
  Dim lRetVal As Long
  
' -------------------------------------------------------
' get the handle of the "Main Menu" window so the
' dialog box can be centered
' -------------------------------------------------------
  lRetVal = FindWindow(0&, "Main Menu")
  hDummyHandle = lRetVal
    
' -------------------------------------------------------------
' Initialize the arrays and convert from Unicode, if not
' already converted
' -------------------------------------------------------------
  bytTo = StrConv(ToPath & vbNullChar & vbNullChar, vbFromUnicode)
  bytFrom = StrConv(FromPath & vbNullChar & vbNullChar, vbFromUnicode)
  bytTitle = StrConv(TitleText & vbNullChar, vbFromUnicode)

' -------------------------------------------------------------
' copy the parameters into the byte array
' -------------------------------------------------------------
  CopyMemory sos(0), hDummyHandle, 4     ' set window handle
  
  CopyMemory sos(4), OpFunction, 4       ' set operation function
  
  lngTemp = VarPtr(bytFrom(0))
  CopyMemory sos(8), lngTemp, 4          ' copy address of FROM string
  
  lngTemp = VarPtr(bytTo(0))
  CopyMemory sos(12), lngTemp, 4         ' copy address of TO string
  
  intFlags = OpFlags And &HFFFF&
  CopyMemory sos(16), intFlags, 2        ' copy operation flags
  
  lngTemp = 0
  CopyMemory sos(18), lngTemp, 4         ' set hNameMappings
  
  lngTemp = 1                            ' we want to know about aborts
  CopyMemory sos(22), lngTemp, 4         ' set fAnyoperationsaborted
  
  lngTemp = VarPtr(bytTitle(0))
  CopyMemory sos(26), lngTemp, 4         ' copy address of TITLE string
  
' -------------------------------------------------------------
' Perform the operation
' -------------------------------------------------------------
  lngTemp = SHFileOperationA(sos(0))
  
' -------------------------------------------------------------
' If the return value is not = zero then something went wrong
' -------------------------------------------------------------
  If lngTemp Then
      ShellFileOp = False                ' operation failed
  Else
      If OpFunction = FO_DELETE Then
          
          If Right(FromPath, 1) <> "\" Then
              FromPath = FromPath & "\"
          End If
          
          ' if this path still exist then the user pressed CANCEL
          If ItemExist(FromPath) Then
              ShellFileOp = False
              Exit Function
          End If
      End If
      
      ' call OK... did anything abort?
      CopyMemory lngTemp, sos(22), 4     ' copy fAnyOperationsAborted bytes
      If lngTemp Then
          ShellFileOp = False            ' something is wrong
      Else
          ShellFileOp = True             ' looks good
      End If
  End If

End Function


Public Function BrowseForFolder(Optional sPrompt As String) As String

' ---------------------------------------------------
' Define local variables
' ---------------------------------------------------
  Dim iPos As Integer
  Dim lFolderList As Long
  Dim lRetVal As Long
  Dim hDummyHandle As Long
  Dim sTmpPath As String
  Dim BI As BrowseInfo

' ---------------------------------------------------
' Set up Browse Dialog Box parameters
' ---------------------------------------------------
  If IsMissing(sPrompt) Then
      sPrompt = "Browse for work area"
  End If

' -------------------------------------------------------
' get the handle of the "Main Menu" form.  This is the
' form caption line.
' -------------------------------------------------------
  hDummyHandle = IsTaskActive("Compact/")

' ---------------------------------------------------
' Set up Browse Dialog Box parameters
' ---------------------------------------------------
  With BI
          ' use zero as the default window handle
          .hWndOwner = hDummyHandle
          
          ' sPrompt will be the title on the Browse dialog box
          .lpszTitle = lstrcat(sPrompt, "")
          
          ' Display only directories
          .ulFlags = BIF_RETURNONLYFSDIRS
  End With

' ---------------------------------------------------
' display the browse dialog box
' ---------------------------------------------------
  lRetVal = SetWindowPos(BI.hWndOwner, HWND_TOPMOST, 0, 0, 0, 0, HWND_FLAGS)
  lFolderList = SHBrowseForFolder(BI)
  lRetVal = SetWindowPos(BI.hWndOwner, HWND_NOTOPMOST, 0, 0, 0, 0, HWND_FLAGS)
  
' ---------------------------------------------------
' If a folder was highlighted then format the
' folder name so it can be returned
' ---------------------------------------------------
  If lFolderList Then
      
      ' set up a pre-padded buffer area for the folder name
      sTmpPath = String(MAX_PATH, 0)
      
      ' Get the name of the folder from the list
      lRetVal = SHGetPathFromIDList(lFolderList, sTmpPath)
      
      Call CoTaskMemFree(lFolderList)
      
      ' Strip any null characters from the folder name
      iPos = InStr(sTmpPath, vbNullChar)
      If iPos > 0 Then
          sTmpPath = Left(sTmpPath, iPos - 1)
      End If
  End If

' ---------------------------------------------------
' return the formatted folder name
' ---------------------------------------------------
  BrowseForFolder = sTmpPath

End Function
Public Function IsThisRestricted(sTestPath As String) As Boolean

On Error GoTo Problems_Detected
' --------------------------------------------------------
' Define local variables
' --------------------------------------------------------
  Dim sTmpPath As String
  Dim iFile As Integer
  
' --------------------------------------------------------
' initialize local variables
' --------------------------------------------------------
  iFile = FreeFile
  
' --------------------------------------------------------
' See if there is a file parameter on the end of
' these path strings.  If so, remove it so we can
' see if we have access to their areas.
' --------------------------------------------------------
  If Len(Trim(sTestPath)) > 0 Then
  
      ' add trailing "\" if missing
      If Right(sTestPath, 1) <> "\" Then
          sTestPath = sTestPath & "\"
      End If
          
      ' open the new file on the drive to make sure
      ' we can write/delete
      Open sTestPath & "X" For Output As #iFile
      Close #iFile
      Kill sTestPath & "X"
      '
      IsThisRestricted = False
  Else
      IsThisRestricted = True
  End If
  
  
Normal_Exit:
  On Error GoTo 0
  Exit Function
  
  
Problems_Detected:
  Err.Clear
  IsThisRestricted = True
  GoTo Normal_Exit
  
End Function
Public Sub ShowDriveProperties(sDriveLtr As String, frm As Form)
    
' ---------------------------------------------------
' display the windows property page for a drive
' ---------------------------------------------------
  
' ---------------------------------------------------
' Define local variables
' ---------------------------------------------------
  Dim SEI As SHELLEXECUTEINFO
  Dim lRetVal As Long
   
' ---------------------------------------------------
' set up the variables
' ---------------------------------------------------
  With SEI
       .cbSize = Len(SEI)
       .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
       .hwnd = frm.hwnd
       .lpVerb = "properties"
       .lpFile = sDriveLtr
       .lpParameters = vbNullChar
       .lpDirectory = vbNullChar
       .nShow = 0
       .hInstApp = 0
       .lpIDList = 0
  End With
    
' ---------------------------------------------------
' display the property page
' ---------------------------------------------------
  lRetVal = ShellExecuteEX(SEI)
    
End Sub

