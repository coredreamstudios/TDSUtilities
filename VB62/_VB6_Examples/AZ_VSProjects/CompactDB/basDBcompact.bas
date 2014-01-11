Attribute VB_Name = "basDBCompact"
Option Explicit

' ---------------------------------------------------------
' Compact MDB files
'
' Written by Kenneth Ives              kenaso@home.com
'
' This program will allow the user to select a MDB file
' to compact.  The size of the file is captured and a
' calculation of twice that size is made to determine
' the amount of free space required to compact the
' database.  Half that amount is used for a backup copy
' of the original database and the other half is for
' the compacted database.  if there is not enough space,
' the user is prompted to select another path in which
' to perform this operation or leave the application.
' After the database is compacted, the original is deleted
' and the new version is moved back into the place of the
' original.
' ---------------------------------------------------------

' ---------------------------------------------------------
' Define variables
' ---------------------------------------------------------
  Public sDatabase As String
  Public sCurPosition As String
  Public sWorkDrive As String
  Private sPartialTitle As String
  
  Private bUsedTempDir As Boolean
  Private bFoundApp As Boolean
  
  Private lAppHandle As Long
  Private lStartSize As Long
  Private lEndSize As Long

  Private WS As Workspace
  Private DB As Database
  Private Const TEMP_DIR As String = "TMP_@@@@\"
 
  Private Declare Function CopyFile Lib "kernel32" _
          Alias "CopyFileA" (ByVal lpExistingFileName As String, _
          ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long

  Private Declare Function GetTempPath Lib "kernel32" _
          Alias "GetTempPathA" (ByVal nBufferLength As Long, _
          ByVal lpBuffer As String) As Long

  Private Declare Function GetDiskFreeSpace Lib "kernel32" _
          Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, _
          lpSectorsPerCluster As Long, lpBytesPerSector As Long, _
          lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
  
  Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" _
          Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, _
          lpFreeBytesAvailableToCaller As Currency, lpTotalNumberOfBytes As Currency, _
          lpTotalNumberOfFreeBytes As Currency) As Long

  Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
          (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

  Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
          (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
          
  Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
  
  Private Declare Function EnumWindows Lib "user32" _
          (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long


' -------------------------------------------------------------
' Type layout for disk information
' -------------------------------------------------------------
  Private Type DRV_DATA
       SectorsPerCluster As Long      ' Number of sectors per cluster
       BytesPerSector As Long         ' Size of each sector
       TotalSectors As Long           ' Total number of Sectors
       FreeSectors As Long            ' Number of free Sectors
       UsedSectors As Long            ' Number of used Sectors
       TotalClusters As Long          ' Total number of clusters
       FreeClusters As Long           ' Number of free clusters
       UsedClusters As Long           ' Number of used clusters
       TotalAvailSpace As Double      ' Actual space available to the user
       TotalFreeSpace As Double       ' Free space on the disk
       TotalUsedSpace As Double       ' Used space on the disk
       TotalDiskSize As Double        ' Total disk size
       SpaceTotal As String           ' Total disk space
       SpaceAvailable As String       ' Actual space available to the user
       SpaceFree As String            ' Free space on the disk
       SpaceUsed As String            ' Used space on the disk
       SpaceAvailablePcnt As String   ' Actual space available to the user
       SpaceFreePcnt As String        ' Free space on the disk
       SpaceUsedPcnt As String        ' Used space on the disk
       MaxRootFileEntries As Long     ' Max number of entries at root level
       Type As String                 ' Type of drive
       Number As String               ' Drive number (A=0)
       DriveLetter As String          ' Letter assigned to a drive
       VolumeName As String           ' Name or Label of disk
       SerialNumber As String         ' Serial Number of disk
       FileSysType As String          ' Type of file system
  End Type

  Private DiskInfo As DRV_DATA
Public Function FindApplication(ByVal app_hWnd As Long, ByVal param As Long) As Long

' ---------------------------------------------------------
' Check the title line of all active application windows
' while looking for a match on all or part of the title.
'
' Called from IsTaskActive
' ---------------------------------------------------------

' ---------------------------------------------------------
' Define local variables
' ---------------------------------------------------------
  Dim lLength As Long    ' Length of the title string
  Dim sBuffer As String  ' buffer area to hold the title string
  Dim sTitle As String   ' application title string after formatting

' ---------------------------------------------------------
' Initialize buffer string with blank values
' ---------------------------------------------------------
  sBuffer = Space(256)
  
' ---------------------------------------------------------
' Get the window's title.  (API call)
' ---------------------------------------------------------
  lLength = GetWindowText(app_hWnd, sBuffer, Len(sBuffer))
  sTitle = StrConv(Left(sBuffer, lLength), vbLowerCase)

' ---------------------------------------------------------
' See if this is the target window.
' ---------------------------------------------------------
  If InStr(1, sTitle, sPartialTitle) > 0 Then
      
      ' capture the handle of the application window
      lAppHandle = FindWindow(vbNullString, sTitle)
      Exit Function
  End If
    
' ---------------------------------------------------------
' Continue searching the application windows
' ---------------------------------------------------------
  FindApplication = 1

End Function
Public Function IsTaskActive(SApplName As String) As Long

' -------------------------------------------------------------
' Define local variables
' -------------------------------------------------------------
  Dim lRetVal As Long
  
' ---------------------------------------------------------
' Initialize variables
' ---------------------------------------------------------
  sPartialTitle = StrConv(SApplName, vbLowerCase)
  lAppHandle = 0
  
' ---------------------------------------------------------
' Ask Windows for the list of tasks.
' ---------------------------------------------------------
  lRetVal = EnumWindows(AddressOf FindApplication, 0&)
  
' ---------------------------------------------------------
' If successful then return the application handle
' ---------------------------------------------------------
  IsTaskActive = lAppHandle
  
End Function

Public Function ShrinkToFit(sInText As String, iMaxLen As Integer) As String

' -------------------------------------------------------------
' This function will shorten a directory name to the
' length passed to the Max parameter.
'
' Syntax:
'
'  sTmp = ShrinkToFit("C:\Program Files\Netscape\Navigator\Programs\Bookmark.htm", 30)
'
' Returns:
'  sTmp = "C:\...\Programs\Bookmark.htm"
' -------------------------------------------------------------

' -------------------------------------------------------------
' Define local variables
' -------------------------------------------------------------
  Dim i As Integer
  Dim iStart As Integer
  Dim iLblLen As Integer
  Dim sTmpStr As String

' -------------------------------------------------------------
' Initialize variables
' -------------------------------------------------------------
  sTmpStr = Trim(sInText)
  iLblLen = iMaxLen

' -------------------------------------------------------------
' if the string is equal to or less than the desired length
' then leave this routine
' -------------------------------------------------------------
  If Len(sTmpStr) <= iLblLen Then
      ShrinkToFit = sTmpStr
      Exit Function
  End If

' -------------------------------------------------------------
' Readjust the desired length by six
' -------------------------------------------------------------
  iLblLen = iLblLen - 6
  iStart = (Len(sTmpStr) - iLblLen)

' -------------------------------------------------------------
' parse thru the string and find the first occurance of a
' backslash starting from a calculated starting position
' -------------------------------------------------------------
  For i = iStart To Len(sTmpStr)
      If Mid(sTmpStr, i, 1) = "\" Then Exit For
  Next
      
' -------------------------------------------------------------
' Return the shortened string
' -------------------------------------------------------------
  ShrinkToFit = Left(sTmpStr, 3) & "..." & Right(sTmpStr, Len(sTmpStr) - (i - 1))
  DoEvents
  
End Function
  
Public Sub CompactMDB()

' -----------------------------------------------------------------
' CompactMDB routine is the heart of this module.
'
' Written by Kenneth Ives          kenaso@home.com
'
' This program will allow the user to select a MDB file
' to compact.  The size of the file is captured and a
' calculation of twice that size is made to determine
' the amount of free space required to compact the
' database.  Half that amount is used for a backup copy
' of the original database and the other half is for
' the compacted database.  if there is not enough space,
' the user is prompted to select another path in which
' to perform this operation or leave the application.
' After the database is compacted, the original is deleted
' and the new version is moved back into the place of the
' original.
' -----------------------------------------------------------------

' ------------------------------------------------------------
' Define local variables
' ------------------------------------------------------------
  Dim sMsg As String
  Dim sNewFile As String
  Dim sBakFile As String
  Dim sDBPath As String
  Dim sDBName As String
  Dim sStartSize As String
  Dim sEndSize As String
  Dim sDiff As String
  Dim sRatio As String
  Dim sMsgBoxText As String
  Dim sMsgBoxTitle As String
  
  Dim fintResponse As Integer
  Dim iMsgBoxResp As Integer
  Dim iResponse As Integer
  
  Dim lRetVal As Long
  Dim lDifference As Long
  
  Dim sngRatio As Single
  
  Dim dBufferSize As Double
  Dim dFreeSpace As Double
 
' ------------------------------------------------------------
' Test the drive/path where we will be compacting the
' database and making a backup copy.
' ------------------------------------------------------------
StartOver:
  sDBPath = ExtractPath(sDatabase)        ' save path to database
  sDBName = UCase(ExtractFileName(sDatabase))  ' Save name of the database
  lStartSize = FileLen(sDatabase)         ' get database size before compacting
  dBufferSize = lStartSize * 2            ' Determine the buffer size for compacting
  sCurPosition = CurDir                   ' remember where we are
  
' ------------------------------------------------------------
' See if this drive/path is a restricted area.  Do we have
' update authority?
' ------------------------------------------------------------
  If IsThisRestricted(sWorkDrive) Then
      sMsgBoxText = "This path is a restricted area.   "
      MsgBox sMsgBoxText, vbOKOnly, "Restricted Area"
      GoTo Normal_Exit
  End If
  
' ------------------------------------------------------------
' Get the amount of free space on the drive where the work
' directory is located.
' ------------------------------------------------------------
  GetDiskSpace sWorkDrive
  dFreeSpace = DiskInfo.TotalFreeSpace
  
' ------------------------------------------------------------
' Display a message if not enough free space
' If there is not at least 2 times the size of the database.
'     1.  Make a bakup copy of the database
'     2.  Creation of the new database
' if all is sucessful, these will be deleted.
' ------------------------------------------------------------
  If dBufferSize > dFreeSpace Then
      '
      sMsgBoxTitle = "Need More Free Space"
      
      sMsgBoxText = "There is not enough free space on drive "
      sMsgBoxText = sMsgBoxText & sWorkDrive & ".    " & vbCrLf & vbCrLf
      sMsgBoxText = sMsgBoxText & "Need at least " & Format(dBufferSize, "#,0") & " bytes of free "
      sMsgBoxText = sMsgBoxText & vbCrLf & "space to compact this database." & vbCrLf
      sMsgBoxText = sMsgBoxText & "    " & vbCrLf & "If you elect to free up more "
      sMsgBoxText = sMsgBoxText & "space, then this    " & vbCrLf & "application will "
      sMsgBoxText = sMsgBoxText & "terminate while you perform    " & vbCrLf
      sMsgBoxText = sMsgBoxText & "this action.    " & vbCrLf & "    " & vbCrLf
      sMsgBoxText = sMsgBoxText & "If you want to select another drive,    " & vbCrLf
      sMsgBoxText = sMsgBoxText & "click RETRY.    " & vbCrLf & "    "
      
      iMsgBoxResp = vbRetryCancel + vbQuestion + vbApplicationModal + vbDefaultButton1
      
      iResponse = MsgBox(sMsgBoxText, iMsgBoxResp, sMsgBoxTitle)
 
      Select Case iResponse
             Case vbRetry:  ' Open browse for folder dialog box
                  Find_Work_Drive              ' Select another drive/path
                  If Len(sWorkDrive) = 0 Then  ' if user did not select anything
                      GoTo Normal_Exit         '      then leave
                  Else
                      GoTo StartOver           ' Go back and test this new path
                  End If
                  
             Case vbCancel: GoTo Normal_Exit
     End Select
  End If
  
' ------------------------------------------------------------
' Initialize database variables
' ------------------------------------------------------------
  sNewFile = UCase(Left(sDBName, Len(sDBName) - 3) & "NEW")
  sBakFile = UCase(Left(sDBName, Len(sDBName) - 3) & "BAK")

' ------------------------------------------------------------
' Get rid of any of the old databases
' ------------------------------------------------------------
  DoEvents
  If ItemExist(sDBPath & sNewFile) Then Kill sDBPath & sNewFile
  If ItemExist(sDBPath & sBakFile) Then Kill sDBPath & sBakFile
  If ItemExist(sWorkDrive & sNewFile) Then Kill sWorkDrive & sNewFile
  If ItemExist(sWorkDrive & sBakFile) Then Kill sWorkDrive & sBakFile

' ------------------------------------------------------------
' Make sure we can open this database in exclusive mode
' ------------------------------------------------------------
  If Not Open_Exclusive Then GoTo Normal_Exit
      
' ------------------------------------------------------------
' copy the existing database to the same name with a "BAK"
' extention and copy it to the temp directory.  Verify it
' got there and then delete it from this directory.
' ------------------------------------------------------------
  DoEvents
  Screen.MousePointer = vbHourglass
  
  lRetVal = CopyFile(sDatabase, sWorkDrive & sBakFile, False)
  
  DoEvents
  Screen.MousePointer = vbNormal
  
  If Not ItemExist(sWorkDrive & sBakFile) Then
      sMsgBoxText = "Failed to make a backup copy "
      sMsgBoxText = sMsgBoxText & "of the database." & vbLf & "Try again."
      MsgBox sMsgBoxText, vbOKOnly, "Bad Database Backup"
      GoTo Normal_Exit
  End If
  
On Error GoTo ErrorHandler
' ------------------------------------------------------------
' Change to the database directory
' ------------------------------------------------------------
  ChDrive sDBPath
  ChDir sDBPath
  
' ------------------------------------------------------------
' Compact the database into the current name with an
' extention of "NEW"
' ------------------------------------------------------------
  Screen.MousePointer = vbHourglass
  
  DoEvents      ' Repair the database
  DBEngine.RepairDatabase sDBName
  
  DoEvents      ' Now compact the database
  DBEngine.CompactDatabase sDBName, sWorkDrive & sNewFile
  
  DoEvents
  Screen.MousePointer = vbNormal
 
' ------------------------------------------------------------
' If the new database does not exist, then the compression
' was a failure.  See if the user wants to attempt a repair.
' If so, Do a repair on the database.  Go back to the top of
' this routine and start over.
' ------------------------------------------------------------
  If Not ItemExist(sWorkDrive & sNewFile) Then
      sMsgBoxTitle = "Error Compacting Database"
      
      sMsgBoxText = "ERR: " & CStr(Err) & vbLf & Err.Description & vbLf & vbLf
      sMsgBoxText = sMsgBoxText & "==> " & sDatabase & vbLf & vbLf
      sMsgBoxText = sMsgBoxText & "Do you want to try to repair the database?  "
      
      iMsgBoxResp = vbYesNo + vbQuestion + vbApplicationModal + vbDefaultButton1
      
      iResponse = MsgBox(sMsgBoxText, iMsgBoxResp, sMsgBoxTitle)
 
      Select Case iResponse
             Case vbYes:  ' repair the database again and start over
                  If ItemExist(sDBPath & sNewFile) Then Kill sDBPath & sNewFile
                  If ItemExist(sDBPath & sBakFile) Then Kill sDBPath & sBakFile
                  If ItemExist(sWorkDrive & sNewFile) Then Kill sWorkDrive & sNewFile
                  If ItemExist(sWorkDrive & sBakFile) Then Kill sWorkDrive & sBakFile
                  DoEvents

                  DBEngine.RepairDatabase sDBName
                  GoTo StartOver
                  
             Case vbNo: GoTo Normal_Exit
      End Select
  End If
   
' ------------------------------------------------------------
' Delete the original database because we successfully
' completed the previous steps.
' ------------------------------------------------------------
  If ItemExist(sDatabase) Then
      Kill sDatabase
  End If

' ------------------------------------------------------------
' move the new database to the original
' ------------------------------------------------------------
  DoEvents
  Screen.MousePointer = vbHourglass
  
  lRetVal = CopyFile(sWorkDrive & sNewFile, sDatabase, False)
  
  DoEvents
  Screen.MousePointer = vbNormal
  
  If Not ItemExist(sDatabase) Then
      sMsgBoxText = "Failed to replace database.  Contact support.   "
      MsgBox sMsgBoxText, vbOKOnly, "Bad Database Replace"
      GoTo Normal_Exit
  End If
  
' ------------------------------------------------------------
' Get rid of any of the old databases
' ------------------------------------------------------------
  DoEvents
  If ItemExist(sDBPath & sNewFile) Then Kill sDBPath & sNewFile
  If ItemExist(sDBPath & sBakFile) Then Kill sDBPath & sBakFile
  If ItemExist(sWorkDrive & sNewFile) Then Kill sWorkDrive & sNewFile
  If ItemExist(sWorkDrive & sBakFile) Then Kill sWorkDrive & sBakFile
  DoEvents

' ------------------------------------------------------------
' If we had to create the temp directory, then remove it.
' ------------------------------------------------------------
  Screen.MousePointer = vbHourglass
  If bUsedTempDir Then
      If ItemExist(sWorkDrive) Then
          DelTree32 sWorkDrive
      End If
  End If
  Screen.MousePointer = vbNormal
  
' ------------------------------------------------------------
' get the database size after compacting
' ------------------------------------------------------------
  lEndSize = FileLen(sDatabase)
  
' ----------------------------------------------------
' calculate the difference in size and compute the
' compression ratio
' ----------------------------------------------------
  lDifference = lStartSize - lEndSize
  sngRatio = CSng(((lStartSize - lEndSize) / lStartSize))

' ----------------------------------------------------
' Format the database results
' ----------------------------------------------------
  sStartSize = Format(Format(lStartSize, "#,0"), "!@@@@@@@@@@@")
  sEndSize = Format(Format(lEndSize, "#,0"), "!@@@@@@@@@@@")
  sDiff = Format(Format(lDifference, "#,0"), "!@@@@@@@@@@@")
  sRatio = Format(Format(sngRatio, "Percent"), "!@@@@@@@@@@@@")
  
' ---------------------------------------------------
' Display a message showing how the database
' being compressed saved x amount of space.
' ---------------------------------------------------
  DoEvents
  sMsgBoxText = vbLf & sDatabase & vbTab & vbTab & vbLf & vbLf
  sMsgBoxText = sMsgBoxText & "Original File Size" & vbTab & vbTab & sStartSize & " Bytes" & Space(5) & vbLf
  sMsgBoxText = sMsgBoxText & "Compressed File Size" & vbTab & sEndSize & " Bytes" & Space(5) & vbLf
  sMsgBoxText = sMsgBoxText & "File Space Freed Up" & vbTab & sDiff & " Bytes" & Space(5) & vbLf
  sMsgBoxText = sMsgBoxText & "Compression Percentage" & vbTab & sRatio
  
  iMsgBoxResp = vbOKOnly + vbInformation + vbApplicationModal + vbDefaultButton1
      
  MsgBox sMsgBoxText, iMsgBoxResp, "Database Maintenance Info"
  
  
Normal_Exit:
  StopTheProgram
  Exit Sub


' =========================================================
'      E R R O R   P R O C E S S I N G   S E C T I O N
' =========================================================

ErrorHandler:
' ------------------------------------------------------------
' Display a message the operation was a failure
' ------------------------------------------------------------
  DoEvents
  Screen.MousePointer = vbNormal
  
  sMsgBoxText = "ERR: " & CStr(Err) & vbLf & Err.Description & vbLf & vbLf
  sMsgBoxText = sMsgBoxText & "DB:  " & sDatabase & vbLf & vbLf & "Contact support personnel."
  iMsgBoxResp = vbOKOnly + vbCritical + vbApplicationModal + vbDefaultButton1
  MsgBox sMsgBoxText, iMsgBoxResp, "Error Compacting Database"
  StopTheProgram
  
End Sub
Private Sub Find_Work_Drive()

' ------------------------------------------------------------
' Display a list of all available drives
' ------------------------------------------------------------
  bUsedTempDir = False
  sWorkDrive = ""
  sWorkDrive = BrowseForFolder         ' Select another drive/path

' ------------------------------------------------------------
' If this is the root directory of a drive then create a
' temporary subdirectory to work in.  Never do your work
' in the root level.  Too easy to corrupt the directory
' structure.
' ------------------------------------------------------------
  If Len(Trim(sWorkDrive)) = 3 Then
      sWorkDrive = sWorkDrive & TEMP_DIR
      MkDir sWorkDrive
      bUsedTempDir = True
  End If
  
End Sub

Public Function GetDiskSpace(sDriveLtr As String) As Boolean

' -----------------------------------------------------------------
' Written by Kenneth Ives          kenaso@home.com
'
' returns data about the selected drive.  Passed is the
' drive letter; the other variables are filled in here.
'
' Syntax:  GetDiskSpace "A:\"
' -----------------------------------------------------------------
  
  On Error GoTo GetDiskSpace_Errors
' -----------------------------------------------------------------
' Define local variables
' -----------------------------------------------------------------
  Dim lRetVal As Long
  Dim cSpaceAvailable As Currency
  Dim cSpaceTotal As Currency
  Dim cSpaceFree As Currency
  Dim cSpaceUsed As Currency
  Dim OpSys As New OperSysInfo
  Dim bFat16 As Boolean
  Dim bFat32 As Boolean
  Dim sTmpLetter As String
  
' -----------------------------------------------------------------
' Initialize local variables
' -----------------------------------------------------------------
  cSpaceTotal = 0
  cSpaceFree = 0
  cSpaceUsed = 0
  cSpaceAvailable = 0
  bFat16 = False
  bFat32 = False
  
' -------------------------------------------------------------
' Prepare the drive letter
' -------------------------------------------------------------
  sTmpLetter = Left(sDriveLtr, 1)
  sTmpLetter = sTmpLetter & ":\"
  
' -------------------------------------------------------------
' Determine the operating system
'  OpSys.MajorVersion    4
'  OpSys.MinorVersion    10
'  OpSys.BuildNumber     1998
'  OpSys.PlatformID      1
'  OpSys.CSDVersion
'  OpSys.Platform        Windows 98
'  OpSys.Version         Windows 98 v4.10, Build 1998
'  OpSys.IsWinNT         False
'  OpSys.IsWin95         False
'  OpSys.IsWin98         True
' -------------------------------------------------------------
  If OpSys.IsWinNT Then
      bFat16 = True
  ElseIf OpSys.IsWin95 Then
      If OpSys.BuildNumber = 950 Then
          bFat16 = True     ' Standard Windows 95
      ElseIf OpSys.BuildNumber = 1111 Then
          bFat32 = True     ' OSR2 of Windows 95
      End If
  ElseIf OpSys.IsWin98 Then
      If OpSys.BuildNumber = 1998 Or OpSys.BuildNumber = 2222 Then
          bFat32 = True
      End If
  End If

' -----------------------------------------------------------------
' Make the API call to get the numeric data about a drive
' Gather information about FAT32 drive
' -----------------------------------------------------------------
  If bFat32 Then
      lRetVal = GetDiskFreeSpaceEx(sTmpLetter, cSpaceAvailable, cSpaceTotal, cSpaceFree)
      
      If lRetVal > 0 Then
          ' show the results, multiplying the returned
          ' value by 10000 to adjust for the 4 decimal
          ' places that the currency data type returns.
          cSpaceTotal = cSpaceTotal * 10000
          cSpaceFree = cSpaceFree * 10000
          cSpaceAvailable = cSpaceAvailable * 10000
          cSpaceUsed = (cSpaceTotal - cSpaceFree)
      End If
  End If
       
' -----------------------------------------------------------------
' Make the API call to get the numeric data about a drive
' Gather information about FAT32 drive
' -----------------------------------------------------------------
  With DiskInfo
   
       If bFat16 Then
           ' Get the drive space allocations
           lRetVal = GetDiskFreeSpace(sTmpLetter, .SectorsPerCluster, .BytesPerSector, _
                                      .FreeClusters, .TotalClusters)
             
           ' Collect space information on the drive
           If lRetVal > 0 Then
               .UsedClusters = (.TotalClusters - .FreeClusters)
               .TotalSectors = (.TotalClusters * .SectorsPerCluster)
               .FreeSectors = (.TotalSectors - (.UsedClusters * .SectorsPerCluster))
               .UsedSectors = (.TotalSectors - .FreeSectors)
               cSpaceTotal = .BytesPerSector * .SectorsPerCluster * .TotalClusters
               cSpaceFree = .BytesPerSector * .SectorsPerCluster * .FreeClusters
               cSpaceUsed = cSpaceTotal - cSpaceFree
           End If
       End If
       
       .TotalDiskSize = CDbl(cSpaceTotal)
       .TotalAvailSpace = CDbl(cSpaceAvailable)
       .TotalUsedSpace = CDbl(cSpaceUsed)
       .TotalFreeSpace = CDbl(cSpaceFree)
       
       If cSpaceTotal > 0 Then
            ' Format the display sizes
           .SpaceTotal = ReFormatSize(cSpaceTotal)
           .SpaceFree = ReFormatSize(cSpaceFree)
           .SpaceAvailable = ReFormatSize(cSpaceAvailable)
           .SpaceUsed = ReFormatSize(cSpaceUsed)
           .SpaceFreePcnt = Format(Format(.TotalFreeSpace / .TotalDiskSize, "Percent"), "@@@@@@@")
           .SpaceUsedPcnt = Format(Format(.TotalUsedSpace / .TotalDiskSize, "Percent"), "@@@@@@@")
           .SpaceAvailablePcnt = Format(Format(.TotalFreeSpace / .TotalDiskSize, "Percent"), "@@@@@@@")
       Else
           .TotalDiskSize = 0
           .TotalAvailSpace = 0
           .TotalUsedSpace = 0
           .TotalFreeSpace = 0
           .SpaceTotal = "0"
           .SpaceFree = "0"
           .SpaceUsed = "0"
           .SpaceAvailable = "0"
           .SpaceFreePcnt = ""
           .SpaceUsedPcnt = ""
           .SpaceAvailablePcnt = ""
           .UsedClusters = 0
           .TotalSectors = 0
           .FreeSectors = 0
           .UsedSectors = 0
       End If
         
   End With
  
' -----------------------------------------------------------------
' If we got to here, things went Okay
' -----------------------------------------------------------------
  GetDiskSpace = True
  Exit Function
  

GetDiskSpace_Errors:
' -----------------------------------------------------------------
' If we got to here, Something went wrong
' -----------------------------------------------------------------
  GetDiskSpace = False

End Function


Private Function ExtractFileExt(sFileName As String) As String

' -------------------------------------------------------------------
' Written by Kenneth Ives     kenaso@home.com
'
' This routine takes a filename and searches backwards for a "."
' and returns the extension.
'
' Syntax:     ExtractFileExt "C:\Program Files\MyFile.doc"
'
' Returns:    "doc"
' -------------------------------------------------------------------

' -------------------------------------------------------------------
' Define local variables
' -------------------------------------------------------------------
  Dim iPos As Integer
  Dim sTmpStr As String

' -------------------------------------------------------------------
' Initialize local variables
' -------------------------------------------------------------------
  sTmpStr = ""
  sFileName = Trim(sFileName)

' -------------------------------------------------------------------
' Parse the path string starting from the last character backwards
' until you find the first "."
' -------------------------------------------------------------------
  iPos = InStr(1, sFileName, ".")
  If iPos > 0 Then
      sTmpStr = Mid(sFileName, iPos + 1)
  End If

' -------------------------------------------------------------------
' Return the extension, if any
' -------------------------------------------------------------------
  ExtractFileExt = sTmpStr

End Function
Public Function ExtractFileName(sFileName As String) As String

' -------------------------------------------------------------------
' Written by Kenneth Ives     kenaso@home.com
'
' This returns just a file name from a full/partial path.
'
' Syntax:     ExtractFileName "C:\Program Files\MyFile.doc"
'
' Returns:    "MyFile.doc"
' -------------------------------------------------------------------

' -------------------------------------------------------------------
' Define local variables
' -------------------------------------------------------------------
  Dim i As Long
  Dim sTmpStr As String
  Dim sTmpFile As String
  
' -------------------------------------------------------------------
' Initialize local variables
' -------------------------------------------------------------------
  sTmpFile = sFileName

  If InStr(1, sTmpFile, "\") > 0 Then GoTo Start_Parsing
  If InStr(1, sTmpFile, ":") > 0 Then GoTo Start_Parsing
  sTmpStr = sTmpFile
  GoTo Normal_Exit
  
Start_Parsing:
' -------------------------------------------------------------------
' Parse the path string starting from the last character backwards
' until you find the first "\"
' -------------------------------------------------------------------
  sTmpStr = ""
  
  If InStr(1, sTmpFile, ":") > 0 Then
      sTmpFile = Mid(sTmpFile, (InStr(1, sTmpFile, ":")) + 1)
  End If
  
  If InStr(1, sTmpFile, "\") > 0 Then
      For i = Len(sTmpFile) To 1 Step -1
          If Mid(sTmpFile, i, 1) = "\" Then
              sTmpStr = Mid(sTmpFile, i + 1)
              Exit For
          End If
      Next
  Else
      sTmpStr = sTmpFile
  End If
  
Normal_Exit:
' -------------------------------------------------------------------
' Return the filename without the path
' -------------------------------------------------------------------
  ExtractFileName = sTmpStr

End Function
Public Function ExtractPath(sFileName As String) As String

' -------------------------------------------------------------------
' Written by Kenneth Ives     kenaso@home.com
'
' This returns just a path name from a full path.
'
' Syntax:     ExtractPath "C:\Program Files\MyFile.doc"
'
' Returns:    "C:\Program Files"
' -------------------------------------------------------------------

' -------------------------------------------------------------------
' Define local variables
' -------------------------------------------------------------------
  Dim i As Long
  Dim sTmpStr As String

' -------------------------------------------------------------------
' Initialize local variables
' -------------------------------------------------------------------
  sTmpStr = ""

' -------------------------------------------------------------------
' Parse the path string starting from the last character backwards
' until you find the first "\"
' -------------------------------------------------------------------
  For i = Len(sFileName) To 1 Step -1
       If Mid(sFileName, i, 1) = "\" Then
          '
          ' Return the path without the file name
          sTmpStr = Mid(sFileName, 1, i)
          Exit For
       End If
  Next
    
' -------------------------------------------------------------------
' Return the path without the file name
' -------------------------------------------------------------------
  ExtractPath = sTmpStr

End Function

Private Function ReFormatSize(cSize As Variant) As String
   
' -------------------------------------------------------------
' Define local variables
' -------------------------------------------------------------
  Dim sRet As String
  Dim cTestSize As Currency
  Dim cMB_Size As Currency
  Dim cKB_Size As Currency
  Dim cGB_Size As Currency
   
  Const KB& = 1024
  Const MB& = KB& * KB&
  Const GB& = MB& * KB&
   
' -------------------------------------------------------------
' Initialize local variables
' -------------------------------------------------------------
  cTestSize = (cSize / KB&)
  cKB_Size = cTestSize
  cMB_Size = (cSize / MB&)
  cGB_Size = (cSize / GB&)
   
' -------------------------------------------------------------
' If less than 1kb then set test value to 1
' -------------------------------------------------------------
  If cSize < KB& Then cTestSize = 1
    
' -------------------------------------------------------------
' Format the abbreviated size output
' -------------------------------------------------------------
  Select Case cTestSize
         Case Is = 1: sRet = Format(cTestSize, "0") & "KB"
         Case Is < 10: sRet = Format(cKB_Size, "0.00") & "KB"
         Case Is < 100: sRet = Format(cKB_Size, "0.0") & "KB"
         Case Is < 1000: sRet = Format(cKB_Size, "0") & "KB"
         Case Is < 10000: sRet = Format(cMB_Size, "0.00") & "MB"
         Case Is < 100000: sRet = Format(cMB_Size, "0.0") & "MB"
         Case Is < 1000000: sRet = Format(cMB_Size, "0") & "MB"
         Case Is < 10000000: sRet = Format(cGB_Size, "0.00") & "GB"
  End Select
   
' -------------------------------------------------------------
' Reformat the return string
' -------------------------------------------------------------
  sRet = sRet & " (" & Format(cSize, "#,0") & " bytes)"
  ReFormatSize = sRet
   
End Function

Public Function GetTempDir() As String

' ----------------------------------------------------------------
'  Returns the path to the temp directory.
' ----------------------------------------------------------------

' ----------------------------------------------------------------
' Define local variables
' ----------------------------------------------------------------
  Dim sBuffer As String
  Dim lBufferLen As Long
  Dim lRetVal As Long
  
' ----------------------------------------------------------------
' Initialize local variables
' ----------------------------------------------------------------
  lBufferLen = 256
  sBuffer = Space(lBufferLen)
  
' ----------------------------------------------------------------
' Get the path to the temp file
' ----------------------------------------------------------------
  lRetVal = GetTempPath(lBufferLen, sBuffer)
  
' ----------------------------------------------------------------
' cleanup the returned data
' ----------------------------------------------------------------
  If lRetVal Then
      ' if non-zero returned, remove leading and
      ' trailing blanks, then remove chr(0) if
      ' one exist
      sBuffer = Trim(sBuffer)
      sBuffer = IIf(InStr(1, sBuffer, Chr(0)) = 0, sBuffer, Left(sBuffer, Len(sBuffer) - 1))
  Else
      ' if not found, create a temp directory
      MkDir TEMP_DIR
      sBuffer = TEMP_DIR
  End If
  
' ----------------------------------------------------------------
' Add a backslash, if missing, prior to returning the data
' ----------------------------------------------------------------
  If Right(sBuffer, 1) <> "\" Then
      sBuffer = sBuffer & "\"
  End If
  
' ----------------------------------------------------------------
' Return the information
' ----------------------------------------------------------------
  GetTempDir = sBuffer

End Function

Public Function Open_Exclusive() As Boolean

' -------------------------------------------------------------------
' Written by Kenneth Ives     kenaso@home.com
'
' See if we can get an exclusive hold of this database.  It is
' mandatory if we are going to compact it.
' -------------------------------------------------------------------
  
  On Error GoTo Open_Exclusive_Errors
' ---------------------------------------------------------
' Define local variables
' ---------------------------------------------------------
  Dim WS_Tmp As Workspace
  Dim DB_Tmp As Database
  
' ---------------------------------------------------------
' Initialize local variables
' ---------------------------------------------------------
  Set WS_Tmp = Nothing
  Set DB_Tmp = Nothing
  
' ---------------------------------------------------------
' Open the MS Access Database in exclusive mode
' ---------------------------------------------------------
  Set WS_Tmp = CreateWorkspace("", "admin", "", dbUseJet)
  Set DB_Tmp = WS_Tmp.OpenDatabase(sDatabase, True, False)
  DB_Tmp.Close
  WS_Tmp.Close
  Open_Exclusive = True
  
Normal_Exit:
' ---------------------------------------------------------
' close everything and leave
' ---------------------------------------------------------
  Set DB_Tmp = Nothing
  Set WS_Tmp = Nothing
  On Error GoTo 0
  Exit Function
  
  
Open_Exclusive_Errors:
' ---------------------------------------------------------
' Could not open in exclusive mode
' ---------------------------------------------------------
  MsgBox "Database has been opened by someone else.", _
         vbOKOnly, "Cannot continue"
  
  Open_Exclusive = False
  GoTo Normal_Exit
  
End Function

Public Sub Main()

' ---------------------------------------------------------
' Set up the path where all of the mail processing
' will take place.
' ---------------------------------------------------------
  ChDrive App.Path
  ChDir App.Path
      
' ---------------------------------------------------------
' See if there is another instance of this program running
' ---------------------------------------------------------
  IsAnotherInstance "compmdb"
  
' ---------------------------------------------------------
' Get the work area drive
' ---------------------------------------------------------
  frmDBMaint.Hide
  Find_Work_Drive
  
' ---------------------------------------------------------
' If we have a work drive selected then start the
' maintenance
' ---------------------------------------------------------
  If Len(Trim(sWorkDrive)) = 0 Then
      MsgBox "Application terminated because no work area was selected.", _
             vbOKOnly + vbExclamation, "No Work Area"
      StopTheProgram
  Else
      frmDBMaint.Reset_frmDBMaint
  End If
  
End Sub
Public Sub StopTheProgram()

' ---------------------------------------------------------
' Upload all forms from memory and termiante this
' application
' ---------------------------------------------------------
  Unload_All_Forms
  End
  
End Sub
Public Sub Unload_All_Forms()

' ---------------------------------------------------------
' Written by Kenneth Ives          kenaso@home.com
'
' Unload all forms before terminating an application
' The calling module will call this routine and usually
' executes END when it returns.
' ---------------------------------------------------------

' ---------------------------------------------------------
' Define local variables
' ---------------------------------------------------------
  Dim frm As Form
  
' ---------------------------------------------------------
' If the form.name property is not the same as the form
' calling this routine, then unload it and free up memory.
' ---------------------------------------------------------
  For Each frm In Forms
      frm.Hide
      Unload frm
      Set frm = Nothing
  Next
  
End Sub


