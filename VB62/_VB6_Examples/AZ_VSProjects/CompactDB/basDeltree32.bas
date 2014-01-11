Attribute VB_Name = "basDelTree32"
Option Explicit

' -------------------------------------------------------------
'                  DelTree32 routine
'
' This code was originally written by Rod Stephens at
' http://www.vb-helper.com/HowTo/
'
' Modified and documented by Kenneth Ives   kenaso@home.com
'
' Deletes an entire directory tree regardless of the
' attributes of the directory or it's contents.  It will
' handle long and short filenames.
'
' BE VERY CAREFUL.  There are no edit checks for sensitive
' drives or data.  There are no warnings built in, you will
' have to add them yourself.  I will not be responsible for
' your lack safe testing.  Remember the five P's before you
' begin any venture.
'
'     "Proper Planning Prevents Poor Performance"
' -------------------------------------------------------------

' -------------------------------------------------------------
' This is usually tied to a CANCEL button on a form.
' i.e.  cmdCancel_Click
' -------------------------------------------------------------
  Public STOP_PRESSED As Boolean

Public Function DelTree32(sPathToDel As String) As Boolean

' -------------------------------------------------------------
'                  DelTree32 routine
'
' This code was originally written by Rod Stephens at
' http://www.vb-helper.com/HowTo/
'
' Modified and documented by Kenneth Ives      kenaso@home.com
'
' Deletes an entire directory tree regardless of the
' attributes of the directory or it's contents.  It will
' handle long and short filenames.
'
' Syntax:
'
'      Delete everything on drive A:
'
'           DelTree32 "A:"
'
'  Delete everything starting at "c:\Dir_1\subDir_1" and below
'  to include "\subDir_1"
'
'           DelTree32 "c:\Dir_1\subDir_1"
' -------------------------------------------------------------

' -------------------------------------------------------------
' Define local variables
' -------------------------------------------------------------
  Dim file_name As String
  Dim files As Collection
  Dim i As Integer
  
' -------------------------------------------------------------
' Initialize variables
' -------------------------------------------------------------
  Set files = New Collection
  
' -------------------------------------------------------------
' if there is a trailing backslash then remove it.
' -------------------------------------------------------------
  If Right(sPathToDel, 1) = "\" Then
      sPathToDel = Left(sPathToDel, Len(sPathToDel) - 1)
  End If
  
' -------------------------------------------------------------
' Get a list of path & filenames from this folder on down.
' -------------------------------------------------------------
  file_name = Dir(sPathToDel & "\*.*", vbNormal Or vbReadOnly Or vbHidden Or _
                                       vbSystem Or vbArchive Or vbDirectory)
    
' -------------------------------------------------------------
' Loop thru the directory structure and add the
' path & filename to the collection.
' -------------------------------------------------------------
  Do While Len(file_name) > 0
      
      If (file_name <> "..") And (file_name <> ".") Then
          ' add to the collection
          files.Add sPathToDel & "\" & file_name
      End If
      
      file_name = Dir()   ' is there anything left?
      DoEvents            ' allow other processes to happen
            
      ' This is usually tied to a CANCEL button on a form
      ' i.e.  cmdCancel_Click
      If STOP_PRESSED Then
          DelTree32 = False
          Exit Function
      End If
  Loop

' -------------------------------------------------------------
' Loop thru the collection and delete the files
' and directories
' -------------------------------------------------------------
  For i = 1 To files.Count
      
      ' move the path & filename to a variable
      file_name = files(i)
      
      ' See if it is a directory.
      If GetAttr(file_name) And vbDirectory Then
          ' This is a directory.
          ' Delete everything in it.
          DelTree32 file_name
      
      ' This is a file.  Delete it.
      Else
          SetAttr file_name, vbNormal  ' reset the attributes to normal
          Kill file_name               ' delete the file
      End If
      
      DoEvents                         ' allow other processes to happen
      
      ' This is usually tied to a CANCEL button on a form
      ' i.e.  cmdCancel_Click
      If STOP_PRESSED Then
          DelTree32 = False
          Exit Function
      End If
  Next

' -------------------------------------------------------------
' If this is the root directory, then leave.  Cannot delete
' the root directory.
' -------------------------------------------------------------
  If Len(sPathToDel) > 2 Then
      RmDir sPathToDel
  End If
  
  DelTree32 = True
  
End Function


Public Sub TEST2()

' -----------------------------------------------------------
' This test written by Kenneth Ives          kenaso@home.com
'
' Rename this routine to "Main" and press F5 to execute.
'
' Test deleting multiple folders and files with different
' attributes
' -----------------------------------------------------------
  DelTree32 "D:\Test_Lvl"
  End

End Sub

Sub TEST1()

' -----------------------------------------------------------
' This test written by Kenneth Ives          kenaso@home.com
'
' Rename this routine to "Main" and press F5 to execute.
'
' Test deleting multiple folders and files with different
' attributes
' -----------------------------------------------------------

' -----------------------------------------------------------
' Create the test directories
' -----------------------------------------------------------
  MkDir "D:\a_Test"
  MkDir "D:\a_Test\Level2"

' -----------------------------------------------------------
' Use the VB function FileCopy to copy the config.sys file
' into the lower level directory
' -----------------------------------------------------------
  FileCopy "c:\config.sys", "d:\a_test\Level2\config.sys"
  FileCopy "c:\config.sys", "d:\a_test\Level2\config.bak"

' -----------------------------------------------------------
' Use the VB function SetAttr to set the attributes of the
' directories and the files in the test area
' -----------------------------------------------------------
  SetAttr "d:\a_test\Level2", vbReadOnly
  SetAttr "d:\a_test\Level2\config.sys", vbReadOnly + vbHidden + vbSystem
  SetAttr "d:\a_test\Level2\config.bak", vbNormal

' -----------------------------------------------------------
' Use the VB command Stop to pause the execution here so
' the user can open Windows Explorer to view the properties
' of the directories and files.  Press F5 to continue.
' -----------------------------------------------------------
  Stop

' -----------------------------------------------------------
' Erase this directory completely and then terminate this
' test.
' -----------------------------------------------------------
  DelTree32 "D:\A_Test"
  End

End Sub
