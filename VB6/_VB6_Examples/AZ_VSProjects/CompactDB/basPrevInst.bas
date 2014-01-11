Attribute VB_Name = "basPrevInst"
Option Explicit

' ---------------------------------------------------------
' This BAS module was found at the VBNet web pages.
' http://www.mvps.org/vbnet/index.html
'
' modified by Kenneth Ives     kenaso@home.com
' ---------------------------------------------------------

' ---------------------------------------------------------
' required for the RestorePreviousInstance routine
' ---------------------------------------------------------
  Private Const SW_SHOWMINIMIZED = 2
  Private Const SW_SHOWNORMAL = 1
  Private Const SW_SHOWNOACTIVATE = 4
  Private Const SW_RESTORE = 9

  Private Type RECT
     Left As Long
     Top As Long
     Right As Long
     Bottom As Long
  End Type

  Private Type POINTAPI
     x As Long
     Y As Long
  End Type

  Private Type WINDOWPLACEMENT
     Length As Long
     flags As Long
     showCmd As Long
     ptMinPosition As POINTAPI
     ptMaxPosition As POINTAPI
     rcNormalPosition As RECT
  End Type

' ---------------------------------------------------------
' Declares required for verifying a previous instance
' of program executiion
' ---------------------------------------------------------
  Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
  Private Declare Function SetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
  Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
  Private Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
  Private Declare Function SendMessageArray Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

' ---------------------------------------------------------
' required just for debugging puproses
' ---------------------------------------------------------
  Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
  Private Const LB_SETTABSTOPS = &H192

Public Sub IsAnotherInstance(SApplName As String)

' ---------------------------------------------------------
' Call this module from the Sub Main()
' ---------------------------------------------------------

' ---------------------------------------------------------
' Define local variable
' ---------------------------------------------------------
  Dim savetitle As String
  
' ---------------------------------------------------------
' Check for a previous instance of a program running
' ---------------------------------------------------------
  If App.PrevInstance Then
      '
      ' change the new instance title to prevent it
      ' from being located instead of the original
      ' instance.  Note however that as this is in
      ' a BAS module and not the form load sub,
      ' change "pgm_name" to the name of the application
      ' that you do not a dupliate instance of.
      savetitle = App.Title
      App.Title = SApplName ' name of executable here(w/o .exe)
  
      '-------------------------------------------------------------------
      ' some debug stuff - remove for live use
      'MsgBox "about to re-activate the original instance of " & savetitle
      '-------------------------------------------------------------------
   
      RestorePreviousInstance savetitle
      End

  End If

End Sub


Private Sub RestorePreviousInstance(prevTitle As String)

' ---------------------------------------------------------
' Define local variable
' ---------------------------------------------------------
  Dim lRetVal As Long
  Dim hPrevWin As Long
  Dim lpString As String
  Dim currWinP As WINDOWPLACEMENT
     
' ---------------------------------------------------------
' VB3 & VB4 use class name "ThunderRTForm"
' VB5 uses class name "ThunderRT5Form"
' VB6 uses class name "ThunderRT6FormDC"
'
' Including the class name for the compiled EXE class
' prevents the routine from finding and attempting
' to activate the project form of the same name.
' ---------------------------------------------------------
  hPrevWin = FindWindow("ThunderRT6FormDC", prevTitle)
   
  DoEvents
   
' ---------------------------------------------------------
' If found
' ---------------------------------------------------------
  If hPrevWin > 0 Then
  
      '-------------------------------------------------------------------
      ' some debug stuff - remove for live use
      ' this is just to verify that the title
      ' found was the title intended.
      '
      ' lpString = Space(256)
      ' lRetVal = GetWindowText(hPrevWin, lpString, 256)
      ' MsgBox "GetWindowText verifies the title as - " & Left(lpString, s)
      '-------------------------------------------------------------------
     
      ' get the current window state of the previous instance
      currWinP.Length = Len(currWinP)
      lRetVal = GetWindowPlacement(hPrevWin, currWinP)
               
      ' if the currWinP.showCmd member indicates that
      ' the window is currently minimized, it needs
      ' to be restored, so ...
      If currWinP.showCmd = SW_SHOWMINIMIZED Then
          currWinP.Length = Len(currWinP)
          currWinP.flags = 0&
          currWinP.showCmd = SW_SHOWNORMAL
          lRetVal = SetWindowPlacement(hPrevWin, currWinP)
      End If
       
      ' bring the window to the front and make
      ' the active window.  Without this, it
      ' may remain behind other windows.
      lRetVal = SetForegroundWindow(hPrevWin)
      DoEvents
    
 ' -------------------------------------------------------------------
 ' More debug stuff.  Comment out the ELSE condition for live use
 '
 ' Else
 '     MsgBox "FindWindow failed on " & prevTitle
 ' -------------------------------------------------------------------
 
  End If
   
End Sub

