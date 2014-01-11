Attribute VB_Name = "basCntrForm"
Option Explicit

' -----------------------------------------------------------------------
' Center Form Module
'
'   Purpose: This module was built because I was too lazy to have to
'            tell a sub-routine that the MDIChild had a parent.
'            This module should work in 16 & 32-bit environments.
'
'    Author: Jeffery S. Kofsky
'     Email: jeffk@cheney.net
'
'       Use:  Add this module to your project, then
'             in the Form_Load event add the following line:
'                CenterForm Me
'
'             That's it!  The center routine will check to see if 'Me'
'             is the MDIChild property is True, then center the form
'             in the available parent area,
'             ELSE it will center it in the available desktop area.
'
'  Comments:  I hope that works well for you.  Any comments can by sent
'             to my email address.
'
'    NOTICE:  This routine is provided as is.  It works well for me, but
'             use it at your own risk.  It is freely distributable, but
'             please leave all of these comments intact.
'
' Revisions:  Ver 1     Initial version.
'             Ver 1.1   Automatic detection of a MDIChild's parent.
'             Ver 1.2   Detection and adjustment for the Win95 TaskBar.
'                       Change to a regular module for easier of use.
' -----------------------------------------------------------------------
   
' ---------------------------------------------------------
' Type declaration
' ---------------------------------------------------------
  Private Type RECT
      Left     As Long
      Top      As Long
      Right    As Long
      Bottom   As Long
  End Type
   
' ---------------------------------------------------------
' API Declarations
' ---------------------------------------------------------
  Private Declare Function GetDesktopWindow Lib "user32" () As Long
  Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
  Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, _
                  lpRect As RECT) As Long
                  
  Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
                  (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
  
  Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
                  lpRect As RECT) As Long

' ---------------------------------------------------------
' return variable declaration
' ---------------------------------------------------------
  Private apiRetVal As Long

Public Sub CenterForm(frm As Form)
   
' ---------------------------------------------------------
' Syntax:    CenterForm "Form1"
' ---------------------------------------------------------

' ---------------------------------------------------------
' Define local variable
' ---------------------------------------------------------
  Dim ClientRect    As RECT           'Holds the area that the form is to be centered in
  Dim TaskBarRect   As RECT           'Holds the TaskBar area if in Win95
  Dim X             As Variant        'temp LeftPosition
  Dim Y             As Variant        'temp TopPosition

' ---------------------------------------------------------
' Check if the form is a MDIChild.
' ---------------------------------------------------------
  If frm.MDIChild Then
      '
      ' Center it in the MDIParent.
      GetClientRect GetParent(frm.hwnd), ClientRect
  Else  'Center it in the available desktop area.
      '
      ' Get the Desktop area
      Call GetClientRect(GetDesktopWindow(), ClientRect)
      '
      ' Check for the Task Bar.
      apiRetVal = FindWindow("Shell_TrayWnd", vbNullString)
      '
      ' If there is a taskbar, ie WIN95 then adjust the ClientRect.
      If apiRetVal Then
         Call GetWindowRect(apiRetVal, TaskBarRect)
         '
         If (TaskBarRect.Right - TaskBarRect.Left) > (TaskBarRect.Bottom - TaskBarRect.Top) Then
            '
            ' TaskBar at the Top of Screen.
            If TaskBarRect.Top <= 0 Then
               ClientRect.Top = ClientRect.Top + TaskBarRect.Bottom
            '
            ' TaskBar at the Bottom of Screen.
            Else
               ClientRect.Bottom = ClientRect.Bottom - (TaskBarRect.Bottom - TaskBarRect.Top)
            End If
         Else
            '
            ' TaskBar is on the Left side of the Screen.
            If TaskBarRect.Left <= 0 Then
               ClientRect.Left = ClientRect.Left + TaskBarRect.Right
            '
            ' TaskBar is on the Right side of the Screen.
            Else
               ClientRect.Right = ClientRect.Right - (TaskBarRect.Right - TaskBarRect.Left)
            End If
         End If   '[TaskBar on Top of Screen?]
      End If      '[if apiRetVal]
  End If

' ---------------------------------------------------------
' Center the Form
' ---------------------------------------------------------
  With frm
       X = (((ClientRect.Right - ClientRect.Left) * Screen.TwipsPerPixelX) - .Width) / 2
       Y = (((ClientRect.Bottom - ClientRect.Top) * Screen.TwipsPerPixelY) - .Height) / 2
       .Move X, Y
  End With

End Sub

