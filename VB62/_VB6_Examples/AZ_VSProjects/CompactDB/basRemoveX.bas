Attribute VB_Name = "basRemoveX"
Option Explicit

' ---------------------------------------------------------
' Written by Randy Birch      http://www.mvps.org/vbnet/
'
' Remove the "X" from the window and menu
'
' A VB developer may find themselves developing an application
' who 's integrity is crucial, and therefore must prevent the
' user from accidentally terminating the application during its
' life, while still displaying the system menu.  And while
' Visual Basic does provide two places to cancel an impending
' close - the QueryUnload and Unload subs -  such a sensitive
' application may need to totally prevent even activation of
' the shutdown.
'
' Although it is not possible to simply disable the Close button
' while the Close system menu option is present, just a few
' lines of API code will remove the system menu Close option
' and in doing so permanently disable the titlebar close button.
' ---------------------------------------------------------------
  Private Const MF_BYPOSITION = &H400
  Private Const MF_REMOVE = &H1000

  Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
  Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
  Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, _
                  ByVal bRevert As Long) As Long
                  
  Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, _
                  ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Sub RemoveX(frm As Form)

' ---------------------------------------------------------------
' For completeness, you may want to confirm that the
' menuItemCount matches the value you expect before performing
' the removal.  For example, on a normal form with a full
' system menu, menuItemCount will return seven.
'
' But what if your application is an MDI app and you want to
' disable the close button on the parent?  Just pass
' MDIForm1.hwnd as the form hwnd parameter in the
' GetSystemMenu() call.
' ---------------------------------------------------------------
   
' ---------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------
  Dim hMenu As Long
  Dim menuItemCount As Long

' ---------------------------------------------------------------
' Obtain the handle to the form's system menu
' ---------------------------------------------------------------
  hMenu = GetSystemMenu(frm.hwnd, 0)
  
  If hMenu Then
      
     ' Obtain the number of items in the menu
      menuItemCount = GetMenuItemCount(hMenu)
    
     ' Remove the system menu Close menu item.
     ' The menu item is 0-based, so the last
     ' item on the menu is menuItemCount - 1
      Call RemoveMenu(hMenu, menuItemCount - 1, MF_REMOVE Or MF_BYPOSITION)
   
     ' Remove the system menu separator line
      Call RemoveMenu(hMenu, menuItemCount - 2, MF_REMOVE Or MF_BYPOSITION)
    
     ' Force a redraw of the menu. This
     ' refreshes the titlebar, dimming the X
      Call DrawMenuBar(frm.hwnd)

   End If
   
End Sub

