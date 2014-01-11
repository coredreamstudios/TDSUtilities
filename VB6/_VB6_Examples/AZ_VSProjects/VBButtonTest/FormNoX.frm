VERSION 5.00
Begin VB.Form NoX 
   Caption         =   "Form3"
   ClientHeight    =   1455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4500
   LinkTopic       =   "Form3"
   ScaleHeight     =   1455
   ScaleWidth      =   4500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
End
Attribute VB_Name = "NoX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Add the following to a form - shown here is the code to disable the X called from the form's Load sub:

'--------------------------------------------------------------------------------
 
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2002 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const MF_BYPOSITION = &H400
Private Const MF_REMOVE = &H1000

Private Declare Function DrawMenuBar Lib "user32" _
      (ByVal hWnd As Long) As Long
      
Private Declare Function GetMenuItemCount Lib "user32" _
      (ByVal hMenu As Long) As Long
      
Private Declare Function GetSystemMenu Lib "user32" _
      (ByVal hWnd As Long, _
       ByVal bRevert As Long) As Long
       
Private Declare Function RemoveMenu Lib "user32" _
      (ByVal hMenu As Long, _
       ByVal nPosition As Long, _
       ByVal wFlags As Long) As Long

Private Sub Command1_Click()
    
    Unload Me
    
End Sub

Private Sub Form_Load()

   Dim hMenu As Long
   Dim menuItemCount As Long

  'Obtain the handle to the form's system menu
   hMenu = GetSystemMenu(Me.hWnd, 0)
  
   If hMenu Then
      
     'Obtain the number of items in the menu
      menuItemCount = GetMenuItemCount(hMenu)
    
     'Remove the system menu Close menu item.
     'The menu item is 0-based, so the last
     'item on the menu is menuItemCount - 1
      Call RemoveMenu(hMenu, menuItemCount - 1, _
                      MF_REMOVE Or MF_BYPOSITION)
   
     'Remove the system menu separator line
      Call RemoveMenu(hMenu, menuItemCount - 2, _
                      MF_REMOVE Or MF_BYPOSITION)
    
     'Force a redraw of the menu. This
     'refreshes the titlebar, dimming the X
      Call DrawMenuBar(Me.hWnd)

   End If
   
End Sub
'--end block--'
 


