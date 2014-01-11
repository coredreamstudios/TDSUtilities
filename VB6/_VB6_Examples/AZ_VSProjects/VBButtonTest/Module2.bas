Attribute VB_Name = "DragWindowNoTitle"
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
Public Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Public Type POINTAPI
   x As Long
   y As Long
End Type

Public Const COLOR_ACTIVECAPTION = 2
Public Const SM_CXDLGFRAME = 7
Public Const SM_CYDLGFRAME = 8

Public Declare Function GetWindowRect Lib "user32" _
    (ByVal hWnd As Long, lpRect As RECT) As Long
    
Public Declare Function GetSysColor Lib "user32" _
    (ByVal nIndex As Long) As Long
    
Public Declare Function GetSystemMetrics Lib "user32" _
    (ByVal nIndex As Long) As Long

Public Declare Function DrawFocusRect Lib "user32" _
    (ByVal hdc As Long, lpRect As RECT) As Long
    
Public Declare Function ClientToScreen Lib "user32" _
    (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
    
Public Declare Function GetDC Lib "user32" _
    (ByVal hWnd As Long) As Long
    
Public Declare Function ReleaseDC Lib "user32" _
    (ByVal hWnd As Long, ByVal hdc As Long) As Long
'--end block--'


