Attribute VB_Name = "DisplayModes"
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
Public Declare Function EnumDisplaySettings Lib "user32" _
    Alias "EnumDisplaySettingsA" _
   (ByVal lpszDeviceName As Long, _
    ByVal iModeNum As Long, _
    lpDevMode As Any) As Long
          
Public Declare Function GetDeviceCaps Lib "gdi32" _
   (ByVal hdc As Long, _
    ByVal nIndex As Long) As Long
   
Public Declare Function SendMessage Lib "user32" _
    Alias "SendMessageA" _
   (ByVal hWnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long
   
Public Const LVS_EX_FULLROWSELECT As Long = &H20
Public Const LVM_FIRST As Long = &H1000
Public Const LVM_SETEXTENDEDLISTVIEWSTYLE As Long = (LVM_FIRST + 54)

Public Const LOGPIXELSX As Long = 88
Public Const LOGPIXELSY As Long = 90
Public Const BITSPIXEL As Long = 12
Public Const HORZRES As Long = 8
Public Const VERTRES As Long = 10

Public Const CCDEVICENAME As Long = 32
Public Const CCFORMNAME As Long = 32

Public Const DM_BITSPERPEL As Long = &H40000
Public Const DM_PELSWIDTH As Long = &H80000
Public Const DM_PELSHEIGHT As Long = &H100000
Public Const DM_DISPLAYFLAGS As Long = &H200000

Public Type DEVMODE
   dmDeviceName      As String * CCDEVICENAME
   dmSpecVersion     As Integer
   dmDriverVersion   As Integer
   dmSize            As Integer
   dmDriverExtra     As Integer
   dmFields          As Long
   dmOrientation     As Integer
   dmPaperSize       As Integer
   dmPaperLength     As Integer
   dmPaperWidth      As Integer
   dmScale           As Integer
   dmCopies          As Integer
   dmDefaultSource   As Integer
   dmPrintQuality    As Integer
   dmColor           As Integer
   dmDuplex          As Integer
   dmYResolution     As Integer
   dmTTOption        As Integer
   dmCollate         As Integer
   dmFormName        As String * CCFORMNAME
   dmUnusedPadding   As Integer
   dmBitsPerPel      As Integer
   dmPelsWidth       As Long
   dmPelsHeight      As Long
   dmDisplayFlags    As Long
   dmDisplayFrequency As Long
End Type
'--end block--'


