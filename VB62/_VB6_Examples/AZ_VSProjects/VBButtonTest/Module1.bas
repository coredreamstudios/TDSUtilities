Attribute VB_Name = "Module1"
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
Public defWindowProc As Long

Public Const GWL_WNDPROC As Long = (-4)
Public Const WM_NCHITTEST As Long = &H84
Public Const HTCAPTION As Long = 2
Public Const HTCLIENT As Long = 1

Public Declare Function GetWindowLong Lib "user32" _
    Alias "GetWindowLongA" _
   (ByVal hWnd As Long, _
    ByVal nIndex As Long) As Long
    
Public Declare Function SetWindowLong Lib "user32" _
    Alias "SetWindowLongA" _
   (ByVal hWnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long
   
Public Declare Function CallWindowProc Lib "user32" _
    Alias "CallWindowProcA" _
   (ByVal lpPrevWndFunc As Long, _
    ByVal hWnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long


Public Function WindowProc(ByVal hWnd As Long, _
                    ByVal Msg As Long, _
                    ByVal wParam As Long, _
                    ByVal lParam As Long) As Long

   'Subclass form to trap messages
    On Error Resume Next
    
    Dim retVal As Long     'Where on the form is the mouse?
    Dim Mouse_X As Long    'Mouse coordinates
    Dim Mouse_Y As Long
    
   'First let the original Window Procedure process the message.
   'CallWindowProc returns the part of the form the mouse is on.
    retVal = CallWindowProc(defWindowProc, _
                            hWnd, _
                            Msg, _
                            wParam, _
                            lParam)
    
  'What message received?
   Select Case Msg
      Case WM_NCHITTEST      'Every mouse action
      
        'in lParam there are the mouse co-ordinates
        'for info only
         Mouse_X = LoWord(lParam)
         Mouse_Y = HiWord(lParam)
         GraphicForm.Label1.Caption = "X: " & Mouse_X & vbCrLf & _
                                      "Y: " & Mouse_Y
            
        'If mouse on client area, tell Windows the mouse is
        'on the caption bar!
         If retVal = HTCLIENT Then  'action on client area
            retVal = HTCAPTION      'tell Windows its the caption
         End If
            
        Case Else
           'Other WM_xxx messages you want to intercept
           'Act as appropriate
            
    End Select
        
       'return the value HTCAPTION to Windows
        WindowProc = retVal
        
End Function


Public Function HiWord(dw As Long) As Integer

    If dw And &H80000000 Then
          HiWord = (dw \ 65535) - 1
    Else: HiWord = dw \ 65535
    End If
    
End Function
  

Public Function LoWord(dw As Long) As Integer

    If dw And &H8000& Then
          LoWord = &H8000& Or (dw And &H7FFF&)
    Else: LoWord = dw And &HFFFF&
    End If
    
End Function
'--end block--'


