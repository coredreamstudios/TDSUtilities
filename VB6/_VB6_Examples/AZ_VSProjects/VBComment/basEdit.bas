Attribute VB_Name = "basEdit"
Option Explicit

' ***************************************************************************
' Project:
'
' Module:        basEdit
'
' Description:   These are the common edit routines you will find in most
'                word processors.  (Copy, Cut, Paste)
'
' ===========================================================================
'    DATE      NAME                      DESCRIPTION
' -----------  ------------------------  ------------------------------------
' 02-JUL-1998  Kenneth Ives              Module created by kenaso@home.com
' ***************************************************************************

Public Sub Edit_Paste()

' ***************************************************************************
' Routine:       Edit_Paste
'
' Description:   Copy whatever text is being held in the clipboard and then
'                paste it in the text box. See Keydown event for the text
'                boxes to see an example of the code calling this routine.
'
' Parameters:
'
' Return Values:
'
' Special Logic:
'
' ===========================================================================
'    DATE      NAME                      DESCRIPTION
' -----------  ------------------------  ------------------------------------
' 02-JUL-1998  Kenneth Ives              Module created by kenaso@home.com
' ***************************************************************************

' ------------------------------------------------------------
' Verify this is a text box that the cursor is over
' ------------------------------------------------------------
  If TypeOf Screen.ActiveControl Is TextBox Then
      '
      ' unload clipboard into the textbox
      Screen.ActiveControl.SelText = Clipboard.GetText()
  End If

End Sub

Public Function CenterText(ByVal sTemp As String, Optional vntNumChars As Variant)
    
' ------------------------------------------------------------
' Define local variables
' ------------------------------------------------------------
  Dim nNumChars As Integer
  Dim nLeftOver As Integer
  Dim nAddToRight As Integer
    
' ------------------------------------------------------------
' See if there is a max string length passed
' ------------------------------------------------------------
  If Not IsMissing(vntNumChars) Then
      nNumChars = vntNumChars
  Else
      nNumChars = 80  ' default to 80
  End If

' ------------------------------------------------------------
' Subtract the length of the incoming string from the max
' length allowed
' ------------------------------------------------------------
  nLeftOver = nNumChars - Len(sTemp)
    
' ------------------------------------------------------------
' If there is something left over then calculate half of that
' value and prefixthe string with that numbe rof blank spaces.
' ------------------------------------------------------------
  If nLeftOver > 0 Then
      sTemp = Space(nLeftOver \ 2) & sTemp
  End If
    
' ------------------------------------------------------------
' Calculate number of spaces to the right
' ------------------------------------------------------------
  nAddToRight = nNumChars - Len(sTemp)
  
' ------------------------------------------------------------
' Append that number of blank spaces to the right
' ------------------------------------------------------------
  If nAddToRight > 0 Then
      sTemp = sTemp & Space(nAddToRight)
  End If
  
' ------------------------------------------------------------
' Return the centered text string.  In most cases we remove
' the trailing blanks.  In some unique cases we may want to
' remove the RTRIM function and retain those trailing blank
' spaces.
' ------------------------------------------------------------
  CenterText = RTrim(sTemp)
 
End Function

Public Sub Edit_Copy()

' ***************************************************************************
' Routine:       Edit_Copy
'
' Description:   Copy highlighted text to the clipboard. See Keydown event
'                for the text boxes to see an example of the code calling
'                this routine.
'
' Parameters:
'
' Return Values:
'
' Special Logic: When the user highlights text with the cursor and presses
'                CTRL+C to perform a copy function.  The highlighted text
'                is then loaded into the clipboard.
'
' ===========================================================================
'    DATE      NAME                      DESCRIPTION
' -----------  ------------------------  ------------------------------------
' 02-JUL-1998  Kenneth Ives              Module created by kenaso@home.com
' ***************************************************************************

' ------------------------------------------------------------
' Verify this is a text box that the cursor is over
' ------------------------------------------------------------
  If TypeOf Screen.ActiveControl Is TextBox Then
      '
      ' clear the clipboard
      Clipboard.Clear
      '
      ' load clipboard with the highlighted text
      Clipboard.SetText Screen.ActiveControl.SelText
  End If
  
End Sub

Public Sub Edit_Cut()

' ***************************************************************************
' Routine:       Edit_Cut
'
' Description:   Copy highlighted text to the clipboard and then remove it
'                from the text box. See Keydown event for the text boxes to
'                see an example of the code calling this routine.
'
' Parameters:
'
' Return Values:
'
' Special Logic: When the user highlights text with the cursor and presses
'                CTRL+X to perform a cutting function.  The highlighted text
'                is then moved to the clipboard.
'
' ===========================================================================
'    DATE      NAME                      DESCRIPTION
' -----------  ------------------------  ------------------------------------
' 02-JUL-1998  Kenneth Ives              Module created by kenaso@home.com
' ***************************************************************************

' ------------------------------------------------------------
' Verify this is a text box that the cursor is over
' ------------------------------------------------------------
  If TypeOf Screen.ActiveControl Is TextBox Then
      '
      ' clear the clipboard
      Clipboard.Clear
      '
      ' load clipboard with the highlighted text
      Clipboard.SetText Screen.ActiveControl.SelText
      '
      ' empty the textbox
      Screen.ActiveControl.SelText = ""
  End If

End Sub

Public Sub Edit_Delete()

' ***************************************************************************
' Routine:       Edit_Delete
'
' Description:   Copy highlighted text to the clipboard and then remove it
'                from the text box. See Keydown event for the text boxes to
'                see an example of the code calling this routine.
'
' Parameters:
'
' Return Values:
'
' Special Logic: When the user highlights text with the cursor and presses
'                CTRL+X to perform a cutting function.  The highlighted text
'                is then moved to the clipboard and the clipboard is emptied
'
' ===========================================================================
'    DATE      NAME                      DESCRIPTION
' -----------  ------------------------  ------------------------------------
' 02-JUL-1998  Kenneth Ives              Module created by kenaso@home.com
' ***************************************************************************


' ------------------------------------------------------------
' Verify this is a text box that the cursor is over
' ------------------------------------------------------------
  If TypeOf Screen.ActiveControl Is TextBox Then
      '
      ' remove the highlighted text from the textbox
      Screen.ActiveControl.SelText = ""
  End If
  
End Sub
