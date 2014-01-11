Attribute VB_Name = "basDelay"
Option Explicit

' ***************************************************************************
' Project:
'
' Module:        basDelay
'
' Description:   Three different ways to pause an application
'
' ===========================================================================
'    DATE      NAME                      DESCRIPTION
' -----------  ------------------------  ------------------------------------
' 23-DEC-1999  Kenneth Ives              Module created by kenaso@home.com
' ***************************************************************************

' ------------------------------------------------------------
' This is a rough translation of the GetTickCount API. The
' tick count of a PC is only valid for the first 49.7 days
' since it was last rebooted.  When you capture the tick
' count, you are capturing the total number of milliseconds
' since the PC was last rebooted.
' ------------------------------------------------------------
  Private Declare Function GetTickCount Lib "kernel32" () As Long
  
' ------------------------------------------------------------
' Delay method #1
' Declare for Sleep API call.  This will stop the application
' entirely while this API is executing.
'
' Syntax:   Sleep 60000&     ' sleep for 1 minute
' ------------------------------------------------------------
  Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub Delay(iAmtOfDelay As Integer, _
                 Optional sTypeOfDelay As String = "s")

' ***************************************************************************
' Routine:       Delay  (Method #2)
'
' Description:   This routine will cause a delay for the time requested,
'                yet will not interfere with the program progress like the
'                Sleep API.  This routine does not rely on the Timer event,
'                a timer control, or the tickcount of the last reboot.  All
'                of these have their drawbacks when dealing with time
'                comparisons.
'
'                Timer control is only valid from midnight of that
'                particular day.  If you want to delay 5 minutes and the
'                time is 23:59:00.  You will never reach the finish time
'                because at midnight, the timer control is reset.
'
'                The timer event is based on a single precision caluclation
'                of the date and time.  to the left of the decimal is the
'                date and to the right is the time.  Somewhere in the 24
'                hour cycle, this is reset.  I suspect midnight.  this is
'                good for immediate testing but not for comparisons.
'
'                I use the system date and time by calling the VB function
'                Now().  As long as the machine is running, it will have a
'                system date and time stamp that is being updated.
'
' Parameters:    iAmtOfDelay - numeric amount of time to delay
'                sTypeOfDelay - (Optional) This is the type of delay
'                accepted in the DateAdd function.  Default is seconds.
'
'                         "s" - seconds      "d" - days
'                         "n" - minutes      "m" - months
'                         "h" - hours        "yyyy" - years
'
' Return Values:
'
' Special Logic:
'
' ===========================================================================
'    DATE      NAME                      DESCRIPTION
' -----------  ------------------------  ------------------------------------
' 23-DEC-1999  Kenneth Ives              Module created by kenaso@home.com
' ***************************************************************************

' -----------------------------------------------------------
' Define local variables
' -----------------------------------------------------------
  Dim vDelayTime As Variant
  Dim vCurrTime As Variant
  
' -----------------------------------------------------------
' Determine the length of time to delay using the VB DateAdd
' function.
'
'    "s" - seconds      "d" - days
'    "n" - minutes      "m" - months
'    "h" - hours        "yyyy" - years
'
' We are adding the amount of delay to the current time and
' formatting the output.
' -----------------------------------------------------------
  vDelayTime = Format(DateAdd(sTypeOfDelay, iAmtOfDelay, Now), "hh:mm:ss")
  
' -----------------------------------------------------------
' Loop thru and continualy check the curent time with the
' calculated time so we know when to leave
' -----------------------------------------------------------
  Do
      vCurrTime = Format(Now, "hh:mm:ss")
      
      ' if the string1 is greater than string2,
      ' a one will be returned
      If StrComp(vCurrTime, vDelayTime) = 1 Then
          Exit Do
      End If
      
      DoEvents
      DoEvents
  Loop

End Sub

Public Sub Mini_Delay(lDelayAmt As Long)

' ***************************************************************************
' Routine:       Mini_Delay  (Method #3)
'
' Description:   This routine uses the tickcount as a countdown referenece.
'                I use this when I do not want to use partial seconds.
'                          1000 milliseconds = 1 second
'                          750 = 3/4 second
'
' Parameters:    lDelayAmt - numeric amount of time to delay
'
' Return Values:
'
' Special Logic:
'
' ===========================================================================
'    DATE      NAME                      DESCRIPTION
' -----------  ------------------------  ------------------------------------
' 23-DEC-1999  Kenneth Ives              Module created by kenaso@home.com
' ***************************************************************************

' -----------------------------------------------------------
' Define local variables
' -----------------------------------------------------------
  Dim lNewTime As Long
  Dim lCurrent As Long
  
' -----------------------------------------------------------
' Calculate the new waiting time
' -----------------------------------------------------------
  lNewTime = GetTickCount + lDelayAmt
  
' -----------------------------------------------------------
' Loop thru and continualy check the curent time with the
' calculated time so we know when to leave
' -----------------------------------------------------------
  Do
      lCurrent = GetTickCount      ' get the current millisecond count
       
      ' if the current millisecond count has not
      ' caught up with the delay amount then
      ' we will try again.
      If lCurrent >= lNewTime Then
          Exit Do
      End If
      
      ' allow other processes to happen
      DoEvents
      DoEvents
  Loop

End Sub

