VERSION 5.00
Begin VB.Form frmSampler 
   Caption         =   "Form1"
   ClientHeight    =   4980
   ClientLeft      =   2190
   ClientTop       =   2685
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   9945
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   540
      Left            =   5955
      TabIndex        =   23
      Top             =   4305
      Width           =   3840
   End
   Begin VB.ComboBox cboEnvVar 
      Height          =   315
      ItemData        =   "frmSampler.frx":0000
      Left            =   7665
      List            =   "frmSampler.frx":000A
      TabIndex        =   22
      Text            =   "cboEnvVar"
      ToolTipText     =   "Enter the name of an environment variable here to see it's value."
      Top             =   1680
      Width           =   2130
   End
   Begin VB.TextBox txtEnvPath 
      Height          =   1605
      Left            =   5955
      MultiLine       =   -1  'True
      TabIndex        =   20
      Top             =   2145
      Width           =   3840
   End
   Begin VB.Frame Frame2 
      Caption         =   "System Info"
      Height          =   4680
      Left            =   165
      TabIndex        =   3
      Top             =   165
      Width           =   5430
      Begin VB.TextBox txtMinAppAddress 
         Height          =   375
         Left            =   2865
         TabIndex        =   18
         Top             =   3990
         Width           =   2295
      End
      Begin VB.TextBox txtMaxAppAddress 
         Height          =   375
         Left            =   2865
         TabIndex        =   16
         Top             =   3465
         Width           =   2295
      End
      Begin VB.TextBox txtProcType 
         Height          =   375
         Left            =   2865
         TabIndex        =   14
         Top             =   2945
         Width           =   2295
      End
      Begin VB.TextBox txtPageSize 
         Height          =   375
         Left            =   2865
         TabIndex        =   12
         Top             =   2425
         Width           =   2295
      End
      Begin VB.TextBox txtOEMID 
         Height          =   375
         Left            =   2865
         TabIndex        =   10
         Top             =   1905
         Width           =   2295
      End
      Begin VB.TextBox txtAllocGran 
         Height          =   375
         Left            =   2865
         TabIndex        =   8
         Top             =   865
         Width           =   2295
      End
      Begin VB.TextBox txtNumOfProcs 
         Height          =   375
         Left            =   2865
         TabIndex        =   6
         Top             =   1385
         Width           =   2295
      End
      Begin VB.TextBox txtProcMask 
         Height          =   375
         Left            =   2865
         TabIndex        =   4
         Top             =   345
         Width           =   2295
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Minimum Application Address:"
         Height          =   255
         Left            =   300
         TabIndex        =   19
         Top             =   4065
         Width           =   2400
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Maximum Application Address:"
         Height          =   255
         Left            =   270
         TabIndex        =   17
         Top             =   3543
         Width           =   2430
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Processor Type:"
         Height          =   255
         Left            =   1005
         TabIndex        =   15
         Top             =   3025
         Width           =   1695
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Page Size:"
         Height          =   255
         Left            =   1005
         TabIndex        =   13
         Top             =   2507
         Width           =   1695
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "OEM ID:"
         Height          =   255
         Left            =   1005
         TabIndex        =   11
         Top             =   1989
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Allocation Granularity:"
         Height          =   255
         Left            =   1005
         TabIndex        =   9
         Top             =   953
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Number of Processors:"
         Height          =   255
         Left            =   1005
         TabIndex        =   7
         Top             =   1471
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Processor Mask:"
         Height          =   255
         Left            =   1005
         TabIndex        =   5
         Top             =   435
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   5925
      TabIndex        =   0
      Top             =   165
      Width           =   3135
      Begin VB.TextBox txtBlinkRate 
         Height          =   375
         Left            =   1920
         TabIndex        =   1
         ToolTipText     =   "Enter a number of milliseconds for the cursor blink rate."
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Cursor Blink Rate (ms):"
         Height          =   255
         Left            =   105
         TabIndex        =   2
         Top             =   330
         Width           =   1695
      End
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "Environment Variable:"
      Height          =   255
      Left            =   5955
      TabIndex        =   21
      Top             =   1725
      Width           =   1575
   End
End
Attribute VB_Name = "frmSampler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------------------------------------------------------------------------------------------
'----------------------------------------------------------WARNING--------------------------------------------------------
'When running this program in debug mode (in the IDE) DO NOT use the stop button to kill the program.
'Use the menu item 'close' or use the 'X' in the control box on the form.  Otherwise you could couse your
'system to act weird untill you re-boot.
'------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------------------------------------
'Name:          Sampler
'Author:        Jonathan Morrison    <jonathanm@mindspring.com>
'Date:          8/15/1998
'------------------------------------------------------------------------------------------------------------------------------

Private Sub cboEnvVar_Change()

Dim strEnvVarName As String
Dim strVarBuffer As String * 4096

'Save the name of the environment variable to get.
strEnvVarName = cboEnvVar.Text

'Try to get the variables value.
GetEnvironmentVariable strEnvVarName, strVarBuffer, Len(strVarBuffer)

'Ignore if variable not found.
txtEnvPath.Text = Trim$(strVarBuffer)

End Sub

Private Sub cboEnvVar_Click()

Dim strEnvVarName As String
Dim strVarBuffer As String * 4096

'Save the name of the environment variable to get.
strEnvVarName = cboEnvVar.Text

'Try to get the variables value.
GetEnvironmentVariable strEnvVarName, strVarBuffer, Len(strVarBuffer)

'Ignore if variable not found.
txtEnvPath.Text = Trim$(strVarBuffer)

End Sub

Private Sub cmdClose_Click()

Unload Me

End Sub

Private Sub Form_Load()

Dim siInfo As SYSTEM_INFO
Dim strUserName As String
Dim strUserBuffer As String * 1024
Dim strComputerName As String
Dim strComputerBuffer As String * 1024

Dim lngCount As Long
Dim lngHours As Long
Dim lngMinutes As Long

lngCount = GetTickCount

lngHours = ((lngCount / 1000) / 60) / 60

lngMinutes = ((lngCount / 1000) / 60) Mod 60

MsgBox "Your system has been running for " & lngHours & " hours and " & lngMinutes & _
            " minutes.", vbInformation, "Sampler"

'Get the computer name.
'Be sure to send a fixed length string that is large enough to hold the name.
'Also be sure thet the second parameter, 'nSize', is including all of the null characters.
GetComputerName strComputerBuffer, Len(strComputerBuffer)

'Take the returned string and trim everything after the first null character.
strComputerName = Left$(strComputerBuffer, (InStr(1, strComputerBuffer, vbNullChar)) - 1)

'Get the user name.
'Be sure to send a fixed length string that is large enough to hold the name.
'Also be sure thet the second parameter, 'nSize', is including all of the null characters.
GetUserName strUserBuffer, Len(strUserBuffer)

'Take the returned string and trim everything after the first null character.
strUserName = Left$(strUserBuffer, (InStr(1, strUserBuffer, vbNullChar)) - 1)

'Make sure we got a string back and display them.
If strUserName <> "" And strComputerName <> "" Then
    Me.Caption = "User: " & strUserName & " on Computer: \\" & strComputerName
Else
    Me.Caption = "No network information found...."
End If

'Let Windows fill the SYSTEM_INFO structure for us.
GetSystemInfo siInfo

'Load the text boxes with the returned information.
txtProcMask.Text = siInfo.dwActiveProcessorMask
txtAllocGran.Text = siInfo.dwAllocationGranularity
txtNumOfProcs.Text = siInfo.dwNumberOrfProcessors
txtOEMID.Text = siInfo.dwOemID
txtPageSize.Text = siInfo.dwPageSize
txtProcType.Text = siInfo.dwProcessorType
txtMaxAppAddress.Text = "0x" & Hex(siInfo.lpMaximumApplicationAddress) 'since these are memory addresses we convert to hex.
txtMinAppAddress.Text = "0x" & Hex(siInfo.lpMinimumApplicationAddress) 'since these are memory addresses we convert to hex.

cboEnvVar.ListIndex = 0

'Initialize the cursor blink rate.
txtBlinkRate.Text = 500
txtBlinkRate.SelLength = Len(txtBlinkRate.Text)

End Sub

Private Sub Form_Unload(Cancel As Integer)

'Return the blink rate to normal.
SetCaretBlinkTime 500&

End Sub

Private Sub txtBlinkRate_Change()

'Make sure that there is a valid rate in the box.  If so..... set the cursor blink rate.
If txtBlinkRate.Text <> "" Then
    SetCaretBlinkTime CLng(txtBlinkRate.Text)
End If

End Sub

