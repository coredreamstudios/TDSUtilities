VERSION 5.00
Begin VB.Form frmComments 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VB Comment Tool"
   ClientHeight    =   5685
   ClientLeft      =   1620
   ClientTop       =   1440
   ClientWidth     =   7950
   Icon            =   "frmAddIn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   7950
   Begin VB.TextBox txtDelay 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2175
      MaxLength       =   2
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   5250
      Width           =   465
   End
   Begin VB.PictureBox picMsg 
      BackColor       =   &H0000FF00&
      Height          =   3690
      Left            =   1350
      ScaleHeight     =   3630
      ScaleWidth      =   5355
      TabIndex        =   18
      Top             =   825
      Visible         =   0   'False
      Width           =   5415
      Begin VB.Label lblMsg 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2865
         Left            =   300
         TabIndex        =   19
         Top             =   375
         Width           =   4740
      End
   End
   Begin VB.TextBox txtComments 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   7
      Left            =   1350
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   4425
      Width           =   6390
   End
   Begin VB.TextBox txtComments 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   6
      Left            =   1350
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   3750
      Width           =   6390
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   1
      Left            =   6825
      Picture         =   "frmAddIn.frx":030A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   17
      Top             =   225
      Width           =   480
   End
   Begin VB.ComboBox cboType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmAddIn.frx":0614
      Left            =   5625
      List            =   "frmAddIn.frx":061B
      Style           =   2  'Dropdown List
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   825
      Width           =   2115
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   0
      Left            =   600
      Picture         =   "frmAddIn.frx":0628
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   14
      Top             =   225
      Width           =   480
   End
   Begin VB.TextBox txtComments 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   5
      Left            =   1350
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   3075
      Width           =   6390
   End
   Begin VB.TextBox txtComments 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   4
      Left            =   1350
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2400
      Width           =   6390
   End
   Begin VB.TextBox txtComments 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   3
      Left            =   1350
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1725
      Width           =   6390
   End
   Begin VB.TextBox txtComments 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   4125
      MaxLength       =   25
      TabIndex        =   1
      Top             =   1275
      Width           =   3615
   End
   Begin VB.TextBox txtComments 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   1350
      TabIndex        =   0
      Top             =   1275
      Width           =   1440
   End
   Begin VB.TextBox txtComments 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Index           =   0
      Left            =   150
      MultiLine       =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   900
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.CommandButton cmdChoice 
      Caption         =   "Cancel"
      Height          =   375
      Index           =   2
      Left            =   6450
      TabIndex        =   9
      Top             =   5175
      Width           =   1215
   End
   Begin VB.CommandButton cmdChoice 
      Caption         =   "Save"
      Height          =   375
      Index           =   1
      Left            =   5025
      TabIndex        =   8
      Top             =   5175
      Width           =   1215
   End
   Begin VB.CommandButton cmdChoice 
      Caption         =   "Clear Boxes"
      Height          =   375
      Index           =   0
      Left            =   3600
      TabIndex        =   7
      Top             =   5175
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Type of Comment"
      Height          =   240
      Index           =   3
      Left            =   3825
      TabIndex        =   27
      Top             =   900
      Width           =   1665
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Number of seconds to display the SAVE window."
      Height          =   465
      Left            =   225
      TabIndex        =   25
      Top             =   5175
      Width           =   2040
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Height          =   390
      Index           =   4
      Left            =   225
      TabIndex        =   24
      Top             =   4500
      Width           =   915
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Height          =   390
      Index           =   3
      Left            =   225
      TabIndex        =   23
      Top             =   3900
      Width           =   915
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Height          =   390
      Index           =   2
      Left            =   225
      TabIndex        =   22
      Top             =   3150
      Width           =   915
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Height          =   390
      Index           =   1
      Left            =   225
      TabIndex        =   21
      Top             =   2475
      Width           =   915
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Height          =   390
      Index           =   0
      Left            =   225
      TabIndex        =   20
      Top             =   1875
      Width           =   915
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Freeware by Kenneth Ives"
      Height          =   240
      Index           =   1
      Left            =   2850
      TabIndex        =   15
      Top             =   75
      Width           =   2340
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "VB Comment Tool"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   2
      Left            =   150
      TabIndex        =   13
      Top             =   300
      Width           =   7590
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Written by:"
      Height          =   240
      Index           =   5
      Left            =   2925
      TabIndex        =   12
      Top             =   1350
      Width           =   1065
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   315
      Index           =   0
      Left            =   225
      TabIndex        =   11
      Top             =   1350
      Width           =   690
   End
End
Attribute VB_Name = "frmComments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ***************************************************************************
' Project:       vbCmts.dll (VB6 add-in)
'
' Module:        frmComments
'
' Description:   This is the main screen for this application.  A user can
'                create their own unique comments and paste them into their
'                code.  This application was created using the VB6 AddIn
'                Wizard.  Read all comments first to understand what is
'                happening.
'
'                The comments entered in this program are from this
'                application
'
'                To get the menu name you desire when you click the Add-In
'                menu option, you must go into the Project window.  This is
'                usually locate on the right side.  Go under the Designers
'                folder and double click Connect.dsr  Type in the name you
'                want to see in the menu in the top box.
'
'                Thanks to John P. Cunningham for the added enhancments.
' ===========================================================================
'    DATE      NAME                      DESCRIPTION
' -----------  ------------------------  ------------------------------------
' 23-DEC-1999  Kenneth Ives              Module created by kenaso@home.com
' 12-JAN-2000  John P Cunningham         Increased the "Updated by" text
'                                        length from 14 to 25 characters.
'                                        Created and added a menu icon.
'                                        mail to:  johnpc@ids.net
' ***************************************************************************

' ------------------------------------------------------------
' These variables were placed here by VB Add-In wizard
' ------------------------------------------------------------
  Public VBInstance As VBIDE.VBE
  Public Connect As Connect

' ------------------------------------------------------------
' Define private variables
' ------------------------------------------------------------
  Private sName As String
  Private sDate As String
  Private iDelay As Integer
  Private bInitialLoad As Boolean

' ------------------------------------------------------------
' Create a temporary file name
' ------------------------------------------------------------
  Private Declare Function GetTempFileName Lib "kernel32" _
                  Alias "GetTempFileNameA" (ByVal lpszPath As String, _
                  ByVal lpPrefixString As String, ByVal wUnique As Long, _
                  ByVal lpTempFileName As String) As Long

Private Sub cboType_Click()
  
' ------------------------------------------------------------
' Depending on what type of comments are selected, we will
' define the default boxes and labels here.
' ------------------------------------------------------------
  
' ------------------------------------------------------------
' Define local variables
' ------------------------------------------------------------
  Dim sTmpComboText As String
  
' ------------------------------------------------------------
' Initialize variables
' ------------------------------------------------------------
  sTmpComboText = StrConv(cboType.Text, vbLowerCase)
  txtComments(1) = sDate
    
' ------------------------------------------------------------
' Based on which item is selected, lock/unlock certain boxes
' ------------------------------------------------------------
  Select Case sTmpComboText
         
         Case "module initial"
              ' unlock these boxes
              txtComments(2).BackColor = vbWhite
              txtComments(2).Enabled = True
              txtComments(3).BackColor = vbWhite
              txtComments(3).Enabled = True
              txtComments(4).BackColor = vbWhite
              txtComments(4).Enabled = True
              txtComments(5).BackColor = vbWhite
              txtComments(5).Enabled = True
         
              ' lock these boxes
              txtComments(6) = ""
              txtComments(6).BackColor = &HE0E0E0      ' light gray
              txtComments(6).Enabled = False
              txtComments(7) = ""
              txtComments(7).BackColor = &HE0E0E0      ' light gray
              txtComments(7).Enabled = False
         
              ' set the title labels
              txtComments(2) = sName
              lblTitle(0) = "Project" & vbCrLf & "Name"
              lblTitle(1) = "Module" & vbCrLf & "Name"
              lblTitle(2) = "Description"
              lblTitle(3) = ""
              lblTitle(4) = ""
              lblTitle(5) = "Updated by"
              
         Case "routine initial"
              ' unlock these boxes
              txtComments(2).BackColor = vbWhite
              txtComments(2).Enabled = True
              txtComments(3).BackColor = vbWhite
              txtComments(3).Enabled = True
              txtComments(4).BackColor = vbWhite
              txtComments(4).Enabled = True
              txtComments(5).BackColor = vbWhite
              txtComments(5).Enabled = True
              txtComments(6).BackColor = vbWhite
              txtComments(6).Enabled = True
              txtComments(7).BackColor = vbWhite
              txtComments(7).Enabled = True
         
              ' set the title labels
              txtComments(2) = sName
              lblTitle(0) = "Routine" & vbCrLf & "Name"
              lblTitle(1) = "Description"
              lblTitle(2) = "Parameters"
              lblTitle(3) = "Return" & vbCrLf & "Values"
              lblTitle(4) = "Special" & vbCrLf & "Logic"
              lblTitle(5) = "Updated by"
         
         Case "module append", "routine append"
              ' unlock this box
              txtComments(2).BackColor = vbWhite
              txtComments(2).Enabled = True
              txtComments(3).BackColor = vbWhite
              txtComments(3).Enabled = True
         
              ' lock these boxes
              txtComments(4) = ""
              txtComments(4).BackColor = &HE0E0E0      ' light gray
              txtComments(4).Enabled = False
              txtComments(5) = ""
              txtComments(5).BackColor = &HE0E0E0      ' light gray
              txtComments(5).Enabled = False
              txtComments(6) = ""
              txtComments(6).BackColor = &HE0E0E0      ' light gray
              txtComments(6).Enabled = False
              txtComments(7) = ""
              txtComments(7).BackColor = &HE0E0E0      ' light gray
              txtComments(7).Enabled = False
         
              ' set the title labels
              txtComments(2) = sName
              lblTitle(0) = "Description" & vbCrLf & "of Updates"
              lblTitle(1) = ""
              lblTitle(2) = ""
              lblTitle(3) = ""
              lblTitle(4) = ""
              lblTitle(5) = "Updated by"
              
         Case "general initial", "general append"
              ' unlock this box
              txtComments(3).BackColor = vbWhite
              txtComments(3).Enabled = True
         
              ' lock these boxes
              txtComments(2) = ""
              txtComments(2).BackColor = &HE0E0E0      ' light gray
              txtComments(2).Enabled = False
              txtComments(4) = ""
              txtComments(4).BackColor = &HE0E0E0      ' light gray
              txtComments(4).Enabled = False
              txtComments(5) = ""
              txtComments(5).BackColor = &HE0E0E0      ' light gray
              txtComments(5).Enabled = False
              txtComments(6) = ""
              txtComments(6).BackColor = &HE0E0E0      ' light gray
              txtComments(6).Enabled = False
              txtComments(7) = ""
              txtComments(7).BackColor = &HE0E0E0      ' light gray
              txtComments(7).Enabled = False
         
              ' set the title labels
              lblTitle(0) = "What is this" & vbCrLf & "code doing?"
              lblTitle(1) = ""
              lblTitle(2) = ""
              lblTitle(3) = ""
              lblTitle(4) = ""
              lblTitle(5) = ""
  End Select

End Sub

Private Sub cmdChoice_Click(Index As Integer)

' ------------------------------------------------------------
' Based on which button is pressed
' ------------------------------------------------------------
  Select Case Index
         Case 0: ClearBoxes          ' Empty all text boxes
         Case 1: CopyInputData       ' Copy input data to the cursor position
         Case 2: Form_Unload False   ' unload this application completely
  End Select
  
End Sub

Private Sub CopyInputData()
  
' ***************************************************************************
' Routine:       CopyInputData
'
' Description:   This routine will copy the input data to a temp file and
'                then to the clipboard.  The temp file will then be deleted.
'
' Parameters:
'
' Return Values:
'
' Special Logic:
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 23-DEC-1999  Kenneth Ives     Routine created by kenaso@home.com
' ***************************************************************************

' ------------------------------------------------------------
' Define local variables
' ------------------------------------------------------------
  Dim lRetVal As Long
  Dim iFile As Integer
  Dim i As Integer
  Dim sTmp As String
  Dim sFileName As String
  Dim arData(6) As String
  Dim sTmpComboText As String
  Dim REC_LENGTH As Integer
  
' ------------------------------------------------------------
' Initialize variables
' ------------------------------------------------------------
  sTmp = ""                   ' empty variable
  sFileName = Space(260)      ' plenty of space for the path and filename
  Erase arData()              ' empty array
  REC_LENGTH = 75
  
' ------------------------------------------------------------
' Get a temporary file name and save everything up to the
' first null value.
' ------------------------------------------------------------
  lRetVal = GetTempFileName(App.Path, "VBC", 0&, sFileName)
  sFileName = Left(sFileName, InStr(sFileName, Chr(0)) - 1)
  
' ------------------------------------------------------------
' Open a new temp file
' ------------------------------------------------------------
  iFile = FreeFile
  Open sFileName For Output As #iFile
  
' ------------------------------------------------------------
' Update the temp file after formatting the data
' ------------------------------------------------------------
  sTmpComboText = StrConv(cboType.Text, vbLowerCase)
  
  Select Case sTmpComboText
         Case "module initial"
                If Len(Trim(sName)) = 0 Then
                    Close #iFile
                    MsgBox "Need the name of the person entering these comments.", _
                            vbCritical + vbOKOnly, "VB Comment Tool"
                    txtComments(2).SetFocus
                    Exit Sub
                Else
                    ' Format the name of the developer
                    i = 16 - Len(sName)
                End If
                
                If Len(Trim(txtComments(4))) = 0 Then
                    Close #iFile
                    MsgBox "Need the name of this module.", _
                            vbCritical + vbOKOnly, "VB Comment Tool"
                    txtComments(4).SetFocus
                    Exit Sub
                End If
                
                ' format the data to fit within the boundries
                arData(0) = PrepareData(1, "' Project:       ", txtComments(3))
                arData(1) = PrepareData(1, "' Module:        ", txtComments(4))
                arData(2) = PrepareData(1, "' Description:   ", txtComments(5))
         
                ' Update the temp file.  By using the PRINT statement, w/o a
                ' trailing semi-colon, we force a carriage return and linefeed
                ' at the end of each line.
                Print #iFile, "' " & String(REC_LENGTH, 42)
                Print #iFile, arData(0)
                Print #iFile, "' "
                Print #iFile, arData(1)
                Print #iFile, "' "
                Print #iFile, arData(2)
                Print #iFile, "' "
                Print #iFile, "' " & String(REC_LENGTH, 61)
                Print #iFile, "'    DATE      NAME             DESCRIPTION"
                Print #iFile, "' -----------  ---------------  ---------------------------------------------"
                Print #iFile, "' " & txtComments(1) & "  " & sName & Space(i) & " Module created"
                Print #iFile, "' " & String(REC_LENGTH, 42)
                Close #iFile
         
         Case "routine initial"
                If Len(Trim(sName)) = 0 Then
                    Close #iFile
                    MsgBox "Need the name of the person entering these comments.", _
                            vbCritical + vbOKOnly, "VB Comment Tool"
                    txtComments(2).SetFocus
                    Exit Sub
                Else
                    ' Format the name of the developer
                    i = 16 - Len(sName)
                End If
                
                If Len(Trim(txtComments(3))) = 0 Then
                    Close #iFile
                    MsgBox "Need the name of this routine.", _
                            vbCritical + vbOKOnly, "VB Comment Tool"
                    txtComments(3).SetFocus
                    Exit Sub
                End If
                
                ' format the data to fit within the boundries
                arData(0) = PrepareData(1, "' Routine:       ", txtComments(3))
                arData(1) = PrepareData(1, "' Description:   ", txtComments(4))
                arData(2) = PrepareData(1, "' Parameters:    ", txtComments(5))
                arData(3) = PrepareData(1, "' Return Values: ", txtComments(6))
                arData(4) = PrepareData(1, "' Special Logic: ", txtComments(7))
  
                ' Update the temp file.  By using the PRINT statement, w/o a
                ' trailing semi-colon, we force a carriage return and linefeed
                ' at the end of each line.
                Print #iFile, "' " & String(REC_LENGTH, 42)
                Print #iFile, arData(0)
                Print #iFile, "' "
                Print #iFile, arData(1)
                Print #iFile, "' "
                Print #iFile, arData(2)
                Print #iFile, "' "
                Print #iFile, arData(3)
                Print #iFile, "' "
                Print #iFile, arData(4)
                Print #iFile, "' "
                Print #iFile, "' " & String(REC_LENGTH, 61)
                Print #iFile, "'    DATE      NAME             DESCRIPTION"
                Print #iFile, "' -----------  ---------------  ---------------------------------------------"
                Print #iFile, "' " & txtComments(1) & "  " & sName & Space(i) & " Routine created"
                Print #iFile, "' " & String(REC_LENGTH, 42)
                Close #iFile
         
         Case "module append", "routine append"
                If Len(Trim(sName)) = 0 Then
                    Close #iFile
                    MsgBox "Need the name of the person entering these comments.", _
                            vbCritical + vbOKOnly, "VB Comment Tool"
                    txtComments(2).SetFocus
                    Exit Sub
                Else
                    ' Format the name of the developer
                    i = 16 - Len(sName)
                End If
                
                ' format the data to fit within the boundries
                arData(0) = PrepareData(2, "", txtComments(3))
  
                ' Update the temp file.  By using the PRINT statement, w/o a
                ' trailing semi-colon, we force a carriage return and linefeed
                ' at the end of each line.
                Print #iFile, "' " & txtComments(1) & "  " & sName & Space(i) & " " & arData(0)
                Close #iFile
         
         Case "general initial"
                ' format the data to fit within the boundries
                arData(0) = PrepareData(3, "", txtComments(3))
                
                ' Update the temp file.  By using the PRINT statement, w/o a
                ' trailing semi-colon, we force a carriage return and linefeed
                ' at the end of each line.
                REC_LENGTH = 60
                Print #iFile, "' " & String(REC_LENGTH, 45)
                Print #iFile, arData(0)
                Print #iFile, "' " & String(REC_LENGTH, 45)
                Close #iFile
         
         Case "general append"
                ' format the data to fit within the boundries
                arData(0) = PrepareData(3, "", txtComments(3))
                
                ' Update the temp file.  By using the PRINT statement, w/o a
                ' trailing semi-colon, we force a carriage return and linefeed
                ' at the end of each line.
                Print #iFile, arData(0)
                Close #iFile
  End Select
  
' ------------------------------------------------------------
' The code is placed in a text box and then copied to the
' clipboard
' ------------------------------------------------------------
  iFile = FreeFile
  Open sFileName For Input As #iFile
  
' ------------------------------------------------------------
' load the hidden textbox with the formatted code
' ------------------------------------------------------------
  txtComments(0) = Input(LOF(iFile), #iFile)
  Close #iFile
  
' ------------------------------------------------------------
' Empty the clipboard and save this code
' ------------------------------------------------------------
  With Clipboard
       .Clear
       .SetText txtComments(0)
  End With
  
' ------------------------------------------------------------
' Display a message for the user
' ------------------------------------------------------------
  If iDelay > 0 Then
      picMsg.Visible = True
      Delay iDelay
      picMsg.Visible = False
  End If
  
' ------------------------------------------------------------
' Delete this temp file
' ------------------------------------------------------------
  Kill sFileName

' ------------------------------------------------------------
' Hide this application window and empty the hidden text box
' ------------------------------------------------------------
  Erase arData()      ' empty the array
  ClearBoxes          ' Empty the text boxes
  Connect.Hide        ' Hide this form

End Sub

Private Sub ClearBoxes()
  
' ***************************************************************************
' Routine:       ClearBoxes
'
' Description:   Clears the text boxes
'
' Parameters:
'
' Return Values:
'
' Special Logic: Just clear the comment area.
'
' ===========================================================================
'    DATE      NAME                      DESCRIPTION
' -----------  ------------------------  ------------------------------------
' 23-DEC-1999  Kenneth Ives              Module created by kenaso@home.com
' ***************************************************************************

' ------------------------------------------------------------
' Define local variables
' ------------------------------------------------------------
  Dim i As Integer
  
' ------------------------------------------------------------
' See if we have a name and date on hand
' ------------------------------------------------------------
  If Len(sDate) > 0 Then
      sDate = StrConv(sDate, vbUpperCase)
      txtComments(1) = sDate
  Else
      txtComments(1) = StrConv(Format(Now, "dd-mmm-yyyy"), vbUpperCase)
      sDate = txtComments(1)
  End If
  
  If Len(sName) > 0 Then
      txtComments(2) = sName
  Else
      sName = txtComments(2)
  End If
  
' ------------------------------------------------------------
' Clear all text boxes
' ------------------------------------------------------------
  For i = 3 To 7
      txtComments(i) = ""
  Next
   
End Sub

Private Function PrepareData(iType As Integer, sDescription As String, _
                             sInputText As String) As String

' ***************************************************************************
' Routine:       PrepareData
'
' Description:   If the input text boxes have large amounts of data in them,
'                this routine will perform line breaks so as to keep the text
'                legible.
'
' Parameters:    sDescription - Line description (i.e. "' Input Parms:  ")
'                sInputText - This is the data from one of the text boxes
'                that might have to have line breaks performed.
'
' Return Values: Formatted text lines with line breaks
'
' Special Logic:
'
' ===========================================================================
'    DATE      NAME                      DESCRIPTION
' -----------  ------------------------  ------------------------------------
' 23-DEC-1999  Kenneth Ives              Module created by kenaso@home.com
' ***************************************************************************

' ------------------------------------------------------------
' Define local variables
' ------------------------------------------------------------
  Dim i As Integer
  Dim iPos As Integer
  Dim iLen As Integer
  Dim sTmp1 As String
  Dim sTmp2 As String
  Dim sTmp3 As String
  Dim sChar As String
  Dim sOutput As String
  Dim sPrefix As String
  Dim STR_LENGTH As Integer
  Dim bFirstTime As Boolean

' ------------------------------------------------------------
' Initialize variables
' ------------------------------------------------------------
  sTmp1 = Trim(sInputText)
  sTmp2 = ""
  sTmp3 = ""
  iPos = 1
  
' ------------------------------------------------------------
' Determine the type of indent on the line format
' ------------------------------------------------------------
  Select Case iType
         Case 1:   ' Initial module or routine
              sPrefix = "' " & Space(15)
              STR_LENGTH = 58
              sOutput = sDescription
         
         Case 2:   ' append module or routine
              sPrefix = "' " & Space(30)
              STR_LENGTH = 45
              sOutput = sDescription
         
         Case 3:   ' General comments in the code
              sPrefix = "' "
              STR_LENGTH = 58
              sOutput = sPrefix
              sDescription = sPrefix
  End Select
  
' ------------------------------------------------------------
' If no input data string then leave
' ------------------------------------------------------------
  If Len(sTmp1) = 0 Then
      PrepareData = sOutput
      Exit Function
  End If
  
' ------------------------------------------------------------
' If input data string is less than the record length, we
' will still search for natural breaks where the user may
' have pressed ENTER
' ------------------------------------------------------------
  If Len(sTmp1) <= STR_LENGTH Then
      
      sTmp1 = sOutput & sTmp1                   ' Prefix the original string
      
      ' Search for embedded carriage returns and linefeeds
      iPos = InStr(1, sTmp1, vbCrLf)
  
      Do While iPos > 0
          sTmp2 = Left(sTmp1, iPos + 1)         ' save the first part to the break
          sTmp3 = Mid(sTmp1, iPos + 2)          ' save the last part after the break
          sTmp1 = sTmp2 & sPrefix & sTmp3       ' insert the comment in the middle
          iPos = InStr(iPos + 2, sTmp1, vbCrLf) ' look for the next break
      Loop
      sOutput = sTmp1                           ' Save reformatted data
      GoTo Normal_Exit
  End If
  
' ------------------------------------------------------------
' if the data string is greater than the record length then
' we will loop thru the data string and perform natural
' line breaks.
' ------------------------------------------------------------
  bFirstTime = True
  sTmp3 = ""
  sOutput = ""
      
  Do
      sTmp2 = Left(sTmp1, STR_LENGTH + 1)           ' capture the record length + 1
      iLen = Len(sTmp2)                             ' Save length of data string
                   
      If iLen <= STR_LENGTH Then
      
          sTmp1 = sPrefix & sTmp2                   ' Prefix the original string
          
          ' Search for embedded carriage returns and linefeeds
          iPos = InStr(1, sTmp1, vbCrLf)
        
          Do While iPos > 0
              sTmp2 = Left(sTmp1, iPos + 1)         ' save the first part to the break
              sTmp3 = Mid(sTmp1, iPos + 2)          ' save the last part after the break
              sTmp1 = sTmp2 & sPrefix & sTmp3       ' insert the comment in the middle
              iPos = InStr(iPos + 2, sTmp1, vbCrLf) ' look for the next break
          Loop
          sTmp3 = sTmp1                             ' Save reformatted data
          sTmp1 = ""                                ' empty original variable
      Else
          DoEvents
          ' parse backwards thru the data string
          ' and find a good spot to break this line
          For i = iLen To 1 Step -1
              ' look for a carriage return linefeed combination
              iPos = InStr(1, sTmp2, vbCrLf)
              If iPos > 0 Then
                  If bFirstTime Then
                      bFirstTime = False
                      sTmp3 = sDescription & Left(sTmp2, iPos + 1)
                  Else
                      sTmp3 = sPrefix & Left(sTmp2, iPos + 1)
                  End If
                  
                  sTmp1 = Mid(sTmp1, iPos + 2) ' resize the input string
                  Exit For
              End If
              
              sChar = Mid(sTmp2, i, 1)  ' capture 1 char at a time
              
              ' look for a valid line break character
              If InStr(" ,;:/\|?.><]})=-", sChar) > 0 Then
                  If bFirstTime Then
                      bFirstTime = False
                      sTmp3 = sDescription & Left(sTmp2, i) & vbCrLf
                  Else
                      sTmp3 = sPrefix & Left(sTmp2, i) & vbCrLf
                  End If
                  
                  sTmp1 = Mid(sTmp1, i + 1) ' resize the input string
                  Exit For
              End If
          Next
                    
          ' if this is one long string of unbreakable data,
          ' capture just the first 55 characters
          If Len(sTmp3) = 0 Then
              If bFirstTime Then
                  bFirstTime = False
                  sTmp3 = sDescription & Left(sTmp2, STR_LENGTH) & vbCrLf
              Else
                  sTmp3 = sPrefix & Left(sTmp2, STR_LENGTH) & vbCrLf
              End If
              
              sTmp1 = Mid(sTmp1, STR_LENGTH + 1) ' reformat the input string
          End If
      End If
      
      sOutput = sOutput & sTmp3     ' append to output string
      sTmp3 = ""                    ' empty the holding area
      
      If Len(sTmp1) <= 0 Then       ' if no more data then leave
          Exit Do
      End If
      
  Loop
      
Normal_Exit:
' -------------------------------------------------------------
' Return the formatted output
' -------------------------------------------------------------
  PrepareData = sOutput

End Function

Private Sub Form_Load()

' -------------------------------------------------------------
' Define local variables
' -------------------------------------------------------------
  Dim sMsg As String

' -------------------------------------------------------------
' Center form on the screen and clear the text boxes
' -------------------------------------------------------------
  CenterForm frmComments
  ClearBoxes

' -------------------------------------------------------------
' Preload the combo box
' -------------------------------------------------------------
  With cboType
       .Clear
       .AddItem "Module Initial"
       .AddItem "Module Append"
       .AddItem "Routine Initial"
       .AddItem "Routine Append"
       .AddItem "General Initial"
       .AddItem "General Append"
       .ListIndex = 0
  End With
  cboType_Click   ' set up the default display
  
' -------------------------------------------------------------
' Set the defaults
' -------------------------------------------------------------
  bInitialLoad = True
  iDelay = 2      ' default delay value
  txtDelay = 2
  sName = ""
  sDate = ""
  
' -------------------------------------------------------------
' Prepare the paste message
' -------------------------------------------------------------
  sMsg = vbCrLf
  sMsg = sMsg & "    Your data has been copied to the" & vbCrLf
  sMsg = sMsg & "    clipboard.  To paste, select the" & vbCrLf
  sMsg = sMsg & "    position in your VB application" & vbCrLf
  sMsg = sMsg & "    with your cursor.  Then use" & vbCrLf
  sMsg = sMsg & "                   CTRL+V" & vbCrLf
  sMsg = sMsg & "    to insert the data at that position."
  
  lblMsg = sMsg
  
End Sub

Private Sub Form_Unload(Cancel As Integer)

' -------------------------------------------------------------
' Shut this application down completely
' -------------------------------------------------------------
  Clipboard.Clear              ' empty the data in the clipboard
  Unload Me                    ' deactivate this Form object
  Set frmComments = Nothing    ' release this object from memory
  
End Sub

Private Sub txtComments_GotFocus(Index As Integer)

' -------------------------------------------------------------
' During the initial loading of this form, do not highlight
' anything.
' -------------------------------------------------------------
  If bInitialLoad Then
      bInitialLoad = False
      Exit Sub
  End If
  
' -------------------------------------------------------------
' Highlight all the text in the box
' -------------------------------------------------------------
  Select Case Index
         Case 1, 2:
              With txtComments(Index)
                   .SelStart = 0
                   .SelLength = Len(.Text)
              End With
  End Select
  
End Sub

Private Sub txtComments_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

' -------------------------------------------------------------
' Define local variables
' -------------------------------------------------------------
  Dim CtrlDown As Integer
  Dim PressedKey As Integer
  
' -------------------------------------------------------------
' Initialize  variables
' -------------------------------------------------------------
  CtrlDown = (Shift And vbCtrlMask) > 0   ' Define control key
  PressedKey = Asc(UCase(Chr(KeyCode)))   ' Convert to uppercase
    
' -------------------------------------------------------------
' Check to see if it is okay to make changes
' -------------------------------------------------------------
  If CtrlDown And PressedKey = 88 Then
      ' Ctrl + X was pressed
      Edit_Cut
  ElseIf CtrlDown And PressedKey = 67 Then
      ' Ctrl + C was pressed
      Edit_Copy
  ElseIf CtrlDown And PressedKey = 86 Then
      ' Ctrl + V was pressed
      Edit_Paste
  ElseIf PressedKey = 16 Then
      ' Delete key was pressed
      Edit_Delete
  End If

End Sub

Private Sub txtComments_KeyPress(Index As Integer, KeyAscii As Integer)

' -------------------------------------------------------------
' nullify keystroke of invalid characters
' -------------------------------------------------------------
  Select Case KeyAscii
         Case 8, 13, 32 To 126: Exit Sub      ' good values
         Case 0 To 7, 9 To 12, 14 To 31, 127 To 255: KeyAscii = 0 ' bad values
  End Select
  
End Sub

Private Sub txtComments_LostFocus(Index As Integer)

' -------------------------------------------------------------
' Save the users name during this session
' -------------------------------------------------------------
  Select Case Index
         Case 1:  ' Date of modification
              If Len(Trim(txtComments(1))) = 0 Then
                  txtComments(1) = StrConv(Format(Now, "dd-MMM-yyyy"), vbUpperCase)
              ElseIf Not IsDate(txtComments(1)) Then
                  txtComments(1) = StrConv(Format(Now, "dd-MMM-yyyy"), vbUpperCase)
              Else
                  txtComments(1) = StrConv(txtComments(1), vbUpperCase)
              End If
              
              If Len(Trim(txtComments(1))) > 0 Then
                  sDate = Trim(txtComments(1))
              Else
                  txtComments(1) = sDate
              End If
              
         Case 2:  ' Name of current user
              If Len(Trim(txtComments(2))) > 0 Then
                  sName = Trim(txtComments(2))
              Else
                  txtComments(2) = sName
              End If
  End Select
  
End Sub

Private Sub txtDelay_GotFocus()

' -------------------------------------------------------------
' Highlight all the text in the box
' -------------------------------------------------------------
  With txtDelay
       .SelStart = 0
       .SelLength = Len(.Text)
  End With
  
End Sub

Private Sub txtDelay_KeyPress(KeyAscii As Integer)

' -------------------------------------------------------------
' nullify keystroke of invalid characters
' -------------------------------------------------------------
  Select Case KeyAscii
         Case 8, 48 To 57: Exit Sub       ' good values
         Case Else: KeyAscii = 0 ' bad values
  End Select
  
End Sub

Private Sub txtDelay_LostFocus()
  
' -------------------------------------------------------------
' Make sure we have a delay value
' -------------------------------------------------------------
  If Len(Trim(txtDelay)) = 0 Then
      iDelay = 2      ' default value
      txtDelay = iDelay
  Else
      iDelay = CInt(txtDelay)
  End If
  
End Sub
