VERSION 5.00
Begin VB.Form frmCheckDB 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Scripter"
   ClientHeight    =   9135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   12975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11640
      TabIndex        =   28
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame FmeDatabase 
      Caption         =   "Database Details"
      Height          =   4695
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   5535
      Begin VB.ListBox lstSPs 
         Height          =   3180
         Left            =   2880
         TabIndex        =   14
         Top             =   1320
         Width           =   2415
      End
      Begin VB.ListBox lstTables 
         Height          =   3180
         Left            =   240
         TabIndex        =   12
         Top             =   1320
         Width           =   2415
      End
      Begin VB.ComboBox cmbName 
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   240
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         Caption         =   "Choose Database"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1515
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         Caption         =   "Tables"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   585
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         Caption         =   "Stored Procedures"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   2880
         TabIndex        =   13
         Top             =   1080
         Width           =   1590
      End
   End
   Begin VB.Frame fmeScript 
      Caption         =   "Generated Script"
      Height          =   8415
      Left            =   5760
      TabIndex        =   26
      Top             =   600
      Width           =   7095
      Begin VB.TextBox txtScript 
         Height          =   7815
         Index           =   0
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Text            =   "frmCheckDB.frx":0000
         Top             =   360
         Width           =   6615
      End
   End
   Begin VB.Frame Fme 
      Caption         =   "Server && Login Details"
      Height          =   1455
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Begin VB.CommandButton cmdLogin 
         Caption         =   "&Login"
         Height          =   375
         Left            =   4080
         TabIndex        =   7
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtServer 
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Text            =   "txtServer"
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtPwd 
         Height          =   285
         Left            =   3840
         TabIndex        =   6
         Text            =   "txtPwd"
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtUID 
         Height          =   285
         Left            =   2280
         TabIndex        =   4
         Text            =   "txtUID"
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   3840
         TabIndex        =   5
         Top             =   360
         Width           =   825
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         Caption         =   "UID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   2280
         TabIndex        =   3
         Top             =   360
         Width           =   345
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         Caption         =   "Server"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   570
      End
   End
   Begin VB.Frame fmeWhatToScript 
      Caption         =   "What To Script"
      Height          =   2535
      Left            =   120
      TabIndex        =   15
      Top             =   6480
      Width           =   5535
      Begin VB.CommandButton cmdGenerateScript 
         Caption         =   "Generate Script"
         Height          =   375
         Left            =   3120
         TabIndex        =   25
         Top             =   2040
         Width           =   1935
      End
      Begin VB.CheckBox chkWhat 
         Caption         =   "CREATE TABLE Statements"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox chkWhat 
         Caption         =   "Stored Procedures"
         Height          =   255
         Index           =   9
         Left            =   3240
         TabIndex        =   24
         Top             =   1560
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkWhat 
         Caption         =   "CREATE DATABASE Statement"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   17
         Top             =   840
         Width           =   2775
      End
      Begin VB.CheckBox chkWhat 
         Caption         =   "Triggers"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   18
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkWhat 
         Caption         =   "Primary Key Constraints"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   19
         Top             =   1560
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkWhat 
         Caption         =   "Foreign Key Constraints"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   20
         Top             =   1920
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkWhat 
         Caption         =   "Indexes"
         Height          =   255
         Index           =   6
         Left            =   3240
         TabIndex        =   21
         Top             =   480
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox chkWhat 
         Caption         =   "DROP statements"
         Height          =   255
         Index           =   7
         Left            =   3240
         TabIndex        =   22
         Top             =   840
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox chkWhat 
         Caption         =   "Object Permissions"
         Height          =   255
         Index           =   8
         Left            =   3240
         TabIndex        =   23
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmCheckDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'* Designed and written by Zarr Team, http://www.zarr.net/vb
'* Please feel free to use this in your own application and we'd
'* appreciate it if you mention Zarr's VB Website at http://www.zarr.net/vb
'***************************************************************************

Option Explicit

Private bIsLoggedOn As Boolean
Private S As SQLDMO.SQLServer

Private Sub chkWhat_Click(Index As Integer)
  If Index = 1 Then chkWhat(1).Value = vbChecked
End Sub

Private Sub cmbName_Click()
  Dim oT As New SQLDMO.Table
  Dim oS As New SQLDMO.StoredProcedure
  Dim oD As New SQLDMO.Database
  Dim iTableWhat As Long
  Dim iSPWhat As Long
  Dim strScript As String
  
  ' if a database has not been clicked on then don't continue
  If cmbName.ItemData(cmbName.ListIndex) = 0 Then
    lstTables.Clear
    lstSPs.Clear
    Exit Sub
  End If

  Me.MousePointer = vbHourglass

  ' set a reference to the selected database
  Set oD = S.Databases(cmbName.ItemData(cmbName.ListIndex) - 1)

  ' loop through each user table in the database, outputting the table name
  With lstTables
    .Clear
    For Each oT In oD.Tables
      If Not oT.SystemObject Then
        .AddItem oT.Owner & "." & oT.Name
        .ItemData(.NewIndex) = .NewIndex + 1
      End If
    Next
  End With
  
  ' loop through each user SP in the database, outputting the SP name
  With lstSPs
    .Clear
    For Each oS In oD.StoredProcedures
      If Not oS.SystemObject Then
        .AddItem oS.Owner & "." & oS.Name
        .ItemData(.NewIndex) = .NewIndex + 1
      End If
    Next
  End With
  
  ' tidy up...
  txtScript(0) = strScript

  Me.MousePointer = vbNormal

End Sub

Private Sub cmdQuit_Click()
  ' if END clicked, unload form and finish
  Unload Me
  End
End Sub

Private Sub cmdLogin_Click()
  Dim oDB As SQLDMO.Database
  
  ' if an error occurs below, then its probably because login details are wrong, so jump to error handler
  On Error GoTo ErrorHandler:
  
  ' try connecting to the SQL Server using the login details entered by user
  Set S = New SQLDMO.SQLServer
  S.Connect txtServer.Text, txtUID.Text, txtPwd.Text
  
  ' if connected then add each database name to the list
  With cmbName
    .Clear
    .AddItem "(Choose)"
    .ItemData(.NewIndex) = 0
    For Each oDB In S.Databases
      .AddItem oDB.Name
      .ItemData(.NewIndex) = .NewIndex + 1
    Next
    .ListIndex = 0
  End With
  
  ' tidy up
  txtScript(0) = ""
  bIsLoggedOn = True
  Call SetFieldStatus

  Exit Sub
  
  
  
ErrorHandler:
  ' code will get here if error above - will generally be if login details wrong
  bIsLoggedOn = False
  Call SetFieldStatus
  MsgBox "Sorry, could not connect to the specified server with those login details.", vbInformation, ""
  
End Sub

Private Sub Form_Load()

  ' get the save the setting in the registry for loading next time
  txtServer = GetSetting("ZarrScript", "Prev", "Server", "")
  txtUID = GetSetting("ZarrScript", "Prev", "UID", "")
  txtPwd = GetSetting("ZarrScript", "Prev", "Pwd", "")
  
  
  ' call procedure to set screen fields to be enabled/disabled as appropriate
  bIsLoggedOn = False
  Call SetFieldStatus
  
End Sub

Private Sub cmdGenerateScript_Click()
  Dim oT As New SQLDMO.Table
  Dim oS As New SQLDMO.StoredProcedure
  Dim oD As New SQLDMO.Database
  Dim iTableWhat As Long
  Dim iSPWhat As Long
  Dim strScript As String
  Const SQL_FILENAME = "C:\SQLScript.sql"
  
  ' if no database selected, don't continue
  If cmbName.ItemData(cmbName.ListIndex) = 0 Then Exit Sub

  Me.MousePointer = vbHourglass

  ' work out what boxes have been ticked and set the parameter for use later on
  iTableWhat = SQLDMO.SQLDMOScript_Default
  iSPWhat = SQLDMO.SQLDMOScript_Default
  If chkWhat(3).Value = vbChecked Then iTableWhat = iTableWhat + SQLDMOScript_Triggers
  If chkWhat(4).Value = vbChecked Then iTableWhat = iTableWhat + SQLDMOScript_DRI_PrimaryKey
  If chkWhat(5).Value = vbChecked Then iTableWhat = iTableWhat + SQLDMOScript_DRI_ForeignKeys + SQLDMO.SQLDMOScript_DRI_UniqueKeys
  If chkWhat(6).Value = vbChecked Then iTableWhat = iTableWhat + SQLDMOScript_Indexes
  If chkWhat(7).Value = vbChecked Then
    iTableWhat = iTableWhat + SQLDMOScript_Drops
    iSPWhat = iSPWhat + SQLDMOScript_Drops
  End If
  If chkWhat(8).Value = vbChecked Then
    iTableWhat = iTableWhat + SQLDMOScript_ObjectPermissions
    iSPWhat = iSPWhat + SQLDMOScript_ObjectPermissions
  End If
    
  ' reference the selected database
  Set oD = S.Databases(cmbName.ItemData(cmbName.ListIndex) - 1)
    
  ' if the 'Check Database' has been ticked, script the database creation
  If chkWhat(2).Value = vbChecked Then
    strScript = oD.Script
  Else
    txtScript(0) = ""
  End If
  
  ' loop through each of the USER tables concatenating the script for the table
  For Each oT In oD.Tables
    If Not oT.SystemObject Then
      strScript = strScript & vbCrLf & oT.Script(iTableWhat)
    End If
  Next
  
  ' if the user ticked to script the SPs, loop through each of the USER SPs concatenating the script for each SP
  If chkWhat(9).Value = vbChecked Then
    For Each oS In oD.StoredProcedures
      If Not oS.SystemObject Then
        strScript = strScript & vbCrLf & oS.Script(iSPWhat)
      End If
    Next
  End If
  
  ' set the field on the form to contain the script
  
  txtScript(0).Text = strScript

  Open "C:\SQLScript.sql" For Output As #1
  Print #1, strScript
  Close #1
  
  Me.MousePointer = vbNormal

  MsgBox "SQL Script written to " & SQL_FILENAME & vbCrLf & vbCrLf & "Extract written to textbox", vbInformation, "Complete"
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  ' save the setting in the registry for loading next time
  SaveSetting "ZarrScript", "Prev", "Server", txtServer.Text
  SaveSetting "ZarrScript", "Prev", "UID", txtUID.Text
  SaveSetting "ZarrScript", "Prev", "Pwd", txtPwd.Text
End Sub

Private Sub SetFieldStatus()
  Dim iCount As Byte

  ' set properties of various controls depending on whether we are logged onto server or not
  fmeScript.Enabled = bIsLoggedOn
  FmeDatabase.Enabled = bIsLoggedOn
  fmeWhatToScript.Enabled = bIsLoggedOn
  
  For iCount = 1 To 9
    chkWhat(iCount).Enabled = bIsLoggedOn
  Next

  Select Case bIsLoggedOn
    Case True
      With cmbName
        .BackColor = vbWindowBackground
      End With
      lstTables.BackColor = vbWindowBackground
      lstSPs.BackColor = vbWindowBackground
      txtScript(0).BackColor = vbWindowBackground
      
    Case False
      With cmbName
        .Clear
        .BackColor = vbInactiveBorder
      End With
      lstTables.Clear
      lstTables.BackColor = vbInactiveBorder
      lstSPs.Clear
      lstSPs.BackColor = vbInactiveBorder
      txtScript(0) = ""
      txtScript(0).BackColor = vbInactiveBorder
  End Select
  
End Sub
