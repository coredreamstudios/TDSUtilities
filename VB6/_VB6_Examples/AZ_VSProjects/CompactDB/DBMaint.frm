VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDBMaint 
   BackColor       =   &H00FF0000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compact/Repair Database v2.0"
   ClientHeight    =   1560
   ClientLeft      =   2775
   ClientTop       =   3375
   ClientWidth     =   4455
   Icon            =   "DBMaint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   4455
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   75
      Top             =   1275
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   990
      Left            =   150
      ScaleHeight     =   930
      ScaleWidth      =   4080
      TabIndex        =   0
      Top             =   150
      Width           =   4140
      Begin VB.Label lblDBName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Left            =   150
         TabIndex        =   2
         Top             =   525
         Width           =   3765
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Now performing maintenance on:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   75
         TabIndex        =   1
         Top             =   150
         Width           =   3915
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Freeware by Kenneth Ives"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2175
      TabIndex        =   3
      Top             =   1200
      Width           =   2115
   End
End
Attribute VB_Name = "frmDBMaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function GetDatabase() As Boolean

' ----------------------------------------------------------
' Define local variables
' ----------------------------------------------------------
  Dim sAppPath As String
  Dim sFilters As String
  Dim sCmdLine As String
  Dim iCmdLnLen As Integer

' ----------------------------------------------------------
' Initialize variables
' ----------------------------------------------------------
  sAppPath = "C:\"
  sFilters = "Access Files (*.mdb)|*.mdb" & "All Files (*.*)|*.*"
  sDatabase = ""

' ----------------------------------------------------------
' Get command line arguments.
' ----------------------------------------------------------
  sCmdLine = Command()
  iCmdLnLen = Len(sCmdLine)
 
' ----------------------------------------------------------
' See if there is the name of a database on the
' command line
' ----------------------------------------------------------
  If iCmdLnLen > 0 Then
      sDatabase = sCmdLine
      GoTo Normal_Exit
  End If
  
' ----------------------------------------------------------
' Get the location of the database.  Display the
' File Open dialog box
' ----------------------------------------------------------
  ' Set CancelError is True
  frmDBMaint.CDialog.CancelError = True
  On Error GoTo Cancel_Was_Pressed
  With frmDBMaint.CDialog
       .DialogTitle = "Select database to compact"
       .DefaultExt = "*.mdb"
       .Filter = sFilters
       .flags = cdlOFNHideReadOnly
       .InitDir = sAppPath
       .FilterIndex = 1                 ' Specify default filter
       .FileName = "*.mdb"
       .ShowOpen                        ' Display the Open dialog box
  End With
  
' ---------------------------------------------------------
' Save the name of the item selected
' ---------------------------------------------------------
  sDatabase = CDialog.FileName
  
  
Normal_Exit:
  GetDatabase = True
  Exit Function
  
  
Cancel_Was_Pressed:
  On Error GoTo 0
  GetDatabase = False
  sDatabase = ""

End Function
Public Sub Reset_frmDBMaint()

  
' ------------------------------------------------------------
' Center this form
' ------------------------------------------------------------
  RemoveX frmDBMaint
  CenterForm frmDBMaint
  
' ----------------------------------------------------------
' Get the database
' ----------------------------------------------------------
  If GetDatabase Then
      If ValidDatabase Then
          ' Show the form and start compacting the database
          With frmDBMaint
               .lblDBName.Caption = ShrinkToFit(sDatabase, 40)
               .Show vbModeless
               .Refresh
          End With
          '
          CompactMDB    ' compact the database
      End If
  End If
  
' ----------------------------------------------------------
' finished
' ----------------------------------------------------------
  Unload Me
  End
  
End Sub

Private Function ValidDatabase() As Boolean

' ----------------------------------------------------------
' Remove trainling spaces
' ----------------------------------------------------------
  sDatabase = Trim(sDatabase)
  
' ----------------------------------------------------------
' Is there something there
' ----------------------------------------------------------
  If Len(sDatabase) = 0 Then
      MsgBox "No database selected", vbOKOnly, "No MDB selected"
      ValidDatabase = False
      Exit Function
  End If
  
' ----------------------------------------------------------
' Is this a database
' ----------------------------------------------------------
  If UCase(Right(sDatabase, 4)) <> ".MDB" Then
      MsgBox "This is not a database." & vbLf & UCase(sDatabase), vbOKOnly, "No MDB selected"
      ValidDatabase = False
      Exit Function
  End If

' ----------------------------------------------------------
' Does the database exist
' ----------------------------------------------------------
  If Not ItemExist(sDatabase) Then
      MsgBox "Database cannot be found at this location." & vbLf & sDatabase, _
             vbOKOnly, "No MDB selected"
      ValidDatabase = False
      Exit Function
  End If
  
' ----------------------------------------------------------
' We are ready to go
' ----------------------------------------------------------
  ValidDatabase = True
  
End Function


