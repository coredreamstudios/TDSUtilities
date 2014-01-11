VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9645
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   9645
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      DataField       =   "LNAME"
      DataMember      =   "OracleNames"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   4560
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      DataField       =   "FNAME"
      DataMember      =   "OracleNames"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   4200
      Width           =   2055
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form2.frx":0000
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   6800
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataMember      =   "OracleNames"
      Caption         =   "Testing the Oracle DE Again"
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "FNAME"
         Caption         =   "FNAME"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "LNAME"
         Caption         =   "LNAME"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "PHONE"
         Caption         =   "PHONE"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim de As DataEnvironment1

Private Sub Command1_Click()
        
    'de.Connection2.Execute ("INSERT INTO TEST_NAMES (FNAME , LNAME) VALUES ('" & Text1.Text & "' , '" & Text2.Text & "')")
    
    'Set de = Nothing
    'Set de = New DataEnvironment1
    
    'de.OracleNames
    'grid1.ClearStructure
    'grid1.Refresh
    
    'de.rsOracleNames.MoveLast
    
    'MsgBox de.rsOracleNames.Fields("LNAME")
    
End Sub

Private Sub Form_Initialize()
    
    Dim x As Integer
    
    DataEnvironment1.Connection2.Properties("Prompt") = 1
    
    x = 0
    
    Do
    Debug.Print CStr(DataEnvironment1.Connection1.Properties.Item(x).Name) & "    " & CStr(DataEnvironment1.Connection1.Properties.Item(x).Value)
    Debug.Print x
    x = x + 1
    Loop Until x = 93
    
    DataEnvironment1.OracleNames
    
End Sub

Private Sub Form_Load()
    
    'Set de = DataEnvironment1
    
    'Set de = Nothing
    
End Sub

Private Sub Form_Terminate()
    
    'see if this event fires
    
    Set Form2 = Nothing
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'If de.rsOracleNames.State = adStateOpen Then
    '    de.rsOracleNames.Close
    'End If
    
    'Set de = Nothing
    
    Unload Form2
    
End Sub
