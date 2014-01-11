VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16830
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   16830
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   6360
      TabIndex        =   4
      Top             =   6840
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   735
      Left            =   3240
      TabIndex        =   3
      Top             =   6840
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   6840
      Width           =   2895
   End
   Begin MSDataGridLib.DataGrid dg1 
      Height          =   6495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   16575
      _ExtentX        =   29236
      _ExtentY        =   11456
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   14040
      TabIndex        =   0
      Top             =   6840
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    
    cn.ConnectionString = "Provider=SQLNCLI10;" _
         & "SERVER=184.168.115.147\SQLEXPRESS;" _
         & "Database=EMC;" _
         & "DataTypeCompatibility=80;" _
         & "User Id=sa;" _
         & "Password=Ll2dXs22jJZH;"
         
    cn.Open
    
    rs.Open "SELECT * FROM Ingredient", cn, adOpenKeyset, adLockOptimistic
    
    'MsgBox rs.Fields("StoreNumber")
    
    Set dg1.DataSource = rs
End Sub

Private Sub Command2_Click()
    Dim cn As New ADODB.Connection
    Dim cn2 As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    
    cn.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0; Data Source = C:\InStore\OMInStore.mdb;"
    cn.CursorLocation = adUseClient
    cn.Open
    
    rs.CursorLocation = adUseClient
    rs.LockType = adLockBatchOptimistic
    rs.Source = "SELECT * FROM Ingredient"
    
    Set rs.ActiveConnection = cn
    rs.Open
    Set rs.ActiveConnection = Nothing
    
    cn.Close
    Set cn = Nothing
    
    cn2.ConnectionString = "Provider=SQLNCLI10;" _
         & "SERVER=184.168.115.147\SQLEXPRESS;" _
         & "Database=EMC;" _
         & "DataTypeCompatibility=80;" _
         & "User Id=sa;" _
         & "Password=Ll2dXs22jJZH;"
    
    cn2.Open
    
'    If cn2.State = ADODB.adStateOpen Then
'        MsgBox "Open"
'    End If
    
    Set rs.ActiveConnection = cn2
    
    rs.UpdateBatch
    
    cn2.Close
    Set cn2 = Nothing
    
End Sub

Private Sub Command3_Click()
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    
    cn.ConnectionString = "Provider=SQLNCLI10;" _
         & "SERVER=184.168.115.147\SQLEXPRESS;" _
         & "Database=EMC;" _
         & "DataTypeCompatibility=80;" _
         & "User Id=sa;" _
         & "Password=Ll2dXs22jJZH;"
         
    cn.Open
    
    rs.Open "SELECT DISTINCT StoreNumber FROM Ingredient", cn, adOpenKeyset, adLockOptimistic
    
    Do Until rs.EOF
        List1.AddItem rs.Fields("StoreNumber")
        rs.MoveNext
    Loop
    
    rs.Close
    cn.Close
    
    Set rs = Nothing
    Set cn = Nothing
End Sub
