VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "MySQL Test"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   2445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Report"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   3960
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Update"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Next Record"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   2040
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get Data"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   840
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   840
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Last"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "First"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "ID"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ws As DAO.Workspace
Dim db As DAO.Database
Dim rs As DAO.Recordset

Dim id As Long
Dim fnme As String
Dim lnme As String

Private Sub Command1_Click()
    
    Text1.Text = rs.Fields("nameID")
    Text2.Text = rs.Fields("first_name")
    Text3.Text = rs.Fields("last_name")
    
End Sub

Private Sub Command2_Click()
        
    rs.AddNew
    rs.Fields("nameID") = Text1.Text
    rs.Fields("first_name") = Text2.Text
    rs.Fields("last_name") = Text3.Text
    rs.Update
    
    Command3_Click
    
End Sub

Private Sub Command3_Click()
    
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    
End Sub

Private Sub Command4_Click()
    
    rs.MoveNext
    
    Command1_Click
    
End Sub

Private Sub Command5_Click()
    
    id = Text1.Text
    fnme = Text2.Text
    lnme = Text3.Text
    
    db.Execute ("UPDATE master_name SET first_name = '" & fnme & "', last_name = '" & lnme & "' WHERE nameID = '" & id & "'")
    
End Sub

Private Sub Command6_Click()
    
    DataReport1.Show
    
End Sub

Private Sub Form_Load()
    
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    Set ws = CreateWorkspace("msql", "crock9l", "7119771197", dbUseODBC)
    Set db = ws.OpenDatabase("VBMySQLTest")
    Set rs = db.OpenRecordset("master_name")
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    rs.Close
    db.Close
    
    Set rs = Nothing
    Set db = Nothing
    
End Sub
