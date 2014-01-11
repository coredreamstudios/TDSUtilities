VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "SQL ADO Test"
   ClientHeight    =   3615
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   6045
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "<<"
      Height          =   375
      Left            =   480
      TabIndex        =   17
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "<"
      Height          =   375
      Left            =   1440
      TabIndex        =   16
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   ">>"
      Height          =   375
      Left            =   4680
      TabIndex        =   15
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      Height          =   375
      Left            =   3720
      TabIndex        =   14
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox txtPhone 
      DataField       =   "Phone"
      DataMember      =   "SQLTst"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   2070
      TabIndex        =   13
      Top             =   2565
      Width           =   3375
   End
   Begin VB.TextBox txtCity 
      DataField       =   "City"
      DataMember      =   "SQLTst"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   2070
      TabIndex        =   11
      Top             =   2190
      Width           =   2475
   End
   Begin VB.TextBox txtAddress 
      DataField       =   "Address"
      DataMember      =   "SQLTst"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   2070
      TabIndex        =   9
      Top             =   1800
      Width           =   3375
   End
   Begin VB.TextBox txtContactTitle 
      DataField       =   "ContactTitle"
      DataMember      =   "SQLTst"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   2070
      TabIndex        =   7
      Top             =   1425
      Width           =   3375
   End
   Begin VB.TextBox txtContactName 
      DataField       =   "ContactName"
      DataMember      =   "SQLTst"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   2070
      TabIndex        =   5
      Top             =   1050
      Width           =   3375
   End
   Begin VB.TextBox txtCompanyName 
      DataField       =   "CompanyName"
      DataMember      =   "SQLTst"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   2070
      TabIndex        =   3
      Top             =   660
      Width           =   3375
   End
   Begin VB.TextBox txtCustomerID 
      DataField       =   "CustomerID"
      DataMember      =   "SQLTst"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   2070
      TabIndex        =   1
      Top             =   285
      Width           =   825
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Phone:"
      Height          =   255
      Index           =   6
      Left            =   225
      TabIndex        =   12
      Top             =   2610
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "City:"
      Height          =   255
      Index           =   5
      Left            =   225
      TabIndex        =   10
      Top             =   2235
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Address:"
      Height          =   255
      Index           =   4
      Left            =   225
      TabIndex        =   8
      Top             =   1845
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ContactTitle:"
      Height          =   255
      Index           =   3
      Left            =   225
      TabIndex        =   6
      Top             =   1470
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ContactName:"
      Height          =   255
      Index           =   2
      Left            =   225
      TabIndex        =   4
      Top             =   1095
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "CompanyName:"
      Height          =   255
      Index           =   1
      Left            =   225
      TabIndex        =   2
      Top             =   705
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "CustomerID:"
      Height          =   255
      Index           =   0
      Left            =   225
      TabIndex        =   0
      Top             =   330
      Width           =   1815
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuEmp 
         Caption         =   "Show &Employees"
      End
      Begin VB.Menu mnnuShOR 
         Caption         =   "Show &Oracle"
      End
      Begin VB.Menu mnuShOrDa 
         Caption         =   "Show Oracle &Data"
      End
      Begin VB.Menu mnnuInpF 
         Caption         =   "In&put Form"
      End
      Begin VB.Menu mnuTG 
         Caption         =   "&Test Grid"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private de As DataEnvironment1
Public mbEditFlag As Boolean

Private Sub Command1_Click()
    
    On Error GoTo errhandler
    
    de.rsSQLTst.MoveNext
    
    ShowFields
    
    Exit Sub
    
errhandler:         MsgBox Err.Description
                    Resume Next
    
End Sub

Private Sub Command2_Click()
    
    de.rsSQLTst.MoveLast
    
    ShowFields
    
End Sub

Private Sub Command3_Click()
    
    On Error GoTo errhandler
    
    de.rsSQLTst.MovePrevious
    
    ShowFields
    
    Exit Sub
    
errhandler:         MsgBox Err.Description
    
End Sub

Private Sub Command4_Click()
    
    de.rsSQLTst.MoveFirst
    
    ShowFields
    
End Sub

Private Sub Form_Load()
    
    Set de = New DataEnvironment1
    
    de.rsSQLTst.Open
    
    mbEditFlag = False
   
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error GoTo errhandler
    
    de.rsSQLTst.Close
    
    Set de = Nothing
    
    Exit Sub
    
errhandler:         de.rsSQLTst.CancelUpdate
                    Resume Next
    
End Sub

Private Sub mnnuInpF_Click()
    
    Form4.Show
    
End Sub

Private Sub mnnuShOR_Click()
    
    Form2.Show
    
End Sub

Private Sub mnuEmp_Click()
    
    NTWindTst.Show
    
End Sub

Private Sub ShowFields()
    
    On Error GoTo errhandler
    
    Me.txtAddress = de.rsSQLTst.Fields("Address")
    Me.txtCity = de.rsSQLTst.Fields("City")
    Me.txtCompanyName = de.rsSQLTst.Fields("CompanyName")
    Me.txtContactName = de.rsSQLTst.Fields("ContactName")
    Me.txtContactTitle = de.rsSQLTst.Fields("ContactTitle")
    Me.txtCustomerID = de.rsSQLTst.Fields("CustomerID")
    Me.txtPhone = de.rsSQLTst.Fields("Phone")
    
    Exit Sub
    
errhandler:         MsgBox Err.Description
    
End Sub

Private Sub mnuShOrDa_Click()
    
    Form3.Show
    
End Sub

Private Sub mnuTG_Click()
    
    Form5.Show
    
End Sub

Private Sub txtCity_KeyPress(KeyAscii As Integer)
    
    mbEditFlag = True
    
End Sub
