VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   10035
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCallQuery 
      Caption         =   "Call SQL Query"
      Height          =   495
      Left            =   8400
      TabIndex        =   8
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtSQL 
      Height          =   285
      Left            =   2160
      TabIndex        =   7
      Text            =   "SELECT * FROM Categories FOR XML AUTO"
      Top             =   1200
      Width           =   5655
   End
   Begin VB.TextBox txtSP 
      Height          =   285
      Left            =   2160
      TabIndex        =   5
      Text            =   "sp_XML_CategoryList"
      Top             =   720
      Width           =   5655
   End
   Begin VB.TextBox txtConn 
      Height          =   285
      Left            =   2160
      TabIndex        =   2
      Text            =   "Provider=SQLOLEDB;Data Source=SERVER1;Initial Catalog=Northwind;User ID=SA;Password=password;trusted_connection=yes"
      Top             =   240
      Width           =   5655
   End
   Begin VB.TextBox txtXMLReturn 
      Height          =   3615
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1680
      Width           =   9615
   End
   Begin VB.CommandButton cmdCallSP 
      Caption         =   "Call Stored Procedure"
      Height          =   495
      Left            =   8400
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "SQL Query Text"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Stored Procedure Name"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Connection String OR DSN"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCallQuery_Click()
  Dim objSP As New cSPXML
  
  objSP.ConnectString = txtConn.Text        'set connection string
  
  'call the query
  If objSP.CallQueryXML(txtSQL.Text) Then
    txtXMLReturn.Text = objSP.XML
  Else
    MsgBox objSP.Message & vbCr & objSP.ConnErrors
  End If
  
  Set objSP = Nothing
End Sub

Private Sub cmdCallSP_Click()
  Dim objSP As New cSPXML
  
  objSP.ConnectString = txtConn.Text    'set the connection string
  
  'call the stored procedure
  If objSP.CallSPXML(txtSP.Text, "<Root>") Then
    txtXMLReturn.Text = objSP.XML     'return xml
  Else
    MsgBox objSP.Message & vbCr & objSP.ConnErrors
  End If
  
  Set objSP = Nothing
End Sub
