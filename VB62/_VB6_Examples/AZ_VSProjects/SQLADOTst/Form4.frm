VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   2010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4260
   LinkTopic       =   "Form4"
   ScaleHeight     =   2010
   ScaleWidth      =   4260
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   1440
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Last Name"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "First Name"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    Dim de As DataEnvironment1
    
    Set de = New DataEnvironment1
    
    de.Connection2.Open
    
    de.Connection2.Execute ("INSERT INTO TEST_NAMES (FNAME , LNAME) VALUES ('" & Text1.Text & "' , '" & Text2.Text & "')")
    
    Set de = Nothing
    
End Sub
