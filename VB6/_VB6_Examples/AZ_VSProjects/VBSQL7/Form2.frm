VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   2925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2190
   LinkTopic       =   "Form2"
   ScaleHeight     =   2925
   ScaleWidth      =   2190
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      DataField       =   "TestfieldC"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      DataField       =   "TestfieldB"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      DataField       =   "TestFieldA"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.Data Data1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Documents and Settings\crock9l\My Documents\AccessDatabases\TestFirewall.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TestFirewallTable"
      Top             =   2400
      Width           =   1935
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
