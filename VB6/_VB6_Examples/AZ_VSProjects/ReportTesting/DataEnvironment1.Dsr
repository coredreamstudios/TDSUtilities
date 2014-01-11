VERSION 5.00
Begin {C0E45035-5775-11D0-B388-00A0C9055D8E} DataEnvironment1 
   ClientHeight    =   7830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4905
   _ExtentX        =   8652
   _ExtentY        =   13811
   FolderFlags     =   1
   TypeLibGuid     =   "{E03C4223-E719-4652-8A91-B8002BEB4DE1}"
   TypeInfoGuid    =   "{C96BF41D-3F93-4A4B-A07B-13CA65E7032A}"
   TypeInfoCookie  =   0
   Version         =   4
   NumConnections  =   1
   BeginProperty Connection1 
      ConnectionName  =   "Connection1"
      ConnDispId      =   1001
      SourceOfData    =   3
      ConnectionSource=   $"DataEnvironment1.dsx":0000
      Expanded        =   -1  'True
      QuoteChar       =   96
      SeparatorChar   =   46
   EndProperty
   NumRecordsets   =   3
   BeginProperty Recordset1 
      CommandName     =   "Command1"
      CommDispId      =   1002
      RsDispId        =   1006
      CommandText     =   "SELECT LastName, FirstName, EmployeeID FROM Employees ORDER BY LastName"
      ActiveConnectionName=   "Connection1"
      CommandType     =   1
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   3
      BeginProperty Field1 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   202
         Name            =   "LastName"
         Caption         =   "LastName"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   202
         Name            =   "FirstName"
         Caption         =   "FirstName"
      EndProperty
      BeginProperty Field3 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "EmployeeID"
         Caption         =   "EmployeeID"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset2 
      CommandName     =   "Command2"
      CommDispId      =   1007
      RsDispId        =   1008
      CommandText     =   "SELECT ShipCountry, OrderID FROM Orders ORDER BY ShipCountry"
      ActiveConnectionName=   "Connection1"
      CommandType     =   1
      Grouping        =   -1  'True
      GroupingName    =   "Command2_Grouping"
      Expanded        =   -1  'True
      SummaryExpanded =   -1  'True
      DetailExpanded  =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   2
      BeginProperty Field1 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   202
         Name            =   "ShipCountry"
         Caption         =   "ShipCountry"
      EndProperty
      BeginProperty Field2 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "OrderID"
         Caption         =   "OrderID"
      EndProperty
      NumGroups       =   2
      BeginProperty Grouping1 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   202
         Name            =   "ShipCountry"
         Caption         =   "ShipCountry"
      EndProperty
      BeginProperty Grouping2 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "OrderID"
         Caption         =   "OrderID"
      EndProperty
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset3 
      CommandName     =   "Command3"
      CommDispId      =   -1
      RsDispId        =   -1
      CommandText     =   $"DataEnvironment1.dsx":0095
      ActiveConnectionName=   "Connection1"
      CommandType     =   1
      GroupingName    =   "Command3_Grouping"
      RelateToParent  =   -1  'True
      ParentCommandName=   "Command2"
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   5
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "OrderID"
         Caption         =   "OrderID"
      EndProperty
      BeginProperty Field2 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "ProductID"
         Caption         =   "ProductID"
      EndProperty
      BeginProperty Field3 
         Precision       =   5
         Size            =   2
         Scale           =   0
         Type            =   2
         Name            =   "Quantity"
         Caption         =   "Quantity"
      EndProperty
      BeginProperty Field4 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "UnitPrice"
         Caption         =   "UnitPrice"
      EndProperty
      BeginProperty Field5 
         Precision       =   7
         Size            =   4
         Scale           =   0
         Type            =   4
         Name            =   "Discount"
         Caption         =   "Discount"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   1
      BeginProperty Relation1 
         ParentField     =   "OrderID"
         ChildField      =   "OrderID"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      AggregateCount  =   0
   EndProperty
End
Attribute VB_Name = "DataEnvironment1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DataEnvironment_Initialize()

Const Var = adSchemaColumns
Connection1.OpenSchema ("" & Var & "")
End Sub

Private Sub DataEnvironment_Terminat()
End Sub

