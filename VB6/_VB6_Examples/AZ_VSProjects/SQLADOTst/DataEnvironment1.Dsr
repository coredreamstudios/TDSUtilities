VERSION 5.00
Begin {C0E45035-5775-11D0-B388-00A0C9055D8E} DataEnvironment1 
   ClientHeight    =   11460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12840
   _ExtentX        =   22648
   _ExtentY        =   20214
   FolderFlags     =   5
   TypeLibGuid     =   "{3C37D27D-0C46-4B51-A40E-FF21AA70FA6E}"
   TypeInfoGuid    =   "{16751566-E6DC-43D1-A996-84BA438C1C5F}"
   TypeInfoCookie  =   0
   Version         =   4
   NumConnections  =   2
   BeginProperty Connection1 
      ConnectionName  =   "Connection1"
      ConnDispId      =   1001
      SourceOfData    =   3
      ConnectionSource=   "Provider=SQLOLEDB.1;Password=purefun;Persist Security Info=True;User ID=sa;Initial Catalog=Northwind2;Data Source=imd_test_sql"
      Expanded        =   -1  'True
      IsSQL           =   -1  'True
      QuoteChar       =   34
      SeparatorChar   =   46
   EndProperty
   BeginProperty Connection2 
      ConnectionName  =   "Connection2"
      ConnDispId      =   1012
      SourceOfData    =   3
      ConnectionSource=   "Provider=MSDAORA.1;Password=lou;User ID=lou;Data Source=vis;Persist Security Info=True"
      Expanded        =   -1  'True
      QuoteChar       =   34
      SeparatorChar   =   46
   EndProperty
   NumRecordsets   =   3
   BeginProperty Recordset1 
      CommandName     =   "SQLTst"
      CommDispId      =   1002
      RsDispId        =   1010
      CommandText     =   "SELECT CustomerID, CompanyName, ContactName, ContactTitle, Address, City, Phone FROM Customers"
      ActiveConnectionName=   "Connection1"
      CommandType     =   1
      Locktype        =   3
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   7
      BeginProperty Field1 
         Precision       =   0
         Size            =   5
         Scale           =   0
         Type            =   130
         Name            =   "CustomerID"
         Caption         =   "CustomerID"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   202
         Name            =   "CompanyName"
         Caption         =   "CompanyName"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   202
         Name            =   "ContactName"
         Caption         =   "ContactName"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   202
         Name            =   "ContactTitle"
         Caption         =   "ContactTitle"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   60
         Scale           =   0
         Type            =   202
         Name            =   "Address"
         Caption         =   "Address"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   202
         Name            =   "City"
         Caption         =   "City"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   24
         Scale           =   0
         Type            =   202
         Name            =   "Phone"
         Caption         =   "Phone"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset2 
      CommandName     =   "OracleNames"
      CommDispId      =   1013
      RsDispId        =   1021
      CommandText     =   "SELECT FNAME, LNAME, PHONE FROM TEST_NAMES"
      ActiveConnectionName=   "Connection2"
      CommandType     =   1
      Locktype        =   3
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   3
      BeginProperty Field1 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "FNAME"
         Caption         =   "FNAME"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   25
         Scale           =   0
         Type            =   200
         Name            =   "LNAME"
         Caption         =   "LNAME"
      EndProperty
      BeginProperty Field3 
         Precision       =   38
         Size            =   20
         Scale           =   0
         Type            =   139
         Name            =   "PHONE"
         Caption         =   "PHONE"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset3 
      CommandName     =   "TestDragGrid"
      CommDispId      =   1022
      RsDispId        =   1027
      CommandText     =   "dbo.Categories"
      ActiveConnectionName=   "Connection1"
      CommandType     =   2
      dbObjectType    =   1
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   4
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "CategoryID"
         Caption         =   "CategoryID"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   202
         Name            =   "CategoryName"
         Caption         =   "CategoryName"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   1073741823
         Scale           =   0
         Type            =   203
         Name            =   "Description"
         Caption         =   "Description"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2147483647
         Scale           =   0
         Type            =   205
         Name            =   "Picture"
         Caption         =   "Picture"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
End
Attribute VB_Name = "DataEnvironment1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub rsOracleNames_RecordChangeComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    
    'DataEnvironment1.Recordsets(2).Resync adAffectAllChapters, adResyncAllValues
    
End Sub

Private Sub rsSQLTst_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    
    If Form1.mbEditFlag = True Then
    Connection1.BeginTrans
    Connection1.Execute ("UPDATE Customers SET Address = '" & Form1.txtAddress & "', " & _
                                               "City = '" & Form1.txtCity & "', " & _
                                               "CompanyName = '" & Form1.txtCompanyName & "', " & _
                                               "ContactName = '" & Form1.txtContactName & "', " & _
                                               "ContactTitle = '" & Form1.txtContactTitle & "', " & _
                                               "Phone = '" & Form1.txtPhone & "' " & _
                                               "WHERE CustomerID = '" & Form1.txtCustomerID & "'")
    Form1.mbEditFlag = False
    Connection1.CommitTrans
    End If
    
End Sub

Private Sub rsSQLTst_WillChangeField(ByVal cFields As Long, ByVal Fields As Variant, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    
    
    
    'rsSQLTst.Update
    
End Sub

