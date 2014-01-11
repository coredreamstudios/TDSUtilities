VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSPXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'execute XML Returning Stored Procedure or query

Dim maParms() As Variant            'variant array to hold parameters
Dim mlngPrmCtr As Long              'the parameter counter, increases each time a param is set
Dim mlngPrmNum As Long              'this keeps track of how many parameters are supposed to used
Dim mstrConnStr As String           'connection string
Dim mstrMessage As String           'holds (some-what) user friendly error message
Dim mstrConnErrors As String        'holds any connection error details
Dim mlngReturnParameter As Long     'the return paramter of the last stored procedure
Dim mblnUseReturn As Boolean        'should @RETURN parameter be added to the param list, DEFAULT = TRUE
Dim mlngConnTimeout As Long         'connection timeout   DEFAULT  = 10
Dim mstrSPName As String            'the stored procedure name
Dim mstrSQL As String               'the sql to execute as a xml query
Dim mstrRootNodeStart  As String    'the starting root node DEFAULT = <SPLists>
Dim mstrRootNodeEnd As String       'the end of the root node   DEFAULT = </SPLists>

Dim mCONN As ADODB.Connection          'class wide connection object
Dim mCMD As ADODB.Command             'class wide command object
Dim mSTMOut As ADODB.Stream         'ADODB.STREAM - used for string xml and streamed xml back to client

'Private parameter array index constants
Private Const PRM_NAME = 0
Private Const PRM_TYPE = 1
Private Const PRM_DIRECTION = 2
Private Const PRM_SIZE = 3
Private Const PRM_VALUE = 4

'DEFAULTS
Private Const DEF_USE_RETURN_PARAM = True
Private Const DEF_CONN_TIMEOUT = 10
Private Const DEF_ROOT_START = "<SPLists>"
Private Const DEF_ROOT_END = "</SPLists>"

'XML Constants
Private Const XML_START = "<?xml version=""1.0""?>"
Private Const XML_DIALECT As String = "{5D531CB2-E6Ed-11D2-B252-00C04F681B71}"
Private Const XML_SP_QUERY_START As String = "<ROOT xmlns:sql='urn:schemas-microsoft-com:xml-sql'><sql:query>"
Private Const XML_SP_QUERY_END As String = "</sql:query></ROOT>"

Private Sub Class_Initialize()
'* Do initial housekeeping here
  Set mCONN = New ADODB.Connection
  
  'set defaults
  mblnUseReturn = DEF_USE_RETURN_PARAM
  mlngConnTimeout = DEF_CONN_TIMEOUT
  mstrRootNodeStart = DEF_ROOT_START
  mstrRootNodeEnd = DEF_ROOT_END
End Sub

Public Property Let ParamCount(ByVal lngParamCount As Long)
'* This Property reinitializes the parameters and makes it possible to reuse this object
On Error Resume Next
  If lngParamCount > 0 Then
    'redim parameter array
    ReDim maParms(lngParamCount - 1, 4)
    mlngPrmNum = lngParamCount
  Else
    mlngPrmNum = 0
  End If
  
  'reset prm ctr
  mlngPrmCtr = 0
End Property

Public Property Let ConnectString(ByVal strConnect As String)
'* Assign The Connection String
  mstrConnStr = strConnect
End Property

Public Property Get ReturnParameterValue() As Long
'**** this is the Return Parameter value of the last SP that was executed with a return parameter
  ReturnParameterValue = mlngReturnParameter
End Property

Public Property Let UseReturnParameter(ByVal blnUseReturn As Boolean)
'**** this is default to true, will fail if not returned zero
  mblnUseReturn = blnUseReturn
End Property

Public Property Get UseReturnParameter() As Boolean
  UseReturnParameter = mblnUseReturn
End Property

Public Property Let ConnTimeout(ByVal lngConnTimeout As Long)
'the connection time out for the connection object
  mlngConnTimeout = lngConnTimeout
End Property

Public Property Let SPName(ByVal strSPName As String)
'set the stored procedure name, can be reset to use different stored proc
  mstrSPName = strSPName
End Property

Public Property Let SQL(ByVal strSQL As String)
'set the sql query text for retrieving query results
  mstrSQL = strSQL
End Property

Public Property Let RootNode(ByVal strRoot As String)
'sets the root node to wrap around the output xml,
On Error Resume Next
  If strRoot <> "" Then
    If Left(strRoot, 1) = "<" Then      'check if it is enclosed already
      mstrRootNodeStart = strRoot
      mstrRootNodeEnd = "</" & Right(strRoot, Len(strRoot) - 1)       'insert the backslash in the tag
    Else
      mstrRootNodeStart = "<" & strRoot & ">"       'add opening tags
      mstrRootNodeEnd = "</" & strRoot & ">"
    End If
  Else
    mstrRootNodeStart = ""
    mstrRootNodeEnd = ""
  End If
End Property

Public Property Get ConnTimeout() As Long
  ConnTimeout = mlngConnTimeout
End Property

Public Property Get Message() As String
'******** This is a semi-formated string used to display error descriptions, numbers, etc...
  Message = mstrMessage
End Property

Public Property Get ConnErrors() As String
'* returns the connection errors that occured, if any
  ConnErrors = mstrConnErrors
End Property

Public Property Get XML() As String
On Error GoTo Hell
  If mSTMOut.State Then
    mSTMOut.Position = 0    'go to the  begining
    
    If mstrRootNodeStart <> "" Then
      XML = mstrRootNodeStart & mSTMOut.ReadText & mstrRootNodeEnd
    Else
      XML = mSTMOut.ReadText
    End If
  Else
    AddToMessage "XML Output Stream Is Not Open"
  End If
  Exit Property
Hell:
  AddToMessage Err.Number & " - " & Err.Description & vbCr & "In Return XML"
End Property

Public Sub AddParam(Optional ByVal strName As String = "", Optional ByVal eType As DataTypeEnum = adEmpty, _
                    Optional ByVal eDirection As ParameterDirectionEnum = adParamInput, _
                    Optional ByVal lngSize As Long, Optional ByVal varValue As Variant)
On Error GoTo Hell
'add another parameter to the internal parameter array
  If mlngPrmCtr < mlngPrmNum Then
    'add the parameter
    maParms(mlngPrmCtr, PRM_NAME) = strName
    maParms(mlngPrmCtr, PRM_TYPE) = eType
    maParms(mlngPrmCtr, PRM_DIRECTION) = eDirection
    maParms(mlngPrmCtr, PRM_SIZE) = lngSize
    maParms(mlngPrmCtr, PRM_VALUE) = varValue
    mlngPrmCtr = mlngPrmCtr + 1
  Else
    AddToMessage "All Of The Parameters Have Been Added Already"
  End If
  
  Exit Sub
Hell:
  AddToMessage CStr(Err.Number) & " - " & Err.Description & vbCr & "In Add Param"
  Exit Sub
End Sub

Public Function CallSPXML(Optional ByVal strProc As String, Optional ByVal strRootNode As String = "") As Boolean
On Error GoTo CallSPXML_Error:
  Dim ctr As Long
  Dim strTempErr As String
  
  If strProc <> "" Then SPName = strProc
  RootNode = strRootNode

  If Not CheckConnect Then Exit Function        'exit if bad connection string
  
  'check param count vs. param number
  If mlngPrmCtr <> mlngPrmNum Then
    'the number of parameters added does not equal the number of params defined
    AddToMessage "Not All Of The Parameters Have Been Set"
    Exit Function
  End If
  
  Set mCMD = New ADODB.Command
  Set mSTMOut = New ADODB.Stream
  
  If Not mCONN.State = adStateOpen Then
    Set mCONN = New ADODB.Connection
    mCONN.ConnectionTimeout = mlngConnTimeout
    mCONN.CursorLocation = adUseClient
    mCONN.Open mstrConnStr      'open connection if not already open
  End If

  Set mCMD.ActiveConnection = mCONN
  mCMD.CommandType = adCmdStoredProc
  mCMD.CommandText = mstrSPName

  'add the return param, if needed (default = True)
  If mblnUseReturn Then
    mCMD.Parameters.Append mCMD.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue)
  End If

  'check for parameters, add if they are there
  If mlngPrmNum > 0 Then
    For ctr = 0 To UBound(maParms, 1)
      mCMD.Parameters.Append mCMD.CreateParameter(maParms(ctr, PRM_NAME), maParms(ctr, PRM_TYPE), maParms(ctr, PRM_DIRECTION), maParms(ctr, PRM_SIZE), maParms(ctr, PRM_VALUE))
    Next
  End If

  mSTMOut.Open               'open output stream
  mCMD.Properties("Output Stream").Value = mSTMOut
  mCMD.Execute , , adExecuteStream        'adExecuteStream = 1024

  'get return parameter value, if used
  If mblnUseReturn Then
    mlngReturnParameter = mCMD.Parameters.Item(0)
    
    If mlngReturnParameter <> 0 Then
      'returned something besides zero
      CallSPXML = False
      AddToMessage "Returned - " & mlngReturnParameter
      Resume CallSPXML_Exit
    End If
  End If

  CallSPXML = True
  Exit Function

CallSPXML_Exit:

  'the command is not set to nothing here because it might be used for return params
  
  Exit Function

CallSPXML_Error:
  strTempErr = Err.Number & " - " & Err.Description
  
  If Not GetConnErrors Then
    'regular run time error
    AddToMessage strTempErr
  End If

  'add return parameter value
  If mblnUseReturn And mCMD.State = adStateOpen Then
    'retrieve return parameter (will always be first parameter added)
    mlngReturnParameter = mCMD.Parameters.Item(0)
    AddToMessage "Returned = " & mlngReturnParameter
  End If
  
  AddToMessage "In Call Stored Procedure - XML"
  CallSPXML = False
  Resume CallSPXML_Exit
End Function

Public Function CallQueryXML(Optional ByVal strSQL As String) As Boolean
On Error GoTo CallQueryXML_Error:
  Dim stmQueryIn As New ADODB.Stream
  Dim ctr As Long
  Dim strTempErr As String
  
  If strSQL <> "" Then SQL = strSQL
  RootNode = ""              'reset root node to "" because it will already have root tage around it
  
  If Not CheckConnect Then GoTo CallQueryXML_Exit        'exit if bad connection string

  Set mCMD = New ADODB.Command
  Set mSTMOut = New ADODB.Stream
  
  If Not mCONN.State = adStateOpen Then
    Set mCONN = New ADODB.Connection
    mCONN.ConnectionTimeout = mlngConnTimeout
    mCONN.CursorLocation = adUseClient
    mCONN.Open mstrConnStr      'open connection if not already open
  End If

  stmQueryIn.Open           'open the inbound stream, add sql
    
  'add the sql query, with the XML needed for SQL Server - constants can be found at the top of the page
  stmQueryIn.WriteText XML_SP_QUERY_START & mstrSQL & XML_SP_QUERY_END
  stmQueryIn.Position = 0
  
  Set mCMD.ActiveConnection = mCONN            'set the active connection
  mCMD.CommandType = adCmdText                 'use admCMDdText as the command type
  
  Set mCMD.CommandStream = stmQueryIn       'set the input stream
  mCMD.Dialect = XML_DIALECT                'set the dialect for xml returning command
  
  mSTMOut.Open                                        'open output stream
  mCMD.Properties("Output Stream").Value = mSTMOut      'set the output stream
  mCMD.Execute , , adExecuteStream                      'executes as a stream

  CallQueryXML = True
  Exit Function

CallQueryXML_Exit:
  'clean up inbound stream
  If stmQueryIn.State Then stmQueryIn.Close
  Set stmQueryIn = Nothing
  
  'this frees the command because it will not be used for return params
  'it is always set to a new command for each query  executed
  Set mCMD = Nothing
  
  Exit Function

CallQueryXML_Error:
  strTempErr = Err.Number & " - " & Err.Description
  
  If Not GetConnErrors Then
    'regular run time error
    AddToMessage strTempErr
  End If

  AddToMessage "In Call Query - XML"
  CallQueryXML = False
  Resume CallQueryXML_Exit
End Function

Public Property Get ReturnVal(ByVal strKey As String) As Variant
On Error GoTo Hell
  Dim strTempErr As String
  
  ReturnVal = mCMD.Parameters.Item(strKey)
  Exit Property
Hell:
  strTempErr = Err.Number & " - " & Err.Description
  If Not GetConnErrors Then
    AddToMessage strTempErr & vbCr & "In Return Parameter Value"
  End If
End Property

Private Sub Class_Terminate()
'* De-reference any instantiated ADODB objects:
On Error Resume Next
  If mSTMOut.State Then mSTMOut.Close
  Set mSTMOut = Nothing
  
  Set mCMD = Nothing
  
  If mCONN.State Then mCONN.Close
  Set mCONN = Nothing
End Sub

Private Sub AddToMessage(ByVal strMessage As String)
'add a string to the message
On Error GoTo Hell
  If Len(mstrMessage) Then
    mstrMessage = mstrMessage & vbCr & strMessage
  Else
    mstrMessage = strMessage
  End If
  Exit Sub
Hell:
  mstrMessage = mstrMessage & Err.Number & "-" & Err.Description & vbCr & "Add To Message"
End Sub

Private Sub AddToConnErrors(ByVal strConnErr As String)
'add a string to the message
On Error GoTo Hell
  If Len(mstrConnErrors) Then
    mstrConnErrors = mstrConnErrors & vbCr & strConnErr
  Else
    mstrConnErrors = strConnErr
  End If
  Exit Sub
Hell:
  mstrMessage = mstrMessage & Err.Number & "-" & Err.Description & vbCr & "Add To Connection Errors"
End Sub

Private Function CheckConnect() As Boolean
On Error Resume Next
  Const MSG_NO_CONN As String = "No Connection String"
  
  'check the connection string
  If Len(mstrConnStr) > 0 Then
    CheckConnect = True
  Else
    'no connection string
    CheckConnect = False
    AddToMessage MSG_NO_CONN
  End If
End Function

Private Function GetConnErrors() As Boolean
'this will add all the connection (if any) to the message, returns true if errors found
On Error Resume Next
  Dim ctr As Long
  
  Const MSG_NOT_XML_RETURNING_SP = "The Stored Procedure Does Not Return XML"
  Const MSG_DUP_PRIMARY_KEY As String = "Duplicate Primary Key - ADO Error"
  Const MSG_FK_VIOLATION As String = "Foreign Key Violation - ADO Error"
  Const MSG_COULD_NOT_CONNECT As String = "Could Not Connect To Server"
          
  If mCONN.Errors.Count > 0 Then
    'there were connection errors
    AddToMessage "Connection Errors" & vbCr & vbCr
    For ctr = 0 To mCONN.Errors.Count - 1
      Select Case mCONN.Errors(ctr).Number
        Case 265946
          AddToMessage MSG_NOT_XML_RETURNING_SP
          AddToConnErrors mCONN.Errors(ctr).Number & " - " & mCONN.Errors(ctr).Description & vbCr
        Case -2147467259    'Connection error
          AddToMessage MSG_COULD_NOT_CONNECT
          AddToConnErrors mCONN.Errors(ctr).Number & " - " & mCONN.Errors(ctr).Description & vbCr
        Case -2147217873    'duplicate primary key
          AddToMessage MSG_DUP_PRIMARY_KEY
          AddToConnErrors mCONN.Errors(ctr).Number & " - " & mCONN.Errors(ctr).Description & vbCr
        Case -2147217873
          AddToMessage MSG_FK_VIOLATION
          AddToConnErrors mCONN.Errors(ctr).Number & " - " & mCONN.Errors(ctr).Description & vbCr
        Case Else
          AddToConnErrors mCONN.Errors(ctr).Number & " - " & mCONN.Errors(ctr).Description & vbCr
      End Select
    Next
    GetConnErrors = True
  End If
End Function

