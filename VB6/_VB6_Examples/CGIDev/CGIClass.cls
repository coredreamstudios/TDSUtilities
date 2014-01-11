VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGIClass"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

  Private Const STD_INPUT_HANDLE = -10&
  Private Const STD_OUTPUT_HANDLE = -11&

  Private Const CGI_AUTH_TYPE         As String = "AUTH_TYPE"
  Private Const CGI_CONTENT_LENGTH    As String = "CONTENT_LENGTH"
  Private Const CGI_CONTENT_TYPE      As String = "CONTENT_TYPE"
  Private Const CGI_GATEWAY_INTERFACE As String = "GATEWAY_INTERFACE"
  Private Const CGI_HTTP_ACCEPT       As String = "HTTP_ACCEPT"
  Private Const CGI_HTTP_REFERER      As String = "HTTP_REFERER"
  Private Const CGI_HTTP_USER_AGENT   As String = "HTTP_USER_AGENT"
  Private Const CGI_PATH_INFO         As String = "PATH_INFO"
  Private Const CGI_PATH_TRANSLATED   As String = "PATH_TRANSLATED"
  Private Const CGI_QUERY_STRING      As String = "QUERY_STRING"
  Private Const CGI_REMOTE_ADDR       As String = "REMOTE_ADDR"
  Private Const CGI_REMOTE_HOST       As String = "REMOTE_HOST"
  Private Const CGI_REMOTE_USER       As String = "REMOTE_USER"
  Private Const CGI_REQUEST_METHOD    As String = "REQUEST_METHOD"
  Private Const CGI_SCRIPT_NAME       As String = "SCRIPT_NAME"
  Private Const CGI_SERVER_NAME       As String = "SERVER_NAME"
  Private Const CGI_SERVER_PORT       As String = "SERVER_PORT"
  Private Const CGI_SERVER_PROTOCOL   As String = "SERVER_PROTOCOL"
  Private Const CGI_SERVER_SOFTWARE   As String = "SERVER_SOFTWARE"

  Private Declare Function Sleep Lib "kernel32" _
  (ByVal dwMilliseconds As Long) As Long

  Private Declare Function stdin Lib "kernel32" Alias "GetStdHandle" _
  (Optional ByVal Handletype As Long = STD_INPUT_HANDLE) As Long

  Private Declare Function stdout Lib "kernel32" Alias "GetStdHandle" _
  (Optional ByVal Handletype As Long = STD_OUTPUT_HANDLE) As Long

  Private Declare Function ReadFile Lib "kernel32" _
  (ByVal hFile As Long, ByVal lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, _
  lpNumberOfBytesRead As Long, Optional ByVal lpOverlapped As Long = 0&) As Long

  Private Declare Function WriteFile Lib "kernel32" _
  (ByVal hFile As Long, ByVal lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, _
  lpNumberOfBytesWritten As Long, Optional ByVal lpOverlapped As Long = 0&) As Long
    
    Private Declare Function GetStdHandle Lib "kernel32" _
    (ByVal nStdHandle As Long) As Long

Private Declare Function SetFilePointer Lib "kernel32" _
   (ByVal hFile As Long, _
   ByVal lDistanceToMove As Long, _
   lpDistanceToMoveHigh As Long, _
   ByVal dwMoveMethod As Long) As Long

Private Declare Function SetEndOfFile Lib "kernel32" _
   (ByVal hFile As Long) As Long
   
Private Type pair
     name As String
     Value As String
End Type

Private tpair() As pair

  Sub Main()

      Dim sReadBuffer As String
      Dim sWriteBuffer As String
      Dim lBytesRead As Long
      Dim lBytesWritten As Long
      Dim hStdIn As Long
      Dim hstdout As Long
      Dim iPos As Integer
      ReDim tpair(0)
      
      ' sleep for one minute so the debugger can attach and set a break
      ' point on line below
      ' Sleep 60000
      SendFunc2
      ''''''''''''sReadBuffer = String$(CLng(Environ$(CGI_CONTENT_LENGTH)), 0)
      
      ' Get STDIN handle
      '''''''''''hStdIn = stdin()
      ' Read client's input
      '''''''''''ReadFile hStdIn, sReadBuffer, Len(sReadBuffer), lBytesRead

       'Find '=' in the name/value pair and parse the buffer
      ''''''''iPos = InStr(sReadBuffer, "=")
      ''''''''sReadBuffer = Mid$(sReadBuffer, iPos + 1)
      'MsgBox sReadBuffer, , "sreadbuffer"
      
      ' Construct and send response to the client
      '''''''''''''''sWriteBuffer = "HTTP/1.0 200 OK" & vbCrLf & "Content-Type: text/html" & vbCrLf & vbCrLf & "Hello, this is a test of the VB Cgi Web Class"
      '''''''''''''''hstdout = stdout()
    
      'response
      'GetData
      'Form
      'SendFunc
         
      'hstdout = stdout()
      '''''''''''''WriteFile hstdout, sWriteBuffer, Len(sWriteBuffer) + 1, lBytesWritten
      'WriteFile hstdout, sReadBuffer, Len(sReadBuffer), lBytesWritten
    
  End Sub
  
  
Sub Send(s As String)

    '======================
    ' Send output to STDOUT
    '======================
    Dim lBytesWritten As Long
    Dim hstdout As Long
    
    hstdout = stdout()
    
    s = (s & vbCrLf)
    WriteFile hstdout, s, Len(s), lBytesWritten, ByVal 0&
    
End Sub

Public Sub SendB(s As String)

    '============================================
    ' Send output to STDOUT without vbCrLf.
    ' Use when sending binary data. For example,
    ' images sent with "Content-type image/jpeg".
    '============================================
    Dim lBytesWritten As Long
    Dim hstdout As Long
    
    WriteFile hstdout, s, Len(s), lBytesWritten, ByVal 0&
    
End Sub

Public Sub GetFormData()
'====================================================
' Get the CGI data from STDIN and/or from QueryString
' Store name/value pairs
'====================================================
Dim sBuff      As String    ' buffer to receive POST method data
Dim lBytesRead As Long      ' actual bytes read by ReadFile()
Dim rc         As Long      ' return code
Dim lContentlength As Long
Dim hStdIn As Long
Dim sFormData As String

'MsgBox "in get form data"
' Method POST - get CGI data from STDIN
' Method GET  - get CGI data from QueryString environment variable
'
hStdIn = stdin()
'MsgBox Environ$(CGI_REQUEST_METHOD), , "reqmeth"
If Environ$(CGI_REQUEST_METHOD) = "POST" Then
   sBuff = String$(lContentlength, Chr$(0))
   rc = ReadFile(hStdIn, ByVal sBuff, lContentlength, lBytesRead, ByVal 0&)
   'ReadFile hStdIn, sBuff, Len(sBuff), lBytesRead
   sFormData = Left$(sBuff, lBytesRead)
   'MsgBox sFormData, , "sformdata"
   ' Make sure posted data is url-encoded
   ' Multipart content types, for example, are not necessarily encoded.
   '
   'MsgBox Environ$(CGI_CONTENT_TYPE)
   'If InStr(0, Environ$(CGI_CONTENT_TYPE), "www-form-urlencoded", 1) Then
    '  StorePairs sFormData
   'End If
End If
'MsgBox Environ$(CGI_QUERY_STRING), , "qstring"
StorePairs Environ$(CGI_QUERY_STRING)
'MsgBox "end of get form data"
End Sub

Public Sub StorePairs(sData As String)
'=====================================================================
' Parse and decode form data and/or query string
' Data is received from server as "name=value&name=value&...name=value"
' Names and values are URL-encoded
'
' Store name/value pairs in array tPair(), and decode them
'
' Note: if an element in the query string does not contain an "=",
'       then it will not be stored.
'
' /cgi-bin/pgm.exe?parm=1   "1" gets stored and can be
'                               retrieved with getCgiValue("parm")
' /cgi-bin/pgm.exe?1        "1" does not get stored, but can be
'                               retrieved with urlDecode(CGI_QueryString)
'
'======================================================================
Dim pointer    As Long      ' sData position pointer
Dim n          As Long      ' name/value pair counter
Dim delim1     As Long      ' position of "="
Dim delim2     As Long      ' position of "&"
Dim lastPair   As Long      ' size of tPair() array
Dim lPairs     As Long      ' number of name=value pairs in sData

lastPair = UBound(tpair)    ' current size of tPair()

pointer = 1
Do
  delim1 = InStr(pointer, sData, "=")
  If delim1 = 0 Then Exit Do
  pointer = delim1 + 1
  lPairs = lPairs + 1
Loop

If lPairs = 0 Then Exit Sub  'nothing to add

' redim tPair() based on the number of pairs found in sData
ReDim Preserve tpair(lastPair + lPairs) As pair

' assign values to tPair().name and tPair().value
pointer = 1
For n = (lastPair + 1) To UBound(tpair)
   delim1 = InStr(pointer, sData, "=") ' find next equal sign
   If delim1 = 0 Then Exit For         ' parse complete

   tpair(n).name = UrlDecode(Mid$(sData, pointer, delim1 - pointer))
   
   delim2 = InStr(delim1, sData, "&")

   ' if no trailing ampersand, we are at the end of data
   If delim2 = 0 Then delim2 = Len(sData) + 1
 
   ' value is between the "=" and the "&"
   tpair(n).Value = UrlDecode(Mid$(sData, delim1 + 1, delim2 - delim1 - 1))
   pointer = delim2 + 1
Next n
End Sub

Public Function UrlDecode(ByVal sEncoded As String) As String
'========================================================
' Accept url-encoded string
' Return decoded string
'========================================================

Dim pointer    As Long      ' sEncoded position pointer
Dim pos        As Long      ' position of InStr target

If sEncoded = "" Then Exit Function

' convert "+" to space
pointer = 1
Do
   pos = InStr(pointer, sEncoded, "+")
   If pos = 0 Then Exit Do
   Mid$(sEncoded, pos, 1) = " "
   pointer = pos + 1
Loop
    
' convert "%xx" to character
pointer = 1

On Error GoTo errorUrlDecode

Do
   pos = InStr(pointer, sEncoded, "%")
   If pos = 0 Then Exit Do
   
   Mid$(sEncoded, pos, 1) = Chr$("&H" & (Mid$(sEncoded, pos + 1, 2)))
   sEncoded = Left$(sEncoded, pos) _
             & Mid$(sEncoded, pos + 3)
   pointer = pos + 1
Loop
On Error GoTo 0     'reset error handling
UrlDecode = sEncoded
Exit Function

errorUrlDecode:
'--------------------------------------------------------------------
' If this function was mistakenly called with the following:
'    UrlDecode("100% natural")
' a type mismatch error would be raised when trying to convert
' the 2 characters after "%" from hex to character.
' Instead, a more descriptive error message will be generated.
'--------------------------------------------------------------------
If Err.Number = 13 Then      'Type Mismatch error
   Err.Clear
   Err.Raise 65001, , "Invalid data passed to UrlDecode() function."
Else
   Err.Raise Err.Number
End If
Resume Next
End Function
 
 Public Function GetCgiValue(cgiName As String) As String
'====================================================================
' Accept the name of a pair
' Return the value matching the name
'
' tPair(0) is always empty.
' An empty string will be returned
'    if cgiName is not defined in the form (programmer error)
'    or, a select type form item was used, but no item was selected.
'
' Multiple values, separated by a semi-colon, will be returned
'     if the form item uses the "multiple" option
'     and, more than one selection was chosen.
'     The calling procedure must parse this string as needed.
'====================================================================
Dim n As Integer
 
For n = 1 To UBound(tpair)
    If UCase$(cgiName) = UCase$(tpair(n).name) Then
       If GetCgiValue = "" Then
          GetCgiValue = tpair(n).Value
       Else             ' allow for multiple selections
          GetCgiValue = GetCgiValue & ";" & tpair(n).Value
       End If
    End If
Next n
End Function

Public Sub SendHeader() '(sTitle As String)

    Send "Status: 200 OK"
    Send "Content-type: text/html" & vbCrLf
    'Send "<HTML><HEAD><TITLE>" & sTitle & "</TITLE></HEAD>"
    
End Sub

Public Sub SendFooter()

    '==================================
    ' standardized footers can be added
    '==================================
    Send "</BODY></HTML>"
    
End Sub

Sub SendFunc2()
    
    SendHeader
    
    Send "<html>"
    Send "<Head>"
    Send "<title>Testing VBCGI</title>"
    Send "</head>"
    Send "<body>"
    Send "<center>"
    Send "<h1>Hello World with the message box</h1>"
    Send "</center>"
    Send "</body>"
    Send "</html>"
    
End Sub

