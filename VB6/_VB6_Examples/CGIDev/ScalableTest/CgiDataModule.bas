Attribute VB_Name = "CgiDataModule"
Option Explicit

Public Sub GetData()
    
    'SendHeader "ADO CGI Data Test"
    'Send ("<body bgcolor=&H00000003450000567 text=red>")
    
    'Dim db As adodb.Connection
    'Dim rs As adodb.Recordset
    'Dim cm As adodb.Command
    'Dim str As String
   
    'Set db = New adodb.Connection
    'Set rs = New adodb.Recordset
    
    'db.CursorLocation = adUseClient
    'db.ConnectionString = "DRIVER={Microsoft Access Driver (*.mdb)};" & _
                                "DBQ=d:\WebDevelopment\Freshens.mdb;"
    'db.Open
    
    'Set rs.ActiveConnection = db
    'rs.CursorType = adOpenDynamic
    'rs.Open ("Employees")
    'rs.MoveFirst
    
    'On Error GoTo errhandler
    
    'Do Until rs.EOF = True
        
    '    str = rs.Fields("FirstName")
    '    Send ("<center>" & str & "<br>")
    '    rs.MoveNext
    '    If rs.EOF = True Then Exit Do
        
    'Loop
    
    'db.Close
    
    'SendFooter
    'Exit Sub
    
errhandler:   SendFooter
              Exit Sub
    
End Sub

Public Sub Form()

    'SendHeader "VB CGI Form Test"
    
    Send ("<form name=test action=/cgi-bin/TestVBCgi>")
    Send ("<input type=text name=txtFName>  Enter Your Name<br>")
    Send ("<input type=submit>")
    Send ("<br>")
    Send ("<input type=radio name=r1 value=v1>  Radio 1<br>")
    Send ("<input type=radio checked name=r1 value=v2>  Radio 2<br>")
    Send ("<input type=checkbox name=secure value=yes> Support Secure Connections")
    
    Send ("<p>")
    Send ("<select name=operating_system>")
    Send ("<option value=Mac OS>Mac OS")
    Send ("<option value=Windows>Windows")
    Send ("<option value=Linux>Linux")
    Send ("</select>")
    Send ("</p>")

    Send ("</form>")
    
    SendFooter
    
End Sub

Public Sub response()

    Dim name As String
    
    name = GetCgiValue("txtFName")
    
    'SendHeader "CGI Value Test"
    
    Send ("<body bgcolor=&H00000003450000567 text=red>")
    Send ("<h1>Hello " & name & ", from the VB CGI script</h1>")
    
    SendFooter
    
End Sub
