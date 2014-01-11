Attribute VB_Name = "HelloCGI"
Option Explicit

Sub Main()
    
    Dim s As VBCGI.CGIClass
    
    Set s = New VBCGI.CGIClass
    
    s.SendHeader
    
    s.Send "<html>"
    s.Send "<head>"
    s.Send "</head>"
    s.Send "<body>"
    s.Send "<h1>Hello World from the gis_web_server</h1>"
    s.Send "</body>"
    s.Send "</html>"
    
    Set s = Nothing
    
End Sub
