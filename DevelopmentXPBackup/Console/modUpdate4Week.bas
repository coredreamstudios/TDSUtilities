Attribute VB_Name = "modUpdate4Week"
Public Sub Main()
    
    Dim strInput As String
    
    Con.Initialize
    Con.Title = "Update 4 Week Usage"
    
    If VBA.Command = "startup" Then
        Con.WriteLine "command line parsed"
    End If
    
    Con.WriteLine "Hello World", True
    Con.WriteLine "This is the next lline before the next lline", True
    
    strInput = Con.ReadLine()
    
    Con.ForeColor = conGreenHi
    Con.WriteLine "this is the next text", True
    Con.ForeColor = conWhiteHi
    
    strInput = Con.ReadLine()
    
    Con.ForeColor = conWhite
    
    If Con.Compiled Then Con.PressAnyKey
    
End Sub
