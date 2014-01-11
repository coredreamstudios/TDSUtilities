Attribute VB_Name = "ConsoleApp"
'Requires a reference to Microsoft Scripting Runtime.
Sub Main()
    Dim FSO As New Scripting.FileSystemObject
    Dim sin As Scripting.TextStream
    Dim sout As Scripting.TextStream
    Dim strWord As String
    
    Set sin = FSO.GetStandardStream(StdIn)
    Set sout = FSO.GetStandardStream(StdOut)
    sout.WriteLine "Hello!"
    sout.WriteLine "What's the word?"
    strWord = sin.ReadLine()
    sout.WriteLine "So, the word is " & strWord
    sout.WriteLine "This is the next line of the test app"
    strWord = sin.ReadLine()
    Set sout = Nothing
    Set sin = Nothing
End Sub

