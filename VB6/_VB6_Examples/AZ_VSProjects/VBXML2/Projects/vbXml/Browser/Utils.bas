Attribute VB_Name = "modUtils"
Option Explicit

Public Function OpenFile() As XMLDocument
    Dim oDlg As clsFileDlg
    Dim sFileName As String
    Dim sFileData As String
    Dim hFile As Integer
    
    Set oDlg = New clsFileDlg
    
    If Not oDlg.VBGetOpenFileName(sFileName, "Browser", , , , , _
        "XML Files (*.xml)|*.xml|All (*.*)|*.*") Then
        Exit Function
    End If
    
    hFile = FreeFile
    
    Open sFileName For Input As hFile
    sFileData = Input(LOF(hFile), hFile)
    
    Set OpenFile = New XMLDocument
    Call OpenFile.LoadData(sFileData)
End Function


