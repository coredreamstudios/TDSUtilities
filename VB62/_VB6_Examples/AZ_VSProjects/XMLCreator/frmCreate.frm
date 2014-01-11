VERSION 5.00
Begin VB.Form frmCreateXMLFile 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sample XML File Creator"
   ClientHeight    =   1230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      Height          =   735
      Left            =   3960
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton cmdGenerateXMLFile 
      Caption         =   "&Generate XML File"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frmCreateXMLFIle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGenerateXMLFile_Click()
  Dim aConn As ADODB.Connection
  Dim aRS   As ADODB.Recordset
  
  ' create a new ADO connection to the database we are going to query
  Set aConn = New ADODB.Connection
  aConn.Open "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & App.Path & "\testdb.mdb;"
  
  ' create a new ADO recordset which is a snapshot of fields and records from a table - can be any table and/or any fields
  Set aRS = New ADODB.Recordset
  aRS.Open "SELECT Surname, Firstname, PkID, DateOfBirth, Country FROM tblStudents order by Surname, Firstname ", aConn
  
  ' call routine to actually generate XML files.
  Call GenerateXMLFile(aRS, "student")
  
  ' cleanup the recordset
  aRS.Close
  Set aRS = Nothing
  
  ' cleanup the connection
  aConn.Close
  Set aConn = Nothing
  
  MsgBox "The XML file 'output.xml' has been created.", vbInformation, "Complete"
  
End Sub

Public Sub GenerateXMLFile(ByVal RS As ADODB.Recordset, Optional ByVal RecordName As String = "item")
  Dim F      As Integer
  Dim oField As ADODB.Field
  
  ' get a free handle for the file and open it for output
  F = FreeFile
  Open App.Path & "\output.xml" For Output As #F
  
  ' start of an XML file should have the plural of what even records are inside the file, eg. students, CDs, cars, trains etc.
  Print #F, "<" & RecordName & "s> "
  
  ' loop through each record in the recordset until we get to the end
  While Not RS.EOF
    ' we have a single record at the moment, so write an XML tag to show we are at the start of a new record (a single record, no plural)
    Print #F, "  <" & RecordName & ">"
    
    ' loop through each field in the recordset writing the fieldname, value and closing fieldname tag
    For Each oField In RS.Fields
      Print #F, "    <" & oField.Name & ">" & oField.Value & "</" & oField.Name & ">"
    Next
    
    ' since we have output all the fields for this record, output as closing record tag
    Print #F, "  </" & RecordName & ">"
    
    ' move to the next record in the list
    RS.MoveNext
  Wend
  
  ' output a closing XML file tag
  Print #F, "</" & RecordName & "s> "
  
  ' close the file
  Close #F
  
End Sub

Private Sub cmdQuit_Click()
  End
End Sub
