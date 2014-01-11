VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   1080
      TabIndex        =   0
      Top             =   1800
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    ' Create a simple PDF file using the mjwPDF class
    Dim objPDF As New mjwPDF
    
    ' Set the PDF title and filename
    objPDF.PDFTitle = "Test PDF Document"
    objPDF.PDFFileName = App.Path & "\test.pdf"
    
    ' We must tell the class where the PDF fonts are located
    objPDF.PDFLoadAfm = App.Path & "\Fonts"
    
    ' View the PDF file after we create it
    'objPDF.PDFView = True
    
    ' Begin our PDF document
    objPDF.PDFBeginDoc
        ' Set the font name, size, and style
        objPDF.PDFSetFont FONT_ARIAL, 15, FONT_BOLD
        
        ' Set the text color
        objPDF.PDFSetTextColor = vbBlue
        
        ' Set the text we want to print
        objPDF.PDFTextOut "Hello, World! From mjwPDF (www.vb6.us)"
        objPDF.PDFTextOut "this is the new text", 1, 5
    
    ' End our PDF document (this will save it to the filename)
    objPDF.PDFEndDoc
End Sub

