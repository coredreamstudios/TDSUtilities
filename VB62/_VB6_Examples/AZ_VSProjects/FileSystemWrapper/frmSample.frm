VERSION 5.00
Object = "{A618B6D9-38B5-11D4-965A-0000B497612F}#1.0#0"; "WinDEMOocx.ocx"
Begin VB.Form frmSample 
   Caption         =   "Sample"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7065
   Icon            =   "frmSample.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   7065
   StartUpPosition =   3  'Windows Default
   Begin Win_DEMO.FileControl FileControl1 
      Left            =   4200
      Top             =   1080
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   960
      TabIndex        =   12
      Top             =   3000
      Width           =   2055
   End
   Begin VB.TextBox txttest 
      Height          =   285
      Index           =   5
      Left            =   2160
      TabIndex        =   11
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox txttest 
      Height          =   285
      Index           =   4
      Left            =   2160
      TabIndex        =   10
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox txttest 
      Height          =   285
      Index           =   3
      Left            =   2160
      TabIndex        =   9
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox txttest 
      Height          =   285
      Index           =   2
      Left            =   3840
      TabIndex        =   8
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox txttest 
      Height          =   285
      Index           =   1
      Left            =   2160
      TabIndex        =   7
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox txttest 
      Height          =   285
      Index           =   0
      Left            =   2160
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdTestFileControl 
      Caption         =   "Get All Drives"
      Height          =   255
      Index           =   5
      Left            =   960
      TabIndex        =   5
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton cmdTestFileControl 
      Caption         =   "Windows Sys Dir"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton cmdTestFileControl 
      Caption         =   "Windows Directory"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton cmdTestFileControl 
      Caption         =   "Delete That File"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton cmdTestFileControl 
      Caption         =   "Rename The File"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton cmdTestFileControl 
      Caption         =   "Get The File"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Dim iMessage As Integer

Private Sub cmdTestFileControl_Click(Index As Integer)
Dim colfiles As New Collection
Dim colDrives As New Collection
Dim iCount As Integer
Dim strfoldername As String
Dim boolReturn As Boolean

On Error Resume Next

Select Case Index

 Case 0 ' get file
    If Len(txttest(0)) < 1 Then
      iMessage = MsgBox("You must enter a valid filename in the textbox", vbOKOnly)
      Exit Sub
    End If
  
    FileControl1.GetFile (txttest(0))
  
 Case 1 ' rename
    If Len(txttest(1)) < 1 Then
      iMessage = MsgBox("You must enter a valid existing filename in the textbox", vbOKOnly)
      Exit Sub
    End If
    If Len(txttest(2)) < 1 Then
      iMessage = MsgBox("You must enter a valid new filename in the textbox", vbOKOnly)
      Exit Sub
    End If
    
    Debug.Print FileControl1.RenameFile(txttest(1), txttest(2))
   
 Case 2 ' delete
    If Len(txttest(3)) < 1 Then
      iMessage = MsgBox("You must enter a valid filename in the textbox", vbOKOnly)
      Exit Sub
    End If
  
    boolReturn = FileControl1.DeleteFile(txttest(3), True)
  
 Case 3 'get windows directory
    txttest(4) = FileControl1.WindowsDirectory
    
 Case 4 ' get system directory
    txttest(5) = FileControl1.WindowsSystemDirectory
    
 Case 5 ' get folders
   Set colDrives = FileControl1.GetDrives
   
   List1.Clear
   For iCount = 1 To colDrives.Count
     List1.AddItem colDrives.Item(iCount)
   Next iCount
 
End Select
  
End Sub


Private Sub FileControl1_Error(ErrMsg As String)
  
  iMessage = MsgBox("You have an error: " & ErrMsg, vbOKOnly + vbCritical, "Sample Error Message")
  
End Sub


