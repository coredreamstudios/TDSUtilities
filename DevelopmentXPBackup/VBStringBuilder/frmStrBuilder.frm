VERSION 5.00
Begin VB.Form frmStrBuilder 
   Caption         =   "String Builder Demonstration"
   ClientHeight    =   3345
   ClientLeft      =   3210
   ClientTop       =   2010
   ClientWidth     =   5760
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStrBuilder.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3345
   ScaleWidth      =   5760
   Begin VB.CommandButton cmdOtherTest 
      Caption         =   "Other Tests"
      Height          =   435
      Left            =   1080
      TabIndex        =   7
      Top             =   2280
      Width           =   1335
   End
   Begin VB.OptionButton optType 
      Caption         =   "Use String&Builder Class"
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   6
      Top             =   1260
      Width           =   2415
   End
   Begin VB.OptionButton optType 
      Caption         =   "&Standard VB Strings"
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   5
      Top             =   1020
      Value           =   -1  'True
      Width           =   2415
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "&Go"
      Default         =   -1  'True
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   1620
      Width           =   1335
   End
   Begin VB.TextBox txtAppend 
      Height          =   315
      Left            =   1080
      TabIndex        =   3
      Text            =   "http://vbaccelerator.com"
      Top             =   420
      Width           =   4515
   End
   Begin VB.TextBox txtIterations 
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Text            =   "1000"
      Top             =   60
      Width           =   1995
   End
   Begin VB.Label lblAppend 
      Caption         =   "&Append:"
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   480
      Width           =   1635
   End
   Begin VB.Label lblIterations 
      Caption         =   "&Iterations:"
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   1635
   End
End
Attribute VB_Name = "frmStrBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Private Declare Function timeEndPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Private m_lT As Long

Private Sub StandardAppend(ByVal sAppend As String, ByVal lCount As Long)
Dim l As Long
Dim sTheString As String
   StartTiming
   For l = 1 To lCount
      sTheString = sTheString & sAppend
   Next l
   MsgBox "Standard Method:" & EndTiming, vbInformation
   'Debug.Print sTheString
End Sub

Private Sub ClassAppend(ByVal sAppend As String, ByVal lCount As Long)
Dim l As Long
Dim cTheString As New cStringBuilder
   StartTiming
   For l = 1 To lCount
      cTheString.Append sAppend
   Next l
   MsgBox "Class Method:" & EndTiming, vbInformation
   'Debug.Print cTheString.ToString
End Sub

Private Sub StartTiming()
   timeBeginPeriod 1
   m_lT = timeGetTime
End Sub
Private Function EndTiming() As Long
   EndTiming = timeGetTime() - m_lT
   timeEndPeriod 1
End Function

Private Sub cmdGo_Click()
Dim sAppend As String
Dim lAppend As Long
On Error GoTo errorHandler
   sAppend = txtAppend.Text
   lAppend = CLng(txtIterations.Text)
   If optType(0).Value Then
      StandardAppend sAppend, lAppend
   Else
      ClassAppend sAppend, lAppend
   End If
   Exit Sub
errorHandler:
   MsgBox "Error: " & Err.Description, vbInformation
   Exit Sub
End Sub

Private Sub cmdOtherTest_Click()
   Dim buf As New cStringBuilder
   buf.Append "http://vbaccelerator.com/"
   Debug.Print "'" & buf.ToString & "'"
   buf.Insert 7, "www."
   Debug.Print "'" & buf.ToString & "'"
   buf.Remove 7, 4
   Debug.Print "'" & buf.ToString & "'"
   Debug.Print buf.Find(".com")
   Dim i As Integer
   i = 1
   buf.AppendByVal i
   Debug.Print "'" & buf.ToString & "'"
   
End Sub
