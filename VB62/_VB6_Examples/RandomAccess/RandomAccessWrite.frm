VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form RandomAccessWrite 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Random Access Files"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   5160
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   2400
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtbalance 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox txtLastName 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox txtFirstName 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox txtAccount 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Enter"
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton cmdOpenFile 
      Caption         =   "Open File"
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label lbl4 
      Caption         =   "Balance"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label lbl3 
      Caption         =   "Last Name"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label lbl2 
      Caption         =   "First Name"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Account"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "RandomAccessWrite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  ' Fig. 15.5
   ' Writing data to a random-access file
Option Explicit

Private Type ClientRecord
        accountNumber As Integer
        lastName As String * 15
        firstName As String * 15
        balance As Currency
End Type

Dim mUdtClient As ClientRecord   ' user-defined type

  Private Sub Form_Load()
     cmdEnter.Enabled = False
     cmdDone.Enabled = False
  End Sub

 Sub cmdOpenFile_Click()

     dlgOpen.ShowOpen
     filename = dlgOpen.filename

     If dlgOpen.FileTitle <> "" Then
        ' Open file for writing
        Open filename For Random Access Write As #1 _
           Len = recordLength

        cmdOpenFile.Enabled = False  ' Disable button
        cmdEnter.Enabled = True
        cmdDone.Enabled = True
     Else
        MsgBox ("You must specify a file name")
     End If
  End Sub

  Private Sub cmdEnter_Click()
     mUdtClient.accountNumber = Val(txtAccount.Text)
     mUdtClient.firstName = txtFirstName.Text
     mUdtClient.lastName = txtLastName.Text
     mUdtClient.balance = Val(txtbalance.Text)

     ' Write record to file
     Put #1, mUdtClient.accountNumber, mUdtClient

     Call ClearFields
  End Sub

  Sub cmdDone_Click()
     Close #1
     cmdOpenFile.Enabled = True
     cmdEnter.Enabled = False
     cmdDone.Enabled = False
  End Sub

  Private Sub Form_Terminate()
     Close #1
  End Sub

  Private Sub ClearFields()
     txtAccount.Text = ""
     txtFirstName.Text = ""
     txtLastName.Text = ""
     txtbalance.Text = ""
  End Sub


