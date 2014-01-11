VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form RandomAccessRead 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Random Access Read"
   ClientHeight    =   2850
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   5160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Update Record"
      Height          =   375
      Left            =   3000
      TabIndex        =   12
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   375
      Left            =   3000
      TabIndex        =   11
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton cmdOpenFile 
      Caption         =   "Open File"
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next Record"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox txtAccount 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox txtFirstName 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox txtLastName 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox txtbalance 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   2295
   End
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   2400
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Account"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label lbl2 
      Caption         =   "First Name"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label lbl3 
      Caption         =   "Last Name"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label lbl4 
      Caption         =   "Balance"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuAccounts 
      Caption         =   "&Accounts"
      Begin VB.Menu mnuSelAcct 
         Caption         =   "&Select Account"
      End
   End
End
Attribute VB_Name = "RandomAccessRead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   ' Fig. 15.6
   ' Reading data sequentially from a random-access file
Option Explicit

Private Sub Command1_Click()
        
    acctnum = InputBox("Enter account number to search for.", "Find Account")
    
    Seek #1, 1
           
    Do
        Get #1, , mUdtClient
    Loop Until EOF(1) Or mUdtClient.accountNumber = acctnum
    
    If EOF(1) Then
        MsgBox "Account number not found!", vbExclamation, "Account Error"
    Else
        txtAccount.Text = Str$(mUdtClient.accountNumber)
        txtFirstName.Text = mUdtClient.firstName
        txtLastName.Text = mUdtClient.lastName
        txtbalance.Text = FormatCurrency(mUdtClient.balance)
    End If
    
End Sub

Private Sub Command2_Click()
    
    Close #1
    
    Open filename For Random Access Write As #1 _
           Len = recordLength
           
    If acctnum = "" Then acctnum = txtAccount.Text
    
    mUdtClient.balance = txtbalance.Text
    
    Put #1, acctnum, mUdtClient
    
    Close #1
    
End Sub

  Private Sub Form_Load()
     cmdNext.Enabled = False
     cmdDone.Enabled = False
  End Sub

  Sub cmdOpenFile_Click()

     ' Determine number of bytes in a ClientRecord object
     recordLength = LenB(mUdtClient)

     dlgOpen.ShowOpen
     filename = dlgOpen.filename

     If dlgOpen.FileTitle <> "" Then
        ' Open file for writing
        Open filename For Random Access Read As #1 _
           Len = recordLength
        cmdOpenFile.Enabled = False  ' Disable button
        cmdNext.Enabled = True
        cmdDone.Enabled = True
     Else
        MsgBox ("You must specify a file name")
     End If
  End Sub

  Private Sub cmdNext_Click()
     Dim recordLength As Long
     'Dim filename As String

     ' Determine number of bytes in a ClientRecord object
     recordLength = LenB(mUdtClient)

     dlgOpen.ShowOpen
     filename = dlgOpen.filename
     
     ' Read record from file
     Do
        Get #1, , mUdtClient
     Loop Until EOF(1) Or mUdtClient.accountNumber <> 0

     If EOF(1) Then
        cmdNext.Enabled = False
        Exit Sub
     End If

     If mUdtClient.accountNumber <> 0 Then
        txtAccount.Text = Str$(mUdtClient.accountNumber)
        txtFirstName.Text = mUdtClient.firstName
        txtLastName.Text = mUdtClient.lastName
        'txtbalance.Text = Str$(mUdtClient.balance)
        txtbalance.Text = FormatCurrency(mUdtClient.balance)
     End If
  End Sub

  Sub cmdDone_Click()
     Close #1
     cmdOpenFile.Enabled = True
     cmdNext.Enabled = False
     cmdDone.Enabled = False
     txtAccount.Text = ""
     txtFirstName.Text = ""
     txtLastName.Text = ""
     txtbalance.Text = ""
  End Sub

  Private Sub Form_Terminate()
     Close #1
  End Sub

Private Sub mnuSelAcct_Click()
    
    Form2.Show
    
End Sub
