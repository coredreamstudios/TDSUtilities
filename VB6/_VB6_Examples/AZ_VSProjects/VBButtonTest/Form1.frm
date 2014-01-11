VERSION 5.00
Begin VB.Form SwitchBoard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   3045
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Caption         =   "ADO Chart Form"
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   4560
      Width           =   2055
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Control Test"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   480
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "View Contol Test Form"
      Top             =   3960
      UseMaskColor    =   -1  'True
      Width           =   2055
   End
   Begin VB.CommandButton Command6 
      Caption         =   "SSTab Form"
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   3360
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Tab Strip Form"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Toolbar Form"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Display Modes"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   2160
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Form With X not Active"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Graphical Form"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   2055
   End
End
Attribute VB_Name = "SwitchBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
        
    Me.Hide
    GraphicForm.Show
    
End Sub

Private Sub Command2_Click()
    
    NoX.Show
    
End Sub

Private Sub Command3_Click()
    
    DisplayModesForm.Show
    
End Sub

Private Sub Command4_Click()
    
    Toolbar.Show
    
End Sub

Private Sub Command5_Click()
    
    TabStripForm.Show
    
End Sub

Private Sub Command6_Click()
    
    SSTabForm.Show
    
End Sub

Private Sub Command7_Click()
    
    ControlTest.Show
    
End Sub

Private Sub Command8_Click()
    
    MSChartForm.Show
    
End Sub

Private Sub Form_Load()
    
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
End Sub

Private Sub Form_Terminate()
    
    End
    
End Sub

Private Sub Image1_Click()
    
    Toolbar.Show
    
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Image1.BorderStyle = 1
    
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Image1.BorderStyle = 0
    
End Sub
