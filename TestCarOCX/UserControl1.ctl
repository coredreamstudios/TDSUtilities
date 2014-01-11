VERSION 5.00
Begin VB.UserControl VB6CarOCX 
   ClientHeight    =   6525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10200
   ScaleHeight     =   6525
   ScaleWidth      =   10200
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4800
      Top             =   2760
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   735
      Left            =   4320
      TabIndex        =   0
      Top             =   5280
      Width           =   3735
   End
End
Attribute VB_Name = "VB6CarOCX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim car As VBControlExtender

Dim ctr As Integer

Private Sub Timer1_Timer()
    ' Randomise the timer
    Randomize
    ' Generate random numbers
    car.object.FrontL = Rnd()
    car.object.FrontR = Rnd()
    car.object.RearL = Rnd()
    car.object.RearR = Rnd()
    
    ctr = ctr + 1
    Label1.Caption = ctr
End Sub

Private Sub UserControl_Initialize()
    Set car = Controls.Add("CarControl.Car", "car")
    
    car.Left = 100
    car.Width = Width - 300
    car.Top = 100
    car.Height = Height - 700
    car.Visible = True
End Sub
