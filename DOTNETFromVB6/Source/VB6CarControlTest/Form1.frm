VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timer 
      Interval        =   500
      Left            =   1680
      Top             =   1320
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim car As VBControlExtender

Private Sub Form_Load()
    Set car = Controls.Add("CarControl.Car", "car", Me)
End Sub

Private Sub Form_Resize()
    car.Left = 100
    car.Width = Me.Width - 300
    car.Top = 100
    car.Height = Me.Height - 700
    car.Visible = True
End Sub

Private Sub timer_Timer()
    ' Randomise the timer
    Randomize
    ' Generate random numbers
    car.object.FrontL = Rnd()
    car.object.FrontR = Rnd()
    car.object.RearL = Rnd()
    car.object.RearR = Rnd()
End Sub
