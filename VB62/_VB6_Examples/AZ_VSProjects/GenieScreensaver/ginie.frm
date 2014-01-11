VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000006&
   BorderStyle     =   0  'None
   ClientHeight    =   8565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11580
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   72
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ginie.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "ginie.frx":08CA
   MousePointer    =   99  'Custom
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   840
      Top             =   840
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "jinie"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3735
      Left            =   120
      MouseIcon       =   "ginie.frx":1A8C
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   3720
      Visible         =   0   'False
      Width           =   5415
      WordWrap        =   -1  'True
   End
   Begin VB.Image Img 
      Height          =   1935
      Index           =   2
      Left            =   5640
      MouseIcon       =   "ginie.frx":2C4E
      MousePointer    =   99  'Custom
      Picture         =   "ginie.frx":3E10
      Top             =   2520
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Image Img 
      Height          =   1995
      Index           =   1
      Left            =   5520
      MouseIcon       =   "ginie.frx":4B77
      MousePointer    =   99  'Custom
      Picture         =   "ginie.frx":5D39
      Top             =   360
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Image Img 
      Height          =   1920
      Index           =   0
      Left            =   4320
      MouseIcon       =   "ginie.frx":6B84
      MousePointer    =   99  'Custom
      Picture         =   "ginie.frx":7D46
      Top             =   360
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "TO save YOUR screen"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   0
      MouseIcon       =   "ginie.frx":89A1
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   8040
      Width           =   3615
   End
   Begin VB.Image Image1 
      Height          =   1935
      Left            =   2400
      MouseIcon       =   "ginie.frx":9B63
      MousePointer    =   99  'Custom
      Top             =   2040
      Width           =   1245
   End
   Begin VB.Image backPic 
      Height          =   1695
      Left            =   5160
      Top             =   4080
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''
' Jinie version 1.0.0 the screen saver coded
' by Mokarrabin on 10 the july 1999.
' mokarrabin@mail.com
' http://www.geocities.com/mokarrabin
'



'''''''''''''''''''''''''''''''''''''''''''
' This procedure is for any ordinary day
' of the year to provide a comment by random
' guessing
Public Sub comment()
  Dim choose As Integer
  Randomize
        choose = Int(6 * Rnd + 1)
            If choose = 1 Then
                
                Label2.Caption = "Made in Bangladesh"
            ElseIf choose = 2 Then
                
                Label2.Caption = "make a wish.."
            ElseIf choose = 3 Then
                
                Label2.Caption = "Lets party!"
            ElseIf choose = 4 Then
            
                Label2.Caption = "jinie version 1.0.1"
            ElseIf choose = 5 Then
                Label2.Visible = False
                backPic.Top = 0
                backPic.Left = 0
                backPic.Height = Screen.Height
                backPic.Width = Screen.Width
                backPic.Stretch = True
                backPic.Picture = Img(Int(3 * Rnd + 1) - 1).Picture
                Image1.Visible = False
                backPic.Visible = True
            Else
                Label2.Visible = False
            End If
            
            
            If choose <> 5 Then
                backPic.Visible = False
                Label2.Visible = True
                Image1.Visible = True
            End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''
' This procedure looks for special days of
' the year and provides a comment.

Public Sub dayComment()
    If Day(Now) = 10 And Month(Now) = 7 Then
        Label2.Visible = True
        Label2.Caption = "Today is jinies birth day"
    ElseIf Day(Now) = 21 And Month(Now) = 2 Then
        Label2.Visible = True
        Label2.Caption = "[ Ekushey February ]"
    ElseIf Day(Now) = 26 And Month(Now) = 3 Then
        Label2.Visible = True
        Label2.Caption = "Independence day Bangladesh"
    ElseIf Day(Now) = 16 And Month(Now) = 12 Then
        Label2.Visible = True
        Label2.Caption = "Bijoy dibosh Bangladesh"
    ElseIf Day(Now) = 1 And Month(Now) = 1 Then
        Label2.Visible = True
        Label2.Caption = "HAPPY NEW YEAR"
    ElseIf Day(Now) = 1 And Month(Now) = 5 Then
        Label2.Visible = True
        Label2.Caption = "MAY DAY"
    Else
      Call comment
    End If
End Sub

Private Sub backPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Static count As Integer
      
      If count > 2 Then
         End
      Else
         count = count + 1
      End If
End Sub

Private Sub Form_Load()
    Image1.Picture = Img(0).Picture
    Form1.Height = Screen.Height
    Form1.Width = Screen.Width
       
    Label2.Width = Screen.Width
    Label2.Height = Screen.Height / 2
    Label2.Top = Screen.Height / 2
                
    'If Command$ = "/s" Then
    '   frmConfig.Show   ' display configuration form
    '    Unload Me        ' bypass regular form   End If
    'End If
     
    Call dayComment
    
End Sub

Private Sub Timer1_Timer()
    
    Static X, Y, hit As Integer
    Static xdirec, ydirec As Boolean
    
        Randomize
    If ydirec Then Y = Y + 100 * Int((5 * Rnd) + 1) Else Y = Y - 100 * Int((4 * Rnd) + 1)
        Randomize
    If xdirec Then X = X + 100 * Int((3 * Rnd) + 1) Else X = X - 100 * Int((6 * Rnd) + 1)
    
    If Y > Screen.Height - 1500 Then
        ydirec = False
        
        hit = hit + 1
        If hit = 1 Then
            Image1.Picture = Img(o).Picture
            Label1.Caption = "Jinie"
            Label2.ForeColor = QBColor(Rnd * 10)
            Call dayComment
        ElseIf hit = 5 Then
            Image1.Picture = Img(1).Picture
            Label1.Caption = "Robot"
            Label2.ForeColor = QBColor(Rnd * 10)
            Call dayComment
        ElseIf hit = 10 Then
            Image1.Picture = Img(2).Picture
            Label1.Caption = "Wizard"
            Label2.ForeColor = QBColor(Rnd * 10)
            Call dayComment
        ElseIf hit = 15 Then
            hit = 0
        End If
    End If
    
    
    If (hit = 1 Or hit = 5 Or hit = 10) Then Timer1.Interval = 500 Else Timer1.Interval = 200
    
    If Y < 0 Then ydirec = True
    
    If X > Screen.Width - 1000 Then xdirec = False
    If X < 0 Then xdirec = True
  
    Image1.Top = Y
    Image1.Left = X
    Label1.Left = Screen.Width - X
    
End Sub

Sub Main()
   If App.PrevInstance Then   ' If already running, end the application.
      End
   Else
      Form1.Show             ' Show the screen saver form.
   End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   End
End Sub

' VB3Line: Enter the following two lines as one, single line of code:
Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      Static count As Integer
      
      If count > 2 Then
         End
      Else
         count = count + 1
      End If
End Sub
