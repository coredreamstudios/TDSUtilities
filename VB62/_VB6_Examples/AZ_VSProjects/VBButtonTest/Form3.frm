VERSION 5.00
Begin VB.Form Toolbar 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   2175
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   2175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "X"
      Height          =   255
      Left            =   1920
      TabIndex        =   15
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton dummy 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3120
      TabIndex        =   13
      Top             =   1560
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   2115
      TabIndex        =   12
      Top             =   2880
      Width           =   2175
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Toolbar 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.CommandButton SSCommand1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   11
      Left            =   1080
      TabIndex        =   11
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton SSCommand1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   10
      Left            =   1080
      TabIndex        =   10
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton SSCommand1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   9
      Left            =   1080
      TabIndex        =   9
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton SSCommand1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   8
      Left            =   360
      TabIndex        =   8
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton SSCommand1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   7
      Left            =   360
      TabIndex        =   7
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton SSCommand1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   6
      Left            =   360
      TabIndex        =   6
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton SSCommand1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   5
      Left            =   1080
      TabIndex        =   5
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton SSCommand1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   4
      Left            =   1080
      TabIndex        =   4
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton SSCommand1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   3
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton SSCommand1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   2
      Left            =   1080
      TabIndex        =   2
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton SSCommand1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton SSCommand1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "Toolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tpoint As POINTAPI
Dim temp  As POINTAPI
Dim dpoint As POINTAPI

Dim fbox As RECT
Dim tbox As RECT
Dim oldbox As RECT

Dim TwipsPerPixelX
Dim TwipsPerPixelY


Private Sub Command1_Click()
    
    Unload Me
    
End Sub

Private Sub Form_Load()

   Dim frameHeight As Long
   Dim frameWidth As Long
   Dim btnSize As Integer

   Toolbar.ScaleMode = 3
   
  'compute the width of the left and right dialog frame
   frameHeight = GetSystemMetrics(SM_CYDLGFRAME) * 2
   
  'compute the width of the top and bottom dialog frame
   frameWidth = GetSystemMetrics(SM_CXDLGFRAME) * 2
   
  'get the size of one of the square toolbar buttons
   btnSize = SSCommand1(0).Width
   
  'set the tool window size
   Toolbar.Height = ((btnSize * 4) + _
                      frameHeight + _
                      Picture1.Height + 1) * Screen.TwipsPerPixelY
   Toolbar.Width = ((btnSize * 3) + _
                     frameWidth) * Screen.TwipsPerPixelX

   Toolbar.ScaleMode = 1

   
  'set the mock titlebar color to that of an active window
   Picture1.BackColor = GetSysColor(COLOR_ACTIVECAPTION)
   Picture1.BackColor = vbYellow

End Sub


Private Sub Form_Activate()
   
  'set to focus to the dummy button
   dummy.SetFocus
  
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Picture1_MouseDown Button, Shift, X, Y
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Picture1_MouseMove Button, Shift, X, Y
    
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Picture1_MouseUp Button, Shift, X, Y
    
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Picture1_MouseDown Button, Shift, X, Y
    
End Sub

Private Sub Picture1_MouseDown(Button As Integer, _
                               Shift As Integer, _
                               X As Single, Y As Single)

    BeginFRDrag X, Y

End Sub


Private Sub Picture1_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, Y As Single)

    If Button = 1 Then DoFRDrag X, Y

End Sub


Private Sub Picture1_MouseUp(Button As Integer, _
                             Shift As Integer, _
                             X As Single, Y As Single)

    EndFRDrag X, Y

End Sub


Private Sub BeginFRDrag(X As Single, Y As Single)

    Dim tDc As Long
    Dim sDc As Long
    Dim d As Long
   
   'convert points to POINTAPI struct
    dpoint.X = X
    dpoint.Y = Y
   
   'get screen area of toolbar
    GetWindowRect Toolbar.hwnd, fbox
   'screen RECT of toolbar
    TwipsPerPixelX = Screen.TwipsPerPixelX
    TwipsPerPixelY = Screen.TwipsPerPixelY
   
   'get point of MouseDown in screen coordinates
    temp = dpoint
    ClientToScreen Toolbar.hwnd, temp

    sDc = GetDC(ByVal 0)
    DrawFocusRect sDc, tbox
    d = ReleaseDC(0, sDc)
    oldbox = tbox

End Sub


Private Sub DoFRDrag(X As Single, Y As Single)

    Dim tDc As Long
    Dim sDc As Long
    Dim d As Long
    
    tpoint.X = X
    tpoint.Y = Y

    ClientToScreen Toolbar.hwnd, tpoint

    tbox.Left = (fbox.Left + tpoint.X / TwipsPerPixelX) - temp.X / TwipsPerPixelX
    tbox.Top = (fbox.Top + tpoint.Y / TwipsPerPixelY) - temp.Y / TwipsPerPixelY
    tbox.Right = (fbox.Right + tpoint.X / TwipsPerPixelX) - temp.X / TwipsPerPixelX
    tbox.Bottom = (fbox.Bottom + tpoint.Y / TwipsPerPixelY) - temp.Y / TwipsPerPixelY

    sDc = GetDC(ByVal 0)
    DrawFocusRect sDc, oldbox
    DrawFocusRect sDc, tbox
    d = ReleaseDC(0, sDc)
    oldbox = tbox

End Sub


Private Sub EndFRDrag(X As Single, Y As Single)

    Dim tDc As Long
    Dim sDc As Long
    Dim d As Long
    
    Dim newleft As Single
    Dim newtop As Single

    sDc = GetDC(ByVal 0)
    DrawFocusRect sDc, oldbox
    d = ReleaseDC(0, sDc)

    newleft = X + fbox.Left * TwipsPerPixelX - dpoint.X
    newtop = Y + fbox.Top * TwipsPerPixelY - dpoint.Y
    
    Toolbar.Move newleft, newtop

End Sub

'Add the following code to the cmdToolBar button's _Click event:

Private Sub cmdToolBar_Click()
   
  'set to focus to the dummy button on a click
   dummy.SetFocus
    
  'perform the action for the button index clicked
   'Select Case Index
      'Case 0
      'Case 1     '... and so on
   'End Select
      
End Sub


'If you added a tiny close button to the
'picturebox, add the following code
'to that button's _Click event:

Private Sub cmdToolEnd_Click()

    Unload Me

End Sub
'--end block--'


Private Sub SSComand1_Click(Index As Integer)

End Sub
