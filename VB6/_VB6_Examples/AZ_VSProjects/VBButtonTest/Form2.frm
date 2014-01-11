VERSION 5.00
Begin VB.Form GraphicForm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3015
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6015
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   6015
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   0
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   3015
      ScaleWidth      =   6015
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.CommandButton Command2 
         BackColor       =   &H0000C0C0&
         Caption         =   "_"
         Height          =   310
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000C0C0&
         Caption         =   "X"
         Height          =   310
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2280
         TabIndex        =   2
         Top             =   1680
         Width           =   2450
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2280
         TabIndex        =   1
         Top             =   1200
         Width           =   2450
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         FillColor       =   &H00E0E0E0&
         FillStyle       =   7  'Diagonal Cross
         Height          =   315
         Left            =   1680
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   3615
      End
      Begin VB.Shape Shape3 
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   310
         Left            =   1680
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   3615
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         Height          =   355
         Left            =   1920
         Top             =   2280
         Width           =   2095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1215
      End
      Begin VB.Image Image1 
         Height          =   345
         Left            =   1920
         Picture         =   "Form2.frx":0770
         Top             =   2280
         Width           =   2085
      End
      Begin VB.Shape Shape1 
         Height          =   345
         Left            =   1920
         Top             =   2280
         Width           =   2085
      End
   End
End
Attribute VB_Name = "GraphicForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub Command3_Click()
    
    Me.WindowState = vbMinimized
    
End Sub

Private Sub Command2_Click()
    
    Me.WindowState = vbMinimized
    Me.Caption = "Send Mail"
    
End Sub

Private Sub Form_Activate()
    
    If Me.WindowState = vbNormal Then
        Me.Caption = ""
        Me.Width = 6045
        Me.Height = 3045
    End If
    
End Sub

Private Sub Form_GotFocus()
    
    If Me.WindowState = vbNormal Then
        Me.Caption = ""
        Me.Width = 6045
        Me.Height = 3045
    End If
    
End Sub

Private Sub Form_Load()
    
    'Begin the subclassing of Form1 by passing the
   'address of our new Window Procedure. SetWindowLong
   'returns the address of the original Window Procedure,
   'so we store it in a global variable to restore
   'when stopping the subclassing (typically, in the
   'Unload event).
    'defWindowProc = SetWindowLong(Form2.hWnd, _
                                  GWL_WNDPROC, _
                                  AddressOf WindowProc)
                                  
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Me.ZOrder (0)
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Me.Top = Me.Top + 1
    Me.Left = Me.Left + 1
    
End Sub

Private Sub Form_Paint()
    
    If Me.WindowState = vbNormal Then
        Me.Caption = ""
        Me.Width = 6045
        Me.Height = 3045
    End If
    
End Sub

Private Sub Form_Resize()
    
    If Me.WindowState = vbNormal Then
        Me.Caption = ""
        Me.Width = 6045
        Me.Height = 3045
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'restore the original Window Procedure
   'before unloading the form, or GPF will occur
    'If defWindowProc Then
      ' Call SetWindowLong(Form2.hWnd, _
                          GWL_WNDPROC, _
                          defWindowProc)
       'defWindowProc = 0
    'End If
    
    SwitchBoard.Show

End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Image1.BorderStyle = 1
    
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Image1.BorderStyle = 0
    
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    BeginFRDrag x, y

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 1 Then DoFRDrag x, y

End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    EndFRDrag x, y

End Sub


Private Sub BeginFRDrag(x As Single, y As Single)

    Dim tDc As Long
    Dim sDc As Long
    Dim d As Long
   
   'convert points to POINTAPI struct
    dpoint.x = x
    dpoint.y = y
   
   'get screen area of toolbar
    GetWindowRect GraphicForm.hWnd, fbox
   'screen RECT of toolbar
    TwipsPerPixelX = Screen.TwipsPerPixelX
    TwipsPerPixelY = Screen.TwipsPerPixelY
   
   'get point of MouseDown in screen coordinates
    temp = dpoint
    ClientToScreen GraphicForm.hWnd, temp

    sDc = GetDC(ByVal 0)
    DrawFocusRect sDc, tbox
    d = ReleaseDC(0, sDc)
    oldbox = tbox

End Sub


Private Sub DoFRDrag(x As Single, y As Single)

    Dim tDc As Long
    Dim sDc As Long
    Dim d As Long
    
    tpoint.x = x
    tpoint.y = y

    ClientToScreen GraphicForm.hWnd, tpoint

    tbox.Left = (fbox.Left + tpoint.x / TwipsPerPixelX) - temp.x / TwipsPerPixelX
    tbox.Top = (fbox.Top + tpoint.y / TwipsPerPixelY) - temp.y / TwipsPerPixelY
    tbox.Right = (fbox.Right + tpoint.x / TwipsPerPixelX) - temp.x / TwipsPerPixelX
    tbox.Bottom = (fbox.Bottom + tpoint.y / TwipsPerPixelY) - temp.y / TwipsPerPixelY

    sDc = GetDC(ByVal 0)
    DrawFocusRect sDc, oldbox
    DrawFocusRect sDc, tbox
    d = ReleaseDC(0, sDc)
    oldbox = tbox

End Sub


Private Sub EndFRDrag(x As Single, y As Single)

    Dim tDc As Long
    Dim sDc As Long
    Dim d As Long
    
    Dim newleft As Single
    Dim newtop As Single

    sDc = GetDC(ByVal 0)
    DrawFocusRect sDc, oldbox
    d = ReleaseDC(0, sDc)

    newleft = x + fbox.Left * TwipsPerPixelX - dpoint.x
    newtop = y + fbox.Top * TwipsPerPixelY - dpoint.y
    
    GraphicForm.Move newleft, newtop

End Sub

