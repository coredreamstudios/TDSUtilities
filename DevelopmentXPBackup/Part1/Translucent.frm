VERSION 5.00
Begin VB.Form Translucent 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Translucent"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame UseThese 
      Height          =   2205
      Left            =   472
      TabIndex        =   0
      Top             =   308
      Width           =   4560
      Begin VB.TextBox txtTransparency 
         Height          =   360
         Left            =   2310
         TabIndex        =   2
         Top             =   360
         Width           =   1845
      End
      Begin VB.CommandButton cmdTranslucent 
         Caption         =   "&Set Transparency"
         Height          =   825
         Left            =   345
         TabIndex        =   1
         Top             =   1065
         Width           =   3825
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Transparency (0 - 255)"
         Height          =   360
         Left            =   285
         TabIndex        =   3
         Top             =   420
         Width           =   1725
      End
   End
End
Attribute VB_Name = "Translucent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim g_nTransparency As Integer

Private Sub cmdTranslucent_Click()
    On Error GoTo ErrorRtn
    'value between 0 and 255
    g_nTransparency = txtTransparency.Text
    If g_nTransparency < 0 Then g_nTransparency = 0
    If g_nTransparency > 255 Then g_nTransparency = 255
    SetTranslucent Me.hwnd, g_nTransparency
    Exit Sub
ErrorRtn:
    MsgBox Err.Description & " Source : " & Err.Source
    
    

End Sub




Private Sub Form_Load()


'initialize
    txtTransparency.Text = 150
    g_nTransparency = 150
    Me.BackColor = RGB(0, 0, 255)
    
End Sub
