VERSION 5.00
Begin VB.Form RandomAccessSwitchboard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Switchboard"
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Read Records"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Write Records"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "RandomAccessSwitchboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    RandomAccessWrite.Show
    
End Sub

Private Sub Command2_Click()
    
    RandomAccessRead.Show
    
End Sub
