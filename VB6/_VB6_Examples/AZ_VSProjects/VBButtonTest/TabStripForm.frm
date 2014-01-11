VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form TabStripForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tab Strip Demo"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   2895
      Left            =   360
      TabIndex        =   21
      Top             =   600
      Width           =   6495
      Begin VB.FileListBox File1 
         Height          =   2040
         Left            =   2880
         TabIndex        =   24
         Top             =   720
         Width           =   3495
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   2880
         TabIndex        =   23
         Top             =   240
         Width           =   3495
      End
      Begin VB.DirListBox Dir1 
         Height          =   2565
         Left            =   360
         TabIndex        =   22
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   2655
      Left            =   360
      TabIndex        =   10
      Top             =   720
      Width           =   6495
      Begin VB.CheckBox Check7 
         Caption         =   "Check4"
         Height          =   255
         Left            =   3720
         TabIndex        =   20
         Top             =   1920
         Width           =   975
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Check4"
         Height          =   255
         Left            =   4800
         TabIndex        =   19
         Top             =   1920
         Width           =   975
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Check4"
         Height          =   255
         Left            =   3720
         TabIndex        =   18
         Top             =   2280
         Width           =   975
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Check4"
         Height          =   255
         Left            =   2640
         TabIndex        =   17
         Top             =   1920
         Width           =   975
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Option5"
         Height          =   255
         Left            =   480
         TabIndex        =   16
         Top             =   2160
         Width           =   1335
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Option4"
         Height          =   255
         Left            =   480
         TabIndex        =   15
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2640
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   1200
         Width           =   3135
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   375
         Left            =   480
         TabIndex        =   13
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2640
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   480
         Width           =   3135
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   375
         Left            =   480
         TabIndex        =   11
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2175
      Left            =   480
      TabIndex        =   2
      Top             =   840
      Width           =   6135
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   1695
         Left            =   3000
         TabIndex        =   6
         Top             =   240
         Width           =   2775
         Begin VB.CheckBox Check3 
            Caption         =   "Check3"
            Height          =   255
            Left            =   480
            TabIndex        =   9
            Top             =   1320
            Width           =   1575
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Check2"
            Height          =   255
            Left            =   480
            TabIndex        =   8
            Top             =   840
            Width           =   1575
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Check1"
            Height          =   255
            Left            =   480
            TabIndex        =   7
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Option3"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   1560
         Width           =   1695
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   960
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   3840
      Width           =   1935
   End
   Begin MSComctlLib.TabStrip TS 
      Height          =   3375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   5953
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Options"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Commands"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Frames"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "TabStripForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    
   TS.TabIndex = 1
   TS_Click
    
End Sub

Private Sub TS_Click()
    
    If TS.SelectedItem = "Options" Then
        Frame1.Visible = True
        Frame3.Visible = False
        Frame4.Visible = False
    ElseIf TS.SelectedItem = "Commands" Then
        Frame1.Visible = False
        Frame3.Visible = True
        Frame4.Visible = False
    ElseIf TS.SelectedItem = "Frames" Then
        Frame1.Visible = False
        Frame3.Visible = False
        Frame4.Visible = True
    End If
    
End Sub
