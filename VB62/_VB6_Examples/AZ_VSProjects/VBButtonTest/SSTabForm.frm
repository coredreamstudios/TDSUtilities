VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "VSFLEX3.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form SSTabForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SSTab Demo"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7830
   LinkTopic       =   "Form2"
   ScaleHeight     =   5010
   ScaleWidth      =   7830
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   3720
      Width           =   2055
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   6165
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "SSTabForm.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "List1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "SSTabForm.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TreeView1"
      Tab(1).Control(1)=   "Command2"
      Tab(1).Control(2)=   "Command3"
      Tab(1).Control(3)=   "Command4"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "SSTabForm.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "ProgressBar1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "vsFlexArray1"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin vsFlexLib.vsFlexArray vsFlexArray1 
         Height          =   1455
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   7095
         _Version        =   196608
         _ExtentX        =   12515
         _ExtentY        =   2566
         _StockProps     =   228
         Appearance      =   1
         ConvInfo        =   1413783674
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   3000
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Command3"
         Height          =   375
         Left            =   -70680
         TabIndex        =   6
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   375
         Left            =   -72600
         TabIndex        =   5
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   375
         Left            =   -74640
         TabIndex        =   4
         Top             =   2400
         Width           =   1815
      End
      Begin ComctlLib.TreeView TreeView1 
         Height          =   1335
         Left            =   -74640
         TabIndex        =   3
         Top             =   720
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   2355
         _Version        =   327682
         Style           =   7
         Appearance      =   1
      End
      Begin VB.ListBox List1 
         Height          =   1815
         Left            =   -74640
         TabIndex        =   2
         Top             =   720
         Width           =   6735
      End
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   810
      Left            =   0
      TabIndex        =   9
      Top             =   4200
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   1429
      _CBWidth        =   7815
      _CBHeight       =   810
      _Version        =   "6.7.8988"
      MinHeight1      =   360
      Width1          =   2880
      NewRow1         =   0   'False
      MinHeight2      =   360
      Width2          =   1440
      NewRow2         =   -1  'True
      MinHeight3      =   360
      Width3          =   1440
      NewRow3         =   0   'False
   End
End
Attribute VB_Name = "SSTabForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    Unload Me
    
End Sub
