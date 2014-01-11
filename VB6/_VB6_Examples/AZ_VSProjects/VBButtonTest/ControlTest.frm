VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "agentctl.dll"
Begin VB.Form ControlTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control Test"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   10710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3000
      TabIndex        =   11
      Top             =   6960
      Width           =   2295
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Left            =   7680
      TabIndex        =   9
      Top             =   6720
      Width           =   2775
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   7680
      TabIndex        =   8
      Top             =   4320
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   22609921
      CurrentDate     =   37547
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1215
      Left            =   360
      TabIndex        =   6
      Top             =   4320
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   2143
      _Version        =   393217
      TextRTF         =   $"ControlTest.frx":0000
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   720
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Documents and Settings\crock9l\My Documents\AccessDatabases\citywideprop.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "cityownedprop"
      Top             =   3720
      Width           =   4695
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "ControlTest.frx":007A
      Height          =   2415
      Left            =   360
      OleObjectBlob   =   "ControlTest.frx":008E
      TabIndex        =   4
      Top             =   1080
      Width           =   4935
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   2415
      Left            =   5280
      TabIndex        =   3
      Top             =   1080
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   4260
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSComctlLib.ImageCombo ImageCombo1 
      Height          =   330
      Left            =   6600
      TabIndex        =   2
      Top             =   240
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Text            =   "ImageCombo1"
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   450
      _Version        =   327682
      Max             =   100
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7845
      Width           =   10710
      _ExtentX        =   18891
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   375
      Left            =   3000
      TabIndex        =   10
      Top             =   6360
      Width           =   2295
      Caption         =   "Hello World"
      Size            =   "4048;661"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   240
      Top             =   7080
   End
   Begin VB.OLE OLE1 
      Class           =   "Excel.Sheet.8"
      Height          =   3015
      Left            =   5850
      OleObjectBlob   =   "ControlTest.frx":0A65
      TabIndex        =   7
      Top             =   1050
      Width           =   4575
   End
End
Attribute VB_Name = "ControlTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ch As String

Private Sub Command1_Click()
    
    Dim x As Integer
    
    Agent1.Characters(ch).Speak "Hello Mike Wheeler"
    
    DoEvents
    
    For x = 1 To 1000
    DoEvents
    Next x

    Agent1.Characters(ch).Speak ("I am the Microsoft Agent Control which you can put into your applications.  Can you imagine how much fun it would be to put me into the Performance Measures application?!")
    
    Agent1.Characters(ch).GestureAt 150, 150
    
    Agent1.Characters(ch).MoveTo 600, 600, 2
    
    For x = 1 To 6000
        DoEvents
    Next x
    
    Agent1.Characters(ch).Speak ("I can move around and evaluate whether the city has met its performance specifications or not!  Then I could chew somebody out for the Mayor and save him the trouble!")
    
    For x = 1 To 12000
        DoEvents
    Next x
    
    Agent1.Characters(ch).Think ("I Wonder if he is buying this or not?")
    
End Sub

Private Sub CommandButton1_Click()
    
    Agent1.Connected = True
    'Agent1.PropertySheet.Visible = True
    'Agent1.ShowDefaultCharacterProperties
    
    Agent1.Characters(ch).Activate
    Agent1.Characters(ch).Show
    
End Sub

Private Sub Form_Load()
    
    ch = ("Genie")
    
    Agent1.Connected = True
    'Agent1.PropertySheet.Visible = True
    Agent1.Characters.Load ch
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Agent1.Characters.Unload (ch)
    Agent1.Connected = False
    
End Sub

Private Sub HScroll1_Change()
    
    MonthView1.Value = MonthView1.Value + 1
    
End Sub

Private Sub Slider1_Scroll()
    
    ProgressBar1.Value = Slider1.Value
    
End Sub

