VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form1"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7065
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   7065
   StartUpPosition =   3  'Windows Default
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash s4 
      Height          =   495
      Left            =   4920
      TabIndex        =   4
      Top             =   480
      Width           =   1695
      _cx             =   2990
      _cy             =   873
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash s1 
      Height          =   1095
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3615
      _cx             =   6376
      _cy             =   1931
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash s2 
      Height          =   1215
      Left            =   360
      TabIndex        =   1
      Top             =   1800
      Width           =   3855
      _cx             =   6800
      _cy             =   2143
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash s3 
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   3120
      Width           =   1935
      _cx             =   3413
      _cy             =   661
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   600
      TabIndex        =   3
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
s1.Movie = App.Path & "\new.swf"
s1.Menu = False        'I don't want to see Flash Player Menu
s1.BackgroundColor = &HC0C0C0

s2.Movie = App.Path & "\save.swf"
s2.Menu = False        'I don't want to see Flash Player Menu
s2.BackgroundColor = &HC0C0C0

s3.Movie = App.Path & "\exit.swf"
s3.Menu = False        'I don't want to see Flash Player Menu
s3.BackgroundColor = &HC0C0C0

s4.Movie = App.Path & "\exit.swf"
s4.Menu = False
s4.BackgroundColor = &HC0C0C0

End Sub

Private Sub s1_FSCommand(ByVal command As String, ByVal args As String)
If command = "ButtonClick" Then
MsgBox "You have successfully triggered the Click Event", vbInformation, "Sample Flash Vb Application"
End If

End Sub

Private Sub s2_FSCommand(ByVal command As String, ByVal args As String)
If command = "ButtonClick" Then
MsgBox "Button 2", vbInformation, "Sample Flash Vb Application"
End If

End Sub


Private Sub s3_FSCommand(ByVal command As String, ByVal args As String)
If command = "ButtonClick" Then
MsgBox "Button three", vbInformation, "Sample Flash Vb Application"
End If
End Sub

Private Sub s4_FSCommand(ByVal command As String, ByVal args As String)
If command = "ButtonClick" Then
MsgBox "Hello From the new Flash Button", vbInformation, "Sample Flash Vb Application"
End If
End Sub

