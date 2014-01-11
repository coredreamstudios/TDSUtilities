VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Password Test"
   ClientHeight    =   2640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2640
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Set Password"
      Height          =   615
      Left            =   1080
      TabIndex        =   1
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private first_time As Boolean

Private Sub Command1_Click()
        
    Form_Activate
   
                    
                     'If mod_pw.pass = "" Then
        'Exit Sub
    'Else
      '  MsgBox "Database opened successfully"
      '  mod_pw.pass = ""
    'End If
    
End Sub

Private Sub Command2_Click()
    
    On Error GoTo errhandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = OpenDatabase(databas, True, False, "MS ACCESS;pwd=" & mod_pw.pass)
    
    DBPWChange.Show vbModal
    
    If Not mod_pw.oldpass = "" Then
        db.NewPassword mod_pw.oldpass, mod_pw.newpass
        mod_pw.oldpass = ""
        mod_pw.newpass = ""
        MsgBox "Password successfully changed.", , "Password Change"
    End If
    
    mod_pw.pass = ""
    
    db.Close
    
    Set db = Nothing
    
    Exit Sub
    
errhandler:         If Err.Number = 3031 Then
                        DBLogin.Show vbModal
                        Set db = OpenDatabase(databas, True, False, "MS ACCESS;pwd=" & mod_pw.pass)
                        Resume Next
                    Else
                        MsgBox Err.Number & "    " & Err.Description, , "Error"
                    End If

End Sub

Private Sub Form_Activate()
    
    On Error GoTo errhandler
    
    If first_time = True Then
        
        first_time = False
        
    Else
    
        Dim db As DAO.Database
        Dim rs As DAO.Recordset
        
        'Set db = OpenDatabase(databas, False, False)
        
        Set db = OpenDatabase(databas, True, False, "MS ACCESS;pwd=" & mod_pw.pass)
        
        MsgBox "Database opened successfully"
        
        db.Close
    
    End If
    
    Exit Sub
    
errhandler:         If Err.Number = 3031 Then
                        DBLogin.Show vbModal
                    Else
                        MsgBox Err.Number & "    " & Err.Description, , "Error"
                    End If
    
End Sub

Private Sub Form_Load()
    
    first_time = True
    
End Sub
