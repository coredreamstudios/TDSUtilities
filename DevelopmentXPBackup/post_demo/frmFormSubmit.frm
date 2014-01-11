VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmFormSubmit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Winsock Form Submission"
   ClientHeight    =   10950
   ClientLeft      =   3135
   ClientTop       =   690
   ClientWidth     =   12120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10950
   ScaleWidth      =   12120
   Begin VB.Frame fmeHeaders 
      Caption         =   "Additional Headers"
      Height          =   1935
      Left            =   120
      TabIndex        =   32
      ToolTipText     =   "Use this space to add some custom HTTP headers of your own."
      Top             =   3360
      Width           =   11775
      Begin VB.PictureBox pbxOHeaders 
         BorderStyle     =   0  'None
         Height          =   1575
         Left            =   120
         ScaleHeight     =   1575
         ScaleWidth      =   9495
         TabIndex        =   33
         Top             =   240
         Width           =   9495
         Begin VB.VScrollBar vsbHeaders 
            Enabled         =   0   'False
            Height          =   1335
            Left            =   9000
            Max             =   0
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   120
            Width           =   255
         End
         Begin VB.PictureBox pbxHeaders 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   1575
            Left            =   0
            ScaleHeight     =   1575
            ScaleWidth      =   8895
            TabIndex        =   34
            Top             =   0
            Width           =   8895
            Begin VB.TextBox txtHeaderValue 
               Height          =   375
               Index           =   0
               Left            =   5280
               TabIndex        =   11
               Top             =   120
               Width           =   3495
            End
            Begin VB.TextBox txtHeaderValue 
               Height          =   375
               Index           =   1
               Left            =   5280
               TabIndex        =   13
               Top             =   600
               Width           =   3495
            End
            Begin VB.TextBox txtHeaderValue 
               Height          =   375
               Index           =   2
               Left            =   5280
               TabIndex        =   15
               Top             =   1080
               Width           =   3495
            End
            Begin VB.TextBox txtHeaderName 
               Height          =   375
               Index           =   0
               Left            =   840
               TabIndex        =   10
               Top             =   120
               Width           =   3495
            End
            Begin VB.TextBox txtHeaderName 
               Height          =   375
               Index           =   1
               Left            =   840
               TabIndex        =   12
               Top             =   600
               Width           =   3495
            End
            Begin VB.TextBox txtHeaderName 
               Height          =   375
               Index           =   2
               Left            =   840
               TabIndex        =   14
               Top             =   1080
               Width           =   3495
            End
            Begin VB.Label lblHeaderValue 
               Caption         =   "Value"
               Height          =   375
               Index           =   0
               Left            =   4560
               TabIndex        =   40
               Top             =   120
               Width           =   615
            End
            Begin VB.Label lblHeaderValue 
               Caption         =   "Value"
               Height          =   375
               Index           =   1
               Left            =   4560
               TabIndex        =   39
               Top             =   600
               Width           =   615
            End
            Begin VB.Label lblHeaderValue 
               Caption         =   "Value"
               Height          =   375
               Index           =   2
               Left            =   4560
               TabIndex        =   38
               Top             =   1080
               Width           =   615
            End
            Begin VB.Label lblHeaderName 
               Caption         =   "Name"
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   37
               Top             =   120
               Width           =   615
            End
            Begin VB.Label lblHeaderName 
               Caption         =   "Name"
               Height          =   375
               Index           =   1
               Left            =   120
               TabIndex        =   36
               Top             =   600
               Width           =   615
            End
            Begin VB.Label lblHeaderName 
               Caption         =   "Name"
               Height          =   375
               Index           =   2
               Left            =   120
               TabIndex        =   35
               Top             =   1080
               Width           =   495
            End
         End
      End
      Begin VB.CommandButton cmdMoreHeaders 
         Caption         =   "More"
         Height          =   375
         Left            =   10200
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.TextBox txtRequest 
      Height          =   1575
      Left            =   3360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   5760
      Width           =   6015
   End
   Begin MSWinsockLib.Winsock winsock 
      Left            =   10080
      Top             =   7200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Default         =   -1  'True
      Height          =   735
      Left            =   9960
      TabIndex        =   2
      Top             =   5760
      Width           =   1935
   End
   Begin VB.ComboBox cboRequestMethod 
      Height          =   315
      ItemData        =   "frmFormSubmit.frx":0000
      Left            =   10200
      List            =   "frmFormSubmit.frx":000A
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txtUrl 
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      ToolTipText     =   "Enter the URL here to the resource you want to request."
      Top             =   120
      Width           =   4815
   End
   Begin VB.TextBox txtResponse 
      Height          =   2895
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   7920
      Width           =   11775
   End
   Begin VB.Frame fmeVariables 
      Caption         =   "Submission Variables"
      Height          =   1935
      Left            =   120
      TabIndex        =   20
      ToolTipText     =   "Use this space to submit some custom vairables."
      Top             =   1080
      Width           =   11775
      Begin VB.CommandButton cmdMoreVariables 
         Caption         =   "More"
         Height          =   375
         Left            =   10200
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   480
         Width           =   1215
      End
      Begin VB.PictureBox pbxOVariables 
         BorderStyle     =   0  'None
         Height          =   1575
         Left            =   120
         ScaleHeight     =   1575
         ScaleWidth      =   9375
         TabIndex        =   23
         Top             =   240
         Width           =   9375
         Begin VB.PictureBox pbxVariables 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   1575
            Left            =   0
            ScaleHeight     =   1575
            ScaleWidth      =   8895
            TabIndex        =   25
            Top             =   0
            Width           =   8895
            Begin VB.TextBox txtVariableName 
               Height          =   375
               Index           =   2
               Left            =   840
               TabIndex        =   8
               Top             =   1080
               Width           =   3495
            End
            Begin VB.TextBox txtVariableName 
               Height          =   375
               Index           =   1
               Left            =   840
               TabIndex        =   6
               Top             =   600
               Width           =   3495
            End
            Begin VB.TextBox txtVariableName 
               Height          =   375
               Index           =   0
               Left            =   840
               TabIndex        =   4
               Top             =   120
               Width           =   3495
            End
            Begin VB.TextBox txtVariableValue 
               Height          =   375
               Index           =   2
               Left            =   5280
               TabIndex        =   9
               Top             =   1080
               Width           =   3495
            End
            Begin VB.TextBox txtVariableValue 
               Height          =   375
               Index           =   1
               Left            =   5280
               TabIndex        =   7
               Top             =   600
               Width           =   3495
            End
            Begin VB.TextBox txtVariableValue 
               Height          =   375
               Index           =   0
               Left            =   5280
               TabIndex        =   5
               Top             =   120
               Width           =   3495
            End
            Begin VB.Label lblVariableName 
               Caption         =   "Name"
               Height          =   375
               Index           =   2
               Left            =   120
               TabIndex        =   31
               Top             =   1080
               Width           =   1455
            End
            Begin VB.Label lblVariableName 
               Caption         =   "Name"
               Height          =   375
               Index           =   1
               Left            =   120
               TabIndex        =   30
               Top             =   600
               Width           =   615
            End
            Begin VB.Label lblVariableName 
               Caption         =   "Name"
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   29
               Top             =   120
               Width           =   615
            End
            Begin VB.Label lblVariableValue 
               Caption         =   "Value"
               Height          =   375
               Index           =   2
               Left            =   4560
               TabIndex        =   28
               Top             =   1080
               Width           =   1455
            End
            Begin VB.Label lblVariableValue 
               Caption         =   "Value"
               Height          =   375
               Index           =   1
               Left            =   4560
               TabIndex        =   27
               Top             =   600
               Width           =   1455
            End
            Begin VB.Label lblVariableValue 
               Caption         =   "Value"
               Height          =   375
               Index           =   0
               Left            =   4560
               TabIndex        =   26
               Top             =   120
               Width           =   1455
            End
         End
         Begin VB.VScrollBar vsbVariables 
            Enabled         =   0   'False
            Height          =   1335
            Left            =   9000
            Max             =   0
            TabIndex        =   24
            Top             =   120
            Width           =   255
         End
      End
   End
   Begin VB.Label Label2 
      Caption         =   "HTTP Request"
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   5760
      Width           =   2175
   End
   Begin VB.Label lblRequestMethod 
      Caption         =   "Request method"
      Height          =   255
      Left            =   7920
      TabIndex        =   19
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblUrl 
      Caption         =   "URL"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   240
      Width           =   1695
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   12120
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      Caption         =   "HTTP Reponse"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   7440
      Width           =   2175
   End
End
Attribute VB_Name = "frmFormSubmit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' we set this to true whil a connection is established
Private blnConnected As Boolean

' this function sends the HTTP request
Private Sub cmdSend_Click()
    Dim eUrl As URL
    
    Dim strMethod As String
    Dim strData As String
    Dim strPostData As String
    Dim strHeaders As String
    
    Dim strHTTP As String
    Dim X As Integer
    
    strPostData = ""
    strHeaders = ""
    strMethod = cboRequestMethod.List(cboRequestMethod.ListIndex)
    
    If blnConnected Then Exit Sub
    
    ' get the url
    eUrl = ExtractUrl(txtUrl.Text)
    
    If eUrl.Host = vbNullString Then
        MsgBox "Invalid Host", vbCritical, "ERROR"
    
        Exit Sub
    End If
    
    ' configure winsock
    winsock.Protocol = sckTCPProtocol
    winsock.RemoteHost = eUrl.Host
    
    If eUrl.Scheme = "http" Then
        If eUrl.Port > 0 Then
            winsock.RemotePort = eUrl.Port
        Else
            winsock.RemotePort = 80
        End If
    ElseIf eUrl.Scheme = vbNullString Then
        winsock.RemotePort = 80
    Else
        MsgBox "Invalid protocol schema"
    End If
    
    ' build encoded data the data is url encoded in the form
    ' var1=value&var2=value
    strData = ""
    For X = 0 To txtVariableName.Count - 1
        If txtVariableName(X).Text <> vbNullString Then
        
            strData = strData & URLEncode(txtVariableName(X).Text) & "=" & _
                            URLEncode(txtVariableValue(X).Text) & "&"
        End If
    Next X
    
    If eUrl.Query <> vbNullString Then
        eUrl.URI = eUrl.URI & "?" & eUrl.Query
    End If
    
    ' check if any variables were supplied
    If strData <> vbNullString Then
        strData = Left(strData, Len(strData) - 1)
        
        
        If strMethod = "GET" Then
            ' if this is a GET request then the URL encoded data
            ' is appended to the URI with a ?
            If eUrl.Query <> vbNullString Then
                eUrl.URI = eUrl.URI & "&" & strData
            Else
                eUrl.URI = eUrl.URI & "?" & strData
            End If
        Else
            ' if it is a post request, the data is appended to the
            ' body of the HTTP request and the headers Content-Type
            ' and Content-Length added
            strPostData = strData
            strHeaders = "Content-Type: application/x-www-form-urlencoded" & vbCrLf & _
                         "Content-Length: " & Len(strPostData) & vbCrLf
                         
        End If
    End If
            
    ' get any aditional headers and add them
    For X = 0 To txtHeaderName.Count - 1
        If txtHeaderName(X).Text <> vbNullString Then
        
            strHeaders = strHeaders & txtHeaderName(X).Text & ": " & _
                            txtHeaderValue(X).Text & vbCrLf
        End If
    Next X
    
    ' clear the old HTTP response
    txtResponse.Text = ""
    
    ' build the HTTP request in the form
    '
    ' {REQ METHOD} URI HTTP/1.0
    ' Host: {host}
    ' {headers}
    '
    ' {post data}
    strHTTP = strMethod & " " & eUrl.URI & " HTTP/1.0" & vbCrLf
    strHTTP = strHTTP & "Host: " & eUrl.Host & vbCrLf
    strHTTP = strHTTP & strHeaders
    strHTTP = strHTTP & vbCrLf
    strHTTP = strHTTP & strPostData

    txtRequest.Text = strHTTP
    
    winsock.Connect
    
    ' wait for a connection
    While Not blnConnected
        DoEvents
    Wend
    
    ' send the HTTP request
    winsock.SendData strHTTP
End Sub

Private Sub winsock_Connect()
    blnConnected = True
End Sub

' this event occurs when data is arriving via winsock
Private Sub winsock_DataArrival(ByVal bytesTotal As Long)
    Dim strResponse As String

    winsock.GetData strResponse, vbString, bytesTotal
    
    strResponse = FormatLineEndings(strResponse)
    
    ' we append this to the response box becuase data arrives
    ' in multiple packets
    txtResponse.Text = txtResponse.Text & strResponse
    
End Sub

Private Sub winsock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox Description, vbExclamation, "ERROR"
    
    winsock.Close
End Sub

Private Sub winsock_Close()
    blnConnected = False
    
    winsock.Close
End Sub

' this function converts all line endings to Windows CrLf line endings
Private Function FormatLineEndings(ByVal str As String) As String
    Dim prevChar As String
    Dim nextChar As String
    Dim curChar As String
    
    Dim strRet As String
    
    Dim X As Long
    
    prevChar = ""
    nextChar = ""
    curChar = ""
    strRet = ""
    
    For X = 1 To Len(str)
        prevChar = curChar
        curChar = Mid$(str, X, 1)
                
        If nextChar <> vbNullString And curChar <> nextChar Then
            curChar = curChar & nextChar
            nextChar = ""
        ElseIf curChar = vbLf Then
            If prevChar <> vbCr Then
                curChar = vbCrLf
            End If
            
            nextChar = ""
        ElseIf curChar = vbCr Then
            nextChar = vbLf
        End If
        
        strRet = strRet & curChar
    Next X
    
    FormatLineEndings = strRet
End Function

Private Sub Form_Load()
    cboRequestMethod.ListIndex = 0
    blnConnected = False
End Sub

' the code below has nothing to do with winsock or HTTP and deals only with the
' display and manipulation of controls
Private Sub cmdMoreHeaders_Click()
    Dim intNext As Integer
    Dim lngTop As Long
    
    ' find the next control
    intNext = txtHeaderName.Count
    
    ' find the next top
    lngTop = txtHeaderName(intNext - 1).Top + txtHeaderName(intNext - 1).Height + 80
    
    ' add new controls
    Load lblHeaderName(intNext)
    Load txtHeaderName(intNext)
    Load lblHeaderValue(intNext)
    Load txtHeaderValue(intNext)
    
                                  
    With lblHeaderName(intNext)
        .Top = lngTop
        .Left = lblHeaderName(intNext - 1).Left
        .Visible = True
    End With
    
    With txtHeaderName(intNext)
        .Top = lngTop
        .Left = txtHeaderName(intNext - 1).Left
        .Visible = True
        .Text = ""
    End With
        
    With lblHeaderValue(intNext)
        .Top = lngTop
        .Left = lblHeaderValue(intNext - 1).Left
        .Visible = True
    End With
    
    With txtHeaderValue(intNext)
        .Top = lngTop
        .Left = txtHeaderValue(intNext - 1).Left
        .Visible = True
        .Text = ""
    End With
    
    ' set the new height of the controls container
    pbxHeaders.Height = txtHeaderName(intNext).Top + txtHeaderName(intNext).Height + 80
    
    ' check if we should activate the scroll bar, ie: the outerbox
    ' is higher than the inner box
    If pbxHeaders.Height > pbxOHeaders.Height Then
        With vsbHeaders
            .Enabled = True
            .SmallChange = txtHeaderName(intNext).Height
            .LargeChange = pbxOHeaders.Height
            .Min = 0
            .Max = pbxHeaders.Height - pbxOHeaders.Height
            .Value = .Max
        End With
    End If
End Sub

Private Sub cmdMoreVariables_Click()
    Dim intNext As Integer
    Dim lngTop As Long
    
    ' find the next control
    intNext = txtVariableName.Count
    
    ' find the next top
    lngTop = txtVariableName(intNext - 1).Top + txtVariableName(intNext - 1).Height + 80
    
    ' add new controls
    Load lblVariableName(intNext)
    Load txtVariableName(intNext)
    Load lblVariableValue(intNext)
    Load txtVariableValue(intNext)
    
                                  
    With lblVariableName(intNext)
        .Top = lngTop
        .Left = lblVariableName(intNext - 1).Left
        .Visible = True
    End With
    
    With txtVariableName(intNext)
        .Top = lngTop
        .Left = txtVariableName(intNext - 1).Left
        .Visible = True
        .TabIndex = txtVariableName(intNext - 1).TabIndex + 2
        .Text = ""
    End With
        
    With lblVariableValue(intNext)
        .Top = lngTop
        .Left = lblVariableValue(intNext - 1).Left
        .Visible = True
    End With
    
    With txtVariableValue(intNext)
        .Top = lngTop
        .Left = txtVariableValue(intNext - 1).Left
        .TabIndex = txtVariableValue(intNext - 1).TabIndex + 2
        .Visible = True
        .Text = ""
    End With
    
    ' set the new height of the controls container
    pbxVariables.Height = txtVariableName(intNext).Top + txtVariableName(intNext).Height + 80
    
    ' check if we should activate the scroll bar, ie: the outerbox
    ' is higher than the inner box
    If pbxVariables.Height > pbxOVariables.Height Then
        With vsbVariables
            .Enabled = True
            .SmallChange = txtVariableName(intNext).Height
            .LargeChange = pbxOVariables.Height
            .Min = 0
            .Max = pbxVariables.Height - pbxOVariables.Height
            .Value = .Max
        End With
    End If
End Sub

Private Sub vsbHeaders_Change()
    pbxHeaders.Top = 0 - vsbHeaders.Value
End Sub

Private Sub vsbHeaders_Scroll()
    pbxHeaders.Top = 0 - vsbHeaders.Value
End Sub

Private Sub vsbVariables_Change()
    pbxVariables.Top = 0 - vsbVariables.Value
End Sub

Private Sub vsbVariables_Scroll()
    pbxVariables.Top = 0 - vsbVariables.Value
End Sub
