VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   2895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4110
   LinkTopic       =   "Form2"
   ScaleHeight     =   2895
   ScaleWidth      =   4110
   StartUpPosition =   3  'Windows Default
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   4683
      _Version        =   393216
      FixedCols       =   0
      FocusRect       =   0
      HighLight       =   2
      ScrollBars      =   2
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    
    Dim rw As Long
    
    rw = 1
    
    Grid1.Rows = 2
    
    Grid1.ColWidth(0) = 1000
    Grid1.ColWidth(1) = 2000
    
    Grid1.FormatString = "AcctNum|Last Name"
    
    Close #1
    
    Open filename For Random Access Read As #1 _
           Len = recordLength
           
    Do
        Get #1, , mUdtClient
            If mUdtClient.accountNumber <> 0 Then
                Grid1.Rows = rw + 1
                Grid1.ColWidth(0) = 1000
                Grid1.ColWidth(1) = 2000
                Grid1.TextMatrix(rw, 0) = mUdtClient.accountNumber
                Grid1.TextMatrix(rw, 1) = mUdtClient.lastName
                rw = rw + 1
            End If
    Loop Until EOF(1) = True
    
End Sub
