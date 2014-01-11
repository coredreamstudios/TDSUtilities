VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form DisplayModesForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Display Modes"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7815
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView LV 
      Height          =   2775
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   4895
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Colour Depth"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Resolution"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "BPP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Frequency"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Current"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enum Display"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   1935
   End
End
Attribute VB_Name = "DisplayModesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Add a listview (LV) and a command button to a form. To the listview, add five column headers as indicated, and set the style to report mode. Add the following to the form:
'After running the app, the listview will contain a list of the available resolutions for that system, and the present resolution setting will be highlighted. (Make sure that the Listview's HideSelection property is false.)
'--------------------------------------------------------------------------------
 
Option Explicit

'vars set in load
Dim currHRes As Long
Dim currVRes As Long
Dim currBPP As Long
   
Private Sub Command2_Click()
    
    Unload Me
    
End Sub

Private Sub Form_Load()

  'set the extended listview style
   Call SendMessage(LV.hWnd, _
                    LVM_SETEXTENDEDLISTVIEWSTYLE, _
                    LVS_EX_FULLROWSELECT, ByVal True)
  
  'retrieve the current screen resolution for
  'later comparison against DEVMODE values in
  'CompareSettings().
   currHRes = GetDeviceCaps(hdc, HORZRES)
   currVRes = GetDeviceCaps(hdc, VERTRES)
   currBPP = GetDeviceCaps(hdc, BITSPIXEL)
         
End Sub

Private Sub LVAdd(DM As DEVMODE)
 
   Dim itmX As MSComctlLib.ListItem
   Dim bppType As String
   
   Select Case DM.dmBitsPerPel
      Case 4:      bppType = "16 Color"
      Case 8:      bppType = "256 Color"
      Case 16:     bppType = "High Color"
      Case 24, 32: bppType = "True Color"
   End Select
   
   Set itmX = LV.ListItems.Add(, , bppType)
  
   itmX.SubItems(1) = Format$(DM.dmPelsWidth, " 000 x") & _
                      Format$(DM.dmPelsHeight, " 000")
                      
   itmX.SubItems(2) = Format$(DM.dmBitsPerPel, " 00")
   
   If DM.dmDisplayFrequency = 1 Then
         itmX.SubItems(3) = "Hardware default"
   Else: itmX.SubItems(3) = Format$(DM.dmDisplayFrequency, " 00") & " hz"
   End If
   
   If CompareSettings(DM) Then
     itmX.SubItems(4) = "Current"
     itmX.Selected = True
   End If
   
End Sub


Private Function CompareSettings(DM As DEVMODE) As Boolean
   
  'compares the current screen resolution with
  'the current DEVMODE values. Returns TRUE if
  'the horizontal and vertical resolutions, and
  'the bits per pixel colour depth, are the same.
  
   CompareSettings = (DM.dmBitsPerPel = currBPP) And _
                      DM.dmPelsHeight = currVRes And _
                      DM.dmPelsWidth = currHRes
   
End Function


Private Sub Command1_Click()

   Dim DM As DEVMODE
   Dim dMode As Long
   Dim r As Long
   
  'set the DEVMODE flags and structure size
   DM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
   DM.dmSize = LenB(DM)
      
  'The first mode is 0
   dMode = 0

   Do While EnumDisplaySettings(0&, dMode, DM) > 0
   
     'if the BitsPerPixel is greater than 4
     '(16 colours), then add the item to the list
      If DM.dmBitsPerPel >= 4 Then Call LVAdd(DM)
      
     'increment and call again. Continue until
     'EnumDisplaySettings returns 0 (no more settings)
      dMode = dMode + 1
   
   Loop
   
End Sub

Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

  LV.SortKey = ColumnHeader.Index - 1
  LV.SortOrder = Abs(Not LV.SortOrder = 1)
  LV.Sorted = True
  
End Sub
'--end block--'
 


