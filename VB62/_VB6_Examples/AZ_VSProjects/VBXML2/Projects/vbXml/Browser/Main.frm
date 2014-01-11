VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "XML Browser"
   ClientHeight    =   8385
   ClientLeft      =   1995
   ClientTop       =   840
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   7125
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   4800
      Left            =   5445
      ScaleHeight     =   5012.637
      ScaleMode       =   0  'User
      ScaleWidth      =   260
      TabIndex        =   0
      Top             =   825
      Visible         =   0   'False
      Width           =   156
   End
   Begin MSComctlLib.TabStrip tsTabStrip 
      Height          =   4755
      Left            =   3150
      TabIndex        =   5
      Top             =   2100
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   8387
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Value"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Elements"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Top             =   8100
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolbar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   20
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Back"
            Object.ToolTipText     =   "Back"
            ImageKey        =   "Back"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Forward"
            Object.ToolTipText     =   "Forward"
            ImageKey        =   "Forward"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Up One Level"
            Object.ToolTipText     =   "Up One Level"
            ImageKey        =   "Up One Level"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Undo"
            Object.ToolTipText     =   "Undo"
            ImageKey        =   "Undo"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Redo"
            Object.ToolTipText     =   "Redo"
            ImageKey        =   "Redo"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete"
            ImageKey        =   "Delete"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "View Large Icons"
            Object.ToolTipText     =   "View Large Icons"
            ImageKey        =   "View Large Icons"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "View Small Icons"
            Object.ToolTipText     =   "View Small Icons"
            ImageKey        =   "View Small Icons"
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "View List"
            Object.ToolTipText     =   "View List"
            ImageKey        =   "View List"
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "View Details"
            Object.ToolTipText     =   "View Details"
            ImageKey        =   "View Details"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.TreeView tvTreeView 
      Height          =   4800
      Left            =   30
      TabIndex        =   1
      Top             =   825
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   8467
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   2685
      Top             =   2610
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0000
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0112
            Key             =   "Forward"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0224
            Key             =   "Up One Level"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0336
            Key             =   "New"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0448
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":055A
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":066C
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":077E
            Key             =   "Redo"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0890
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":09A2
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0AB4
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0BC6
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0CD8
            Key             =   "View Large Icons"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0DEA
            Key             =   "View Small Icons"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0EFC
            Key             =   "View List"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":100E
            Key             =   "View Details"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtValue 
      Height          =   525
      Left            =   2970
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   3960
      Visible         =   0   'False
      Width           =   1245
   End
   Begin MSComctlLib.ListView lvListView 
      Height          =   4800
      Left            =   2085
      TabIndex        =   2
      Top             =   825
      Visible         =   0   'False
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   8467
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Attribute Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Attribute Value"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Image imgSplitter 
      Height          =   4785
      Left            =   2010
      MouseIcon       =   "Main.frx":1120
      MousePointer    =   99  'Custom
      Top             =   825
      Width           =   150
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Const NAME_COLUMN = 0
'Const TYPE_COLUMN = 1
'Const SIZE_COLUMN = 2
'Const DATE_COLUMN = 3

Private Const sglSplitLimit = 500
Private Const M_BufferWidth = 60

Private mbMoving As Boolean
Private m_oDoc As XMLDocument
Private m_oCurrentElement As CXmlElement
Private Sub FillNode(oNode As MSComctlLib.Node, oElement As CXmlElement)
    Dim oChild As CXmlElement
    Dim oChNode As MSComctlLib.Node
    Dim lIndex As Long
    Dim sKey As String
    
    oNode.Text = oElement.Name
    For Each oChild In oElement
        sKey = oNode.Key & ":" & lIndex
        
        Set oChNode = tvTreeView.Nodes.Add(oNode, tvwChild, sKey)
        Call FillNode(oChNode, oChild)
        lIndex = lIndex + 1
    Next
End Sub

Private Sub LoadDoc(oDoc As XMLDocument)
    Dim oNode As MSComctlLib.Node
    
    Set m_oDoc = oDoc
    
    tvTreeView.Nodes.Clear
    Set oNode = tvTreeView.Nodes.Add(, , ":0")
    Set m_oCurrentElement = m_oDoc.Root
    
    Call FillNode(oNode, m_oCurrentElement)
    Set tvTreeView.SelectedItem = oNode
End Sub

Private Sub Form_Load()
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub

Private Sub mnuFileOpen_Click()
    Dim oDoc As XMLDocument
    
    Set oDoc = OpenFile
    
    If Not oDoc Is Nothing Then Call LoadDoc(oDoc)
End Sub


Private Sub tbToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "Back"
            'ToDo: Add 'Back' button code.
            MsgBox "Add 'Back' button code."
        Case "Forward"
            'ToDo: Add 'Forward' button code.
            MsgBox "Add 'Forward' button code."
        Case "Up One Level"
            'ToDo: Add 'Up One Level' button code.
            MsgBox "Add 'Up One Level' button code."
        Case "New"
            'ToDo: Add 'New' button code.
            MsgBox "Add 'New' button code."
        Case "Open"
            'ToDo: Add 'Open' button code.
            MsgBox "Add 'Open' button code."
        Case "Save"
            'ToDo: Add 'Save' button code.
            MsgBox "Add 'Save' button code."
        Case "Undo"
            'ToDo: Add 'Undo' button code.
            MsgBox "Add 'Undo' button code."
        Case "Redo"
            'ToDo: Add 'Redo' button code.
            MsgBox "Add 'Redo' button code."
        Case "Cut"
            'ToDo: Add 'Cut' button code.
            MsgBox "Add 'Cut' button code."
        Case "Copy"
            'ToDo: Add 'Copy' button code.
            MsgBox "Add 'Copy' button code."
        Case "Paste"
            'ToDo: Add 'Paste' button code.
            MsgBox "Add 'Paste' button code."
        Case "Delete"
            'ToDo: Add 'Delete' button code.
            MsgBox "Add 'Delete' button code."
        Case "View Large Icons"
            'ToDo: Add 'View Large Icons' button code.
            MsgBox "Add 'View Large Icons' button code."
        Case "View Small Icons"
            'ToDo: Add 'View Small Icons' button code.
            MsgBox "Add 'View Small Icons' button code."
        Case "View List"
            'ToDo: Add 'View List' button code.
            MsgBox "Add 'View List' button code."
        Case "View Details"
            'ToDo: Add 'View Details' button code.
            MsgBox "Add 'View Details' button code."
    End Select
End Sub



Private Sub Form_Resize()
  On Error Resume Next
  If Me.Width < 3000 Then Me.Width = 3000
  SizeControls imgSplitter.Left
End Sub

Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  With imgSplitter
    picSplitter.Move .Left, .TOp, .Width - 20, .Height - 20
  End With
  picSplitter.Visible = True
  mbMoving = True
End Sub

Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim sglPos As Single
  
  If mbMoving Then
    sglPos = x + imgSplitter.Left
    If sglPos < sglSplitLimit Then
      picSplitter.Left = sglSplitLimit
    ElseIf sglPos > Me.Width - sglSplitLimit Then
      picSplitter.Left = Me.Width - sglSplitLimit
    Else
      picSplitter.Left = sglPos
    End If
  End If
End Sub
Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  SizeControls picSplitter.Left
  picSplitter.Visible = False
  mbMoving = False
End Sub

Sub SizeControls(Optional x As Single)
    Dim lTop As Long
    Dim lHeight As Long
    Dim oCtrl As Control
    
    On Error Resume Next
    
    'set the width
    If x < 1500 Then x = 1500
    If x > (Width - 1500) Then _
        x = Width - 1500
    
    lTop = IIf(tbToolbar.Visible, tbToolbar.Height, 0) + M_BufferWidth
    lHeight = ScaleHeight - (IIf(sbStatusBar.Visible, sbStatusBar.Height, 0) + lTop + M_BufferWidth)
    
    Call tvTreeView.Move(M_BufferWidth, lTop, x, lHeight)
    Call tsTabStrip.Move(x + (M_BufferWidth * 2), lTop, ScaleWidth - (x + (M_BufferWidth * 2)), lHeight)
    
    imgSplitter.Left = x
    
    imgSplitter.TOp = tvTreeView.TOp
    imgSplitter.Height = tvTreeView.Height
    
    With tsTabStrip
        If .SelectedItem.Index = 1 Then
            Set oCtrl = txtValue
            lvListView.Visible = False
        Else
            Set oCtrl = lvListView
            txtValue.Visible = False
        End If
        
        oCtrl.Visible = True
        oCtrl.ZOrder
        Call oCtrl.Move(.ClientLeft, .ClientTop, .ClientWidth, .ClientHeight)
    End With
End Sub

Private Sub tsTabStrip_Click()
    Call SizeControls(imgSplitter.Left)
End Sub


Private Sub tvTreeView_DragDrop(Source As Control, x As Single, y As Single)
    If Source = imgSplitter Then SizeControls x
End Sub

Private Sub tvTreeView_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim oAtt As CXmlAttribute
    Dim oElement As CXmlElement
    Dim sKey As String
    Dim vIndices As Variant
    Dim lCount As Long
    Dim oListItem As MSComctlLib.ListItem
    
    lvListView.ListItems.Clear
    
    sKey = Node.Key
    vIndices = Split(sKey, ":")
    
    Set oElement = m_oDoc.Root
    
    For lCount = 2 To UBound(vIndices)
        Set oElement = oElement.Node(vIndices(lCount) + 1)
    Next
    
    txtValue.Text = oElement.Body
    
    For lCount = 1 To oElement.AttributeCount
        Set oAtt = oElement.ElementAttribute(lCount)
        Set oListItem = lvListView.ListItems.Add(, , oAtt.KeyWord)
        oListItem.SubItems(1) = oAtt.Value
    Next

End Sub


