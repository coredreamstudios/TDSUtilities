VERSION 5.00
Object = "{922FB004-DD9A-11D3-BD8D-DAAFCB8D9378}#2.1#0"; "DNVideoX.ocx"
Begin VB.UserControl RBOCX 
   ClientHeight    =   5535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8490
   ScaleHeight     =   5535
   ScaleWidth      =   8490
   Begin DNVideoXLib.DNVideoX DNVideoX1 
      Height          =   4935
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7815
      _Version        =   65536
      _ExtentX        =   13785
      _ExtentY        =   8705
      _StockProps     =   0
   End
End
Attribute VB_Name = "RBOCX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Sub UserControl_Initialize()
    If DNVideoX1.GetVideoDeviceCount < 1 Then
        MsgBox "Build 6"
    End If
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DNVideoX1,DNVideoX1,-1,FreezePreview
Public Function FreezePreview(ByVal Freeze As Long) As Long
Attribute FreezePreview.VB_Description = "Freeze preview video on/off"
    FreezePreview = DNVideoX1.FreezePreview(Freeze)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DNVideoX1,DNVideoX1,-1,CapFilename
Public Property Get CapFilename() As String
Attribute CapFilename.VB_Description = "Filename for captured media file. Extension can be AVI or WMV."
    CapFilename = DNVideoX1.CapFilename
End Property

Public Property Let CapFilename(ByVal New_CapFilename As String)
    DNVideoX1.CapFilename() = New_CapFilename
    PropertyChanged "CapFilename"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DNVideoX1,DNVideoX1,-1,StartCapture
Public Function StartCapture() As Boolean
Attribute StartCapture.VB_Description = "Starts video capture"
    StartCapture = DNVideoX1.StartCapture()
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DNVideoX1,DNVideoX1,-1,StopCapture
Public Function StopCapture() As Boolean
Attribute StopCapture.VB_Description = "Stops video capture"
    StopCapture = DNVideoX1.StopCapture()
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DNVideoX1,DNVideoX1,-1,Connected
Public Property Get Connected() As Boolean
Attribute Connected.VB_Description = "Get/set connection to video device"
    Connected = DNVideoX1.Connected
End Property

Public Property Let Connected(ByVal New_Connected As Boolean)
    DNVideoX1.Connected() = New_Connected
    PropertyChanged "Connected"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DNVideoX1,DNVideoX1,-1,LightOn
Public Function LightOn(ByVal Value As Long) As Long
Attribute LightOn.VB_Description = "Control camera LED. 1=on, 2=off."
    LightOn = DNVideoX1.LightOn(Value)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DNVideoX1,DNVideoX1,-1,SetVideoFormat
Public Function SetVideoFormat(ByVal width As Long, ByVal height As Long) As Boolean
Attribute SetVideoFormat.VB_Description = "Set video image dimensions"
    SetVideoFormat = DNVideoX1.SetVideoFormat(width, height)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DNVideoX1,DNVideoX1,-1,SetTextOverlay
Public Function SetTextOverlay(ByVal Index As Long, ByVal Caption As String, ByVal X As Long, ByVal Y As Long, ByVal FontName As String, ByVal FontSize As Long, ByVal FColor As Long, ByVal BColor As Long) As Long
Attribute SetTextOverlay.VB_Description = "Sets on-video text caption."
    SetTextOverlay = DNVideoX1.SetTextOverlay(Index, Caption, X, Y, FontName, FontSize, FColor, BColor)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DNVideoX1,DNVideoX1,-1,SetZoom
Public Function SetZoom(ByVal Left As Long, ByVal Top As Long, ByVal width As Long, ByVal height As Long) As Long
Attribute SetZoom.VB_Description = "Set zoom rectangle on video. Use all zeros as parameters to this method to reset zoom."
    SetZoom = DNVideoX1.SetZoom(Left, Top, width, height)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DNVideoX1,DNVideoX1,-1,VideoWidth
Public Property Get VideoWidth() As Long
Attribute VideoWidth.VB_Description = "Returns current video width in pixels. This property is read-only."
    VideoWidth = DNVideoX1.VideoWidth
End Property

Public Property Let VideoWidth(ByVal New_VideoWidth As Long)
    DNVideoX1.VideoWidth() = New_VideoWidth
    PropertyChanged "VideoWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DNVideoX1,DNVideoX1,-1,VideoHeight
Public Property Get VideoHeight() As Long
Attribute VideoHeight.VB_Description = "Returns current video height in pixels. This property is read-only."
    VideoHeight = DNVideoX1.VideoHeight
End Property

Public Property Let VideoHeight(ByVal New_VideoHeight As Long)
    DNVideoX1.VideoHeight() = New_VideoHeight
    PropertyChanged "VideoHeight"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    DNVideoX1.CapFilename = PropBag.ReadProperty("CapFilename", "c:\capture.avi")
'    DNVideoX1.UseVideoFilter = PropBag.ReadProperty("UseVideoFilter", 1)
    DNVideoX1.Connected = PropBag.ReadProperty("Connected", False)
    DNVideoX1.VideoWidth = PropBag.ReadProperty("VideoWidth", 0)
    DNVideoX1.VideoHeight = PropBag.ReadProperty("VideoHeight", 0)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("CapFilename", DNVideoX1.CapFilename, "c:\capture.avi")
'    Call PropBag.WriteProperty("UseVideoFilter", DNVideoX1.UseVideoFilter, 1)
    Call PropBag.WriteProperty("Connected", DNVideoX1.Connected, False)
    Call PropBag.WriteProperty("VideoWidth", DNVideoX1.VideoWidth, 0)
    Call PropBag.WriteProperty("VideoHeight", DNVideoX1.VideoHeight, 0)
End Sub
