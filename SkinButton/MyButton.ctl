VERSION 5.00
Begin VB.UserControl SkinnableButton 
   ClientHeight    =   720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   810
   ScaleHeight     =   720
   ScaleWidth      =   810
   ToolboxBitmap   =   "MyButton.ctx":0000
End
Attribute VB_Name = "SkinnableButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Const DoHook = True

Public Enum Button_Styles
    Button = 1
    CheckBox = 2
    OptionBox = 3
End Enum
Public Enum CheckBoxValue
    Uncheck = 1
    Checked = 2
End Enum

Private Enum ImageMode
    Normal = 1
    Hover = 2
    Pressed = 3
End Enum
Private Picture_Normal As Picture
Private Picture_Hover As Picture
Private Picture_Pressed As Picture
Private Picture_Checked_Normal As Picture
Private Picture_Checked_Hover As Picture
Private Picture_Checked_Pressed As Picture
Private m_Style As Button_Styles
Private m_Value As CheckBoxValue
Private m_bAllowHover As Boolean
Private m_bAllowPressed As Boolean
Private m_MouseIsOver As Boolean

Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseOver()
Event MouseOut()
Event Click()
Event DblClick()

'Our Defines
Private LastPictureMode As ImageMode, LastValue As CheckBoxValue
Private MouseIsDown As Boolean
Private ConvertDoubleClickToDown As Boolean

' I M A G E S
Property Set Image_Normal(ByVal New_Picture As Picture)
    Set Picture_Normal = New_Picture
    PropertyChanged "Image_Normal"
End Property
Property Set Image_Hover(ByVal New_Picture As Picture)
    Set Picture_Hover = New_Picture
    PropertyChanged "Image_Hover"
End Property
Property Set Image_Pressed(ByVal New_Picture As Picture)
    Set Picture_Pressed = New_Picture
    PropertyChanged "Image_Pressed"
End Property
Property Set Image_Checked_Normal(ByVal New_Picture As Picture)
    Set Picture_Checked_Normal = New_Picture
End Property
Property Set Image_Checked_Hover(ByVal New_Picture As Picture)
    Set Picture_Checked_Hover = New_Picture
End Property
Property Set Image_Checked_Pressed(ByVal New_Picture As Picture)
    Set Picture_Checked_Pressed = New_Picture
End Property
'***********************
Property Get Image_Normal() As Picture
    Set Image_Normal = Picture_Normal
End Property
Property Get Image_Hover() As Picture
    Set Image_Hover = Picture_Hover
End Property
Property Get Image_Pressed() As Picture
    Set Image_Pressed = Picture_Pressed
End Property
Property Get Image_Checked_Normal() As Picture
    Set Image_Checked_Normal = Picture_Checked_Normal
End Property
Property Get Image_Checked_Hover() As Picture
    Set Image_Checked_Hover = Picture_Checked_Hover
End Property
Property Get Image_Checked_Pressed() As Picture
    Set Image_Checked_Pressed = Picture_Checked_Pressed
End Property
' E N D    I M A G E S
Private Sub SetPicture(Mode As ImageMode)
'Temporary Variable
Dim PicToSet As Picture
    'If Picture is set we don't need to set it again
    If LastPictureMode = Mode And LastValue <> m_Value Then Exit Sub
    'If Button or Checkbox(value=X) because when value is uncheked the checkbox is like normal buttons
    If m_Style = Button Or (m_Style = CheckBox And m_Value = Uncheck) Then
        Select Case Mode
        Case Normal
            Set PicToSet = Picture_Normal
        Case Hover
            If m_bAllowHover Then
                Set PicToSet = Picture_Hover
            Else
                Set PicToSet = Picture_Normal
            End If
        Case Pressed
            If m_bAllowPressed Then
                Set PicToSet = Picture_Pressed
            Else
                If m_bAllowHover Then
                    Set PicToSet = Picture_Hover
                Else
                    Set PicToSet = pictureNormal
                End If
            End If
        End Select
    'it's checkbox and it's Checked
    ElseIf m_Style = CheckBox And (m_Value = Checked) Then
        Select Case Mode
        Case Normal
            Set PicToSet = Picture_Checked_Normal
        Case Hover
            If m_bAllowHover Then
                Set PicToSet = Picture_Checked_Hover
            Else
                Set PicToSet = Picture_Checked_Normal
            End If
        Case Pressed
            If m_bAllowPressed Then
                Set PicToSet = Picture_Checked_Pressed
            Else
                If m_bAllowHover Then
                    Set PicToSet = Picture_Checked_Hover
                Else
                    Set PicToSet = Picture_Checked_Normal
                End If
            End If
        End Select
    End If
    
    'Set Picture to usercontrol
    If Not PicToSet Is Nothing Then Set UserControl.Picture = PicToSet
    LastPictureMode = Mode
    LastValue = m_Value
End Sub

Property Let AllowHover(ByVal New_Value As Boolean)
    m_bAllowHover = New_Value
    PropertyChanged "AllowHover"
End Property
Property Get AllowHover() As Boolean
    AllowHover = m_bAllowHover
End Property

Property Let AllowPressed(ByVal New_Value As Boolean)
    m_bAllowPressed = New_Value
    PropertyChanged "AllowPressed"
End Property
Property Get AllowPressed() As Boolean
    AllowPressed = m_bAllowPressed
End Property

Property Let Style(ByVal New_Value As Button_Styles)
    m_Style = New_Value
    'ChangeStyle
    PropertyChanged "Style"
End Property
Property Get Style() As Button_Styles
    Style = m_Style
End Property

Property Let Value(ByVal New_Value As CheckBoxValue)
    m_Value = New_Value
    PropertyChanged "Value"
End Property
Property Get Value() As CheckBoxValue
    Value = m_Value
End Property
Private Sub HookMe()
    'Hook using Module(modHookButton.bas)
    'But only in RunTime Mode
    If Ambient.UserMode Then HookButton UserControl.hWnd, Me
End Sub
Private Sub UnHookMe()
    'Unhook using module
    'But only in Runtime (in runtime it will be hooked so it's unhook in runtime,too
    'but in DesignTime we have no unhook because we have no hook!
    If Ambient.UserMode Then UnHookButton UserControl.hWnd
End Sub

Private Sub TrackLeaving()
    'Request for MouseLeave event (need to use callbacks and hook window)
    Dim ET As TRACKMOUSEEVENTTYPE
    ET.cbSize = Len(ET)
    ET.hwndTrack = UserControl.hWnd
    ET.dwFlags = TME_LEAVE
    TrackMouseEvent ET
End Sub
Private Sub TrackedMouseOut()
    m_MouseIsOver = False
    RaiseEvent MouseOut
    SetPicture Normal
End Sub

Private Sub UserControl_DblClick()
    If ConvertDoubleClickToDown Then
        ConvertDoubleClickToDown = False
        SetPicture Pressed
        RaiseEvent DblClick
    End If
End Sub

Private Sub UserControl_Hide()
    'We must unhook it.
    If DoHook Then UnHookMe
End Sub

Private Sub UserControl_InitProperties()
    m_Style = Button
    m_bAllowHover = True
    m_bAllowPressed = True
    m_Value = Uncheck
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button And vbLeftButton Then
        MouseIsDown = True
        ConvertDoubleClickToDown = True
        SetPicture Pressed
    Else
        ConvertDoubleClickToDown = False
    End If
    'Raise MouseDown
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Raise Mouse Over Event and set the MouseIsOver
    If Not m_MouseIsOver Then
        m_MouseIsOver = True
        RaiseEvent MouseOver
        SetPicture Hover
    End If
    'Tracking- Request for MouseLeave event if happened
    TrackLeaving
     
    'Raise the Mouse Move Event
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        MouseIsDown = False
        If (X >= 0 And Y >= 0) And (X <= UserControl.Width And Y <= UserControl.Height) Then
            If m_Style = CheckBox Then
                If m_Value = Checked Then m_Value = Uncheck Else m_Value = Checked
                'Dim i As Long
                'For i = 0 To UserControl.ParentControls.Count
                    'If UserControl.ParentControls(i).Value = Checked Then UserControl.ParentControls(i).Value = Unchecked
                'Next
            End If
            SetPicture Hover
            RaiseEvent Click
        End If
    End If
    'Raise MouseUp
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Style = PropBag.ReadProperty("Style", Button)
    m_bAllowHover = PropBag.ReadProperty("AllowHover", True)
    m_bAllowPressed = PropBag.ReadProperty("AllowPressed", True)
    m_Value = PropBag.ReadProperty("Value", Unchecked)
    Set Picture_Normal = PropBag.ReadProperty("Image_Normal", Nothing)
    Set Picture_Hover = PropBag.ReadProperty("Image_Hover", Nothing)
    Set Picture_Pressed = PropBag.ReadProperty("Image_Pressed", Nothing)
    Set Picture_Checked_Normal = PropBag.ReadProperty("Image_Checked_Normal", Nothing)
    Set Picture_Checked_Hover = PropBag.ReadProperty("Image_Checked_Hover", Nothing)
    Set Picture_Checked_Pressed = PropBag.ReadProperty("Image_Checked_Pressed", Nothing)
End Sub

Private Sub UserControl_Show()
    'If Usercontrol show or created it will be hooked
    If DoHook Then HookMe
    
    'Set Picture
    TrackLeaving
    SetPicture Normal
End Sub

Public Sub HookMsg(MsgID As Long)
    'We have Received a Msg
    If MsgID = WM_MOUSELEAVE Then
        'Raise Event
        TrackedMouseOut
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Style", m_Style, Button
    PropBag.WriteProperty "AllowHover", m_bAllowHover, True
    PropBag.WriteProperty "AllowPressed", m_bAllowPressed, True
    PropBag.WriteProperty "Value", m_Value, Unchecked
    PropBag.WriteProperty "Image_Normal", Picture_Normal, Nothing
    PropBag.WriteProperty "Image_Hover", Picture_Hover, Nothing
    PropBag.WriteProperty "Image_Pressed", Picture_Pressed, Nothing
    PropBag.WriteProperty "Image_Checked_Normal", Picture_Checked_Normal, Nothing
    PropBag.WriteProperty "Image_Checked_Hover", Picture_Checked_Hover, Nothing
    PropBag.WriteProperty "Image_Checked_Pressed", Picture_Checked_Pressed, Nothing
End Sub
