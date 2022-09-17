VERSION 5.00
Begin VB.UserControl LED 
   BackStyle       =   0  'Transparent
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   3840
   ToolboxBitmap   =   "uscLED.ctx":0000
   Windowless      =   -1  'True
   Begin VB.Timer tmrBlink 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   420
      Top             =   2160
   End
   Begin Proyect1.ShapeEx ShapeEx1 
      Height          =   432
      Left            =   480
      TabIndex        =   0
      Top             =   420
      Width           =   912
      _ExtentX        =   1609
      _ExtentY        =   762
      FillStyle       =   0
      UseSubclassing  =   0
   End
End
Attribute VB_Name = "LED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event Click()
Event DblClick()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Enum LEDShapeConstants
    ledRound
    ledSquare
    ledRectangle
    ledRoundedSquare
    ledRoundedRectangle
End Enum

Public Enum LEDBlinkTypeConstants
    ledBlinkShorter
    ledBlinkShort
    ledBlinkMedium
    ledBlinkLong
    ledBlinkTwice
End Enum

Public Enum LEDStateConstants
    ledOff
    ledOn
    ledBlinking
End Enum

Public Enum LEDColorConstants
    ledRed
    ledGreen
    ledBlue
    ledYellow
    ledWhite
    ledCustomColor
End Enum

Public Enum LEDStyleConstants
    ledStyle2D
    ledStyle3D
End Enum

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long

' Property defaults
Private Const mdef_Shape = ledRound
Private Const mdef_BlinkPeriod = 700
Private Const mdef_BlinkType = ledBlinkShort
Private Const mdef_BorderWidth = 1
Private Const mdef_BorderColor = &HC0&
Private Const mdef_ColorOn = vbRed
Private Const mdef_ColorOff = &H808080
Private Const mdef_State = ledOn
Private Const mdef_Color = ledRed
Private Const mdef_ToggleOnClick = False
Private Const mdef_Style = ledStyle3D

' Properties
Private mShape As LEDShapeConstants
Private mBlinkPeriod As Long
Private mBlinkType As LEDBlinkTypeConstants
Private mBorderWidth As Long
Private mBorderColor As Long
Private mColorOn As Long
Private mColorOff As Long
Private mState As LEDStateConstants
Private mColor As LEDColorConstants
Private mToggleOnClick As Boolean
Private mStyle As LEDStyleConstants

Private mStateOn As Boolean

Private Sub ShapeEx1_Click()
    If mToggleOnClick Then
        If mState = ledOn Then
            State = ledOff
        ElseIf mState = ledOff Then
            State = ledOn
        End If
    End If
    RaiseEvent Click
End Sub

Private Sub ShapeEx1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub ShapeEx1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub ShapeEx1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub ShapeEx1_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub tmrBlink_Timer()
    Dim t As Long
    Dim b As Long
    
    t = Round((Timer * 1000 Mod mBlinkPeriod) / 100)
    'Debug.Print t

    If mBlinkType = ledBlinkShorter Then
        SetOn = (t = 0)
    ElseIf mBlinkType = ledBlinkShort Then
        SetOn = (t = 0) Or (t = 1)
    ElseIf mBlinkType = ledBlinkTwice Then
        SetOn = (t = 0) Or (t = 2)
    ElseIf mBlinkType = ledBlinkLong Then
        SetOn = (t <> 0) And (t <> 1)
    Else ' medium
        SetOn = Round((Timer * 1000 Mod mBlinkPeriod) / 100) > (mBlinkPeriod / 100 / 2)
    End If
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    If PropertyName = "BackColor" Then
        UserControl.BackColor = Ambient.BackColor
    End If
End Sub

Private Sub UserControl_HitTest(X As Single, Y As Single, HitResult As Integer)
    HitResult = vbHitResultHit
End Sub

Private Sub UserControl_InitProperties()
    mShape = mdef_Shape
    mBlinkPeriod = mdef_BlinkPeriod
    mBlinkType = mdef_BlinkType
    mBorderWidth = mdef_BorderWidth
    mBorderColor = mdef_BorderColor
    mColorOn = mdef_ColorOn
    mColorOff = mdef_ColorOff
    mState = mdef_State
    mColor = mdef_Color
    mToggleOnClick = mdef_ToggleOnClick
    mStyle = mdef_Style
    ShowControl
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or (KeyAscii = vbKeySpace) Then
        ShapeEx1_Click
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mShape = PropBag.ReadProperty("Shape", mdef_Shape)
    mBlinkPeriod = PropBag.ReadProperty("BlinkPeriod", mdef_BlinkPeriod)
    mBlinkType = PropBag.ReadProperty("BlinkType", mdef_BlinkType)
    mBorderWidth = PropBag.ReadProperty("BorderWidth", mdef_BorderWidth)
    mBorderColor = PropBag.ReadProperty("BorderColor", mdef_BorderColor)
    mColorOn = PropBag.ReadProperty("ColorOn", mdef_ColorOn)
    mColorOff = PropBag.ReadProperty("ColorOff", mdef_ColorOff)
    mState = PropBag.ReadProperty("State", mdef_State)
    mColor = PropBag.ReadProperty("Color", mdef_Color)
    mToggleOnClick = PropBag.ReadProperty("ToggleOnClick", mdef_ToggleOnClick)
    mStyle = PropBag.ReadProperty("Style", mdef_Style)
    ShowControl
End Sub

Private Sub UserControl_Resize()
    Dim iWidth As Long
    
    If (UserControl.Height < 7 * Screen.TwipsPerPixelY) Then UserControl.Height = 7 * Screen.TwipsPerPixelY
    Select Case mShape
        Case ledRound, ledSquare, ledRoundedSquare
            iWidth = UserControl.Height
            UserControl.Width = iWidth
        Case Else
            If (UserControl.Width < 7 * Screen.TwipsPerPixelX) Then UserControl.Width = 7 * Screen.TwipsPerPixelX
            iWidth = UserControl.Width
    End Select
    
    ShapeEx1.Move 2 * Screen.TwipsPerPixelX, 2 * Screen.TwipsPerPixelY, iWidth - 4 * Screen.TwipsPerPixelX, UserControl.ScaleHeight - 4 * Screen.TwipsPerPixelY
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Shape", mShape, mdef_Shape
    PropBag.WriteProperty "BlinkPeriod", mBlinkPeriod, mdef_BlinkPeriod
    PropBag.WriteProperty "BlinkType", mBlinkType, mdef_BlinkType
    PropBag.WriteProperty "BorderWidth", mBorderWidth, mdef_BorderWidth
    PropBag.WriteProperty "BorderColor", mBorderColor, mdef_BorderColor
    PropBag.WriteProperty "ColorOn", mColorOn, mdef_ColorOn
    PropBag.WriteProperty "ColorOff", mColorOff, mdef_ColorOff
    PropBag.WriteProperty "State", mState, mdef_State
    PropBag.WriteProperty "Color", mColor, mdef_Color
    PropBag.WriteProperty "ToggleOnClick", mToggleOnClick, mdef_ToggleOnClick
    PropBag.WriteProperty "Style", mStyle, mdef_Style
End Sub


Public Property Get Shape() As LEDShapeConstants
    Shape = mShape
End Property

Public Property Let Shape(nValue As LEDShapeConstants)
    If nValue <> mShape Then
        If (nValue < ledRound) Or (nValue > ledRoundedRectangle) Then Err.Raise 380, TypeName(Me): Exit Property
        mShape = nValue
        UserControl_Resize
        ShowControl
        PropertyChanged "Shape"
    End If
End Property


Public Property Get BlinkPeriod() As Long
    BlinkPeriod = mBlinkPeriod
End Property

Public Property Let BlinkPeriod(nValue As Long)
    If nValue <> mBlinkPeriod Then
        If (nValue < 300) Or (nValue > 60000) Then Err.Raise 380, TypeName(Me): Exit Property
        mBlinkPeriod = nValue
        PropertyChanged "BlinkPeriod"
    End If
End Property


Public Property Get BlinkType() As LEDBlinkTypeConstants
    BlinkType = mBlinkType
End Property

Public Property Let BlinkType(nValue As LEDBlinkTypeConstants)
    If nValue <> mBlinkType Then
        If (nValue < ledBlinkShorter) Or (nValue > ledBlinkTwice) Then Err.Raise 380, TypeName(Me): Exit Property
        mBlinkType = nValue
        PropertyChanged "BlinkType"
    End If
End Property


Public Property Get BorderWidth() As Long
    BorderWidth = mBorderWidth
End Property

Public Property Let BorderWidth(nValue As Long)
    If nValue <> mBorderWidth Then
        If (nValue < 0) Or (nValue > 10) Then Err.Raise 380, TypeName(Me): Exit Property
        mBorderWidth = nValue
        ShowControl
        PropertyChanged "BorderWidth"
    End If
End Property


Public Property Get BorderColor() As OLE_COLOR
    BorderColor = mBorderColor
End Property

Public Property Let BorderColor(nValue As OLE_COLOR)
    If nValue <> mBorderColor Then
        mBorderColor = nValue
        mColor = ledCustomColor
        ShowControl
        PropertyChanged "BorderColor"
    End If
End Property


Public Property Get ColorOn() As OLE_COLOR
    ColorOn = mColorOn
End Property

Public Property Let ColorOn(nValue As OLE_COLOR)
    If nValue <> mColorOn Then
        mColorOn = nValue
        mColor = ledCustomColor
        If mState = ledOn Then
            ShowControl
        End If
        PropertyChanged "ColorOn"
    End If
End Property


Public Property Get ColorOff() As OLE_COLOR
    ColorOff = mColorOff
End Property

Public Property Let ColorOff(nValue As OLE_COLOR)
    If nValue <> mColorOff Then
        mColorOff = nValue
        mColor = ledCustomColor
        If mState = ledOff Then
            ShowControl
        End If
        PropertyChanged "ColorOff"
    End If
End Property


Public Property Get State() As LEDStateConstants
    State = mState
End Property

Public Property Let State(nValue As LEDStateConstants)
    If nValue <> mState Then
        If (nValue < ledOff) Or (nValue > ledBlinking) Then Err.Raise 380, TypeName(Me): Exit Property
        mState = nValue
        SetState
        PropertyChanged "State"
    End If
End Property


Public Property Get Color() As LEDColorConstants
    Color = mColor
End Property

Public Property Let Color(nValue As LEDColorConstants)
    If nValue <> mColor Then
        If (nValue < ledRed) Or (nValue > ledCustomColor) Then Err.Raise 380, TypeName(Me): Exit Property
        mColor = nValue
        ShowControl
        PropertyChanged "Color"
    End If
End Property


Public Property Get ToggleOnClick() As Boolean
    ToggleOnClick = mToggleOnClick
End Property

Public Property Let ToggleOnClick(nValue As Boolean)
    If nValue <> mToggleOnClick Then
        mToggleOnClick = nValue
        ShowControl
        PropertyChanged "ToggleOnClick"
    End If
End Property


Public Property Get Style() As LEDStyleConstants
    Style = mStyle
End Property

Public Property Let Style(nValue As LEDStyleConstants)
    If nValue <> mStyle Then
        mStyle = nValue
        ShowControl
        PropertyChanged "Style"
    End If
End Property


Private Sub ShowControl()
    UserControl.BackColor = Ambient.BackColor
    SetColor
    If mShape = ledSquare Then
        ShapeEx1.Shape = seShapeSquare
    ElseIf mShape = ledRectangle Then
        ShapeEx1.Shape = seShapeRectangle
    ElseIf mShape = ledRoundedSquare Then
        ShapeEx1.Shape = seShapeRoundedSquare
    ElseIf mShape = ledRoundedRectangle Then
        ShapeEx1.Shape = seShapeRoundedRectangle
    Else ' round
        ShapeEx1.Shape = seShapeCircle
    End If
    ShapeEx1.BorderWidth = mBorderWidth
    ShapeEx1.BorderColor = mBorderColor
    
    SetState
End Sub

Private Sub SetColor()
    If mColor >= ledCustomColor Then Exit Sub
    If mColor = ledRed Then
        mBorderColor = &H626479
        mColorOn = vbRed
        mColorOff = &HA5A6B6
    ElseIf mColor = ledGreen Then
        mColorOn = vbGreen
        mBorderColor = &H6B8B6F
        mColorOff = &HA3B8A5
    ElseIf mColor = ledBlue Then
        mColorOn = &HFFCBAE     ' &H00DF684A&
        mBorderColor = &H9F7E79
        mColorOff = &HCCBAB7
    ElseIf mColor = ledYellow Then
        mColorOn = vbYellow
        mBorderColor = &H678F8E
        mColorOff = &HB8CBCB
    Else ' white
        mColorOn = 16777215 ' &H00DF684A&
        mBorderColor = &H8E8E8E     ' &H4EA09C
        mColorOff = &HD7D7D7
    End If
    If mStyle = ledStyle3D Then
        If (mColor = ledBlue) Then
            mColorOff = ShiftColor(mColorOff, vbWhite, 40)
        ElseIf (mColor = ledWhite) Then
            mColorOff = ShiftColor(mColorOff, vbWhite, 200)
        ElseIf (mColor = ledYellow) Then
            mColorOff = ShiftColor(mColorOff, vbWhite, 160)
        Else
            mColorOff = ShiftColor(mColorOff, vbWhite, 80)
        End If
        mBorderColor = ShiftColor(mBorderColor, vbWhite, 180)
    End If
    SetOn = mStateOn
End Sub

Private Sub SetState()
    If mState = ledBlinking Then
        SetOn = True
        If Ambient.UserMode Then tmrBlink.Enabled = True
        'tmrBlink.Enabled = True
    ElseIf mState = ledOff Then
        tmrBlink.Enabled = False
        SetOn = False
    Else ' on
        tmrBlink.Enabled = False
        SetOn = True
    End If
End Sub

Private Property Let SetOn(nValue As Boolean)
    If nValue Then
        ShapeEx1.FillColor = mColorOn
    Else
        ShapeEx1.FillColor = mColorOff
    End If
    mStateOn = nValue
    If mStyle = ledStyle3D Then
        If mStateOn Then
            ShapeEx1.Style3D = seStyle3DBoth
        Else
            ShapeEx1.Style3D = seStyle3DShadow
        End If
    Else
        ShapeEx1.Style3D = seStyle3DNone
    End If
End Property

' From Leandro Ascierto
Private Function ShiftColor(ByVal clrFirst As Long, ByVal clrSecond As Long, ByVal lAlpha As Long) As Long
    Dim clrFore(3)         As Byte
    Dim clrBack(3)         As Byte
 
    OleTranslateColor clrFirst, 0, VarPtr(clrFore(0))
    OleTranslateColor clrSecond, 0, VarPtr(clrBack(0))
    
    clrFore(0) = (clrFore(0) * lAlpha + clrBack(0) * (255 - lAlpha)) / 255
    clrFore(1) = (clrFore(1) * lAlpha + clrBack(1) * (255 - lAlpha)) / 255
    clrFore(2) = (clrFore(2) * lAlpha + clrBack(2) * (255 - lAlpha)) / 255
     
    CopyMemory ShiftColor, clrFore(0), 4
End Function
