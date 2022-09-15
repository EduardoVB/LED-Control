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
      Top             =   420
      Width           =   912
      _ExtentX        =   1609
      _ExtentY        =   762
      FillStyle       =   0
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

Public Enum veLEDShapeConstants
    veLedRound
    veLedSquare
    veLedRectangle
    veLedRoundedSquare
    veLedRoundedRectangle
End Enum

Public Enum veLEDBlinkTypeConstants
    veLedShorter
    veLedShort
    veLedMedium
    veLedLong
    veLedDouble
End Enum

Public Enum veLEDStateConstants
    veLedOff
    veLedOn
    veLedBlinking
End Enum

Public Enum veLEDColorConstants
    veLedRed
    veLedGreen
    veLedBlue
    veLedYellow
    veLedWhite
    veLedCustomColor
End Enum

' Property defaults
Private Const mdef_Shape = veLedRound
Private Const mdef_BlinkRate = 700
Private Const mdef_BlinkType = veLedShort
Private Const mdef_BorderWidth = 2
Private Const mdef_BorderColor = &HC0&
Private Const mdef_ColorOn = vbRed
Private Const mdef_ColorOff = &H808080
Private Const mdef_State = veLedOn
Private Const mdef_Color = veLedRed
Private Const mdef_ToggleOnClick = False

' Properties
Private mShape As veLEDShapeConstants
Private mBlinkRate As Long
Private mBlinkType As veLEDBlinkTypeConstants
Private mBorderWidth As Long
Private mBorderColor As Long
Private mColorOn As Long
Private mColorOff As Long
Private mState As veLEDStateConstants
Private mColor As veLEDColorConstants
Private mToggleOnClick As Boolean

Private Sub ShapeEx1_Click()
    If mToggleOnClick Then
        If mState = veLedOn Then
            State = veLedOff
        ElseIf mState = veLedOff Then
            State = veLedOn
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
    
    t = Round((Timer * 1000 Mod mBlinkRate) / 100)
    'Debug.Print t

    If mBlinkType = veLedShorter Then
        SetOn = (t = 0)
    ElseIf mBlinkType = veLedShort Then
        SetOn = (t = 0) Or (t = 1)
    ElseIf mBlinkType = veLedDouble Then
        SetOn = (t = 0) Or (t = 2)
    ElseIf mBlinkType = veLedLong Then
        SetOn = (t <> 0) And (t <> 1)
    Else ' medium
        SetOn = Round((Timer * 1000 Mod mBlinkRate) / 100) > (mBlinkRate / 100 / 2)
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
    mBlinkRate = mdef_BlinkRate
    mBlinkType = mdef_BlinkType
    mBorderWidth = mdef_BorderWidth
    mBorderColor = mdef_BorderColor
    mColorOn = mdef_ColorOn
    mColorOff = mdef_ColorOff
    mState = mdef_State
    mColor = mdef_Color
    mToggleOnClick = mdef_ToggleOnClick
    ShowControl
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or (KeyAscii = vbKeySpace) Then
        ShapeEx1_Click
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mShape = PropBag.ReadProperty("Shape", mdef_Shape)
    mBlinkRate = PropBag.ReadProperty("BlinkRate", mdef_BlinkRate)
    mBlinkType = PropBag.ReadProperty("BlinkType", mdef_BlinkType)
    mBorderWidth = PropBag.ReadProperty("BorderWidth", mdef_BorderWidth)
    mBorderColor = PropBag.ReadProperty("BorderColor", mdef_BorderColor)
    mColorOn = PropBag.ReadProperty("ColorOn", mdef_ColorOn)
    mColorOff = PropBag.ReadProperty("ColorOff", mdef_ColorOff)
    mState = PropBag.ReadProperty("State", mdef_State)
    mColor = PropBag.ReadProperty("Color", mdef_Color)
    mToggleOnClick = PropBag.ReadProperty("ToggleOnClick", mdef_ToggleOnClick)
    ShowControl
End Sub

Private Sub UserControl_Resize()
    Dim iWidth As Long
    
    If (UserControl.Height < 7 * Screen.TwipsPerPixelY) Then UserControl.Height = 7 * Screen.TwipsPerPixelY
    Select Case mShape
        Case veLedRound, veLedSquare, veLedRoundedSquare
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
    PropBag.WriteProperty "BlinkRate", mBlinkRate, mdef_BlinkRate
    PropBag.WriteProperty "BlinkType", mBlinkType, mdef_BlinkType
    PropBag.WriteProperty "BorderWidth", mBorderWidth, mdef_BorderWidth
    PropBag.WriteProperty "BorderColor", mBorderColor, mdef_BorderColor
    PropBag.WriteProperty "ColorOn", mColorOn, mdef_ColorOn
    PropBag.WriteProperty "ColorOff", mColorOff, mdef_ColorOff
    PropBag.WriteProperty "State", mState, mdef_State
    PropBag.WriteProperty "Color", mColor, mdef_Color
    PropBag.WriteProperty "ToggleOnClick", mToggleOnClick, mdef_ToggleOnClick
End Sub


Public Property Get Shape() As veLEDShapeConstants
    Shape = mShape
End Property

Public Property Let Shape(nValue As veLEDShapeConstants)
    If nValue <> mShape Then
        If (nValue < veLedRound) Or (nValue > veLedRoundedRectangle) Then Err.Raise 380, TypeName(Me): Exit Property
        mShape = nValue
        UserControl_Resize
        ShowControl
        PropertyChanged "Shape"
    End If
End Property


Public Property Get BlinkRate() As Long
    BlinkRate = mBlinkRate
End Property

Public Property Let BlinkRate(nValue As Long)
    If nValue <> mBlinkRate Then
        If (nValue < 300) Or (nValue > 60000) Then Err.Raise 380, TypeName(Me): Exit Property
        mBlinkRate = nValue
        PropertyChanged "BlinkRate"
    End If
End Property


Public Property Get BlinkType() As veLEDBlinkTypeConstants
    BlinkType = mBlinkType
End Property

Public Property Let BlinkType(nValue As veLEDBlinkTypeConstants)
    If nValue <> mBlinkType Then
        If (nValue < veLedShorter) Or (nValue > veLedDouble) Then Err.Raise 380, TypeName(Me): Exit Property
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
        mColor = veLedCustomColor
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
        mColor = veLedCustomColor
        If mState = veLedOn Then
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
        mColor = veLedCustomColor
        If mState = veLedOff Then
            ShowControl
        End If
        PropertyChanged "ColorOff"
    End If
End Property


Public Property Get State() As veLEDStateConstants
    State = mState
End Property

Public Property Let State(nValue As veLEDStateConstants)
    If nValue <> mState Then
        If (nValue < veLedOff) Or (nValue > veLedBlinking) Then Err.Raise 380, TypeName(Me): Exit Property
        mState = nValue
        SetState
        PropertyChanged "State"
    End If
End Property


Public Property Get Color() As veLEDColorConstants
    Color = mColor
End Property

Public Property Let Color(nValue As veLEDColorConstants)
    If nValue <> mColor Then
        If (nValue < veLedRed) Or (nValue > veLedCustomColor) Then Err.Raise 380, TypeName(Me): Exit Property
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


Private Sub ShowControl()
    UserControl.BackColor = Ambient.BackColor
    SetColor
    If mShape = veLedSquare Then
        ShapeEx1.Shape = veShapeSquare
    ElseIf mShape = veLedRectangle Then
        ShapeEx1.Shape = veShapeRectangle
    ElseIf mShape = veLedRoundedSquare Then
        ShapeEx1.Shape = veShapeRoundedSquare
    ElseIf mShape = veLedRoundedRectangle Then
        ShapeEx1.Shape = veShapeRoundedRectangle
    Else ' round
        ShapeEx1.Shape = veShapeCircle
    End If
    ShapeEx1.BorderWidth = mBorderWidth
    ShapeEx1.BorderColor = mBorderColor
    SetState
End Sub

Private Sub SetColor()
    If mColor >= veLedCustomColor Then Exit Sub
    If mColor = veLedRed Then
        mBorderColor = &H626479
        mColorOn = vbRed
        mColorOff = &HA5A6B6
    ElseIf mColor = veLedGreen Then
        mColorOn = vbGreen
        mBorderColor = &H6B8B6F
        mColorOff = &HA3B8A5
    ElseIf mColor = veLedBlue Then
        mColorOn = &HFFCBAE     ' &H00DF684A&
        mBorderColor = &H9F7E79
        mColorOff = &HCCBAB7
    ElseIf mColor = veLedYellow Then
        mColorOn = vbYellow
        mBorderColor = &H678F8E
        mColorOff = &HB8CBCB
    Else ' white
        mColorOn = 16777215 ' &H00DF684A&
        mBorderColor = &H8E8E8E     ' &H4EA09C
        mColorOff = &HD7D7D7
    End If
End Sub

Private Sub SetState()
    If mState = veLedBlinking Then
        SetOn = True
        If Ambient.UserMode Then tmrBlink.Enabled = True
        'tmrBlink.Enabled = True
    ElseIf mState = veLedOff Then
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
End Property
