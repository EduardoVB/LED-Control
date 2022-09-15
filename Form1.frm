VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3252
   ClientLeft      =   2112
   ClientTop       =   2736
   ClientWidth     =   4836
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3252
   ScaleWidth      =   4836
   Begin VB.ComboBox cboStyle 
      Height          =   336
      ItemData        =   "Form1.frx":0000
      Left            =   2340
      List            =   "Form1.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   2640
      Width           =   1992
   End
   Begin VB.ComboBox cboToggleOnClick 
      Height          =   336
      ItemData        =   "Form1.frx":0016
      Left            =   2340
      List            =   "Form1.frx":0020
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   2160
      Width           =   1992
   End
   Begin VB.ComboBox cboShape 
      Height          =   336
      ItemData        =   "Form1.frx":0031
      Left            =   2340
      List            =   "Form1.frx":0044
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1680
      Width           =   1992
   End
   Begin VB.TextBox txtBlinkPeriod 
      Height          =   360
      Left            =   2340
      TabIndex        =   8
      Text            =   "700"
      Top             =   1200
      Width           =   1992
   End
   Begin VB.ComboBox cboBlinkType 
      Height          =   336
      ItemData        =   "Form1.frx":0085
      Left            =   2340
      List            =   "Form1.frx":0098
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   720
      Width           =   1992
   End
   Begin VB.ComboBox cboState 
      Height          =   336
      ItemData        =   "Form1.frx":00C1
      Left            =   2340
      List            =   "Form1.frx":00CE
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   240
      Width           =   1992
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Style:"
      Height          =   312
      Left            =   780
      TabIndex        =   15
      Top             =   2700
      Width           =   1512
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "ToggleOnClick:"
      Height          =   312
      Left            =   780
      TabIndex        =   13
      Top             =   2220
      Width           =   1512
   End
   Begin Proyect1.LED LED1 
      Height          =   324
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   300
      Width           =   324
      _ExtentX        =   572
      _ExtentY        =   572
      BorderColor     =   9342606
      ColorOn         =   16777215
      ColorOff        =   14145495
      Color           =   4
   End
   Begin Proyect1.LED LED1 
      Height          =   324
      Index           =   1
      Left            =   360
      TabIndex        =   3
      Top             =   780
      Width           =   324
      _ExtentX        =   572
      _ExtentY        =   572
      BorderColor     =   6448249
      ColorOff        =   10856118
   End
   Begin Proyect1.LED LED1 
      Height          =   324
      Index           =   2
      Left            =   360
      TabIndex        =   6
      Top             =   1260
      Width           =   324
      _ExtentX        =   572
      _ExtentY        =   572
      BorderColor     =   7048047
      ColorOn         =   65280
      ColorOff        =   10729637
      Color           =   1
   End
   Begin Proyect1.LED LED1 
      Height          =   324
      Index           =   3
      Left            =   360
      TabIndex        =   9
      Top             =   1740
      Width           =   324
      _ExtentX        =   572
      _ExtentY        =   572
      BorderColor     =   6786958
      ColorOn         =   65535
      ColorOff        =   12110795
      Color           =   3
   End
   Begin Proyect1.LED LED1 
      Height          =   324
      Index           =   4
      Left            =   360
      TabIndex        =   12
      Top             =   2220
      Width           =   324
      _ExtentX        =   572
      _ExtentY        =   572
      BorderColor     =   10452601
      ColorOn         =   16763822
      ColorOff        =   13417143
      Color           =   2
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Shape:"
      Height          =   312
      Left            =   1440
      TabIndex        =   10
      Top             =   1740
      Width           =   852
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "BlinkPeriod:"
      Height          =   312
      Left            =   1020
      TabIndex        =   7
      Top             =   1260
      Width           =   1272
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "BlinkType:"
      Height          =   312
      Left            =   1440
      TabIndex        =   4
      Top             =   780
      Width           =   852
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "State: "
      Height          =   312
      Left            =   1440
      TabIndex        =   1
      Top             =   300
      Width           =   852
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboBlinkType_Click()
    Dim c  As Long
    
    For c = LED1.lbound To LED1.UBound
        LED1(c).BlinkType = cboBlinkType.ListIndex
    Next
End Sub

Private Sub cboShape_Click()
    Dim c  As Long
    
    For c = LED1.lbound To LED1.UBound
        LED1(c).Shape = cboShape.ListIndex
        If (cboShape.ListIndex = ledRectangle) Or (cboShape.ListIndex = ledRoundedRectangle) Then
            LED1(c).Width = LED1(c).Height * 0.7
        End If
    Next
End Sub

Private Sub cboState_Click()
    Dim c  As Long
    
    For c = LED1.lbound To LED1.UBound
        LED1(c).State = cboState.ListIndex
    Next
End Sub

Private Sub cboStyle_Click()
    Dim c  As Long
    
    For c = LED1.lbound To LED1.UBound
        LED1(c).Style = cboStyle.ListIndex
        If cboStyle.ListIndex = ledStyle3D Then
            LED1(c).BorderWidth = 1
        Else
            LED1(c).BorderWidth = 2
        End If
    Next
End Sub

Private Sub cboToggleOnClick_Click()
    Dim c  As Long
    
    For c = LED1.lbound To LED1.UBound
        LED1(c).ToggleOnClick = CBool(cboToggleOnClick.ListIndex)
    Next
End Sub

Private Sub Form_Load()
    cboState.ListIndex = 1
    cboBlinkType.ListIndex = 1
    cboShape.ListIndex = 0
    cboToggleOnClick.ListIndex = 0
    cboStyle.ListIndex = 1
End Sub

Private Sub LED1_Click(Index As Integer)
    If LED1(Index).ToggleOnClick Then
        If LED1(Index).State = ledBlinking Then
            MsgBox "While blinking can't be toggled.", vbExclamation
        End If
    End If
End Sub

Private Sub txtBlinkPeriod_Change()
    Dim c  As Long
    Dim iNewVal As Long
    
    iNewVal = Val(txtBlinkPeriod.Text)
    If (iNewVal < 300) Or (iNewVal > 60000) Then
        txtBlinkPeriod.BackColor = &HC0C0FF
    Else
        txtBlinkPeriod.BackColor = vbWindowBackground
        For c = LED1.lbound To LED1.UBound
            LED1(c).BlinkPeriod = iNewVal
        Next
    End If
End Sub
