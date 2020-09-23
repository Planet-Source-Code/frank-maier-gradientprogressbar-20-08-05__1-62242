VERSION 5.00
Begin VB.UserControl GradientProgressBar 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4695
   ScaleHeight     =   25
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   313
End
Attribute VB_Name = "GradientProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================
' Module:        usrGradientProgressBar.ctl
' Author:        Frank Maier - 2005
' Dependencies:  modGradient.bas
' Last revision: 2005.08.20
'================================================

Option Explicit

Private Const TRANSPARENT       As Long = 1

Private Const BDR_RAISEDOUTER   As Long = &H1 'API
Private Const BDR_SUNKENOUTER   As Long = &H2 'API
Private Const BDR_RAISEDINNER   As Long = &H4 'API
Private Const BDR_SUNKENINNER   As Long = &H8 'API

Private Const BF_LEFT           As Long = &H1
Private Const BF_RIGHT          As Long = &H4
Private Const BF_TOP            As Long = &H2
Private Const BF_BOTTOM         As Long = &H8

Private Const BF_SOFT           As Long = &H1000
Private Const BF_RECT           As Long = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Const EDGE_BUMP         As Long = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Private Const EDGE_ETCHED       As Long = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Private Const EDGE_RAISED       As Long = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN       As Long = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)

'Standard-Eigenschaftswerte:
Private Const mconBorderStyle       As Long = 8
Private Const mconStyle             As Long = 0
Private Const mconMaxValue          As Long = 100
Private Const mconMinValue          As Long = 0
Private Const mconValue             As Long = 0
Private Const mconBarColorTop       As Long = 11333035
Private Const mconBarColor          As Long = 3724597
Private Const mconBarColorBottom    As Long = 9496462
Private Const mconBackColor         As Long = 16777215
Private Const mconEnabled           As Long = True
Private Const mconBorderColor       As Long = 6842472
Private Const mconSpace             As Long = 0
Private Const mconOrientation       As Long = 0
Private Const mconMouseChange       As Boolean = False

Public Enum enumOrientation
    Horizontal = 0
    Vertical = 1
End Enum

Public Enum enumStylePrgB
    Bar = &H0
    Bar_Soft = &H1
    Smooth = &H2
End Enum

Public Enum enumBorderStylePrgB
    NoBorder = &H0
    SoftRaised = &H1
    Raised = &H2
    SoftSunken = &H3
    Sunken = &H4
    Bump = &H5
    Etched = &H6
    FlatBorder = &H7
    XP = &H8
End Enum

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'Eigenschaftsvariablen:
Private mlngBorderStyle     As enumBorderStylePrgB
Private mlngStyle           As enumStylePrgB
Private mlngMaxValue        As Long
Private mlngMinValue        As Long
Private mlngValue           As Long
Private moleBarColorTop     As OLE_COLOR
Private moleBarColor        As OLE_COLOR
Private moleBarColorBottom  As OLE_COLOR
Private moleBackColor       As OLE_COLOR
Private moleBorderColor     As OLE_COLOR
Private mblnEnabled         As Boolean
Private mlngSpace           As Long
Private mlngOrientation     As enumOrientation
Private mblnMouseChange     As Boolean

Private mlngDCBar           As Long
Private mlngDCBkg           As Long
Private mrctBar             As RECT
Private mrctBkg             As RECT

Private mblnLoaded          As Boolean
Private mlngLeftBar         As Long
Private mlngTopBar          As Long
Private mblnMouseDown       As Boolean

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Private Declare Sub OleTranslateColor Lib "oleaut32.dll" (ByVal clr As Long, ByVal hpal As Long, ByRef lpcolorref As Long)
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'Ereignisdeklarationen:
Event Click()
Event DblClick()
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event Resize()

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
    
    mblnMouseDown = True
    UserControl_MouseMove Button, Shift, x, y
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
    
    mblnMouseDown = False
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim lngAbs      As Long
Dim lngTemp     As Long

    RaiseEvent MouseMove(Button, Shift, x, y)
    
    If mblnMouseChange Then
        If mblnMouseDown Then
            lngAbs = mlngMaxValue - mlngMinValue
            If mlngOrientation = Horizontal Then
                lngTemp = x - mlngLeftBar
                mlngValue = lngTemp * lngAbs / mrctBar.Right
            Else
                lngTemp = mrctBkg.Bottom - y
                mlngValue = lngTemp * lngAbs / mrctBar.Bottom
            End If
            uscPaint
        End If
    End If
End Sub

Private Sub UserControl_Terminate()
    DeleteDC mlngDCBar
    DeleteDC mlngDCBkg
End Sub

Private Sub UserControl_Resize()
    If mblnLoaded Then
        RaiseEvent Resize
    
        uscCreateBkg
        uscCreateBar
        uscPaint
    End If
End Sub

'Create Background
Private Sub uscCreateBkg()

Dim rctTemp As RECT
Dim lngColorTemp As Long

Dim lngDCTemp As Long
Dim lngDCBitmap As Long

Dim lngBrush    As Long

    If mlngDCBkg <> 0 Then
        DeleteDC mlngDCBkg
    End If
    
    lngDCTemp = GetDC(0)
    lngDCBitmap = CreateCompatibleBitmap(lngDCTemp, ScaleWidth, ScaleHeight)
    mlngDCBkg = CreateCompatibleDC(lngDCTemp)

    SelectObject mlngDCBkg, lngDCBitmap
    
    DeleteObject lngDCBitmap
    DeleteDC lngDCTemp
    
    SetBkMode mlngDCBkg, TRANSPARENT

    'Draw
    SetRect mrctBkg, 0, 0, ScaleWidth, ScaleHeight
    
    'Bkg
    lngBrush = CreateSolidBrush(TranslateColor(moleBackColor))
    FillRect mlngDCBkg, mrctBkg, lngBrush
    DeleteObject lngBrush
    
    'Border
    Select Case mlngBorderStyle
'   Case NoBorder
    Case SoftRaised
        DrawEdge mlngDCBkg, mrctBkg, BDR_RAISEDOUTER, BF_SOFT Or BF_RECT
    Case Raised
        DrawEdge mlngDCBkg, mrctBkg, EDGE_RAISED, BF_SOFT Or BF_RECT
    Case SoftSunken
        DrawEdge mlngDCBkg, mrctBkg, BDR_SUNKENOUTER, BF_SOFT Or BF_RECT
    Case Sunken
        DrawEdge mlngDCBkg, mrctBkg, EDGE_SUNKEN, BF_SOFT Or BF_RECT
    Case Bump
        DrawEdge mlngDCBkg, mrctBkg, EDGE_BUMP, BF_SOFT Or BF_RECT
    Case Etched
        DrawEdge mlngDCBkg, mrctBkg, EDGE_ETCHED, BF_RECT
    Case FlatBorder
        DrawBorder mlngDCBkg, mrctBkg, moleBorderColor
    Case XP
        SetRect rctTemp, 2, 2, mrctBkg.Right - 1, mrctBkg.Bottom - 1
        lngColorTemp = BlendColor(moleBorderColor, moleBackColor, 228)
        DrawBorder mlngDCBkg, rctTemp, lngColorTemp
            
        CopyRect rctTemp, mrctBkg
        rctTemp.Top = 1
        rctTemp.Left = 1
        lngColorTemp = BlendColor(moleBorderColor, moleBackColor, 147)
        DrawBorder mlngDCBkg, rctTemp, lngColorTemp
        
        DrawBorder mlngDCBkg, mrctBkg, moleBorderColor
    End Select
    
End Sub

'Create Bar
Private Sub uscCreateBar()

Dim lngDCTemp   As Long
Dim lngDCBitmap As Long

Dim lngBrush    As Long
Dim lngI        As Long

Dim rctTemp     As RECT

    
    'Calc size
    Select Case mlngBorderStyle
    Case NoBorder
        mlngLeftBar = 0 + mlngSpace
        mlngTopBar = 0 + mlngSpace
        SetRect mrctBar, 0, 0, ScaleWidth - (2 * mlngSpace), ScaleHeight - (2 * mlngSpace)
    Case SoftRaised, SoftSunken, FlatBorder
        mlngLeftBar = 1 + mlngSpace
        mlngTopBar = 1 + mlngSpace
        SetRect mrctBar, 0, 0, ScaleWidth - 2 - (2 * mlngSpace), ScaleHeight - 2 - (2 * mlngSpace)
    Case Raised, Sunken, Bump
        mlngLeftBar = 2 + mlngSpace
        mlngTopBar = 2 + mlngSpace
        SetRect mrctBar, 0, 0, ScaleWidth - 4 - (2 * mlngSpace), ScaleHeight - 4 - (2 * mlngSpace)
    Case Etched
        mlngLeftBar = 1 + mlngSpace
        mlngTopBar = 1 + mlngSpace
        SetRect mrctBar, 0, 0, ScaleWidth - 3 - (2 * mlngSpace), ScaleHeight - 3 - (2 * mlngSpace)
    Case XP
        mlngLeftBar = 3 + mlngSpace
        If mlngOrientation = Horizontal Then
            mlngTopBar = 3 + mlngSpace
            SetRect mrctBar, 0, 0, ScaleWidth - 5 - (2 * mlngSpace), ScaleHeight - 6 - (2 * mlngSpace)
        Else    'Vertical
            mlngTopBar = 2 + mlngSpace
            SetRect mrctBar, 0, 0, ScaleWidth - 5 - (2 * mlngSpace), ScaleHeight - 5 - (2 * mlngSpace)
        End If
    End Select
    
    'Create
    If mlngDCBar <> 0 Then
        DeleteDC mlngDCBar
    End If
    
    lngDCTemp = GetDC(0)
    lngDCBitmap = CreateCompatibleBitmap(lngDCTemp, mrctBar.Right, mrctBar.Bottom)
    mlngDCBar = CreateCompatibleDC(lngDCTemp)
    
    SelectObject mlngDCBar, lngDCBitmap
    
    DeleteObject lngDCBitmap
    DeleteDC lngDCTemp
    
    SetBkMode mlngDCBar, TRANSPARENT
    
    'Fill Background
    lngBrush = CreateSolidBrush(TranslateColor(moleBarColor))
    FillRect mlngDCBar, mrctBar, lngBrush
    DeleteObject lngBrush
    
    'Style
    If mlngOrientation = Horizontal Then
        'Top
        PaintGradient mlngDCBar, 0, 0, mrctBar.Right, 6, TranslateColor(moleBarColorTop), TranslateColor(moleBarColor), gdVertical
        'Bottom
        PaintGradient mlngDCBar, 0, mrctBar.Bottom - 6, mrctBar.Right, 6, TranslateColor(moleBarColor), TranslateColor(moleBarColorBottom), gdVertical
    Else    'Vertical
        'Left
        PaintGradient mlngDCBar, 0, 0, 6, mrctBar.Bottom, TranslateColor(moleBarColorTop), TranslateColor(moleBarColor), gdHorizontal
        'Right
        PaintGradient mlngDCBar, mrctBar.Right - 6, 0, 6, mrctBar.Bottom, TranslateColor(moleBarColor), TranslateColor(moleBarColorBottom), gdHorizontal
    End If
    
    'Divide
    If mlngStyle = Bar Or mlngStyle = Bar_Soft Then
        lngBrush = CreateSolidBrush(TranslateColor(moleBackColor))
                
        CopyRect rctTemp, mrctBar
        
        If mlngOrientation = Horizontal Then
            For lngI = 6 To mrctBar.Right Step 6
                rctTemp.Left = lngI
                rctTemp.Right = lngI + 2
                FillRect mlngDCBar, rctTemp, lngBrush
                lngI = lngI + 2
            Next lngI
        Else    'Vertical
            For lngI = mrctBar.Bottom - 6 To 0 Step -6
                rctTemp.Top = lngI - 2
                rctTemp.Bottom = lngI
                FillRect mlngDCBar, rctTemp, lngBrush
                lngI = lngI - 2
            Next lngI
        End If
        
        DeleteObject lngBrush
    End If
    
End Sub

'Draw the hole Progressbar
Private Sub uscPaint()

Dim lngAbs As Long
Dim lngAbsValue As Long
Dim lngSize As Long

    'Draw Background
    BitBlt UserControl.hDC, 0, 0, mrctBkg.Right, mrctBkg.Bottom, mlngDCBkg, 0, 0, vbSrcCopy
    
    'Clac progress
    lngAbs = mlngMaxValue - mlngMinValue
    lngAbsValue = mlngValue - mlngMinValue
    If Orientation = Horizontal Then
        lngSize = lngAbsValue / lngAbs * mrctBar.Right
    Else    'Vertical
        lngSize = lngAbsValue / lngAbs * mrctBar.Bottom
    End If
    
    If mlngStyle = Bar Then
        If lngSize Mod 8 <> 0 Then
            lngSize = (Int(lngSize / 8) + 1) * 8
        End If
    End If
    
    'Draw Bar
    
    If Orientation = Horizontal Then
        BitBlt UserControl.hDC, mlngLeftBar, mlngTopBar, lngSize, mrctBar.Bottom, mlngDCBar, 0, 0, vbSrcCopy
    Else    'Vertical
        BitBlt UserControl.hDC, mlngLeftBar, mrctBkg.Bottom - lngSize - mlngTopBar, mrctBar.Right, lngSize, mlngDCBar, 0, mrctBar.Bottom - lngSize, vbSrcCopy
    End If
    
    UserControl.Refresh
End Sub

Private Sub UserControl_InitProperties()
    moleBackColor = mconBackColor
    mblnEnabled = mconEnabled
    mlngBorderStyle = mconBorderStyle
    mlngStyle = mconStyle
    mlngMaxValue = mconMaxValue
    mlngMinValue = mconMinValue
    mlngValue = mconValue
    moleBarColorTop = mconBarColorTop
    moleBarColor = mconBarColor
    moleBarColorBottom = mconBarColorBottom
    moleBorderColor = mconBorderColor
    mlngSpace = mconSpace
    mlngOrientation = mconOrientation
    mblnMouseChange = mconMouseChange
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        Set MouseIcon = .ReadProperty("MouseIcon", Nothing)
        UserControl.MousePointer = .ReadProperty("MousePointer", 0)
        mblnMouseChange = .ReadProperty("MouseChange", mconMouseChange)
        
        moleBackColor = .ReadProperty("BackColor", mconBackColor)
        moleBorderColor = .ReadProperty("BorderColor", mconBorderColor)
        moleBarColor = .ReadProperty("BarColor", mconBarColor)
        moleBarColorTop = .ReadProperty("BarColorTop", mconBarColorTop)
        moleBarColorBottom = .ReadProperty("BarColorBottom", mconBarColorBottom)
        
        mlngMaxValue = .ReadProperty("MaxValue", mconMaxValue)
        mlngMinValue = .ReadProperty("MinValue", mconMinValue)
        mlngValue = .ReadProperty("Value", mconValue)
        
        mlngBorderStyle = .ReadProperty("BorderStyle", mconBorderStyle)
        mlngStyle = .ReadProperty("Style", mconStyle)
        mlngSpace = .ReadProperty("Space", mconSpace)
        mlngOrientation = .ReadProperty("Orientation", mconOrientation)
        
        mblnEnabled = .ReadProperty("Enabled", mconEnabled)
    End With
    mblnLoaded = True
    UserControl_Resize
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        .WriteProperty "MouseIcon", MouseIcon, Nothing
        .WriteProperty "MousePointer", UserControl.MousePointer, 0
        .WriteProperty "MouseChange", mblnMouseChange, mconMouseChange
        
        .WriteProperty "BackColor", moleBackColor, mconBackColor
        .WriteProperty "BorderColor", moleBorderColor, mconBorderColor
        .WriteProperty "BarColor", moleBarColor, mconBarColor
        .WriteProperty "BarColorTop", moleBarColorTop, mconBarColorTop
        .WriteProperty "BarColorBottom", moleBarColorBottom, mconBarColorBottom
        
        .WriteProperty "MaxValue", mlngMaxValue, mconMaxValue
        .WriteProperty "MinValue", mlngMinValue, mconMinValue
        .WriteProperty "Value", mlngValue, mconValue

        .WriteProperty "BorderStyle", mlngBorderStyle, mconBorderStyle
        .WriteProperty "Style", mlngStyle, mconStyle
        .WriteProperty "Space", mlngSpace, mconSpace
        .WriteProperty "Orientation", mlngOrientation, mconOrientation
        
        .WriteProperty "Enabled", mblnEnabled, mconEnabled
    End With
End Sub

Public Property Get hDC() As Long
    hDC = UserControl.hDC
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property



Public Property Get MouseIcon() As Picture
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As Integer
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Get MouseChange() As Boolean
    MouseChange = mblnMouseChange
End Property

Public Property Let MouseChange(ByVal New_MouseChange As Boolean)
    mblnMouseChange = New_MouseChange
    PropertyChanged "MouseChange"
End Property


Public Property Get BackColor() As OLE_COLOR
    BackColor = moleBackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    moleBackColor = New_BackColor
    PropertyChanged "BackColor"
    If mblnLoaded Then
        uscCreateBkg
        uscCreateBar
        uscPaint
    End If
End Property

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = moleBorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    moleBorderColor = New_BorderColor
    PropertyChanged "BorderColor"
    If mblnLoaded Then
        uscCreateBkg
        uscCreateBar
        uscPaint
    End If
End Property

Public Property Get BarColor() As OLE_COLOR
    BarColor = moleBarColor
End Property

Public Property Let BarColor(ByVal New_BarColor As OLE_COLOR)
    moleBarColor = New_BarColor
    PropertyChanged "BarColor"
    If mblnLoaded Then
        uscCreateBar
        uscPaint
    End If
End Property

Public Property Get BarColorTop() As OLE_COLOR
    BarColorTop = moleBarColorTop
End Property

Public Property Let BarColorTop(ByVal New_BarColorTop As OLE_COLOR)
    moleBarColorTop = New_BarColorTop
    PropertyChanged "BarColorTop"
    If mblnLoaded Then
        uscCreateBar
        uscPaint
    End If
End Property

Public Property Get BarColorBottom() As OLE_COLOR
    BarColorBottom = moleBarColorBottom
End Property

Public Property Let BarColorBottom(ByVal New_BarColorBottom As OLE_COLOR)
    moleBarColorBottom = New_BarColorBottom
    PropertyChanged "BarColorBottom"
    If mblnLoaded Then
        uscCreateBar
        uscPaint
    End If
End Property



Public Property Get MaxValue() As Long
    MaxValue = mlngMaxValue
End Property

Public Property Let MaxValue(ByVal New_MaxValue As Long)
    mlngMaxValue = New_MaxValue
    PropertyChanged "MaxValue"
    mlngValue = OptimizeValue(mlngValue)
    If mblnLoaded Then
        uscPaint
    End If
End Property

Public Property Get MinValue() As Long
    MinValue = mlngMinValue
End Property

Public Property Let MinValue(ByVal New_MinValue As Long)
    mlngMinValue = New_MinValue
    PropertyChanged "MinValue"
    mlngValue = OptimizeValue(mlngValue)
    If mblnLoaded Then
        uscPaint
    End If
End Property

Public Property Get Value() As Long
    Value = mlngValue
End Property

Public Property Let Value(ByVal New_Value As Long)
    mlngValue = OptimizeValue(New_Value)
    PropertyChanged "Value"
    If mblnLoaded Then
        uscPaint
    End If
End Property



Public Property Get BorderStyle() As enumBorderStylePrgB
    BorderStyle = mlngBorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As enumBorderStylePrgB)
    mlngBorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
    If mblnLoaded Then
        uscCreateBkg
        uscCreateBar
        uscPaint
    End If
End Property

Public Property Get Style() As enumStylePrgB
    Style = mlngStyle
End Property

Public Property Let Style(ByVal New_Style As enumStylePrgB)
    mlngStyle = New_Style
    PropertyChanged "Style"
    If mblnLoaded Then
        uscCreateBar
        uscPaint
    End If
End Property

Public Property Get Space() As Long
    Space = mlngSpace
End Property

Public Property Let Space(ByVal New_Space As Long)
    mlngSpace = New_Space
    PropertyChanged "Space"
    If mblnLoaded Then
        uscCreateBar
        uscPaint
    End If
End Property

Public Property Get Orientation() As enumOrientation
    Orientation = mlngOrientation
End Property

Public Property Let Orientation(ByVal New_Orientation As enumOrientation)
    mlngOrientation = New_Orientation
    PropertyChanged "Orientation"
    If mblnLoaded Then
        uscCreateBar
        uscCreateBkg
        uscPaint
    End If
End Property

Public Property Get Enabled() As Boolean
    Enabled = mblnEnabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    mblnEnabled = New_Enabled
    PropertyChanged "Enabled"
End Property

Private Function OptimizeValue(ByVal lngNewValue As Long) As Long
    If lngNewValue < mlngMinValue Then
        OptimizeValue = mlngMinValue
    ElseIf lngNewValue > mlngMaxValue Then
        OptimizeValue = mlngMaxValue
    Else
        OptimizeValue = lngNewValue
    End If
End Function

Private Function TranslateColor(ByVal oleColor As OLE_COLOR) As Long
    OleTranslateColor oleColor, 0, TranslateColor
End Function

Private Sub DrawBorder(ByVal lngDC As Long, _
                       ByRef rctRect As RECT, _
                       ByVal oleColor As OLE_COLOR)

Dim lngBrush As Long
    
    lngBrush = CreateSolidBrush(TranslateColor(oleColor))
    FrameRect lngDC, rctRect, lngBrush
    DeleteObject lngBrush

End Sub

Private Function BlendColor(ByVal oleColor1 As OLE_COLOR, _
                            ByVal oleColor2 As OLE_COLOR, _
                            ByVal bytBlend As Byte) As Long

Dim lngColor1 As Long
Dim lngColor2 As Long

Dim lngR1    As Long
Dim lngG1    As Long
Dim lngB1    As Long

Dim lngR2    As Long
Dim lngG2    As Long
Dim lngB2    As Long

Dim lngR3    As Long
Dim lngG3    As Long
Dim lngB3    As Long

    lngColor1 = TranslateColor(oleColor1)
    lngColor2 = TranslateColor(oleColor2)

    lngR1 = lngColor1 And &HFF
    lngG1 = (lngColor1 And &HFF00&) \ 256
    lngB1 = (lngColor1 And &HFF0000) \ 65536

    lngR2 = lngColor2 And &HFF
    lngG2 = (lngColor2 And &HFF00&) \ 256
    lngB2 = (lngColor2 And &HFF0000) \ 65536

    lngR3 = lngR1 + ((lngR2 - lngR1) / 255 * bytBlend)
    lngG3 = lngG1 + ((lngG2 - lngG1) / 255 * bytBlend)
    lngB3 = lngB1 + ((lngB2 - lngB1) / 255 * bytBlend)

    BlendColor = RGB(lngR3, lngG3, lngB3)
End Function

