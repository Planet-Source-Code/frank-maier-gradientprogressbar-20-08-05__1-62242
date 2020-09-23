VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "GradientProgressbar"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10005
   LinkTopic       =   "Form1"
   ScaleHeight     =   426
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   667
   StartUpPosition =   3  'Windows-Standard
   Begin ComctlLib.ProgressBar ProgressBarVB5 
      Height          =   495
      Left            =   5040
      TabIndex        =   7
      Top             =   2640
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   873
      _Version        =   327682
      Appearance      =   0
   End
   Begin MSComctlLib.ProgressBar ProgressBarVB6 
      Height          =   375
      Left            =   5040
      TabIndex        =   8
      Top             =   1560
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdSpeedTest 
      Caption         =   "Run"
      Height          =   375
      Left            =   6840
      TabIndex        =   6
      Top             =   3600
      Width           =   1095
   End
   Begin Gradient.GradientProgressBar GradientProgressBar2 
      Height          =   375
      Left            =   5040
      Top             =   480
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   661
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3960
      Top             =   5520
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test It"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   5520
      Width           =   3615
   End
   Begin Gradient.GradientProgressBar GradientProgressBar1 
      Height          =   375
      Index           =   0
      Left            =   120
      Top             =   3960
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      MouseChange     =   -1  'True
      Value           =   100
   End
   Begin Gradient.GradientProgressBar GradientProgressBar1 
      Height          =   375
      Index           =   1
      Left            =   120
      Top             =   120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      MouseChange     =   -1  'True
      Value           =   90
      BorderStyle     =   0
   End
   Begin Gradient.GradientProgressBar GradientProgressBar1 
      Height          =   375
      Index           =   2
      Left            =   120
      Top             =   600
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      MouseChange     =   -1  'True
      BarColor        =   12937777
      BarColorTop     =   12937777
      BarColorBottom  =   12937777
      Value           =   80
      BorderStyle     =   1
      Space           =   5
   End
   Begin Gradient.GradientProgressBar GradientProgressBar1 
      Height          =   375
      Index           =   3
      Left            =   120
      Top             =   1080
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      MouseChange     =   -1  'True
      BarColor        =   65535
      Value           =   70
      BorderStyle     =   2
      Style           =   1
      Space           =   2
   End
   Begin Gradient.GradientProgressBar GradientProgressBar1 
      Height          =   375
      Index           =   4
      Left            =   120
      Top             =   1560
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      MousePointer    =   3
      MouseChange     =   -1  'True
      Value           =   60
      BorderStyle     =   3
      Style           =   2
      Space           =   5
   End
   Begin Gradient.GradientProgressBar GradientProgressBar1 
      Height          =   375
      Index           =   5
      Left            =   120
      Top             =   2040
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      MousePointer    =   7
      MouseChange     =   -1  'True
      BarColor        =   3352786
      BarColorTop     =   11512813
      BarColorBottom  =   9670375
      Value           =   50
      BorderStyle     =   4
      Style           =   1
      Space           =   1
   End
   Begin Gradient.GradientProgressBar GradientProgressBar1 
      Height          =   375
      Index           =   6
      Left            =   120
      Top             =   2520
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      MouseChange     =   -1  'True
      BarColor        =   2675406
      BarColorTop     =   11267564
      BarColorBottom  =   9365477
      Value           =   40
      BorderStyle     =   5
      Style           =   1
      Space           =   1
   End
   Begin Gradient.GradientProgressBar GradientProgressBar1 
      Height          =   375
      Index           =   7
      Left            =   120
      Top             =   3000
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      MouseChange     =   -1  'True
      BarColor        =   13772857
      BarColorTop     =   15575985
      BarColorBottom  =   15175319
      Value           =   30
      BorderStyle     =   6
      Style           =   2
      Space           =   1
   End
   Begin Gradient.GradientProgressBar GradientProgressBar1 
      Height          =   375
      Index           =   8
      Left            =   120
      Top             =   3480
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      MouseChange     =   -1  'True
      Value           =   30
      BorderStyle     =   7
      Space           =   1
   End
   Begin Gradient.GradientProgressBar GradientProgressBar1 
      Height          =   375
      Index           =   9
      Left            =   120
      Top             =   4440
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      MouseChange     =   -1  'True
      BorderColor     =   3724597
      Value           =   30
   End
   Begin Gradient.GradientProgressBar GradientProgressBar1 
      Height          =   375
      Index           =   10
      Left            =   120
      Top             =   4920
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      MouseChange     =   -1  'True
      BackColor       =   65535
      BorderColor     =   255
      BarColor        =   49344
      BarColorTop     =   33023
      BarColorBottom  =   192
      Value           =   30
      Style           =   2
   End
   Begin Gradient.GradientProgressBar GradientProgressBar1 
      Height          =   5175
      Index           =   11
      Left            =   3840
      Top             =   120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   9128
      MouseChange     =   -1  'True
      Value           =   90
      Orientation     =   1
   End
   Begin Gradient.GradientProgressBar GradientProgressBar1 
      Height          =   5175
      Index           =   12
      Left            =   4560
      Top             =   120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   9128
      MouseChange     =   -1  'True
      BackColor       =   12632256
      BorderColor     =   16512
      BarColor        =   49152
      BarColorTop     =   192
      BarColorBottom  =   16711935
      Value           =   90
      Style           =   2
      Orientation     =   1
   End
   Begin Gradient.GradientProgressBar GradientProgressBar1 
      Height          =   5175
      Index           =   13
      Left            =   4200
      Top             =   120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   9128
      MouseChange     =   -1  'True
      Value           =   90
      Style           =   1
      Orientation     =   1
   End
   Begin Gradient.GradientProgressBar GradientProgressBar1 
      Height          =   135
      Index           =   14
      Left            =   0
      Top             =   6240
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   238
      MousePointer    =   3
      MouseChange     =   -1  'True
      BackColor       =   -2147483633
      BarColor        =   16777215
      BarColorTop     =   32768
      BarColorBottom  =   32768
      Value           =   90
      BorderStyle     =   0
   End
   Begin VB.Label lblSpeedVB6 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fest Einfach
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7080
      TabIndex        =   10
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label3 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "VB6 Progressbar: "
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5040
      TabIndex        =   9
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label lblSpeedVB5 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fest Einfach
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7080
      TabIndex        =   5
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label lblSpeedGrp 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fest Einfach
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7080
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label2 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "VB5 Progressbar: "
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5040
      TabIndex        =   3
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label1 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "GradientProgressbar: "
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5040
      TabIndex        =   2
      Top             =   960
      Width           =   1935
   End
   Begin VB.Line Line1 
      X1              =   328
      X2              =   328
      Y1              =   8
      Y2              =   392
   End
   Begin VB.Label lblSpeed 
      Caption         =   "Speed test"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsTiming As cTiming

Private Sub cmdSpeedTest_Click()
Dim lngI As Long
    
    Set clsTiming = New cTiming
    
    'GradientProgressbar
    clsTiming.Reset
    For lngI = 0 To 100
        GradientProgressBar2.Value = lngI
    Next lngI
    lblSpeedGrp.Caption = Round(clsTiming.Elapsed, 3) & "ms"
    
    'VB6 ProgressBar
    clsTiming.Reset
    For lngI = 0 To 100
        ProgressBarVB6.Value = lngI
    Next lngI
    lblSpeedVB6.Caption = Round(clsTiming.Elapsed, 3) & "ms"
    
    'VB5 ProgressBar
    clsTiming.Reset
    For lngI = 0 To 100
        ProgressBarVB5.Value = lngI
    Next lngI
    lblSpeedVB5.Caption = Round(clsTiming.Elapsed, 3) & "ms"
    
    Set clsTiming = Nothing
End Sub

Private Sub cmdTest_Click()
Dim lngI As Long

    If cmdTest.Caption = "Stop" Then
        Timer1.Enabled = False
        cmdTest.Caption = "Test it"
    Else
        For lngI = 0 To GradientProgressBar1.UBound
            GradientProgressBar1(lngI).Value = 0
        Next lngI
    
        Timer1.Enabled = True
        cmdTest.Caption = "Stop"
    End If
End Sub

Private Sub Form_Resize()
    GradientProgressBar1(14).Move 0, ScaleHeight - 10, ScaleWidth, 10
End Sub

Private Sub Timer1_Timer()
Dim lngI As Long

    For lngI = 0 To GradientProgressBar1.UBound
        GradientProgressBar1(lngI).Value = GradientProgressBar1(lngI).Value + 1
    Next lngI
    
    If GradientProgressBar1(0).Value = 100 Then Timer1.Enabled = False
End Sub
