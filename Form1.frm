VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00606060&
   Caption         =   "Flip Digit Demo"
   ClientHeight    =   2730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3855
   LinkTopic       =   "Form1"
   ScaleHeight     =   182
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   257
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Count Tens"
      Height          =   570
      Left            =   900
      TabIndex        =   12
      Top             =   2100
      Width           =   930
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Reset"
      Height          =   480
      Left            =   2550
      TabIndex        =   11
      Top             =   1470
      Width           =   930
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Count Units"
      Height          =   570
      Left            =   1935
      TabIndex        =   10
      Top             =   2100
      Width           =   930
   End
   Begin Project1.FlipDigit FlipDigit8 
      Height          =   645
      Left            =   1380
      TabIndex        =   9
      Top             =   1410
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   1138
      BackColor       =   8421631
   End
   Begin Project1.FlipDigit FlipDigit7 
      Height          =   645
      Left            =   1890
      TabIndex        =   8
      Top             =   1410
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   1138
      BackColor       =   8454143
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   45
      Top             =   2205
   End
   Begin Project1.FlipDigit FlipDigit1 
      Height          =   645
      Left            =   3030
      TabIndex        =   5
      Top             =   240
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   1138
   End
   Begin Project1.FlipDigit FlipDigit6 
      Height          =   645
      Left            =   255
      TabIndex        =   4
      Top             =   240
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   1138
      MaxNum          =   0
   End
   Begin Project1.FlipDigit FlipDigit5 
      Height          =   645
      Left            =   750
      TabIndex        =   3
      Top             =   240
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   1138
      MaxNum          =   12
   End
   Begin Project1.FlipDigit FlipDigit4 
      Height          =   645
      Left            =   1395
      TabIndex        =   2
      Top             =   240
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   1138
      MaxNum          =   5
   End
   Begin Project1.FlipDigit FlipDigit3 
      Height          =   645
      Left            =   1890
      TabIndex        =   1
      Top             =   240
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   1138
   End
   Begin Project1.FlipDigit FlipDigit2 
      Height          =   645
      Left            =   2535
      TabIndex        =   0
      Top             =   240
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   1138
      MaxNum          =   5
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   120
      Left            =   2400
      Top             =   615
      Width           =   105
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   120
      Left            =   2400
      Top             =   375
      Width           =   105
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   120
      Left            =   1260
      Top             =   615
      Width           =   105
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   120
      Left            =   1260
      Top             =   375
      Width           =   105
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Counter Example"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   1035
      TabIndex        =   7
      Top             =   1125
      Width           =   1860
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Clock Example"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1200
      TabIndex        =   6
      Top             =   -15
      Width           =   1710
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Stay on top
'Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'Const conHwndTopmost = -1
'Const conSwpNoActivate = &H10
'Const conSwpShowWindow = &H40

Dim myTime As String
Dim s1 As Long
Dim s2 As Long
Dim m1 As Long
Dim m2 As Long
Dim hr1 As Long
Dim hr2 As Long

Private Sub Command1_Click()
   FlipDigit7.Enabled = True
   If FlipDigit7.Value = 9 Then FlipDigit8.Enabled = True
End Sub

Private Sub Command2_Click()
   FlipDigit7.Reset = True
   FlipDigit8.Reset = True
End Sub

Private Sub Command3_Click()
FlipDigit8.Enabled = True
End Sub

Private Sub Form_Load()
   'Stay on top
  ' SetWindowPos hwnd, conHwndTopmost, 100, 100, Form1.ScaleWidth, Form1.ScaleHeight + 50, conSwpNoActivate Or conSwpShowWindow
   Timer1.Enabled = True
End Sub

Private Sub Form_Paint()
myTime = Format(Now, "hhmmss a/p")
s1 = Mid$(myTime, 6, 1)
FlipDigit1.StartDigit = s1
s2 = Mid$(myTime, 5, 1)
FlipDigit2.StartDigit = s2
m1 = Mid$(myTime, 4, 1)
FlipDigit3.StartDigit = m1
m2 = Mid$(myTime, 3, 1)
FlipDigit4.StartDigit = m2
hr1 = Mid$(myTime, 2, 1)
FlipDigit5.StartDigit = hr1
hr2 = Left$(myTime, 1)
FlipDigit6.StartDigit = hr2
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Timer1.Enabled = False
   Unload Me
End Sub

Private Sub Timer1_Timer()

FlipDigit1.Enabled = True

If FlipDigit1.Value = 9 Then FlipDigit2.Enabled = True

If FlipDigit1.Value = 9 And FlipDigit2.Value = 5 Then FlipDigit3.Enabled = True
If FlipDigit1.Value = 9 And FlipDigit2.Value = 5 And FlipDigit3.Value = 9 Then FlipDigit4.Enabled = True
If FlipDigit1.Value = 9 And FlipDigit2.Value = 5 And FlipDigit3.Value = 9 And FlipDigit4.Value = 5 Then FlipDigit5.Enabled = True
If FlipDigit1.Value = 9 And FlipDigit2.Value = 5 And FlipDigit3.Value = 9 And FlipDigit4.Value = 5 And FlipDigit5.Value = 12 Then FlipDigit6.Enabled = True
End Sub
