VERSION 5.00
Begin VB.UserControl FlipDigit 
   ClientHeight    =   5760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6885
   ScaleHeight     =   5760
   ScaleWidth      =   6885
   Begin VB.PictureBox pic1TO12 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   135
      Picture         =   "FlipDigit.ctx":0000
      ScaleHeight     =   615
      ScaleWidth      =   21600
      TabIndex        =   4
      Top             =   4935
      Width           =   21630
   End
   Begin VB.PictureBox pic0to5 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   135
      Picture         =   "FlipDigit.ctx":11CF
      ScaleHeight     =   615
      ScaleWidth      =   10800
      TabIndex        =   3
      Top             =   3450
      Width           =   10830
   End
   Begin VB.PictureBox picBlktoOne 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   120
      Picture         =   "FlipDigit.ctx":1BA1
      ScaleHeight     =   615
      ScaleWidth      =   3600
      TabIndex        =   2
      Top             =   2715
      Width           =   3630
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   270
      Top             =   2295
   End
   Begin VB.PictureBox picNum 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   105
      Picture         =   "FlipDigit.ctx":1DF3
      ScaleHeight     =   615
      ScaleWidth      =   18000
      TabIndex        =   1
      Top             =   4185
      Width           =   18030
   End
   Begin VB.PictureBox picDisplay 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   450
      TabIndex        =   0
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "FlipDigit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020

Public Enum eMaxNum
    ZeroOne = 0
    ZeroFive = 5
    ZeroNine = 9
    OneTwelve = 12
 End Enum

Dim counternum As Long
Dim counter1 As Integer

Private Const m_def_Enabled = False
Private Const m_def_MaxNum = 9
Private Const m_def_Speed = 125
Private Const m_def_Value = 0
Private Const m_def_BackColor = &H606060
Private Const m_def_Reset = False
Private Const m_def_StartDigit = 0

Dim m_Enabled As Boolean
Dim m_MaxNum As Long
Dim m_Speed As Long
Dim m_Value As Long
Dim m_BackColor As OLE_COLOR
Dim m_Reset As Boolean
Dim m_StartDigit As Long

Private Sub UserControl_InitProperties()
   Let Enabled = m_def_Enabled
   Let MaxNum = m_def_MaxNum
   Let Speed = m_def_Speed
   Let Value = m_def_Value
   Let BackColor = m_def_BackColor
   Let Reset = m_def_Reset
   Let StartDigit = m_def_StartDigit
End Sub

Private Sub Timer1_Timer()
     If Reset = True Then
        counternum = -1
        If m_MaxNum = 0 Then BitBlt picDisplay.hDC, 0, 0, 30, 41, picBlktoOne.hDC, (1 * 120) + (counter1 * 30), 0, SRCCOPY
        If m_MaxNum = 5 Then BitBlt picDisplay.hDC, 0, 0, 30, 41, pic0to5.hDC, (5 * 120) + (counter1 * 30), 0, SRCCOPY
        If m_MaxNum = 9 Then BitBlt picDisplay.hDC, 0, 0, 30, 41, picNum.hDC, (9 * 120) + (counter1 * 30), 0, SRCCOPY
        If m_MaxNum = 12 Then BitBlt picDisplay.hDC, 0, 0, 30, 41, pic1TO12.hDC, (11 * 120) + (counter1 * 30), 0, SRCCOPY
        If counter1 = 4 Then Reset = False
     End If
     
     If m_MaxNum = 0 Then BitBlt picDisplay.hDC, 0, 0, 30, 41, picBlktoOne.hDC, (counternum * 120) + (counter1 * 30), 0, SRCCOPY
     If m_MaxNum = 5 Then BitBlt picDisplay.hDC, 0, 0, 30, 41, pic0to5.hDC, (counternum * 120) + (counter1 * 30), 0, SRCCOPY
     If m_MaxNum = 9 Then BitBlt picDisplay.hDC, 0, 0, 30, 41, picNum.hDC, (counternum * 120) + (counter1 * 30), 0, SRCCOPY
     If m_MaxNum = 12 Then BitBlt picDisplay.hDC, 0, 0, 30, 41, pic1TO12.hDC, (counternum * 120) + (counter1 * 30), 0, SRCCOPY
   
    counter1 = counter1 + 1

    If counter1 > 4 Then
       counternum = counternum + 1      'advance to next whole number
       Timer1.Enabled = False
       counter1 = 1
    End If

    ' blank to one
     If m_MaxNum = 0 Then
        If counternum * 120 = 240 Then
           counternum = 0
           BitBlt picDisplay.hDC, 0, 0, 30, 41, picBlktoOne.hDC, counternum, 0, SRCCOPY
       End If
       Value = counternum
     End If
     
     'zero to 5
     If m_MaxNum = 5 Then
        If counternum * 120 = 720 Then
           counternum = 0
           BitBlt picDisplay.hDC, 0, 0, 30, 41, pic0to5.hDC, counternum, 0, SRCCOPY
        End If
        Value = counternum
     End If
     
     'zero to nine
     If m_MaxNum = 9 Then
        If counternum * 120 = 1200 Then
           counternum = 0
           BitBlt picDisplay.hDC, 0, 0, 30, 41, picNum.hDC, counternum, 0, SRCCOPY
      End If
      Value = counternum
      If Value = 10 Then Value = 0
   End If
   
   'one to twelve
     If m_MaxNum = 12 Then
        If counternum * 120 = 1440 Then
           counternum = 0
           BitBlt picDisplay.hDC, 0, 0, 30, 41, pic1TO12.hDC, counternum, 0, SRCCOPY
      End If
      Value = counternum + 1
   End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   With PropBag
      'Let Enabled = .ReadProperty("Enabled", m_def_Enabled)
      Let MaxNum = .ReadProperty("MaxNum", m_def_MaxNum)
      Let Speed = .ReadProperty("Speed", m_def_Speed)
      Let BackColor = .ReadProperty("BackColor", m_def_BackColor)
      Let Reset = .ReadProperty("Reset", m_def_Reset)
      Let StartDigit = .ReadProperty("StartDigit", m_def_StartDigit)
   End With
   Timer1.Enabled = False
End Sub

Private Sub UserControl_Resize()
   UserControl.Width = picDisplay.Width
   UserControl.Height = picDisplay.Height
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   With PropBag
      '.WriteProperty "Enabled", m_Enabled, m_def_Enabled
      .WriteProperty "MaxNum", m_MaxNum, m_def_MaxNum
      .WriteProperty "Speed", m_Speed, m_def_Speed
      .WriteProperty "BackColor", m_BackColor, m_def_BackColor
      .WriteProperty "Reset", m_Reset, m_def_Reset
      .WriteProperty "StartDigit", m_StartDigit, m_def_StartDigit
   End With
End Sub

Public Property Get BackColor() As OLE_COLOR
   Let BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal NewBackColor As OLE_COLOR)
   Let m_BackColor = NewBackColor
   picBlktoOne.BackColor = m_BackColor
   pic0to5.BackColor = m_BackColor
   picNum.BackColor = m_BackColor
   pic1TO12.BackColor = m_BackColor
   picDisplay.BackColor = m_BackColor
   PropertyChanged "BackColor"
End Property

'Public Property Get Enabled() As Boolean
 '  Let Enabled = m_Enabled
'End Property

Public Property Let Enabled(ByVal NewEnabled As Boolean)
   Let m_Enabled = NewEnabled
   Timer1.Enabled = m_Enabled
   PropertyChanged "Enabled"
End Property

Public Property Get MaxNum() As eMaxNum
   Let MaxNum = m_MaxNum
End Property

Public Property Let MaxNum(ByVal NewMaxNum As eMaxNum)
   Let m_MaxNum = NewMaxNum
     If m_MaxNum = 0 Then picDisplay.Picture = picBlktoOne.Picture
     If m_MaxNum = 5 Then picDisplay.Picture = pic0to5.Picture
     If m_MaxNum = 9 Then picDisplay.Picture = picNum.Picture
     If m_MaxNum = 12 Then picDisplay.Picture = pic1TO12.Picture
   PropertyChanged "MaxNum"
End Property

Public Property Get StartDigit() As Long
   Let StartDigit = m_StartDigit
End Property

Public Property Let StartDigit(ByVal NewStartDigit As Long)
   Let m_StartDigit = NewStartDigit
   If StartDigit <> 0 Then
      If m_MaxNum <> 12 Then counternum = StartDigit - 1
      If m_MaxNum = 12 Then counternum = StartDigit - 2
      Timer1.Enabled = True
   End If
   PropertyChanged "Reset"
End Property

Public Property Get Reset() As Boolean
   Let Reset = m_Reset
End Property

Public Property Let Reset(ByVal NewReset As Boolean)
   Let m_Reset = NewReset
   Timer1.Enabled = True
   PropertyChanged "Reset"
End Property

Public Property Get Speed() As Long
   Let Speed = m_Speed
End Property

Public Property Let Speed(ByVal NewSpeed As Long)
   Let m_Speed = NewSpeed
      Timer1.Interval = m_Speed
   PropertyChanged "Speed"
End Property

Public Property Get Value() As Long
   Let Value = m_Value
End Property

Public Property Let Value(ByVal NewValue As Long)
   Let m_Value = NewValue
   PropertyChanged "Value"
End Property

Private Sub UserControl_Terminate()
   Timer1.Enabled = False
End Sub

