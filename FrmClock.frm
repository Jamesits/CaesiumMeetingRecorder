VERSION 5.00
Object = "{BAACC8BE-5CF7-41EE-BE50-E7D125FEF313}#1.0#0"; "APNGViewer.ocx"
Begin VB.Form FrmClock 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "计时器 - 已停止"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5955
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   5955
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer TimerRefresh 
      Interval        =   100
      Left            =   4320
      Top             =   1800
   End
   Begin APNGViewer.ucAPNG ButtonRestart 
      Height          =   960
      Left            =   3300
      Top             =   1200
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      Data            =   "FrmClock.frx":0000
   End
   Begin APNGViewer.ucAPNG ButtonStop 
      Height          =   960
      Left            =   4800
      Top             =   1200
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      Data            =   "FrmClock.frx":093C
   End
   Begin APNGViewer.ucAPNG ButtonPause 
      Height          =   960
      Left            =   1800
      Top             =   1200
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      Data            =   "FrmClock.frx":0DA9
   End
   Begin APNGViewer.ucAPNG ButtonStart 
      Height          =   960
      Left            =   240
      Top             =   1200
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      Data            =   "FrmClock.frx":1250
   End
   Begin VB.Label LblTime 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   72
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00858585&
      Height          =   1485
      Left            =   120
      TabIndex        =   0
      Top             =   -300
      Width           =   5730
   End
End
Attribute VB_Name = "FrmClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function DwmIsCompositionEnabled Lib "dwmapi.dll" (ByRef enabledptr As Long) As Long
Private Declare Function DwmExtendFrameIntoClientArea Lib "dwmapi.dll" (ByVal hwnd As Long, margin As MARGINS) As Long
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Dim en
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type MARGINS
  m_Left As Long
  m_Right As Long
  m_Top As Long
  m_Button As Long
End Type
Private CProvider As FrmClockProvider

Private Sub ButtonPause_Click()
CProvider.Pause
Me.Caption = "计时器 - 已暂停"
End Sub

Private Sub ButtonRestart_Click()
CProvider.Clear
'CProvider.Start
If CProvider.CEnabled = True Then Me.Caption = "计时器 - 正在运行" Else If CProvider.Hms > 0 Then Me.Caption = "计时器 - 已暂停" Else Me.Caption = "计时器 - 已停止"
End Sub

Private Sub ButtonStart_Click()
CProvider.Start
Me.Caption = "计时器 - 正在运行"
End Sub

Private Sub ButtonStop_Click()
CProvider.Pause
CProvider.Clear
Me.Caption = "计时器 - 已停止"
End Sub

Private Sub Form_Load()
'Dim tempStyle As Long
'tempStyle = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
'SetWindowLong Me.hwnd, GWL_EXSTYLE, tempStyle Or WS_EX_LAYERED
'SetLayeredWindowAttributesByColor Me.hwnd, m_transparencyKey, 0, LWA_COLORKEY
    Dim mg As MARGINS, en As Long
    mg.m_Left = -1
    mg.m_Button = -1
    mg.m_Right = -1
    mg.m_Top = -1
    DwmIsCompositionEnabled en
    If en Then
      DwmExtendFrameIntoClientArea Me.hwnd, mg
    Else
      LblTime.ForeColor = RGB(127, 127, 127)
    End If
Set CProvider = New FrmClockProvider
Load CProvider
'LblTime.ForeColor = RGB(127, 127, 127)
End Sub


Private Sub Form_Paint()
    DwmIsCompositionEnabled en
    If en Then
    Dim hBrush As Long, m_Rect As RECT, hBrushOld As Long
    hBrush = CreateSolidBrush(m_transparencyKey)
    hBrushOld = SelectObject(Me.hDC, hBrush)
    GetClientRect Me.hwnd, m_Rect
    FillRect Me.hDC, m_Rect, hBrush
    SelectObject Me.hDC, hBrushOld
    DeleteObject hBrush
    Else
        LblTime.ForeColor = vbBlack
End If

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, WM_SYSCOMMAND, SC_MOVE + HTCAPTION, 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
Me.Hide
End Sub

Private Sub LblTime_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, WM_SYSCOMMAND, SC_MOVE + HTCAPTION, 0
End Sub

Private Sub TimerRefresh_Timer()
LblTime.Caption = FormatTime(CProvider.Sec)
End Sub
