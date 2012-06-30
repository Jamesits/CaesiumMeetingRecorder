VERSION 5.00
Object = "{BAACC8BE-5CF7-41EE-BE50-E7D125FEF313}#1.0#0"; "APNGViewer.ocx"
Begin VB.Form FrmWizard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ceasium Meeting Recorder Alpha"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6795
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   6795
   StartUpPosition =   2  '屏幕中心
   Begin APNGViewer.ucAPNG ButtonWebsite 
      Height          =   960
      Left            =   480
      ToolTipText     =   "网站"
      Top             =   1920
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      Data            =   "FrmWizard.frx":0000
   End
   Begin APNGViewer.ucAPNG ButtonUpdate 
      Height          =   960
      Left            =   5400
      ToolTipText     =   "更新"
      Top             =   360
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      Data            =   "FrmWizard.frx":0CFB
   End
   Begin APNGViewer.ucAPNG ButtonHelp 
      Height          =   960
      Left            =   2040
      ToolTipText     =   "帮助"
      Top             =   1920
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      Data            =   "FrmWizard.frx":19E3
   End
   Begin APNGViewer.ucAPNG ButtonClose 
      Height          =   960
      Left            =   5400
      ToolTipText     =   "退出"
      Top             =   1920
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      Data            =   "FrmWizard.frx":2107
   End
   Begin APNGViewer.ucAPNG ButtonOpenFile 
      Height          =   960
      Left            =   3720
      ToolTipText     =   "打开"
      Top             =   360
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      Data            =   "FrmWizard.frx":26F7
   End
   Begin APNGViewer.ucAPNG ButtonAbout 
      Height          =   960
      Left            =   3720
      ToolTipText     =   "关于"
      Top             =   1920
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      Data            =   "FrmWizard.frx":2D5B
   End
   Begin APNGViewer.ucAPNG ButtonRec 
      Height          =   960
      Left            =   2040
      ToolTipText     =   "记录"
      Top             =   360
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      Data            =   "FrmWizard.frx":364C
   End
   Begin APNGViewer.ucAPNG ButtonClock 
      Height          =   960
      Left            =   480
      ToolTipText     =   "计时"
      Top             =   360
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      Data            =   "FrmWizard.frx":3BF2
   End
End
Attribute VB_Name = "FrmWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function DwmIsCompositionEnabled Lib "dwmapi.dll" (ByRef enabledptr As Long) As Long
Private Declare Function DwmExtendFrameIntoClientArea Lib "dwmapi.dll" (ByVal hwnd As Long, margin As MARGINS) As Long
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

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

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const WM_SYSCOMMAND = &H112
Private Const SC_MOVE = &HF010&
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private Sub ButtonAbout_Click()
Load FrmAbout
End Sub

Private Sub ButtonClock_Click()
Load FrmClock
 FrmClock.Show
End Sub

Private Sub ButtonClose_Click()
End
End Sub

Private Sub ButtonHelp_Click()
MsgBox "该功能在当前版本中不可用。", vbOKOnly + vbDefaultButton1 + vbInformation + vbApplicationModal, "Ceasium Meeting Recorder"
End Sub

Private Sub ButtonUpdate_Click()
MsgBox "该功能在当前版本中不可用。", vbOKOnly + vbDefaultButton1 + vbInformation + vbApplicationModal, "Ceasium Meeting Recorder"
End Sub

Private Sub ButtonWebsite_Click()
Shell "cmd /c explorer https://sourceforge.net/projects/caesiummr/"
End Sub

Private Sub Form_Load()
Me.Caption = "Ceasium Meeting Recorder " & Versionstring
    Dim mg As MARGINS, en As Long
    mg.m_Left = -1
    mg.m_Button = -1
    mg.m_Right = -1
    mg.m_Top = -1
    DwmIsCompositionEnabled en
    If en Then
      DwmExtendFrameIntoClientArea Me.hwnd, mg
    End If
   ' Me.Show
End Sub

Private Sub Form_Paint()
    Dim hBrush As Long, m_Rect As RECT, hBrushOld As Long
    hBrush = CreateSolidBrush(RGB(0, 0, 0))
    hBrushOld = SelectObject(Me.hdc, hBrush)
    GetClientRect Me.hwnd, m_Rect
    FillRect Me.hdc, m_Rect, hBrush
    SelectObject Me.hdc, hBrushOld
    DeleteObject hBrush
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, WM_SYSCOMMAND, SC_MOVE + HTCAPTION, 0
End Sub


