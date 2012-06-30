VERSION 5.00
Object = "{BAACC8BE-5CF7-41EE-BE50-E7D125FEF313}#1.0#0"; "APNGViewer.ocx"
Begin VB.Form FrmDock 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1470
   ClientLeft      =   8010
   ClientTop       =   0
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   ScaleHeight     =   1470
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   Begin APNGViewer.ucAPNG ButtonExpand 
      Height          =   480
      Left            =   2640
      Top             =   1080
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Data            =   "FrmDock.frx":0000
   End
   Begin APNGViewer.ucAPNG ButtonClose 
      Height          =   960
      Left            =   4800
      ToolTipText     =   "退出"
      Top             =   120
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      Data            =   "FrmDock.frx":06C6
   End
   Begin APNGViewer.ucAPNG ButtonClock 
      Height          =   960
      Left            =   0
      ToolTipText     =   "计时"
      Top             =   120
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      Data            =   "FrmDock.frx":0CB6
   End
   Begin APNGViewer.ucAPNG ButtonRec 
      Height          =   960
      Left            =   1560
      ToolTipText     =   "记录"
      Top             =   120
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      Data            =   "FrmDock.frx":1716
   End
   Begin APNGViewer.ucAPNG ButtonOpenFile 
      Height          =   960
      Left            =   3240
      ToolTipText     =   "打开"
      Top             =   120
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      Data            =   "FrmDock.frx":1CBC
   End
End
Attribute VB_Name = "FrmDock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function DwmIsCompositionEnabled Lib "dwmapi.dll" (ByRef enabledptr As Long) As Long
Private Declare Function DwmExtendFrameIntoClientArea Lib "dwmapi.dll" (ByVal hwnd As Long, margin As MARGINS) As Long
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
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

Private Sub ButtonClock_Click()
Load FrmClock
 FrmClock.Show
End Sub

Private Sub ButtonClose_Click()
quit
End Sub

Private Sub ButtonExpand_Click()
FrmWizard.Show
Me.Hide
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = Int((Screen.Width - Me.Width) / 2)
    Dim mg As MARGINS, en As Long
    mg.m_Left = -1
    mg.m_Button = -1
    mg.m_Right = -1
    mg.m_Top = -1
    DwmIsCompositionEnabled en
    If en Then
      DwmExtendFrameIntoClientArea Me.hwnd, mg
    End If

End Sub

Private Sub Form_Paint()
    DwmIsCompositionEnabled en
    If en Then
    Dim hBrush As Long, m_Rect As RECT, hBrushOld As Long
    hBrush = CreateSolidBrush(RGB(0, 0, 0))
    hBrushOld = SelectObject(Me.hdc, hBrush)
    GetClientRect Me.hwnd, m_Rect
    FillRect Me.hdc, m_Rect, hBrush
    SelectObject Me.hdc, hBrushOld
    DeleteObject hBrush
    End If

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, WM_SYSCOMMAND, SC_MOVE + HTCAPTION, 0
End Sub


