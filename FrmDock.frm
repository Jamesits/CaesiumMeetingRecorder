VERSION 5.00
Object = "{BAACC8BE-5CF7-41EE-BE50-E7D125FEF313}#1.0#0"; "APNGViewer.ocx"
Begin VB.Form FrmDock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "    Ceasium Meeting Recorder - Docking Panel"
   ClientHeight    =   885
   ClientLeft      =   8055
   ClientTop       =   375
   ClientWidth     =   4185
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   885
   ScaleWidth      =   4185
   ShowInTaskbar   =   0   'False
   Begin APNGViewer.ucAPNG ButtonExpand 
      Height          =   480
      Left            =   1800
      Top             =   540
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Data            =   "FrmDock.frx":0000
   End
   Begin APNGViewer.ucAPNG ButtonClose 
      Height          =   960
      Left            =   3360
      ToolTipText     =   "退出"
      Top             =   -120
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      Data            =   "FrmDock.frx":06C6
   End
   Begin APNGViewer.ucAPNG ButtonClock 
      Height          =   960
      Left            =   -60
      ToolTipText     =   "计时"
      Top             =   -120
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      Data            =   "FrmDock.frx":0CB6
   End
   Begin APNGViewer.ucAPNG ButtonRec 
      Height          =   960
      Left            =   1020
      ToolTipText     =   "记录"
      Top             =   -120
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      Data            =   "FrmDock.frx":1716
   End
   Begin APNGViewer.ucAPNG ButtonOpenFile 
      Height          =   960
      Left            =   2220
      ToolTipText     =   "打开"
      Top             =   -120
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
'Me.Caption = " " 'FrmWizard.Caption
'Me.Caption = "Caesium Meeting Recorder"
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
    hBrushOld = SelectObject(Me.hDC, hBrush)
    GetClientRect Me.hwnd, m_Rect
    FillRect Me.hDC, m_Rect, hBrush
    SelectObject Me.hDC, hBrushOld
    DeleteObject hBrush
    End If

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, WM_SYSCOMMAND, SC_MOVE + HTCAPTION, 0
End Sub


