VERSION 5.00
Begin VB.Form FrmClock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "¼ÆÊ±Æ÷"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5955
   BeginProperty Font 
      Name            =   "Î¢ÈíÑÅºÚ"
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
   ScaleHeight     =   3030
   ScaleWidth      =   5955
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   72
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1860
      Left            =   120
      TabIndex        =   0
      Top             =   360
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

Private Sub Form_Load()
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
    Else
        Label1.ForeColor = vbBlack
End If

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, WM_SYSCOMMAND, SC_MOVE + HTCAPTION, 0
End Sub

