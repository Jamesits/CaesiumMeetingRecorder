Attribute VB_Name = "ModWindow"
Option Explicit

Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Const WS_EX_LAYERED = &H80000
Public Const GWL_EXSTYLE = (-20)
Public Const LWA_ALPHA = &H2
Public Const LWA_COLORKEY = &H1
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function DwmIsCompositionEnabled Lib "dwmapi.dll" (ByRef enabledptr As Long) As Long
Public Declare Function DwmExtendFrameIntoClientArea Lib "dwmapi.dll" (ByVal hwnd As Long, margin As MARGINS) As Long
Public Declare Function SetLayeredWindowAttributesByColor Lib "user32" Alias "SetLayeredWindowAttributes" (ByVal hwnd As Long, ByVal crey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Const WM_SYSCOMMAND = &H112
Public Const SC_MOVE = &HF010&
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Public m_transparencyKey
'Public Type RECT
'        Left As Long
'        Top As Long
'        Right As Long
'        Bottom As Long
'End Type

'Public Type MARGINS
'  m_Left As Long
'  m_Right As Long
'  m_Top As Long
'  m_Button As Long
'End Type
'
'
'Public Sub SetWindowGlassOnLoad(FormName As Form)
'm_transparencyKey = RGB(255, 255, 1)
'SetWindowLong FormName.hwnd, GWL_EXSTYLE, GetWindowLong(FormName.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
'SetLayeredWindowAttributesByColor FormName.hwnd, m_transparencyKey, 0, LWA_COLORKEY
'    Dim mg As MARGINS, en As Long
'    mg.m_Left = -1
'    mg.m_Button = -1
'    mg.m_Right = -1
'    mg.m_Top = -1
'    DwmIsCompositionEnabled en
'    If en Then
'      DwmExtendFrameIntoClientArea FormName.hwnd, mg
'    End If
'End Sub
'
'Public Sub SetWindowGlassOnPaint(FormName As Form)
'    Dim hBrush As Long, m_Rect As RECT, hBrushOld As Long
'    hBrush = CreateSolidBrush(RGB(0, 0, 0))
'    hBrushOld = SelectObject(FormName.hdc, hBrush)
'    GetClientRect FormName.hwnd, m_Rect
'    FillRect FormName.hdc, m_Rect, hBrush
'    SelectObject FormName.hdc, hBrushOld
'    DeleteObject hBrush
'End Sub




Public Sub SetWindowAlpha(ByVal hwnd As Long, ByVal Alpha As Byte)
Dim sty As Long
sty = GetWindowLong(hwnd, GWL_EXSTYLE)
sty = sty Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, sty
SetLayeredWindowAttributes hwnd, 0, Alpha, LWA_ALPHA
End Sub

Public Sub SetWindowMaskColor(ByVal hwnd As Long, ByVal Color)
Dim sty As Long
sty = GetWindowLong(hwnd, GWL_EXSTYLE)
sty = sty Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, sty
SetLayeredWindowAttributes hwnd, Color, 0, LWA_COLORKEY
End Sub
