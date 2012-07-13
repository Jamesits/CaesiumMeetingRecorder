VERSION 5.00
Object = "{BAACC8BE-5CF7-41EE-BE50-E7D125FEF313}#1.0#0"; "APNGViewer.ocx"
Begin VB.Form FrmClockXP 
   Caption         =   "计时器 - 已停止"
   ClientHeight    =   2460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   ScaleHeight     =   2460
   ScaleWidth      =   5760
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer TimerRefresh 
      Interval        =   100
      Left            =   4200
      Top             =   2100
   End
   Begin VB.Label LblTime 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   72
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00858585&
      Height          =   1485
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5730
   End
   Begin APNGViewer.ucAPNG ButtonStart 
      Height          =   960
      Left            =   120
      Top             =   1500
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      Data            =   "FrmClockXP.frx":0000
   End
   Begin APNGViewer.ucAPNG ButtonPause 
      Height          =   960
      Left            =   1680
      Top             =   1500
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      Data            =   "FrmClockXP.frx":064A
   End
   Begin APNGViewer.ucAPNG ButtonStop 
      Height          =   960
      Left            =   4680
      Top             =   1500
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      Data            =   "FrmClockXP.frx":0AF1
   End
   Begin APNGViewer.ucAPNG ButtonRestart 
      Height          =   960
      Left            =   3180
      Top             =   1500
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      Data            =   "FrmClockXP.frx":0F5E
   End
End
Attribute VB_Name = "FrmClockXP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Cprovider As New FrmClockProvider

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
LblTime.Caption = FormatTime(Cprovider.Sec)
End Sub


Private Sub ButtonPause_Click()
Cprovider.Pause
Me.Caption = "计时器 - 已暂停"
End Sub

Private Sub ButtonRestart_Click()
Cprovider.Clear
'CProvider.Start
If Cprovider.CEnabled = True Then Me.Caption = "计时器 - 正在运行" Else If Cprovider.Hms > 0 Then Me.Caption = "计时器 - 已暂停" Else Me.Caption = "计时器 - 已停止"
End Sub

Private Sub ButtonStart_Click()
Cprovider.Start
Me.Caption = "计时器 - 正在运行"
End Sub

Private Sub ButtonStop_Click()
Cprovider.Pause
Cprovider.Clear
Me.Caption = "计时器 - 已停止"
End Sub
