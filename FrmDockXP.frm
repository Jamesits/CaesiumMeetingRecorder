VERSION 5.00
Object = "{BAACC8BE-5CF7-41EE-BE50-E7D125FEF313}#1.0#0"; "APNGViewer.ocx"
Begin VB.Form FrmDockXP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ceasium Meeting Recorder - Docking Panel"
   ClientHeight    =   990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   4380
   StartUpPosition =   3  '窗口缺省
   Begin APNGViewer.ucAPNG ButtonOpenFile 
      Height          =   960
      Left            =   2280
      ToolTipText     =   "打开"
      Top             =   0
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      Data            =   "FrmDockXP.frx":0000
   End
   Begin APNGViewer.ucAPNG ButtonRec 
      Height          =   960
      Left            =   1080
      ToolTipText     =   "记录"
      Top             =   0
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      Data            =   "FrmDockXP.frx":0664
   End
   Begin APNGViewer.ucAPNG ButtonClock 
      Height          =   960
      Left            =   0
      ToolTipText     =   "计时"
      Top             =   0
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      Data            =   "FrmDockXP.frx":0C0A
   End
   Begin APNGViewer.ucAPNG ButtonClose 
      Height          =   960
      Left            =   3420
      ToolTipText     =   "退出"
      Top             =   0
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      Data            =   "FrmDockXP.frx":166A
   End
   Begin APNGViewer.ucAPNG ButtonExpand 
      Height          =   480
      Left            =   1800
      Top             =   660
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Data            =   "FrmDockXP.frx":1C5A
   End
End
Attribute VB_Name = "FrmDockXP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ButtonClock_Click()
Load FrmClockXP
 FrmClockXP.Show
End Sub

Private Sub ButtonClose_Click()
quit
End Sub

Private Sub ButtonExpand_Click()
FrmWizardXP.Show
Me.Hide
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = Int((Screen.Width - Me.Width) / 2)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, WM_SYSCOMMAND, SC_MOVE + HTCAPTION, 0
End Sub

