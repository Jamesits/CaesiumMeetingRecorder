VERSION 5.00
Object = "{BAACC8BE-5CF7-41EE-BE50-E7D125FEF313}#1.0#0"; "APNGViewer.ocx"
Begin VB.Form FrmWizardXP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ceasium Meeting Recorder Alpha"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   5895
   StartUpPosition =   3  '窗口缺省
   Begin APNGViewer.ucAPNG ButtonClock 
      Height          =   960
      Left            =   0
      ToolTipText     =   "计时"
      Top             =   0
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      Data            =   "FrmWizardXP.frx":0000
   End
   Begin APNGViewer.ucAPNG ButtonRec 
      Height          =   960
      Left            =   1560
      ToolTipText     =   "记录"
      Top             =   0
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      Data            =   "FrmWizardXP.frx":0A60
   End
   Begin APNGViewer.ucAPNG ButtonAbout 
      Height          =   960
      Left            =   3240
      ToolTipText     =   "关于"
      Top             =   1560
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      Data            =   "FrmWizardXP.frx":1006
   End
   Begin APNGViewer.ucAPNG ButtonOpenFile 
      Height          =   960
      Left            =   3240
      ToolTipText     =   "打开"
      Top             =   0
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      Data            =   "FrmWizardXP.frx":18F7
   End
   Begin APNGViewer.ucAPNG ButtonClose 
      Height          =   960
      Left            =   4920
      ToolTipText     =   "退出"
      Top             =   1560
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      Data            =   "FrmWizardXP.frx":1F5B
   End
   Begin APNGViewer.ucAPNG ButtonHelp 
      Height          =   960
      Left            =   1560
      ToolTipText     =   "帮助"
      Top             =   1560
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      Data            =   "FrmWizardXP.frx":254B
   End
   Begin APNGViewer.ucAPNG ButtonUpdate 
      Height          =   960
      Left            =   4920
      ToolTipText     =   "更新"
      Top             =   0
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      Data            =   "FrmWizardXP.frx":2C6F
   End
   Begin APNGViewer.ucAPNG ButtonWebsite 
      Height          =   960
      Left            =   0
      ToolTipText     =   "网站"
      Top             =   1560
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      Data            =   "FrmWizardXP.frx":3957
   End
   Begin APNGViewer.ucAPNG ButtonDock 
      Height          =   480
      Left            =   2640
      Top             =   2640
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Data            =   "FrmWizardXP.frx":4652
   End
End
Attribute VB_Name = "FrmWizardXP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ButtonAbout_Click()
Load FrmAbout
End Sub

Private Sub ButtonClock_Click()
Load FrmClockXP
 FrmClockXP.Show
End Sub

Private Sub ButtonClose_Click()
quit
End Sub

Private Sub ButtonDock_Click()
Load FrmDockXP
FrmDockXP.Show
Me.Hide
End Sub

Private Sub ButtonHelp_Click()
MsgBox "该功能在当前版本中不可用。", vbOKOnly + vbDefaultButton1 + vbInformation + vbApplicationModal, "Ceasium Meeting Recorder"
End Sub

Private Sub ButtonUpdate_Click()
MsgBox "该功能在当前版本中不可用。", vbOKOnly + vbDefaultButton1 + vbInformation + vbApplicationModal, "Ceasium Meeting Recorder"
End Sub

Private Sub ButtonWebsite_Click()
Shell "cmd /c explorer https://sourceforge.net/projects/caesiummr/" '360大烧饼误报
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, WM_SYSCOMMAND, SC_MOVE + HTCAPTION, 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
quit
End Sub

