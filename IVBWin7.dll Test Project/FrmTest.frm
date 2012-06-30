VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmTest 
   AutoRedraw      =   -1  'True
   BackColor       =   &H000000FF&
   Caption         =   "系统默认标题Test"
   ClientHeight    =   3435
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   Picture         =   "FrmTest.frx":0000
   ScaleHeight     =   229
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   385
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command9 
      Caption         =   "删除JumpList"
      Height          =   495
      Left            =   2400
      TabIndex        =   11
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CommandButton Command8 
      Caption         =   "建立JumpList"
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CommandButton Command7 
      Caption         =   "以管理员权限运行程序"
      Height          =   495
      Left            =   2400
      TabIndex        =   9
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton Command6 
      Caption         =   "按钮盾牌图标"
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   2040
      Width           =   2055
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTest.frx":277D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTest.frx":2AD1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTest.frx":2E25
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTest.frx":3179
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTest.frx":34CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTest.frx":3821
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      LargeChange     =   5
      Left            =   120
      Max             =   20
      Min             =   1
      TabIndex        =   7
      Top             =   360
      Value           =   7
      Width           =   5535
   End
   Begin VB.CommandButton Command5 
      Caption         =   "任务栏可被挡住"
      Height          =   495
      Left            =   4560
      TabIndex        =   6
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "裁切缩略图"
      Height          =   495
      Left            =   3480
      TabIndex        =   5
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "更改提示"
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "小按钮"
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "小图标"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "FrmTest.frx":3B75
      Left            =   120
      List            =   "FrmTest.frx":3B77
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   720
      Width           =   5535
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   10
      Left            =   120
      Max             =   100
      TabIndex        =   0
      Top             =   1080
      Width           =   5535
   End
End
Attribute VB_Name = "FrmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim h As Long
Dim OverlayIconCount As Boolean
Dim OnTop As Boolean
Dim ButtonElevationRequiredState As Boolean

Private Sub Combo1_Click()
    Select Case Combo1.ListIndex
    Case 0
        SetProgressState h, hwnd, TBPF_NOPROGRESS
    Case 1
        SetProgressState h, hwnd, TBPF_INDETERMINATE
    Case 2
        SetProgressState h, hwnd, TBPF_NORMAL
    Case 3
        SetProgressState h, hwnd, TBPF_ERROR
    Case 4
        SetProgressState h, hwnd, TBPF_PAUSED
    End Select
End Sub

Private Sub Command1_Click()
    OverlayIconCount = Not OverlayIconCount
    SetOverlayIcon h, hwnd, Icon.Handle * -OverlayIconCount
End Sub

Private Sub Command2_Click() '这些按钮很奇怪，弹出Msgbox后竟然不会卡住程序……
    Dim t(6) As THUMBBUTTON
    With t(0)
        .dwFlags = THBF_ENABLED '有背景，可按下，按下后预览窗口不消失
        .dwMask = THB_TOOLTIP Or THB_FLAGS Or THB_BITMAP
        .iId = 1
        .iBitmap = 0
        .szTip = "bt1按钮，有背景，可按下，按下后预览窗口不消失" & Chr(0)
    End With
    With t(1)
        .dwFlags = THBF_DISABLED '有背景，不可按下，灰色
        .dwMask = THB_TOOLTIP Or THB_FLAGS Or THB_BITMAP
        .iId = 2
        .iBitmap = 1
        .szTip = "bt2，有背景，不可按下，灰色" & Chr(0)
    End With
    With t(2)
        .dwFlags = THBF_DISMISSONCLICK '有背景，可按下，按下后预览窗口消失
        .dwMask = THB_TOOLTIP Or THB_FLAGS Or THB_BITMAP
        .iId = 3
        .iBitmap = 2
        .szTip = "bt3，有背景，可按下，按下后预览窗口消失" & Chr(0)
    End With
    With t(3)
        .dwFlags = THBF_NOBACKGROUND '无背景，可按下，按下后预览窗口不消失
        .dwMask = THB_TOOLTIP Or THB_FLAGS Or THB_BITMAP
        .iId = 4
        .iBitmap = 3
        .szTip = "bt4，无背景，可按下，按下后预览窗口不消失" & Chr(0)
    End With
    With t(4)
        .dwFlags = THBF_NONINTERACTIVE '有背景，不可按下
        .dwMask = THB_TOOLTIP Or THB_FLAGS Or THB_BITMAP
        .iId = 5
        .iBitmap = 4
        .szTip = "bt5，有背景，不可按下" & Chr(0)
    End With
    With t(5)
        .dwFlags = THBF_DISMISSONCLICK Or THBF_NOBACKGROUND '无背景，可按下，按下后预览窗口消失'''''''其他的自己组合吧：）
        .dwMask = THB_TOOLTIP Or THB_FLAGS Or THB_BITMAP
        .iId = 6
        .iBitmap = 5
        .szTip = "bt6，无背景，可按下，按下后预览窗口消失" & Chr(0)
    End With
    With t(6)
        .dwFlags = THBF_ENABLED
        .dwMask = THB_ICON Or THB_TOOLTIP Or THB_FLAGS
        .hIcon = FrmTest.Icon.Handle
        .iId = 7
        .szTip = "bt7，使用图标" & Chr(0)
    End With
    ThumbBarSetImageList h, hwnd, ImageList1.hImageList
    ThumbBarAddButtons h, hwnd, 7, VarPtr(t(0)) '最多7个，否则系统很生气，后果很严重：）
End Sub

Private Sub Command3_Click()
    SetThumbnailTooltip h, hwnd, StrConv(InputBox("输入新提示", , "Tip"), vbUnicode)
End Sub

Private Sub Command4_Click()
    Dim ClipRECT As RECT
    Do
        ClipRECT.Left = ScaleWidth * Rnd
        ClipRECT.Right = ScaleWidth * Rnd
    Loop Until ClipRECT.Left < ClipRECT.Right - 100
    Do
        ClipRECT.Top = ScaleHeight * Rnd
        ClipRECT.Bottom = ScaleHeight * Rnd
    Loop Until ClipRECT.Top < ClipRECT.Bottom - 50
    SetThumbnailClip h, hwnd, ClipRECT
End Sub

Private Sub Command5_Click()
    OnTop = Not OnTop
    MarkFullscreenWindow h, hwnd, OnTop '注意按钮标题的主语是“任务栏”，此设置会影响其他窗口……
End Sub

Private Sub Command6_Click()
    ButtonElevationRequiredState = Not ButtonElevationRequiredState
    SetButtonElevationRequiredState Command6.hwnd, ButtonElevationRequiredState '只针对有Win7风格的按钮有效(manifest中指定Win7风格)
End Sub

Private Sub Command7_Click()
    RunAsAdministrator Replace(App.Path, "\", "\\") & "\\EXE2.exe", "", "", 1
End Sub

Private Sub Command8_Click()
    Dim JLInfo(4) As JUMPLISTINFO
    With JLInfo(0)
        .WorkingDirectory = App.Path
        .Path = App.Path & "\" & App.EXEName & ".exe"
        .IconIndex = 0
        .Arguments = "a"
        .Title = "About"
        .IconLocation = .Path
    End With
    With JLInfo(1)
        .WorkingDirectory = App.Path
        .Path = App.Path & "\" & App.EXEName & ".exe"
        .IconIndex = 0
        .Arguments = "b"
        .Title = "b"
        .IconLocation = .Path
    End With
    With JLInfo(2)
        .IsSeparator = True
        .WorkingDirectory = App.Path
        .Path = "notepad.exe"
        .IconIndex = 0
        .Title = "记事本"
        .IconLocation = .Path
    End With
    With JLInfo(3)
        .WorkingDirectory = App.Path
        .Path = "notepad.exe"
        .IconIndex = 0
        .Title = "记事本"
        .IconLocation = .Path
    End With
    SetJumpList VarPtr(JLInfo(0)), 3, "dft1"
End Sub

Private Sub Command9_Click()
    SetJumpList 0, -1, "dft1"
End Sub

Private Sub Form_Load() 'ATTENTION！Load中任务栏按钮尚未创建，对任务栏按钮操作无效：（，要等到收到WM_TASKBARBUTTONCREATED才可以操作
    Select Case VBA.Command
    Case "a" '互相通讯自己解决吧……
        About
        End
    Case "b"
        MsgBox "b"
        End
    Case Else
        h = Attach()
        AttachMessageTranslate hwnd, False
        Combo1.AddItem "无进度"
        Combo1.AddItem "滚动"
        Combo1.AddItem "普通"
        Combo1.AddItem "出错"
        Combo1.AddItem "暂停"
        Combo1.ListIndex = 2
        Form2.Show
        SetPreventPinning hwnd, True '必须在SetAppID之前，否则JumpList会错乱
        SetAppID hwnd, "dft1"
        SetAppID Form2.hwnd, "df2"
    End Select
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MsgBox Y
    If Y = -4096 Then MsgBox "bt" & X & "按下", vbInformation
    If Y = -4093 Then MsgBox "WM_TASKBARBUTTONCREATED", vbInformation
End Sub

Private Sub Form_Resize()
    Dim r As RECT
    r.Bottom = ScaleHeight
    r.Right = ScaleWidth
    Cls
    DrawGlowingText hDC, StrConv("发光标题Test", vbUnicode), r, DT_CENTER Or DT_TOP Or DT_SINGLELINE, HScroll2.Value
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Form2
    DetachMessageTranslate hwnd
    Detach h
End Sub

Private Sub HScroll1_Change()
    Dim l1 As ULONGLONG, l2 As ULONGLONG
    l1.Long1 = HScroll1.Value
    l2.Long1 = 100
    SetProgressValue h, hwnd, l1, l2
End Sub

Private Sub HScroll1_Scroll()
    HScroll1_Change
End Sub

Private Sub HScroll2_Change()
    Form_Resize
End Sub

Private Sub HScroll2_Scroll()
    HScroll2_Change
End Sub
