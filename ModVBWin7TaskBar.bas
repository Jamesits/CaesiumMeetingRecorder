Attribute VB_Name = "ModVBWin7TaskBar"
Option Explicit

Public Type ULONGLONG
    Long1 As Long
    Long2 As Long
End Type
Public Enum THUMBBUTTONFLAGS
    THBF_ENABLED = 0
    THBF_DISABLED = &H1
    THBF_DISMISSONCLICK = &H2
    THBF_NOBACKGROUND = &H4
    THBF_HIDDEN = &H8
    THBF_NONINTERACTIVE = &H10
End Enum
Public Enum THUMBBUTTONMASK
    THB_BITMAP = &H1
    THB_ICON = &H2
    THB_TOOLTIP = &H4
    THB_FLAGS = &H8
End Enum
Public Type THUMBBUTTON
    dwMask As Long 'THUMBBUTTONMASK '这两个枚举类型不合适，直接用会造成错误
    iId As Long
    iBitmap As Long
    hIcon As Long
    szTip As String * 260 '导致Byref不能正常传数组的指针，应用Varptr或szTip(519) As Byte，但String比较简单，必须以Chr(0)结尾，否则会出现很多空格—_—b
    dwFlags As Long 'THUMBBUTTONFLAGS
End Type
Public Declare Function Attach Lib "IVBWin7.dll" () As Long
Public Declare Sub Detach Lib "IVBWin7.dll" (ByVal pITaskbarList As Long)

Public Declare Sub AttachMessageTranslate Lib "IVBWin7.dll" (ByVal hwnd As Long, ByVal AutoDetachMessageTranslate As Long)
Public Declare Sub DetachMessageTranslate Lib "IVBWin7.dll" (ByVal hwnd As Long)

Public Declare Sub SetAppID Lib "IVBWin7.dll" (ByVal hwnd As Long, ByVal szAppID As String)
Public Declare Sub SetPreventPinning Lib "IVBWin7.dll" (ByVal hwnd As Long, ByVal Pinning As Long)

Public Declare Function AddTab Lib "IVBWin7.dll" (ByVal pITaskbarList As Long, ByVal hwnd As Long) As Long
Public Declare Function DeleteTab Lib "IVBWin7.dll" (ByVal pITaskbarList As Long, ByVal hwnd As Long) As Long
Public Declare Function ActivateTab Lib "IVBWin7.dll" (ByVal pITaskbarList As Long, ByVal hwnd As Long) As Long
Public Declare Function SetActiveAlt Lib "IVBWin7.dll" (ByVal pITaskbarList As Long, ByVal hwnd As Long) As Long

Public Declare Function MarkFullscreenWindow Lib "IVBWin7.dll" (ByVal pITaskbarList As Long, ByVal hwnd As Long, ByVal fFullscreen As Long) As Long

Public Enum TBPFLAG
    TBPF_NOPROGRESS = 0
    TBPF_INDETERMINATE = &H1
    TBPF_NORMAL = &H2
    TBPF_ERROR = &H4
    TBPF_PAUSED = &H8
End Enum
Public Declare Function SetProgressValue Lib "IVBWin7.dll" (ByVal pITaskbarList As Long, ByVal hwnd As Long, ullCompleted As ULONGLONG, ullTotal As ULONGLONG) As Long
Public Declare Function SetProgressState Lib "IVBWin7.dll" (ByVal pITaskbarList As Long, ByVal hwnd As Long, ByVal tbpFlags As TBPFLAG) As Long
Public Declare Function RegisterTab Lib "IVBWin7.dll" (ByVal pITaskbarList As Long, ByVal hwndTab As Long, ByVal hwndMDI As Long) As Long
Public Declare Function UnregisterTab Lib "IVBWin7.dll" (ByVal pITaskbarList As Long, ByVal hwndTab As Long) As Long
Public Declare Function SetTabOrder Lib "IVBWin7.dll" (ByVal pITaskbarList As Long, ByVal hwndTab As Long, Optional ByVal hwndInsertBefore As Long) As Long
Public Declare Function SetTabActive Lib "IVBWin7.dll" (ByVal pITaskbarList As Long, ByVal hwndTab As Long, ByVal hwndMDI As Long, ByVal dwReserved As Long) As Long
Public Declare Function ThumbBarAddButtons Lib "IVBWin7.dll" (ByVal pITaskbarList As Long, ByVal hwnd As Long, ByVal cButtons As Long, ByVal pButton As Long) As Long
Public Declare Function ThumbBarUpdateButtons Lib "IVBWin7.dll" (ByVal pITaskbarList As Long, ByVal hwnd As Long, ByVal cButtons As Long, ByVal pButton As Long) As Long
Public Declare Function ThumbBarSetImageList Lib "IVBWin7.dll" (ByVal pITaskbarList As Long, ByVal hwnd As Long, ByVal himl As Long) As Long
Public Declare Function SetOverlayIcon Lib "IVBWin7.dll" (ByVal pITaskbarList As Long, ByVal hwnd As Long, Optional ByVal hIcon As Long = 0, Optional ByVal pszDescription As String = "") As Long
Public Declare Function SetThumbnailTooltip Lib "IVBWin7.dll" (ByVal pITaskbarList As Long, ByVal hwnd As Long, ByVal pszTip As String) As Long
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Declare Function SetThumbnailClip Lib "IVBWin7.dll" (ByVal pITaskbarList As Long, ByVal hwnd As Long, prcClip As RECT) As Long

Public Enum STPFLAG
    STPF_NONE = 0
    STPF_USEAPPTHUMBNAILALWAYS = &H1
    STPF_USEAPPTHUMBNAILWHENACTIVE = &H2
    STPF_USEAPPPEEKALWAYS = &H4
    STPF_USEAPPPEEKWHENACTIVE = &H8
End Enum
Public Declare Function SetTabProperties Lib "IVBWin7.dll" (ByVal pITaskbarList As Long, ByVal hwndTab As Long, ByVal stpFlags As STPFLAG) As Long

Public Declare Sub SetButtonElevationRequiredState Lib "IVBWin7.dll" (ByVal hwnd As Long, ByVal fRequired As Long)
Public Declare Function RunAsAdministrator Lib "IVBWin7.dll" (ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Const DT_TOP = &H0
Public Const DT_LEFT = &H0
Public Const DT_CENTER = &H1
Public Const DT_RIGHT = &H2
Public Const DT_VCENTER = &H4
Public Const DT_BOTTOM = &H8
Public Const DT_WORDBREAK = &H10
Public Const DT_SINGLELINE = &H20
Public Const DT_EXPANDTABS = &H40
Public Const DT_TABSTOP = &H80
Public Const DT_NOCLIP = &H100
Public Const DT_EXTERNALLEADING = &H200
Public Const DT_CALCRECT = &H400
Public Const DT_NOPREFIX = &H800
Public Const DT_INTERNAL = &H1000
Public Declare Sub DrawGlowingText Lib "IVBWin7.dll" (ByVal hDC As Long, ByVal szText As String, rcArea As RECT, Optional ByVal dwTextFlags As Long = DT_LEFT Or DT_VCENTER Or DT_SINGLELINE, Optional ByVal iGlowSize As Long = 10)

Public Type JUMPLISTINFO
    IsSeparator As Long
    Path As String '可恶的string……不能byref
    WorkingDirectory As String
    IconLocation As String
    IconIndex As Long
    Arguments As String
    Title As String
End Type
Public Declare Function SetJumpList Lib "IVBWin7.dll" (ByVal pJumpListInfo As Long, ByVal Count As Long, Optional ByVal AppID As String = "") As Long

Public Declare Sub About Lib "IVBWin7.dll" ()
