Attribute VB_Name = "ModVBWin7DWMApi"
' Blur behind data structures
Public Const DWM_BB_ENABLE = &H1                        ' fEnable has been specified
Public Const DWM_BB_BLURREGION = &H2                    ' hRgnBlur has been specified
Public Const DWM_BB_TRANSITIONONMAXIMIZED = &H4         ' fTransitionOnMaximized has been specified

Public Type DWM_BLURBEHIND
    dwFlags As Long
    fEnable As Boolean
    hRgnBlur As Long
    fTransitionOnMaximized As Boolean
End Type

' Window attributes
 Public Enum DWMWINDOWATTRIBUTE
    DWMWA_NCRENDERING_ENABLED = 1      ' [get] Is non-client rendering enabled/disabled
    DWMWA_NCRENDERING_POLICY           ' [set] Non-client rendering policy
    DWMWA_TRANSITIONS_FORCEDISABLED    ' [set] Potentially enable/forcibly disable transitions
    DWMWA_ALLOW_NCPAINT                ' [set] Allow contents rendered in the non-client area to be visible on the DWM-drawn frame.
    DWMWA_CAPTION_BUTTON_BOUNDS        ' [get] Bounds of the caption button area in window-relative space.
    DWMWA_NONCLIENT_RTL_LAYOUT         ' [set] Is non-client content RTL mirrored
    DWMWA_FORCE_ICONIC_REPRESENTATION  ' [set] Force this window to display iconic thumbnails.
    DWMWA_FLIP3D_POLICY                ' [set] Designates how Flip3D will treat the window.
    DWMWA_EXTENDED_FRAME_BOUNDS        ' [get] Gets the extended frame bounds rectangle in screen space
    DWMWA_HAS_ICONIC_BITMAP            ' [set] Indicates an available bitmap when there is no better thumbnail representation.
    DWMWA_DISALLOW_PEEK                ' [set] Don't invoke Peek on the window.
    DWMWA_EXCLUDED_FROM_PEEK           ' [set] LivePreview exclusion information
    DWMWA_LAST
End Enum

' Non-client rendering policy attribute values
 Public Enum DWMNCRENDERINGPOLICY
    DWMNCRP_USEWINDOWSTYLE ' Enable/disable non-client rendering based on window style
    DWMNCRP_DISABLED       ' Disabled non-client rendering; window style is ignored
    DWMNCRP_ENABLED        ' Enabled non-client rendering; window style is ignored
    DWMNCRP_LAST
End Enum

' Values designating how Flip3D treats a given window.
Public Enum DWMFLIP3DWINDOWPOLICY
    DWMFLIP3D_DEFAULT      ' Hide or include the window in Flip3D based on window style and visibility.
    DWMFLIP3D_EXCLUDEBELOW ' Display the window under Flip3D and disabled.
    DWMFLIP3D_EXCLUDEABOVE ' Display the window above Flip3D and enabled.
    DWMFLIP3D_LAST
End Enum


' Thumbnails
    'typedef HANDLE HTHUMBNAIL
    'typedef HTHUMBNAIL* PHTHUMBNAIL '*
    
    Public Const DWM_TNP_RECTDESTINATION = &H1
    Public Const DWM_TNP_RECTSOURCE = &H2
    Public Const DWM_TNP_OPACITY = &H4
    Public Const DWM_TNP_VISIBLE = &H8
    Public Const DWM_TNP_SOURCECLIENTAREAONLY = &H10
    
'Public Type RECT
'        Left As Long
'        Top As Long
'        Right As Long
'        Bottom As Long
'End Type

Public Type DWM_THUMBNAIL_PROPERTIES
    dwFlags As Long
    rcDestination As RECT
    rcSource As RECT
    opacity As Byte
    fVisible As Boolean
    fSourceClientAreaOnly As Boolean
End Type

' Video enabling apis

Public Type DWM_FRAME_COUNT
    Long1 As Long
    Long2 As Long
End Type

Public Type QPC_TIME
    Long1 As Long
    Long2 As Long
End Type

'Public Type ULONGLONG
'    Long1 As Long
'    Long2 As Long
'End Type
    
Public Type UNSIGNED_RATIO
    uiNumerator As Long
    uiDenominator As Long
End Type

Public Type DWM_TIMING_INFO
    cbSize As Long
    
    ' Data on DWM composition overall
    ' Monitor refresh rate
    rateRefresh As UNSIGNED_RATIO
    
    ' Actual period
    qpcRefreshPeriod As QPC_TIME
    
    ' composition rate
    rateCompose As UNSIGNED_RATIO
    
    ' QPC time at a VSync interupt
    qpcVBlank As QPC_TIME
    
    ' DWM refresh count of the last vsync
    ' DWM refresh count is a 64bit number where zero is
    ' the first refresh the DWM woke up to process
    cRefresh As DWM_FRAME_COUNT
    
    ' DX refresh count at the last Vsync Interupt
    ' DX refresh count is a 32bit number with zero
    ' being the first refresh after the card was initialized
    ' DX increments a counter when ever a VSync ISR is processed
    ' It is possible for DX to miss VSyncs
    '
    ' There is not a fixed mapping between DX and DWM refresh counts
    ' because the DX will rollover and may miss VSync interupts
    cDXRefresh As Integer
    
    ' QPC time at a compose time.
    qpcCompose As QPC_TIME
    
    ' Frame number that was composed at qpcCompose
    cFrame As DWM_FRAME_COUNT
    
    ' The present number DX uses to identify renderer frames
    cDXPresent As Integer
    
    ' Refresh count of the frame that was composed at qpcCompose
    cRefreshFrame As DWM_FRAME_COUNT
    
    
    ' DWM frame number that was last submitted
    cFrameSubmitted As DWM_FRAME_COUNT
    
    ' DX Present number that was last submitted
    cDXPresentSubmitted As Integer
    
    ' DWM frame number that was last confirmed presented
    cFrameConfirmed As DWM_FRAME_COUNT
    
    ' DX Present number that was last confirmed presented
    cDXPresentConfirmed As Integer
    
    ' The target refresh count of the last
    ' frame confirmed completed by the GPU
    cRefreshConfirmed As DWM_FRAME_COUNT
    
    ' DX refresh count when the frame was confirmed presented
    cDXRefreshConfirmed As Integer
    
    ' Number of frames the DWM presented late
    ' AKA Glitches
    cFramesLate As DWM_FRAME_COUNT
    ' the number of composition frames that
    ' have been issued but not confirmed completed
    cFramesOutstanding As Integer
    
    
    ' Following fields are only relavent when anis specified
    ' Display frame
    
    
    ' Last frame displayed
    cFrameDisplayed As DWM_FRAME_COUNT
    
    ' QPC time of the composition pass when the frame was displayed
    qpcFrameDisplayed As QPC_TIME
    
    ' Count of the VSync when the frame should have become visible
    cRefreshFrameDisplayed As DWM_FRAME_COUNT
    
    ' Complete frames: DX has notified the DWM that the frame is done rendering
    
    ' ID of the the last frame marked complete (starts at 0)
    cFrameComplete As DWM_FRAME_COUNT
    
    ' QPC time when the last frame was marked complete
    qpcFrameComplete As QPC_TIME
    
    ' Pending frames:
    ' The application has been submitted to DX but not completed by the GPU
    ' ID of the the last frame marked pending (starts at 0)
    cFramePending As DWM_FRAME_COUNT
    
    ' QPC time when the last frame was marked pending
    qpcFramePending As QPC_TIME
    
    ' number of unique frames displayed
    cFramesDisplayed As DWM_FRAME_COUNT
    
    ' number of new completed frames that have been received
    cFramesComplete As DWM_FRAME_COUNT
    
    ' number of new frames submitted to DX but not yet complete
    cFramesPending As DWM_FRAME_COUNT
    
    ' number of frames available but not displayed, used or dropped
    cFramesAvailable As DWM_FRAME_COUNT
    
    ' number of rendered frames that were never
    ' displayed because composition occured too late
    cFramesDropped As DWM_FRAME_COUNT
    ' number of times an old frame was composed
    ' when a new frame should have been used
    ' but was not available
    cFramesMissed As DWM_FRAME_COUNT
    ' the refresh at which the next frame is
    ' scheduled to be displayed
    cRefreshNextDisplayed As DWM_FRAME_COUNT
    
    ' the refresh at which the next DX present is
    ' scheduled to be displayed
    cRefreshNextPresented As DWM_FRAME_COUNT
    
    ' The total number of refreshes worth of content
    ' for thisthat have been displayed by the DWM
    ' since DwmSetPresentParameters was called
    cRefreshesDisplayed As DWM_FRAME_COUNT
    ' The total number of refreshes worth of content
    ' that have been presented by the application
    ' since DwmSetPresentParameters was called
    cRefreshesPresented As DWM_FRAME_COUNT
    
    
    ' The actual refresh # when content for this
    ' window started to be displayed
    ' it may be different than that requested
    ' DwmSetPresentParameters
    cRefreshStarted As DWM_FRAME_COUNT
    
    ' Total number of pixels DX redirected
    ' to the DWM.
    ' If Queueing is used the full buffer
    ' is transfered on each present.
    ' If not queuing it is possible only
    ' a dirty region is updated
    cPixelsReceived As ULONGLONG
    
    ' Total number of pixels drawn.
    ' Does not take into account if
    ' if the window is only partial drawn
    ' do to clipping or dirty rect management
    cPixelsDrawn As ULONGLONG
    
    ' The number of buffers in the flipchain
    ' that are empty.   An application can
    ' present that number of times and guarantee
    ' it won't be blocked waiting for a buffer to
    ' become empty to present to
    cBuffersEmpty As DWM_FRAME_COUNT
    
End Type

Public Enum DWM_SOURCE_FRAME_SAMPLING
    ' Use the first source frame that
    ' includes the first refresh of the output frame
    DWM_SOURCE_FRAME_SAMPLING_POINT
    ' use the source frame that includes the most
    ' refreshes of out the output frame
    ' in case of multiple source frames with the
    ' same coverage the last will be used
    DWM_SOURCE_FRAME_SAMPLING_COVERAGE
    ' Sentinel value
    DWM_SOURCE_FRAME_SAMPLING_LAST
End Enum

Public Const c_DwmMaxQueuedBuffers As Long = 8
Public Const c_DwmMaxMonitors As Long = 16
Public Const c_DwmMaxAdapters As Long = 1

Public Type DWM_PRESENT_PARAMETERS
    cbSize As Long
    fQueue As Boolean
    cRefreshStart As DWM_FRAME_COUNT
    cBuffer As Integer
    fUseSourceRate As Boolean
    rateSource As UNSIGNED_RATIO
    cRefreshesPerFrame As Integer
    eSampling As DWM_SOURCE_FRAME_SAMPLING
End Type

Public Type MARGINS
    cxLeftWidth As Long 'width of left border that retains its size
    cxRightWidth As Long 'width of right border that retains its size
    cyTopHeight As Long 'height of top border that retains its size
    cyBottomHeight As Long 'height of bottom border that retains its size
End Type

Public Const DWM_FRAME_DURATION_DEFAULT = -1

'Public Declare Function DwmDefWindowProc Lib "dwmapi.dll" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long, plResult As LRESULT) As Boolean

Public Declare Function DwmEnableBlurBehindWindow Lib "dwmapi.dll" (ByVal hwnd As Long, pBlurBehind As DWM_BLURBEHIND) As Long

Public Const DWM_EC_DISABLECOMPOSITION = 0
Public Const DWM_EC_ENABLECOMPOSITION = 1

Public Type tagSIZE
    cx As Long
    cy As Long
End Type

Public Declare Function DwmEnableComposition Lib "dwmapi.dll" (ByVal uCompositionAction As Long) As Long

Public Declare Function DwmEnableMMCSS Lib "dwmapi.dll" (ByVal fEnableMMCSS As Boolean) As Long

Public Declare Function DwmExtendFrameIntoClientArea Lib "dwmapi.dll" (ByVal hwnd As Long, pMarInset As MARGINS) As Long

Public Declare Function DwmGetColorizationColor Lib "dwmapi.dll" (pcrColorization As Long, pfOpaqueBlend As Boolean) As Long

Public Declare Function DwmGetCompositionTimingInfo Lib "dwmapi.dll" (ByVal hwnd As Long, pTimingInfo As DWM_TIMING_INFO) As Long

Public Declare Function DwmGetWindowAttribute Lib "dwmapi.dll" (ByVal hwnd As Long, ByVal dwAttribute As Long, pvAttribute As Any, ByVal cbAttribute As Long) As Long

Public Declare Function DwmIsCompositionEnabled Lib "dwmapi.dll" (pfEnabled As Boolean) As Long

Public Declare Function DwmModifyPreviousDxFrameDuration Lib "dwmapi.dll" (ByVal hwnd As Long, ByVal cRefreshes As Long, ByVal fRelative As Boolean) As Long

Public Declare Function DwmQueryThumbnailSourceSize Lib "dwmapi.dll" (ByVal hThumbnail As Long, pSize As tagSIZE) As Long

Public Declare Function DwmRegisterThumbnail Lib "dwmapi.dll" (ByVal hwndDestination As Long, ByVal hwndSource As Long, phThumbnailId As Long) As Long

Public Declare Function DwmSetDxFrameDuration Lib "dwmapi.dll" (ByVal hwnd As Long, ByVal cRefreshes As Long) As Long

Public Declare Function DwmSetPresentParameters Lib "dwmapi.dll" (ByVal hwnd As Long, pPresentParams As DWM_PRESENT_PARAMETERS) As Long

Public Declare Function DwmSetWindowAttribute Lib "dwmapi.dll" (ByVal hwnd As Long, ByVal dwAttribute As Long, pvAttribute As Any, ByVal cbAttribute As Long) As Long

Public Declare Function DwmUnregisterThumbnail Lib "dwmapi.dll" (ByVal hThumbnailId As Long) As Long

Public Declare Function DwmUpdateThumbnailProperties Lib "dwmapi.dll" (ByVal hThumbnailId As Long, ptnProperties As DWM_THUMBNAIL_PROPERTIES) As Long

Public Const DWM_SIT_DISPLAYFRAME = &H1           ' Display a window frame around the provided bitmap

Public Declare Function DwmSetIconicThumbnail Lib "dwmapi.dll" (ByVal hwnd As Long, ByVal hbmp As Long, ByVal dwSITFlags As Long) As Long

'Public Declare Function DwmSetIconicLivePreviewBitmap Lib "dwmapi.dll" (ByVal hwnd As Long, ByVal hbmp As Long, pptClient As Point, ByVal dwSITFlags As Long) As Long

Public Declare Function DwmInvalidateIconicBitmaps Lib "dwmapi.dll" (ByVal hwnd As Long) As Long

Public Declare Function DwmAttachMilContent Lib "dwmapi.dll" (ByVal hwnd As Long) As Long

Public Declare Function DwmDetachMilContent Lib "dwmapi.dll" (ByVal hwnd As Long) As Long

Public Declare Sub DwmFlush Lib "dwmapi.dll" ()

Public Type MilMatrix3x2D
    S_11 As Double
    S_12 As Double
    S_21 As Double
    S_22 As Double
    DX As Double
    DY As Double
End Type
    
Public Declare Function DwmGetGraphicsStreamTransformHint Lib "dwmapi.dll" (uIndex, pTransform As MilMatrix3x2D) As Long
    
'Public Declare Function DwmGetGraphicsStreamClient Lib "dwmapi.dll" (uIndex, pClientUuid As UUID) As Long
    
Public Declare Function DwmGetTransportAttributes Lib "dwmapi.dll" (pfIsRemoting As Boolean, pfIsConnected As Boolean, pDwGeneration As Long) As Long
