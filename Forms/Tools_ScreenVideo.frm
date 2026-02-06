VERSION 5.00
Begin VB.Form FormScreenVideo 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   " Select capture area"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9600
   ClipControls    =   0   'False
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   Icon            =   "Tools_ScreenVideo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdLabel lblInfo 
      Height          =   540
      Left            =   120
      Top             =   6600
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   953
      Alignment       =   2
      Caption         =   ""
      Layout          =   1
   End
   Begin PhotoDemon.pdButton cmdExit 
      Height          =   495
      Left            =   8280
      TabIndex        =   1
      Top             =   6600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "Cancel"
   End
   Begin PhotoDemon.pdButton cmdStart 
      Height          =   495
      Left            =   5880
      TabIndex        =   0
      Top             =   6600
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      Caption         =   "Start recording"
   End
End
Attribute VB_Name = "FormScreenVideo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Animated screen capture dialog
'Copyright 2020-2026 by Tanner Helland
'Created: 01/July/20
'Last updated: 01/October/21
'Last update: add support for WebP as a target format (in addition to existing APNG support)
'
'PD can write both animated PNGs and animated WebP files.  These formats are a great fit
' for animated screen captures (24-bit color!).  PhotoDemon provides a rudimentary screen
' recorder that can dump frames directly to either format, or cache them internally for
' subsequent loading into PhotoDemon (as a generic animated image container which you can then
' export however you want, even to GIF with all its limitations).
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'WAPI
Private Enum Win32_CombineRgnResult
    crr_Error = 0
    crr_NullRegion = 1
    crr_SimpleRegion = 2
    crr_ComplexRegion = 3
End Enum

#If False Then
    Private Const crr_Error = 0, crr_NullRegion = 1, crr_SimpleRegion = 2, crr_ComplexRegion = 3
#End If

Private Enum Win32_CombineRgnType
    crt_And = 1
    crt_Or = 2
    crt_Xor = 3
    crt_Diff = 4
    crt_Copy = 5
End Enum

#If False Then
    Private Const crt_And = 1, crt_Or = 2, crt_Xor = 3, crt_Diff = 4, crt_Copy = 5
#End If

Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Win32_CombineRgnType) As Win32_CombineRgnResult
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObj As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal wFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedrawImmediately As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

'END WAPI

'Magic magenta is used to mark the transparent section of the recording frame
Private Const KEY_COLOR As Long = &HFF00FF

'There are several ways to create a window with a "cut-out" region.  They all have trade-offs,
' especially when attempting to support XP through Win 10.  This value is set when the dialog
' is first loaded.
Private Enum PD_TransparentWindow
    tw_GDIRegion = 0
    tw_LayeredWindow = 1
End Enum

#If False Then
    Private Const tw_GDIRegion = 0, tw_LayeredWindow = 1
#End If

Private m_WindowMethod As PD_TransparentWindow

'The rectangle (in screen coords) of the window that summoned us (if any);
' this is used to position this window the first time it is launched.
Private m_parentRect As winRect, m_myRect As winRect

'A timer triggers capture events
Private WithEvents m_Timer As pdTimer
Attribute m_Timer.VB_VarHelpID = -1

'Target format when dumping frames to file.  Must be either PDIF_PNG or PDIF_WEBP.
Private m_FileFormat As PD_IMAGE_FORMAT

'A pdPNG instance handles the actual PNG writing (if dumping directly to an APNG file)
Private m_PNG As pdPNG

'Similarly, a pdWebP instance handles WebP encoding
Private m_WebP As pdWebP

'Destination file, if one is selected (check for null before using)
Private m_DstFilename As String

'Target maximum frame rate (as frames-per-second)
Private m_FPS As Double

'Current actual FPS (a total of all frame times; divide by timer hits NOT frame count,
' as duplicate frames are auto-suspended)
Private m_NetFrameTime As Currency, m_lastFrameTime As Currency, m_TimerHits As Long

'Repeat count of the final APNG
Private m_LoopCount As Long

'DEFLATE compression level (goes to 12, c/o libdeflate)
Private m_PNGCompressionLevel As Long

'WebP quality level [0, 100].  100 = lossless.  libwebp uses a float so we do too.
Private m_WebPQuality As Single

'Whether to include mouse cursor position and/or clicks in the animation
Private m_ShowCursor As Boolean, m_ShowClicks As Boolean

'Seconds to countdown (if any) before starting the recording
Private m_CountdownTime As Long

'Whether to save the image directly to disk, or load it into PD for further editing
Private m_SaveImmediatelyToDisk As Boolean

'Capture rects; once populated (at the start of the capture), these *cannot* be changed
Private m_CaptureRectClient As RectL, m_CaptureRectScreen As RectL

'If a capture event is ACTIVE, this will be set to TRUE
Private m_CaptureActive As Boolean

'If a full capture was performed successfully, this will be set to TRUE.
' At unload time, we use this to flag whether we should save the current screen position
' or not.
Private m_CaptureSuccessful As Boolean

'Capture DIBs.  Reused on successive frames for perf reasons.  Separate 24- and 32-bpp DIBs
' are used because we perform the actual capture using a 24-bpp DIB (there is a measurable
' perf benefit to this on Win 10), but the PNG export engine requires 32-bpp DIBs because
' it needs to perform heuristics to see if there's benefits to writing 32-bpp frames if
' transparency provides meaningful benefits during frame optimizations.  (Don't sweat the
' details!)  Anyway, the point is that all captured frames are captured and cached as 24-bpp,
' but they are upsampled to 32-bpp before getting passed to the APNG encoder.
Private m_captureDIB24 As pdDIB, m_captureDIB32 As pdDIB

'To improve performance, we cache the last-captured frame and compare it against subsequent
' frames.  Duplicate frames are auto-skipped.  (This is a very common occurrence during
' a screen capture, and it can save a *lot* of resources.)  The capture code automatically
' switches between two separate capture DIBs; this provides a fast way to perform duplicate
' detection within a fixed memory budget.
Private m_captureDIB24_2 As pdDIB

'Time stamp of when the recording started; used to determine elapsed time for the UI
Private m_StartTimeMS As Currency

'Emergency flag set when the dialog is canceled; if an APNG is actively being saved,
' we use this flag to terminate any current operations.
Private m_Cancel As Boolean

'Captured frames are stored as a collection of lz4-compressed arrays.  I don't currently have
' access to a DEFLATE library that can compress screen-capture-sized frames (e.g. 1024x768) in
' real-time on an XP-era PC, which would be required for dumping frames directly to an APNG file.
' lz4 ends up being a better general-purpose solution here, although it requires us to "play back"
' the stored frames when the capture ends.
Private Type PD_APNGFrameCapture
    fcTimeStamp As Currency
    frameSizeOrig As Long
    frameSizeCompressed As Long
    frameData() As Byte
End Type

'Frame count is tracked manually; certain things need to be handled differently on
' e.g. the first frame vs subsequent frames.
Private Const INIT_FRAME_BUFFER As Long = 64
Private m_FrameCount As Long
Private m_Frames() As PD_APNGFrameCapture

'For perf reasons, a persistent compression buffer is used; it is auto-enlarged to
' a "worst-case" size before capture begins.
Private m_CompressionBuffer() As Byte

'Countdown timer before starting.  Note that this may *not* be used if the countdown
' delay is set to 0.
Private WithEvents m_CountdownTimer As pdTimerCountdown
Attribute m_CountdownTimer.VB_VarHelpID = -1

'Various events are handled via API, not VB; this helps us support high-DPI displays
Private WithEvents m_Resize As pdWindowSize
Attribute m_Resize.VB_VarHelpID = -1
Private WithEvents m_Painter As pdWindowPainter
Attribute m_Painter.VB_VarHelpID = -1

'Window settings (particularly position) are saved/restored on each window launch
Private WithEvents m_lastUsedSettings As pdLastUsedSettings
Attribute m_lastUsedSettings.VB_VarHelpID = -1

'This dialog must be invoked via this function.  It preps a bunch of internal values that must exist
' for the recorder to function.
Public Sub ShowDialog(ByVal ptrToParentRect As Long, ByRef listOfSettings As String)
    
    'Before doing anything else, determine how we're going to "cut-out" a portion of this window.
    ' (The method used is the result of a ton of testing across multiple Windows versions.  Change at
    ' your own peril.)
    m_WindowMethod = tw_GDIRegion
    
    'Only one parameter is passed as-is (a pointer to the preferences window rect).
    ' All others are stored inside a standard PD parameter string.
    If (ptrToParentRect <> 0) Then CopyMemoryStrict VarPtr(m_parentRect), ptrToParentRect, LenB(m_parentRect)
    
    Dim cSettings As pdSerialize
    Set cSettings = New pdSerialize
    cSettings.SetParamString listOfSettings
    
    With cSettings
        m_FPS = .GetLong("frame-rate", 10)
        m_CountdownTime = .GetLong("countdown", 0)
        m_ShowCursor = .GetBool("show-cursor", True)
        m_ShowClicks = .GetBool("show-clicks", True)
        m_LoopCount = .GetLong("loop-count", 0)
        
        'Compression levels are validated before forwarding to their respective libraries
        m_PNGCompressionLevel = .GetLong("png-compression", 6)
        If (m_PNGCompressionLevel < 1) Then m_PNGCompressionLevel = 1
        If (m_PNGCompressionLevel > Compression.GetMaxCompressionLevel(cf_Zlib)) Then m_PNGCompressionLevel = Compression.GetMaxCompressionLevel(cf_Zlib)
        m_WebPQuality = .GetSingle("webp-quality", 100!)
        If (m_WebPQuality < 0!) Then m_WebPQuality = 0!
        If (m_WebPQuality > 100!) Then m_WebPQuality = 100!
        m_SaveImmediatelyToDisk = .GetBool("save-to-disk", True)
        
        'Only APNG and WebP are currently supported as recording targets
        Select Case .GetString("file-format", "webp")
            Case "png"
                m_FileFormat = PDIF_PNG
            Case Else
                m_FileFormat = PDIF_WEBP
        End Select
        
    End With
    
    'Initialize a last-used settings object.  (Because this isn't a "standard" PhotoDemon dialog,
    ' it doesn't get things like last-used settings for free.)
    Set m_lastUsedSettings = New pdLastUsedSettings
    m_lastUsedSettings.SetParentForm Me
    
    'Restore any last-used settings; note that this will include window position
    m_lastUsedSettings.LoadAllControlValues
    
    'Prepare the dialog (inc. setting up window transparency and moving the window to
    ' its last-used location)
    PrepWindowForRecording
    
    'Display this dialog as a MODELESS window (critical for always-on-top behavior!)
    Me.Show vbModeless
    
End Sub

Private Sub PrepWindowForRecording()

    If PDMain.IsProgramRunning Then
        
        'Subclassers are used for resize and paint events
        Set m_Resize = New pdWindowSize
        m_Resize.AttachToHWnd Me.hWnd, True
        Set m_Painter = New pdWindowPainter
        m_Painter.StartPainter Me.hWnd, True
        
        'Layered window approach!
        If (m_WindowMethod = tw_LayeredWindow) Then
        
            'Mark the underlying window as a layered window
            Const GWL_EXSTYLE As Long = -20
            Const WS_EX_LAYERED As Long = &H80000
            SetWindowLong Me.hWnd, GWL_EXSTYLE, GetWindowLong(Me.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED
        
        End If
        
        'Apply any icons
        cmdStart.AssignImage "macro_record", Nothing, Interface.FixDPI(20), Interface.FixDPI(20)
        
        'When applying theming, note that we request to paint our window manually;
        ' normally PD handles this centrally, but this window has special needs.
        Interface.ApplyThemeAndTranslations Me, False
        
        'Ask the system to let us paint at least once before the form is actually displayed
        ForceWindowRepaint
        
        'Validate our initial positioning rectangle, and adjust as necessary for off-screen positions
        Dim screenRect As RectL
        g_Displays.GetDesktopRect screenRect
        
        With m_myRect
            If (.x1 < screenRect.Left) Then .x1 = screenRect.Left
            If (.y1 < screenRect.Top) Then .y1 = screenRect.Top
            If (.x2 > screenRect.Right) Then .x1 = screenRect.Right - (.x2 - .x1)
            If (.y2 > screenRect.Bottom) Then .y1 = screenRect.Bottom - (.y2 - .y1)
        End With
        
        'Mark this window as "always on-top", then position it.
        ' (Note that out of an abundance of caution, we also pass SWP_FRAMECHANGED because
        ' we may have messed with window bits via SetWindowLong earlier in the function.)
        Const HWND_TOPMOST As Long = -1&, SWP_FRAMECHANGED As Long = &H20&
        With m_myRect
            g_WindowManager.SetWindowPos_API Me.hWnd, HWND_TOPMOST, .x1, .y1, .x2 - .x1, .y2 - .y1, SWP_FRAMECHANGED
        End With
        
    End If
    
End Sub

Private Sub cmdExit_Click()
    m_Cancel = True
    StopTimer_Forcibly
    Unload Me
End Sub

Private Sub cmdStart_Click()
    
    'If a capture is already active, STOP the timer and construct the APNG file
    If m_CaptureActive Then
        Capture_Stop
        
    'If a capture is NOT active, START the capture timer (after prompting the user for an export path)
    Else
        
        'Update the screen coord version of the transparent window
        m_CaptureRectScreen = m_CaptureRectClient
        g_WindowManager.GetClientToScreen_Universal Me.hWnd, VarPtr(m_CaptureRectScreen.Left)
        g_WindowManager.GetClientToScreen_Universal Me.hWnd, VarPtr(m_CaptureRectScreen.Right)
        
        'If all preliminary checks passed, activate the capture timer.  Note that this
        ' *will* forcibly overwrite the file at the destination location, if one exists.
        ' (Also, note that we deliberately request a slightly shorter interval (5%) than
        ' we require - timer events are not that precise, especially because of coalescing
        ' on Win 7+.  A slightly reduced interval gives us some breathing room and
        ' increases our chances of hitting the user's requested frame rate.)
        Dim tInterval As Long
        tInterval = Int(1000# / m_FPS + 0.5)
        tInterval = Int((tInterval * 95#) / 100# + 0.5)
        
        Set m_Timer = New pdTimer
        m_Timer.Interval = tInterval
        
        'Prep the capture DIB
        Set m_captureDIB24 = New pdDIB
        m_captureDIB24.CreateBlank m_CaptureRectClient.Right - m_CaptureRectClient.Left, m_CaptureRectClient.Bottom - m_CaptureRectClient.Top, 24, vbBlack, 255
        
        'Prep the "last frame" DIB which we'll use to look for duplicate frames
        Set m_captureDIB24_2 = New pdDIB
        m_captureDIB24_2.CreateFromExistingDIB m_captureDIB24
        
        'Prepare a persistent compression buffer
        Dim cmpBufferSize As Long
        cmpBufferSize = Compression.GetWorstCaseSize(m_captureDIB24.GetDIBStride * m_captureDIB24.GetDIBHeight, cf_Lz4)
        ReDim m_CompressionBuffer(0 To cmpBufferSize - 1) As Byte
        
        'Initialize the frame collection
        m_FrameCount = 0
        ReDim m_Frames(0 To INIT_FRAME_BUFFER - 1) As PD_APNGFrameCapture
        
        'Change the start button to a STOP button
        cmdStart.AssignImage "macro_stop", Nothing, Interface.FixDPI(20), Interface.FixDPI(20)
        cmdStart.Caption = g_Language.TranslateMessage("End recording")
        cmdExit.Caption = g_Language.TranslateMessage("Cancel")
        
        'If a countdown timer was specified, start one now; otherwise, start recording immediately
        If (m_CountdownTime > 0) Then
            Set m_CountdownTimer = New pdTimerCountdown
            m_CountdownTimer.SetIntervalTimeInMS 1000
            m_CountdownTimer.SetCountdownTimeInMS m_CountdownTime * 1000
            lblInfo.Caption = g_Language.TranslateMessage("Recording will start in %1...", m_CountdownTime)
            m_CountdownTimer.StartCountdown
        Else
            Capture_Start
        End If
        
    End If
        
End Sub

'Start the capture timer
Private Sub Capture_Start()
    m_CaptureActive = True
    m_Timer.StartTimer
End Sub

'Stop the active capture
Private Sub Capture_Stop()

    If m_CaptureActive Then
        
        'Immediately stop the capture timer
        m_CaptureActive = False
        If (Not m_Timer Is Nothing) Then m_Timer.StopTimer
        
        'Saving to disk may take awhile.  We don't want to remain on-top while this occurs, so disable
        ' top-most behavior until saving finishes.
        
        'Mark this window as "always on-top", then position it.
        Dim currentRect As winRect
        If (Not g_WindowManager Is Nothing) Then g_WindowManager.GetWindowRect_API Me.hWnd, currentRect
        Const HWND_TOPMOST As Long = -1&, HWND_NOTOPMOST As Long = -2&
        Const SWP_NOMOVE As Long = &H2&, SWP_NORESIZE As Long = &H1&, SWP_FRAMECHANGED As Long = &H20&
        
        With currentRect
            g_WindowManager.SetWindowPos_API Me.hWnd, HWND_NOTOPMOST, .x1, .y1, .x2 - .x1, .y2 - .y1, SWP_NOMOVE Or SWP_NORESIZE Or SWP_FRAMECHANGED
        End With
        
        'Next, our behavior varies depending on the user's export settings.
        If m_SaveImmediatelyToDisk Then
            FinishRecording_ToDisk
        Else
            FinishRecording_ToPD
        End If
        
        'Immediately trigger a save of the current screen position.
        ' (Unlike other dialogs, we don't save position at export time - we save it after
        ' a successful capture!)
        m_CaptureSuccessful = True
        If (Not m_lastUsedSettings Is Nothing) Then m_lastUsedSettings.SaveAllControlValues
        
        'Restore original topmost behavior
        If (Not g_WindowManager Is Nothing) Then g_WindowManager.GetWindowRect_API Me.hWnd, currentRect
        With currentRect
            g_WindowManager.SetWindowPos_API Me.hWnd, HWND_TOPMOST, .x1, .y1, .x2 - .x1, .y2 - .y1, SWP_NOMOVE Or SWP_NORESIZE Or SWP_FRAMECHANGED
        End With
        
        'If the user is loading this file into PD, unload this dialog immediately
        If (Not m_SaveImmediatelyToDisk) Then
            StopTimer_Forcibly
            Unload Me
        End If
        
    End If
        
End Sub

Private Sub FinishRecording_ToDisk()

    'Notify the animation engine that we're handling the export locally
    Animation.SetAnimationTmpFile vbNullString
    
    'We first need to prompt the user for a destination filename
    
    'Start by validating m_dstFilename, which will be filled with the user's past
    ' destination filename (if one exists), or a default capture filename in the
    ' user's current "Save image" folder
    If ((LenB(m_DstFilename) = 0) Or (Not Files.PathExists(Files.FileGetPath(m_DstFilename)))) Then
    
        'm_dstFilename is bad.  Attempt to populate it with default values.
        Dim tmpPath As String, tmpFilename As String, tmpFileExtension As String
        tmpPath = UserPrefs.GetPref_String("Paths", "Save Image", vbNullString)
        tmpFilename = g_Language.TranslateMessage("capture")
        
        If (m_FileFormat = PDIF_PNG) Then
            tmpFileExtension = "png"
        ElseIf (m_FileFormat = PDIF_WEBP) Then
            tmpFileExtension = "webp"
        End If
        m_DstFilename = tmpPath & IncrementFilename(tmpPath, tmpFilename, tmpFileExtension) & "." & tmpFileExtension
    
    End If
    
    'Use a standard common-dialog to prompt for filename
    Dim cSave As pdOpenSaveDialog
    Set cSave = New pdOpenSaveDialog
    
    Dim okToProceed As Boolean, sFile As String
    sFile = m_DstFilename
    If (m_FileFormat = PDIF_PNG) Then
        okToProceed = cSave.GetSaveFileName(sFile, Files.FileGetName(m_DstFilename), True, "Animated PNG (.png)|*.png;*.apng", 1, Files.FileGetPath(m_DstFilename), "Save image", ".png", Me.hWnd)
    ElseIf (m_FileFormat = PDIF_WEBP) Then
        okToProceed = cSave.GetSaveFileName(sFile, Files.FileGetName(m_DstFilename), True, "Animated WebP (.webp)|*.webp", 1, Files.FileGetPath(m_DstFilename), "Save image", ".webp", Me.hWnd)
    End If
        
    'The user can cancel the common-dialog - that's fine; it just means we don't save
    ' any of the current settings (or close the window).
    If okToProceed Then
        
        'Save the current export path as the latest "save image" path
        m_DstFilename = sFile
        UserPrefs.SetPref_String "Paths", "Save Image", Files.FileGetPath(m_DstFilename)
        UserPrefs.SetPref_Boolean "Saving", "Has Saved A File", True
        
        'To avoid confusion, set the "start recording" button caption to "please wait".
        ' It will get formally reset after the image export ends.
        cmdStart.Caption = g_Language.TranslateMessage("Please wait")
        
        'Start dumping frames to file
        If (m_FileFormat = PDIF_PNG) Then
            WriteAPNG
        ElseIf (m_FileFormat = PDIF_WEBP) Then
            WriteWebP
        End If
        
        'If the caller canceled recording, delete any in-progress file efforts
        If m_Cancel Then
        
            Files.FileDeleteIfExists m_DstFilename
        
        'If recording was successful, add this image to PD's recent files list
        ' (which greatly simplifies the process of re-opening it for further edits).
        Else
            
            'Re-extract the first frame's pixel data
            Compression.DecompressPtrToPtr m_captureDIB24.GetDIBPointer, m_Frames(0).frameSizeOrig, VarPtr(m_Frames(0).frameData(0)), m_Frames(0).frameSizeCompressed, cf_Lz4
            
            'Convert it to 32-bpp with a solid alpha channel
            If (Not m_captureDIB32 Is Nothing) Then
                GDI.BitBltWrapper m_captureDIB32.GetDIBDC, 0, 0, m_captureDIB32.GetDIBWidth, m_captureDIB32.GetDIBHeight, m_captureDIB24.GetDIBDC, 0, 0, vbSrcCopy
                m_captureDIB32.ForceNewAlpha 255
            End If
            
            g_RecentFiles.AddFileToList m_DstFilename, Nothing, m_captureDIB32
            
        End If
        
        'Free the memory used by the first frame (now that we've generated a thumbnail)
        Erase m_Frames(0).frameData
        
        'Note that the save was successful
        lblInfo.Caption = g_Language.TranslateMessage("Save complete.")
    
    Else
        lblInfo.Caption = g_Language.TranslateMessage("Save canceled.")
    End If
    
    'Reset this button's caption and notify the user that we're finished
    cmdStart.AssignImage "macro_record", Nothing, Interface.FixDPI(20), Interface.FixDPI(20)
    cmdStart.Caption = g_Language.TranslateMessage("Start recording")
    cmdExit.Caption = g_Language.TranslateMessage("Exit")
    
End Sub

Private Sub WriteAPNG()

    'Start an APNG streamer
    If (m_PNG Is Nothing) Then Set m_PNG = New pdPNG
    If (m_PNG.SaveAPNG_Streaming_Start(m_DstFilename, m_captureDIB24.GetDIBWidth, m_captureDIB24.GetDIBHeight) < png_Failure) Then
        
        'Now comes the fun part: loading all cached frames, and passing them off to the APNG writer
        ' so that it can produce a usable APNG file!
        Dim i As Long
        For i = 0 To m_FrameCount - 1
            
            'Periodically check for emergency cancellation
            If m_Cancel Then GoTo EndImmediately
            
            lblInfo.Caption = g_Language.TranslateMessage("Saving animation frame %1 of %2...", i + 1, m_FrameCount)
            lblInfo.RequestRefresh
            
            'Extract this frame into the capture DIB, then immediately free its compressed memory
            Compression.DecompressPtrToPtr m_captureDIB24.GetDIBPointer, m_Frames(i).frameSizeOrig, VarPtr(m_Frames(i).frameData(0)), m_Frames(i).frameSizeCompressed, cf_Lz4
            If m_Cancel Then GoTo EndImmediately
            
            '(Note that we deliberately do *not* free the first frame - we want to save it
            ' to generate a file thumbnail before exiting.)
            If (i <> 0) Then Erase m_Frames(i).frameData
            
            'Convert the 24-bpp DIB to 32-bpp before handing it off to the APNG encoder
            If (m_captureDIB32 Is Nothing) Then
                Set m_captureDIB32 = New pdDIB
                m_captureDIB32.CreateBlank m_captureDIB24.GetDIBWidth, m_captureDIB24.GetDIBHeight, 32, 0, 255
                m_captureDIB32.SetInitialAlphaPremultiplicationState True
            End If
            
            GDI.BitBltWrapper m_captureDIB32.GetDIBDC, 0, 0, m_captureDIB32.GetDIBWidth, m_captureDIB32.GetDIBHeight, m_captureDIB24.GetDIBDC, 0, 0, vbSrcCopy
            If m_Cancel Then GoTo EndImmediately
            m_captureDIB32.ForceNewAlpha 255
            
            'Pass the frame off to the PNG encoder
            If m_Cancel Then GoTo EndImmediately
            m_PNG.SaveAPNG_Streaming_Frame m_captureDIB32, m_Frames(i).fcTimeStamp, m_PNGCompressionLevel
            
            'Every few frames, notify the OS that we're still alive
            If ((i And 3) = 0) Then VBHacks.DoEvents_SingleHwnd Me.hWnd
            
        Next i
        
    Else
        PDDebug.LogAction "WARNING!  APNG screen capture failed for unknown reason.  Consult debug log."
    End If
    
EndImmediately:
    
    'Notify the PNG encoder that the stream has ended
    If (Not m_PNG Is Nothing) Then m_PNG.SaveAPNG_Streaming_Stop m_LoopCount
    Set m_PNG = Nothing
        
End Sub

Private Sub WriteWebP()

    'Start a WebP streamer
    If (m_WebP Is Nothing) Then Set m_WebP = New pdWebP
    If m_WebP.SaveStreamingWebP_Start(m_captureDIB24.GetDIBWidth, m_captureDIB24.GetDIBHeight, m_LoopCount, m_WebPQuality) Then
        
        'Now comes the fun part: loading all cached frames, and passing them off to the APNG writer
        ' so that it can produce a usable APNG file!
        Dim i As Long
        For i = 0 To m_FrameCount - 1
            
            'Periodically check for emergency cancellation
            If m_Cancel Then GoTo EndImmediately
            
            lblInfo.Caption = g_Language.TranslateMessage("Saving animation frame %1 of %2...", i + 1, m_FrameCount)
            lblInfo.RequestRefresh
            
            'Extract this frame into the capture DIB, then immediately free its compressed memory
            Compression.DecompressPtrToPtr m_captureDIB24.GetDIBPointer, m_Frames(i).frameSizeOrig, VarPtr(m_Frames(i).frameData(0)), m_Frames(i).frameSizeCompressed, cf_Lz4
            If m_Cancel Then GoTo EndImmediately
            
            '(Note that we deliberately do *not* free the first frame - we want to save it
            ' to generate a file thumbnail before exiting.)
            If (i <> 0) Then Erase m_Frames(i).frameData
            
            'WebP encoding uses timestamps instead of frame times.  This means that individual frame times
            ' are not calculated until the *next* frame arrives.
            m_WebP.SaveStreamingWebP_AddFrame m_captureDIB24, m_Frames(i).fcTimeStamp
            
            'Every few frames, notify the OS that we're still alive
            If ((i And 3) = 0) Then VBHacks.DoEvents_SingleHwnd Me.hWnd
            
        Next i
        
    Else
        PDDebug.LogAction "WARNING!  WebP screen capture failed for unknown reason.  Consult debug log."
    End If
    
EndImmediately:
    
    'Notify the WebP encoder that the stream has ended
    If (Not m_WebP Is Nothing) Then m_WebP.SaveStreamingWebP_Stop m_Frames(m_FrameCount - 1).fcTimeStamp + 3000, m_DstFilename
    Set m_WebP = Nothing
        
End Sub

Private Sub FinishRecording_ToPD()

    'Ideally, we'd construct a new pdImage object purely "in-memory" and load *that* into
    ' the program.  But I don't really have a good way to do this at present, as all image-loading
    ' code is built around discrete files.
    '
    'So instead, I'm gonna cheat and do things the easy way: by constructing a "temporary" pdImage
    ' object, saving it out to a temp file, then using PD's standard "load file" function to create
    ' a new image.  (This ensures that the new image behaves like any other freshly loaded file.)
    Dim tmpImage As pdImage
    PDImages.GetDefaultPDImageObject tmpImage
    
    'Add each recorded layer one-at-a-time, and free its associated memory as we go.
    Dim newLayerID As Long
    
    Dim i As Long
    For i = 0 To m_FrameCount - 1
    
        lblInfo.Caption = g_Language.TranslateMessage("Saving animation frame %1 of %2...", i + 1, m_FrameCount)
        lblInfo.RequestRefresh
        
        'Extract the original frame data into the capture DIB (which is already sized correctly),
        ' then immediately free its compressed memory.
        Compression.DecompressPtrToPtr m_captureDIB24.GetDIBPointer, m_Frames(i).frameSizeOrig, VarPtr(m_Frames(i).frameData(0)), m_Frames(i).frameSizeCompressed, cf_Lz4
        Erase m_Frames(i).frameData
        
        'Convert the 24-bpp DIB to 32-bpp before constructing a layer from it
        If (m_captureDIB32 Is Nothing) Then
            Set m_captureDIB32 = New pdDIB
            m_captureDIB32.CreateBlank m_captureDIB24.GetDIBWidth, m_captureDIB24.GetDIBHeight, 32, 0, 255
            m_captureDIB32.SetInitialAlphaPremultiplicationState True
        End If
        
        GDI.BitBltWrapper m_captureDIB32.GetDIBDC, 0, 0, m_captureDIB32.GetDIBWidth, m_captureDIB32.GetDIBHeight, m_captureDIB24.GetDIBDC, 0, 0, vbSrcCopy
        m_captureDIB32.ForceNewAlpha 255
        
        'Add this frame to the image
        newLayerID = tmpImage.CreateBlankLayer()
        
        'Calculate frame time (noting that the final frame is written differently, since it doesn't
        ' represent a difference between frames).
        Dim frameMS As Long
        If (i < m_FrameCount - 1) Then
            frameMS = m_Frames(i + 1).fcTimeStamp - m_Frames(i).fcTimeStamp
        
        'Display the last frame for some arbitrary amount of time (currently 3 seconds)
        Else
            frameMS = 3000
        End If
        
        Dim lyrName As String
        lyrName = g_Language.TranslateMessage("Frame %1", i + 1)
        lyrName = lyrName & " (" & Trim$(Str$(frameMS)) & "ms)"
        tmpImage.GetLayerByID(newLayerID).InitializeNewLayer PDL_Image, lyrName, m_captureDIB32, True
        tmpImage.GetLayerByID(newLayerID).SetLayerFrameTimeInMS frameMS
        tmpImage.GetLayerByID(newLayerID).SetLayerVisibility (i = 0)
        
    Next i
    
    lblInfo.Caption = g_Language.TranslateMessage("Finalizing image...")
    lblInfo.RequestRefresh
    
    'With all layers added, we need to populate a few remaining image attributes
    tmpImage.SetActiveLayerByIndex 0
    tmpImage.UpdateSize
    tmpImage.SetDPI 96, 96
    
    tmpImage.SetOriginalFileFormat PDIF_UNKNOWN
    tmpImage.SetCurrentFileFormat PDIF_UNKNOWN
    tmpImage.SetOriginalColorDepth 32
    tmpImage.SetOriginalGrayscale False
    tmpImage.SetOriginalAlpha True
    tmpImage.SetAnimated True
    
    tmpImage.ImgStorage.AddEntry "CurrentLocationOnDisk", vbNullString
    tmpImage.ImgStorage.AddEntry "OriginalFileName", vbNullString
    tmpImage.ImgStorage.AddEntry "OriginalFileExtension", vbNullString
    
    'Write the temp image out to file
    Dim tmpExportFile As String
    tmpExportFile = UserPrefs.GetTempPath & "screen_capture.pdi"
    Saving.SavePDI_Image tmpImage, tmpExportFile, True, cf_Lz4, cf_Lz4, False
    Animation.SetAnimationTmpFile tmpExportFile
    
End Sub

'Whenever the form is resized, we must repaint it with the transparent key color
Private Sub PaintForm(ByVal targetDC As Long)
    
    If PDMain.IsProgramRunning() Then
        
        'Retrieve a client rect and convert it to screen coordinates
        Dim myRectClient As winRect, myRectScreen As winRect
        g_WindowManager.GetClientWinRect Me.hWnd, myRectClient
        myRectScreen = myRectClient
        g_WindowManager.GetClientToScreen_Universal Me.hWnd, VarPtr(myRectScreen.x1)
        g_WindowManager.GetClientToScreen_Universal Me.hWnd, VarPtr(myRectScreen.x2)
        
        'Wrap a surface object around the window DC
        Dim cSurface As pd2DSurface
        Set cSurface = New pd2DSurface
        cSurface.WrapSurfaceAroundDC targetDC
        cSurface.SetSurfaceAntialiasing P2_AA_None
        cSurface.SetSurfaceCompositing P2_CM_Overwrite
        
        'Paint the background with the default theme background color
        ' (and make the paint region arbitrarily large)
        Dim cBrush As pd2DBrush
        Set cBrush = New pd2DBrush
        cBrush.SetBrushColor g_Themer.GetGenericUIColor(UI_Background)
        PD2D.FillRectangleI cSurface, cBrush, 0, 0, (myRectClient.x2 - myRectClient.x1) * 2, (myRectClient.y2 - myRectClient.y1) * 2
        
        'Padding is arbitrary; we want it large enough to create a clean border, but not so large
        ' that we waste precious screen real estate
        Dim borderPadding As Long
        borderPadding = 8
        
        'Populate the capture rect; this is the region of this window (client coords)
        ' that will be made transparent, e.g. mouse events will "click through" this region
        With m_CaptureRectClient
            
            'Pad accordingly
            .Left = borderPadding
            .Right = myRectClient.x2 - borderPadding
            
            'Same for top/bottom
            .Top = myRectClient.y1 + borderPadding
            .Bottom = cmdStart.GetTop - borderPadding
            
        End With
        
        'Paint the capture rect with the transparent key color (layered windows only)
        If (m_WindowMethod = tw_LayeredWindow) Then
            cBrush.SetBrushColor KEY_COLOR
            PD2D.FillRectangleI_FromRectL cSurface, cBrush, m_CaptureRectClient
        End If
        
        Set cBrush = Nothing
        
        'Finally, paint a 1px black border around the capture area
        Dim cPen As pd2DPen
        Set cPen = New pd2DPen
        cPen.SetPenColor RGB(0, 0, 0)
        cPen.SetPenLineJoin P2_LJ_Miter
        cPen.SetPenWidth 1!
        PD2D.DrawRectangleI_AbsoluteCoords cSurface, cPen, m_CaptureRectClient.Left - 1, m_CaptureRectClient.Top - 1, m_CaptureRectClient.Right, m_CaptureRectClient.Bottom
        Set cPen = Nothing
        
        Set cSurface = Nothing
        
        If (m_WindowMethod = tw_LayeredWindow) Then
        
            'Update window attributes to note the key color
            Const LWA_COLORKEY As Long = &O1&
            SetLayeredWindowAttributes Me.hWnd, KEY_COLOR, 255, LWA_COLORKEY
        
        ElseIf (m_WindowMethod = tw_GDIRegion) Then
            
            'The GDI regions used by this function include non-client areas (ugh).  As such,
            ' we need to manually create a new window
            'Create two regions: one for the full window rect, and one for the capture area
            Dim rgn1 As Long, rgn2 As Long, myWinRect As winRect
            If (Not g_WindowManager Is Nothing) Then g_WindowManager.GetWindowRect_API Me.hWnd, myWinRect
            rgn1 = CreateRectRgn(0, 0, myWinRect.x2 - myWinRect.x1, myWinRect.y2 - myWinRect.y1)
            rgn2 = CreateRectRgn((myRectScreen.x1 - myWinRect.x1) + m_CaptureRectClient.Left, (myRectScreen.y1 - myWinRect.y1) + m_CaptureRectClient.Top, (myRectScreen.x1 - myWinRect.x1) + m_CaptureRectClient.Right, (myRectScreen.y1 - myWinRect.y1) + m_CaptureRectClient.Bottom)
            
            'Merge them, then delete the original copies
            Dim rgn3 As Long
            rgn3 = CreateRectRgn(0, 0, 0, 0)
            CombineRgn rgn3, rgn1, rgn2, crt_Diff
            DeleteObject rgn1
            DeleteObject rgn2
            
            'Assign the window region (and importantly, do *not* delete it - the system owns it)
            SetWindowRgn Me.hWnd, rgn3, 1
            
        End If
        
    End If

End Sub

'On timer events, capture the current frame by calling this function
Private Sub CaptureFrameNow()
    
    'Make sure we have room to store this frame
    If (m_FrameCount > UBound(m_Frames)) Then ReDim Preserve m_Frames(0 To m_FrameCount * 2 - 1) As PD_APNGFrameCapture
    
    '*Immediately* before capture, note the current time
    Dim capTime As Currency
    capTime = VBHacks.GetHighResTimeInMSEx()
    
    'Capture the frame in question into a pdDIB object; note that we alternate
    ' which DIB we use for capture; this allows us to compare back-to-back frames
    ' in real-time without need for a manual copy op
    Dim captureTarget As pdDIB
    If ((m_FrameCount And &H1&) <> 0) Then
        Set captureTarget = m_captureDIB24_2
    Else
        Set captureTarget = m_captureDIB24
    End If
    
    ScreenCapture.GetPartialDesktopAsDIB captureTarget, m_CaptureRectScreen, m_ShowCursor, m_ShowClicks
    
    'Before saving this frame, check for duplicate frames.  This is very common during
    ' a screen capture event, and we can save a lot of memory by skipping these frames.
    ' (Note that the way we track time between frames handles these skips naturally,
    ' no extra work required!)
    Dim keepFrame As Boolean
    keepFrame = (m_FrameCount = 0)
    
    If (Not keepFrame) Then
        keepFrame = Not VBHacks.MemCmp(m_captureDIB24.GetDIBPointer, m_captureDIB24_2.GetDIBPointer, m_captureDIB24.GetDIBStride * m_captureDIB24.GetDIBHeight)
    End If
    
    If keepFrame Then
        
        'Compress the captured frame into the temporary frame buffer
        Dim cmpSize As Long
        cmpSize = UBound(m_CompressionBuffer) + 1
        Compression.CompressPtrToPtr VarPtr(m_CompressionBuffer(0)), cmpSize, captureTarget.GetDIBPointer, m_captureDIB24.GetDIBStride * m_captureDIB24.GetDIBHeight, cf_Lz4
        
        'Allocate memory for this frame and store all data
        With m_Frames(m_FrameCount)
            .fcTimeStamp = capTime
            .frameSizeOrig = m_captureDIB24.GetDIBStride * m_captureDIB24.GetDIBHeight
            .frameSizeCompressed = cmpSize
            ReDim .frameData(0 To cmpSize - 1) As Byte
            CopyMemoryStrict VarPtr(.frameData(0)), VarPtr(m_CompressionBuffer(0)), cmpSize
        End With
        
        'PDDebug.LogAction "Total frame time: " & VBHacks.GetTimeDiffNowAsString(capTime) & ", compression reduced size by " & CStr(100# * (1# - (cmpSize / (m_captureDIB24.GetDIBStride * m_captureDIB24.GetDIBHeight)))) & "%"
        
        'Increment frame count
        m_FrameCount = m_FrameCount + 1
    
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'Forcibly stop the current capture (if any)
    StopTimer_Forcibly
    
    'Free the last-used settings manager.  (Note that unlike most dialogs, we do NOT
    ' save dialog settings here - instead, we save it after a successful capture event.)
    m_lastUsedSettings.SetParentForm Nothing
    
    'Restore the main PD window
    FormMain.WindowState = vbNormal
    
    'If we're supposed to load our recording into PD, do so now
    Animation.CreateNewPDImageFromAnimation
    
End Sub

Private Sub m_CountdownTimer_CountdownFinished()
    Capture_Start
End Sub

Private Sub m_CountdownTimer_UpdateTimeRemaining(ByVal timeRemainingInMS As Long)
    lblInfo.Caption = g_Language.TranslateMessage("Recording will start in %1...", Int((timeRemainingInMS + 500) \ 1000))
End Sub

Private Sub m_LastUsedSettings_AddCustomPresetData()
    
    'Start by retrieving our current window rect (in screen coordinates!)
    Dim myRect As winRect
    If (Not g_WindowManager Is Nothing) Then
    
        g_WindowManager.GetWindowRect_API Me.hWnd, myRect
    
        'Save the current rect to file
        With m_lastUsedSettings
            .AddPresetData "window-x1", Trim$(Str$(myRect.x1))
            .AddPresetData "window-x2", Trim$(Str$(myRect.x2))
            .AddPresetData "window-y1", Trim$(Str$(myRect.y1))
            .AddPresetData "window-y2", Trim$(Str$(myRect.y2))
        End With
        
    End If
    
    'Also save the current destination filename, if any
    If ((LenB(m_DstFilename) <> 0) And Files.PathExists(Files.FileGetPath(m_DstFilename))) Then m_lastUsedSettings.AddPresetData "dst-capture-filename", m_DstFilename

End Sub

Private Sub m_LastUsedSettings_ReadCustomPresetData()

    'Look for previously saved window state
    If m_lastUsedSettings.DoesPresetExist("window-x1") Then
    
        'Screen coordinates are stored individually
        With m_lastUsedSettings
            m_myRect.x1 = .RetrievePresetData("window-x1", m_parentRect.x1)
            m_myRect.x2 = .RetrievePresetData("window-x2", m_parentRect.x2)
            m_myRect.y1 = .RetrievePresetData("window-y1", m_parentRect.y1)
            m_myRect.y2 = .RetrievePresetData("window-y2", m_parentRect.y2)
        End With
        
    'Previous window state doesn't exist; try to create a good default position
    Else
        
        'Make sure our parent rect (if any was loaded) is valid
        If ((m_parentRect.x2 - m_parentRect.x1) > 0) And ((m_parentRect.y2 - m_parentRect.y1) > 0) Then
            
            'Position this window identically to our parent window; this makes it likely that
            ' the user will see where this window loaded!
            m_myRect = m_parentRect
        
        'If we weren't passed a parent rect.... idk.  This should never happen.
        ' As a failsafe, steal FormMain's rect.
        Else
            Dim mainRect As winRect
            If (Not g_WindowManager Is Nothing) Then g_WindowManager.GetWindowRect_API FormMain.hWnd, mainRect
            m_myRect = mainRect
        End If
        
    End If
    
    'Look for a previously saved destination filename.  If one does not exist,
    ' we want to populate the destination path with a good default suggestion.
    Dim fileExtension As String
    If (m_FileFormat = PDIF_PNG) Then
        fileExtension = "png"
    Else
        fileExtension = "webp"
    End If
    
    If (Not m_lastUsedSettings.DoesPresetExist("dst-capture-filename")) Then
    
        Dim tmpPath As String, tmpFilename As String, tmpCombined As String
        tmpPath = UserPrefs.GetPref_String("Paths", "Save Image", vbNullString)
        tmpFilename = g_Language.TranslateMessage("capture")
        tmpCombined = tmpPath & IncrementFilename(tmpPath, tmpFilename, fileExtension) & "." & fileExtension
        
        m_DstFilename = tmpCombined
        
    Else
    
        'Retrieve the previous filename, and ensure the extension is updated to match whatever the
        ' *CURRENT* format is for recording (png or webp)
        m_DstFilename = m_lastUsedSettings.RetrievePresetData("dst-capture-filename", UserPrefs.GetPref_String("Paths", "Save Image", vbNullString) & "capture." & fileExtension)
        m_DstFilename = Files.FileGetPath(m_DstFilename) & Files.FileGetName(m_DstFilename, True) & "." & fileExtension
    
    End If
    
End Sub

'When the system sends a WM_ERASEBKGND message, do a quick fill with the current theme
' background color; this prevents white flickering along form edges during a resize.
Private Sub m_Painter_EraseBkgnd()
    
    Dim tmpDC As Long
    tmpDC = m_Painter.GetPaintStructDC()
    
    Dim tmpRect As winRect
    g_WindowManager.GetClientWinRect Me.hWnd, tmpRect
    
    With tmpRect
        GDI.FillRectToDC tmpDC, .x1, .y1, .x2 - .x1, .y2 - .y1, g_Themer.GetGenericUIColor(UI_Background)
    End With
    
End Sub

'Respond to all paint messages, obviously!
Private Sub m_Painter_PaintWindow(ByVal winLeft As Long, ByVal winTop As Long, ByVal winWidth As Long, ByVal winHeight As Long)
    Dim tmpDC As Long
    tmpDC = m_Painter.GetPaintStructDC()
    PaintForm tmpDC
End Sub

Private Sub m_Resize_WindowResize(ByVal newWidth As Long, ByVal newHeight As Long)

    'Before resizing, ensure we are *not* capturing - if we are, cancel everything
    StopTimer_Forcibly
    
    'Reposition the "start" and "stop" buttons
    Dim btnPadding As Long
    btnPadding = Interface.FixDPI(6)
    
    Dim myClientRect As winRect
    If (Not g_WindowManager Is Nothing) Then g_WindowManager.GetClientWinRect Me.hWnd, myClientRect
    
    cmdExit.SetLeft (myClientRect.x2 - myClientRect.x1) - (btnPadding + cmdExit.GetWidth)
    cmdExit.SetTop (myClientRect.y2 - myClientRect.y1) - (btnPadding + cmdExit.GetHeight)
    
    cmdStart.SetLeft (myClientRect.x2 - myClientRect.x1) - (btnPadding + cmdExit.GetWidth) - (btnPadding + cmdStart.GetWidth)
    cmdStart.SetTop (myClientRect.y2 - myClientRect.y1) - (btnPadding + cmdStart.GetHeight)
    
    'Reposition the "info" label
    lblInfo.SetPositionAndSize Interface.FixDPI(8), cmdStart.GetTop, PDMath.Max2Int(cmdStart.GetLeft - (btnPadding + Interface.FixDPI(8)), btnPadding), lblInfo.GetHeight
    
    'Force a repaint of the back buffer (to ensure the window transparency gets updated too!)
    ForceWindowRepaint
    
End Sub

Private Sub m_Timer_Timer()
    
    'Update the "time elapsed" window
    Dim timeElapsedInSeconds As Currency
    
    If (m_FrameCount = 0) Then
        VBHacks.GetHighResTimeInMS m_StartTimeMS
        timeElapsedInSeconds = 0
        m_NetFrameTime = 0
        m_TimerHits = 0
    Else
        timeElapsedInSeconds = (VBHacks.GetHighResTimeInMSEx() - m_StartTimeMS) / 1000
    End If
    
    'Capture the current frame
    If m_CaptureActive Then CaptureFrameNow
    
    'Estimate FPS
    Dim estimatedFPS As Double
    If (m_TimerHits > 0) Then
        m_NetFrameTime = m_NetFrameTime + (VBHacks.GetHighResTimeInMSEx() - m_lastFrameTime)
        estimatedFPS = m_NetFrameTime / m_TimerHits
        If (estimatedFPS > 0) Then estimatedFPS = (1000# / estimatedFPS)
    Else
        estimatedFPS = m_FPS
    End If
    
    m_TimerHits = m_TimerHits + 1
    VBHacks.GetHighResTimeInMS m_lastFrameTime
    
    'Convert total recording time to usable minutes/seconds values
    Dim timeElapsedInMinutes As Currency
    timeElapsedInMinutes = Int(timeElapsedInSeconds / 60#)
    timeElapsedInSeconds = timeElapsedInSeconds - (timeElapsedInMinutes * 60)
    
    'Display recording speed and elapsed time to the user
    lblInfo.Caption = g_Language.TranslateMessage("Recording - %1:%2 @ %3 fps", Format$(timeElapsedInMinutes, "00"), Format$(timeElapsedInSeconds, "00"), Format$(estimatedFPS, "0.0"))
    
End Sub

'Want to emergency stop capture for whatever reason?  Call this function.
Private Sub StopTimer_Forcibly()
    
    If m_CaptureActive Then
        
        'Cancel any future timer events
        If (Not m_Timer Is Nothing) Then m_Timer.StopTimer
        Set m_Timer = Nothing
        
        'Cancel the PNG writer
        If (Not m_PNG Is Nothing) Then m_PNG.SaveAPNG_Streaming_Cancel
        Set m_PNG = Nothing
        
        'Cancel the WebP writer
        If (Not m_WebP Is Nothing) Then m_WebP.SaveStreamingWebP_Cancel
        Set m_WebP = Nothing
        
        'Erase the destination file (if one exists)
        If Files.FileExists(m_DstFilename) Then Files.FileDelete m_DstFilename
        m_DstFilename = vbNullString
        
        'Deactivate capture mode
        m_CaptureActive = False
        
    End If
    
End Sub

'This dialog uses a lazy window repainter; forced repaints simply invalidate the window
Private Sub ForceWindowRepaint()
    Const RDW_INVALIDATE As Long = &H1
    RedrawWindow Me.hWnd, 0&, 0&, RDW_INVALIDATE
End Sub
