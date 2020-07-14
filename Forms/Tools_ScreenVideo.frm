VERSION 5.00
Begin VB.Form FormRecordAPNG 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   " Animated screen capture"
   ClientHeight    =   6765
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8565
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
   ScaleHeight     =   451
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   571
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdButton cmdExit 
      Height          =   495
      Left            =   7200
      TabIndex        =   1
      Top             =   6120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "Cancel"
   End
   Begin PhotoDemon.pdButton cmdStart 
      Height          =   495
      Left            =   4560
      TabIndex        =   0
      Top             =   6120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      Caption         =   "Start"
   End
End
Attribute VB_Name = "FormRecordAPNG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Animated screen capture dialog
'Copyright 2020-2020 by Tanner Helland
'Created: 01/July/20
'Last updated: 11/July/20
'Last update: get frame optimizations up and running
'
'PD can write animated PNGs.  APNGs seem like a great fit for animated screen captures.
' Let's see if we can merge the two, eh?
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'The current window must be converted to a layered window to make regions of it transparent
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RectL) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal wFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Const KEY_COLOR As Long = &HFF00FF
Private m_BackBuffer As pdDIB

'A timer triggers capture events
Private WithEvents m_Timer As pdTimer
Attribute m_Timer.VB_VarHelpID = -1

'A pdPNG instance handles the actual PNG writing
Private m_PNG As pdPNG

'Destination file, if one is selected (check for null before using)
Private m_DstFilename As String

'Target maximum frame rate (as frames-per-second)
Private m_FPS As Double

'Capture rects; once populated (at the start of the capture), these *cannot* be changed
Private m_CaptureRectClient As RectL, m_CaptureRectScreen As RectL

'If a capture event is ACTIVE, this will be set to TRUE
Private m_CaptureActive As Boolean

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

'Captured frames are stored as a collection of lz4-compressed arrays.  I don't currently have
' access to a DEFLATE library that can compress screen-capture-sized frames (e.g. 1024x768) in
' real-time on an XP-era PC.  lz4 is a better solution here, although it requires us to
' "play back" the frames when the capture ends.
'
'(This type is now declared publicly, so that we can pass the data directly to the APNG encoder.)
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

'UI elements on this dialog are automatically reflowed as the window is moved/resized.
' We track their positioning rects here; these values are populated at Form_Load().
Private Type UI_Data
    hWnd As Long
    ptTopLeft As PointAPI
End Type

Private m_numUIElements As Long, m_UIElements() As UI_Data

'Resize and paint events are handled via API, not VB; this helps us support high-DPI displays
Private WithEvents m_Resize As pdWindowSize
Attribute m_Resize.VB_VarHelpID = -1
Private WithEvents m_Painter As pdWindowPainter
Attribute m_Painter.VB_VarHelpID = -1

'This dialog must be invoked via this function.  It preps a bunch of internal values that must exist
' for the recorder to function.
Public Sub ShowDialog(ByRef dstFilename As String, ByVal dstFrameRateFPS As Double)
    
    'Cache all passed values
    m_DstFilename = dstFilename
    m_FPS = dstFrameRateFPS
    
    'Prepare the dialog (inc. setting up window transparency)
    PrepWindowForRecording
    
    'Display this dialog as a MODELESS window (critical for always-on-top behavior!)
    Me.Show vbModeless
    
End Sub

Private Sub PrepWindowForRecording()

    If PDMain.IsProgramRunning Then
        
        'Mark the underlying window as a layered window
        Const GWL_EXSTYLE As Long = -20
        Const WS_EX_LAYERED As Long = &H80000
        SetWindowLong Me.hWnd, GWL_EXSTYLE, GetWindowLong(Me.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED
        
        'Subclassers are used for resize and paint events
        Set m_Resize = New pdWindowSize
        m_Resize.AttachToHWnd Me.hWnd, True
        Set m_Painter = New pdWindowPainter
        m_Painter.StartPainter Me.hWnd, True
        
        'We'll momentarily build a collection of all UI elements on the form; we use this
        ' to anchor their position if/when the form is resized.
        Const MAX_NUM_UIELEMENTS As Long = 16
        ReDim m_UIElements(0 To MAX_NUM_UIELEMENTS - 1) As UI_Data
        
        'Retrieve the current client rect of our window
        Dim myRect As winRect
        g_WindowManager.GetClientWinRect Me.hWnd, myRect
        
        'Enumerate every control on the form and cache its position offsets relative to the
        ' *bottom* of this dialog.
        Dim eControl As Control, tmpRect As winRect
        For Each eControl In Me.Controls
            
            If (m_numUIElements > UBound(m_UIElements)) Then ReDim Preserve m_UIElements(0 To m_numUIElements * 2 - 1) As UI_Data
            
            'Get the top-left coordinate of each control in *screen* coordinates
            g_WindowManager.GetWindowRect_API eControl.hWnd, tmpRect
            m_UIElements(m_numUIElements).hWnd = eControl.hWnd
            m_UIElements(m_numUIElements).ptTopLeft.x = tmpRect.x1
            m_UIElements(m_numUIElements).ptTopLeft.y = tmpRect.y1
            
            'Convert that to *client* coordinates
            g_WindowManager.GetScreenToClient Me.hWnd, m_UIElements(m_numUIElements).ptTopLeft
            m_UIElements(m_numUIElements).ptTopLeft.y = m_UIElements(m_numUIElements).ptTopLeft.y - (myRect.y2 - myRect.y1)
            
            m_numUIElements = m_numUIElements + 1
            
        Next eControl
        
        'When applying theming, note that we request to paint our window manually;
        ' normally PD handles this centrally, but this window has special needs.
        Interface.ApplyThemeAndTranslations Me, False
        
        'Ask the system to let us paint at least once before the form is actually displayed
        ForceWindowRepaint
        
        'Mark this window as "always on-top"
        Const HWND_TOPMOST As Long = -1&
        Const SWP_NOMOVE As Long = &H2&
        Const SWP_NOSIZE As Long = &H1&
        g_WindowManager.SetWindowPos_API Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
        
    End If
    
End Sub

Private Sub cmdExit_Click()
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
        Set m_Timer = New pdTimer
        m_Timer.Interval = Int(1000# / m_FPS + 0.5)
        
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
        
        'Initialize the PNG writer
        Set m_PNG = New pdPNG
        If (m_PNG.SaveAPNG_Streaming_Start(m_DstFilename, m_captureDIB24.GetDIBWidth, m_captureDIB24.GetDIBHeight) < png_Failure) Then
            
            'Change the start button to a STOP button
            cmdStart.Caption = g_Language.TranslateMessage("Stop")
            
            'Start the capture timer
            m_CaptureActive = True
            m_Timer.StartTimer
        
        Else
            PDDebug.LogAction "WARNING!  APNG screen capture failed for unknown reason.  Consult debug log."
        End If
        
    End If
        
End Sub

'Stop the active capture
Private Sub Capture_Stop()

    If m_CaptureActive Then
    
        'Immediately stop the capture timer
        m_CaptureActive = False
        If (Not m_Timer Is Nothing) Then m_Timer.StopTimer
        
        'Now comes the fun part: loading all cached frames, and passing them off to the APNG writer
        ' so that it can produce a usable APNG file!
        Dim i As Long
        For i = 0 To m_FrameCount - 1
            
            g_WindowManager.SetWindowCaptionW Me.hWnd, g_Language.TranslateMessage("Processing frame %1 of %2", i + 1, m_FrameCount)
            VBHacks.DoEvents_SingleHwnd Me.hWnd
            
            'Extract this frame into the capture DIB, then immediately free its compressed memory
            Compression.DecompressPtrToPtr m_captureDIB24.GetDIBPointer, m_Frames(i).frameSizeOrig, VarPtr(m_Frames(i).frameData(0)), m_Frames(i).frameSizeCompressed, cf_Lz4
            Erase m_Frames(i).frameData
            
            'Convert the 24-bpp DIB to 32-bpp before handing it off to the APNG encoder
            If (m_captureDIB32 Is Nothing) Then
                Set m_captureDIB32 = New pdDIB
                m_captureDIB32.CreateBlank m_captureDIB24.GetDIBWidth, m_captureDIB24.GetDIBHeight, 32, 0, 255
                m_captureDIB32.SetInitialAlphaPremultiplicationState True
            End If
            
            GDI.BitBltWrapper m_captureDIB32.GetDIBDC, 0, 0, m_captureDIB32.GetDIBWidth, m_captureDIB32.GetDIBHeight, m_captureDIB24.GetDIBDC, 0, 0, vbSrcCopy
            m_captureDIB32.ForceNewAlpha 255
            
            'Pass the frame off to the PNG encoder
            m_PNG.SaveAPNG_Streaming_Frame m_captureDIB32, m_Frames(i).fcTimeStamp
            
        Next i
        
        'Notify the PNG encoder that the stream has ended
        If (Not m_PNG Is Nothing) Then m_PNG.SaveAPNG_Streaming_Stop 0
        
        'Reset this button's caption
        cmdStart.Caption = g_Language.TranslateMessage("Start")
        
    End If
        
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
        
        'Build the backing surface
        If (m_BackBuffer Is Nothing) Then Set m_BackBuffer = New pdDIB
        m_BackBuffer.CreateBlank myRectClient.x2 - myRectClient.x1, myRectClient.y2 - myRectClient.y1, 24, 0
        
        'Wrap a surface object around the newly created back buffer
        Dim cSurface As pd2DSurface
        Set cSurface = New pd2DSurface
        cSurface.WrapSurfaceAroundPDDIB m_BackBuffer
        cSurface.SetSurfaceAntialiasing P2_AA_None
        cSurface.SetSurfaceCompositing P2_CM_Overwrite
        
        'Paint the background with the default theme background color
        Dim cBrush As pd2DBrush
        Set cBrush = New pd2DBrush
        cBrush.SetBrushColor g_Themer.GetGenericUIColor(UI_Background)
        PD2D.FillRectangleI cSurface, cBrush, 0, 0, m_BackBuffer.GetDIBWidth, m_BackBuffer.GetDIBHeight
        
        'Padding is arbitrary; we want it large enough to create a clean border, but not so large
        ' that we waste precious screen real estate
        Dim borderPadding As Long
        borderPadding = 8
        
        'Populate the capture rect; this is the region of this window (client coords)
        ' that will be made transparent, e.g. mouse events will "click through" this region
        With m_CaptureRectClient
            
            'Pad accordingly
            .Left = myRectClient.x1 + borderPadding
            .Right = myRectClient.x2 - borderPadding
            
            'Same for top/bottom
            .Top = myRectClient.y1 + borderPadding
            .Bottom = myRectClient.y1 + cmdStart.GetTop - borderPadding
            
        End With
        
        'Paint the capture rect with the transparent key color
        cBrush.SetBrushColor KEY_COLOR
        PD2D.FillRectangleI_FromRectL cSurface, cBrush, m_CaptureRectClient
        
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
        
        'Paint the backbuffer onto the specified DC
        GDI.BitBltWrapper targetDC, 0, 0, m_BackBuffer.GetDIBWidth, m_BackBuffer.GetDIBHeight, m_BackBuffer.GetDIBDC, 0, 0, vbSrcCopy
        
        'Update window attributes to note the key color
        Const LWA_COLORKEY As Long = &O1&
        SetLayeredWindowAttributes Me.hWnd, KEY_COLOR, 255, LWA_COLORKEY
        
    End If

End Sub

'On timer events, capture the current frame by calling this function
Private Sub CaptureFrameNow()
    
    'Ensure we have an active APNG instance
    If (Not m_PNG Is Nothing) Then
        
        'Make sure we have room to store this frame
        If (m_FrameCount > UBound(m_Frames)) Then ReDim Preserve m_Frames(0 To m_FrameCount * 2 - 1) As PD_APNGFrameCapture
        
        '*Immediately* before capture, note the current time
        Dim capTime As Currency, testTime As Currency
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
        
        ScreenCapture.GetPartialDesktopAsDIB captureTarget, m_CaptureRectScreen, True
        
        'Capture can take a non-trivial amount of time on Vista+ due to compositor changes in DWM.
        ' To try and accurately mirror the moment that was actually captured, set the capture time
        ' to the halfway point between capture start and end.
        Dim capTime2 As Currency
        capTime2 = VBHacks.GetHighResTimeInMSEx()
        
        'Safety check for clock rollover
        If (capTime < capTime2) Then capTime = (capTime + capTime2) / 2
        
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
        
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'Forcibly stop the current capture (if any)
    StopTimer_Forcibly
    
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
    
    'Make sure all UI elements were initialized correctly
    If (m_numUIElements > 0) Then
    
        'Reposition all controls according to their original offset
        Dim myRect As winRect
        g_WindowManager.GetClientWinRect Me.hWnd, myRect
        
        'Various SetWindowPos constants
        Const SWP_NOACTIVATE As Long = &H10&
        Const SWP_NOSIZE As Long = &H1&
        Const SWP_NOZORDER As Long = &H4&

        Dim wFlags As Long
        wFlags = SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOZORDER
        
        'Enumerate every control on the form and cache its position offsets relative to the
        ' *bottom* of this dialog.
        Dim i As Long, uiRect As winRect
        For i = 0 To m_numUIElements - 1
            g_WindowManager.GetWindowRect_API m_UIElements(i).hWnd, uiRect
            g_WindowManager.SetWindowPos_API m_UIElements(i).hWnd, 0&, m_UIElements(i).ptTopLeft.x, myRect.y2 + m_UIElements(i).ptTopLeft.y, 0&, 0&, wFlags
        Next i
        
        'Force a repaint
        ForceWindowRepaint
        
    End If
    
End Sub

Private Sub m_Timer_Timer()
    If m_CaptureActive Then CaptureFrameNow
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
