VERSION 5.00
Begin VB.Form dialog_ExportAnimatedPNG 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Animation options"
   ClientHeight    =   7230
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12060
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
   Icon            =   "File_Export_AnimatedPNG.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   482
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   804
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   375
      Left            =   6240
      Top             =   3960
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   661
      Caption         =   "advanced settings"
      FontSize        =   12
   End
   Begin PhotoDemon.pdButtonStrip btsLoop 
      Height          =   975
      Left            =   6240
      TabIndex        =   5
      Top             =   120
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   1720
      Caption         =   "repeat"
   End
   Begin PhotoDemon.pdButtonToolbox btnPlay 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   6000
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      StickyToggle    =   -1  'True
   End
   Begin PhotoDemon.pdSliderStandalone sldFrame 
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   6000
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   661
   End
   Begin PhotoDemon.pdPictureBox picPreview 
      Height          =   5775
      Left            =   120
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   10186
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   6480
      Width           =   12060
      _ExtentX        =   21273
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdButtonToolbox btnPlay 
      Height          =   375
      Index           =   1
      Left            =   5640
      TabIndex        =   3
      Top             =   6000
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      StickyToggle    =   -1  'True
   End
   Begin PhotoDemon.pdSlider sldLoop 
      Height          =   735
      Left            =   6600
      TabIndex        =   4
      Top             =   1200
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1296
      Caption         =   "repeat count"
      FontSizeCaption =   10
      Min             =   1
      Max             =   65535
      ScaleStyle      =   2
      Value           =   1
      DefaultValue    =   1
   End
   Begin PhotoDemon.pdButtonStrip btsFrameTimes 
      Height          =   975
      Left            =   6240
      TabIndex        =   6
      Top             =   2040
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   1720
      Caption         =   "animation speed"
   End
   Begin PhotoDemon.pdSlider sldFrameTime 
      Height          =   735
      Left            =   6600
      TabIndex        =   7
      Top             =   3120
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1296
      FontSizeCaption =   10
      Max             =   100000
      ScaleStyle      =   1
      ScaleExponent   =   4
      Value           =   100
      GradientColorRight=   1703935
      NotchPosition   =   2
      NotchValueCustom=   100
   End
   Begin PhotoDemon.pdSlider sldCompression 
      Height          =   735
      Left            =   6600
      TabIndex        =   8
      Top             =   4440
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1296
      Caption         =   "compression level"
      FontSizeCaption =   10
      Max             =   12
      Value           =   9
      GradientColorRight=   1703935
      NotchPosition   =   2
      NotchValueCustom=   9
   End
End
Attribute VB_Name = "dialog_ExportAnimatedPNG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Animated PNG export dialog
'Copyright 2012-2026 by Tanner Helland
'Created: 26/August/19
'Last updated: 28/October/21
'Last update: prep dialog for compatibility with batch processor
'
'In v8.0, PhotoDemon gained the ability to export animated PNG files.  This dialog exposes relevant
' export parameters to the user.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'This form can (and should!) be notified of the image being exported.  The only exception to this rule is invoking
' the dialog from the batch process dialog, as no image is associated with that preview.
Private m_SrcImage As pdImage

'OK or CANCEL result
Private m_UserDialogAnswer As VbMsgBoxResult

'Final format-specific XML packet, with all format-specific settings defined as tag+value pairs
Private m_FormatParamString As String

'Final metadata XML packet, with all metadata settings defined as tag+value pairs.  Currently unused as ExifTool
' cannot write any BMP-specific data.
Private m_MetadataParamString As String

'To avoid circular updates on animation state changes, we use this tracker
Private m_DoNotUpdate As Boolean

'A dedicated animation timer is used; it auto-corrects for frame time variations during rendering
Private WithEvents m_Timer As pdTimerAnimation
Attribute m_Timer.VB_VarHelpID = -1

Private m_Thumbs As pdSpriteSheet
Private m_Frames() As PD_AnimationFrame
Private m_FrameCount As Long
Private m_AniThumbBounds As RectF
Private m_FrameTimesUndefined As Boolean

'Animation updates are rendered to a temporary DIB, which is then forwarded to the preview window
Private m_AniFrame As pdDIB

'The user's answer is returned via this property
Public Function GetDialogResult() As VbMsgBoxResult
    GetDialogResult = m_UserDialogAnswer
End Function

Public Function GetFormatParams() As String
    GetFormatParams = m_FormatParamString
End Function

Public Function GetMetadataParams() As String
    GetMetadataParams = m_MetadataParamString
End Function

'The ShowDialog routine presents the user with this form.
Public Sub ShowDialog(Optional ByRef srcImage As pdImage = Nothing)

    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    m_UserDialogAnswer = vbCancel
    
    Message "Waiting for user to specify export options... "
    
    'Prep any UI elements
    btsLoop.AddItem "none", 0
    btsLoop.AddItem "forever", 1
    btsLoop.AddItem "custom", 2
    btsLoop.ListIndex = 0
    
    btsFrameTimes.AddItem "fixed", 0
    btsFrameTimes.AddItem "pull from layer names", 1
    btsFrameTimes.ListIndex = 0
    
    'Prep a preview (if any)
    Set m_SrcImage = srcImage
    If (Not m_SrcImage Is Nothing) Then
        
        'Get loop behavior
        SyncLoopButton m_SrcImage.ImgStorage.GetEntry_Long("animation-loop-count", 1)
        If (btsLoop.ListIndex = 1) Then btnPlay(1).Value = True
        m_Timer.SetRepeat btnPlay(1).Value
        
        'Update animation frames (so the user can preview them!)
        UpdateAnimationSettings
        
    End If
    
    'Next, prepare various controls on the metadata panel
    'mtdManager.SetParentImage m_SrcImage, PDIF_PNG
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True, picPreview.hWnd
    UpdateAgainstCurrentTheme
    
    'With theming handled, reflow the interface one final time before displaying the window
    ReflowInterface
    
    'Render the first frame of the animation
    RenderAnimationFrame
    
    'Display the dialog
    ShowPDDialog vbModal, Me, True

End Sub

Private Sub btnPlay_Click(Index As Integer, ByVal Shift As ShiftConstants)

    'Failsafe check for batch process mode (which won't supply a source image)
    If (m_FrameCount <= 0) Or (m_SrcImage Is Nothing) Then Exit Sub
    
    Select Case Index
    
        'Play/pause
        Case 0
            
            If btnPlay(Index).Value Then
                
                'Reset the current animation frame, as necessary
                If (m_Timer.GetCurrentFrame() >= m_FrameCount - 1) Then m_Timer.SetCurrentFrame 0
                
                'Relay animation settings to the animation timer
                RelayAnimationSettings
                
                'The animation timer handles the rest!
                m_Timer.StartTimer
                
            Else
                m_Timer.StopTimer
            End If
                
        '1x/repeat
        Case 1
            m_Timer.SetRepeat btnPlay(Index).Value
    
    End Select

End Sub

Private Sub btsFrameTimes_Click(ByVal buttonIndex As Long)
    NotifyNewFrameTimes
    ReflowInterface
End Sub

Private Sub btsLoop_Click(ByVal buttonIndex As Long)
    ReflowInterface
End Sub

Private Sub cmdBar_CancelClick()
    m_Timer.StopTimer
    m_UserDialogAnswer = vbCancel
    Me.Visible = False
End Sub

Private Sub cmdBar_OKClick()
    m_Timer.StopTimer
    m_FormatParamString = GetExportParamString
    'm_MetadataParamString = mtdManager.GetMetadataSettings
    m_UserDialogAnswer = vbOK
    Me.Visible = False
End Sub

Private Sub cmdBar_ReadCustomPresetData()

    'If a source image exists, preferentially use its animation loop setting
    If (Not m_SrcImage Is Nothing) Then
        If m_SrcImage.ImgStorage.DoesKeyExist("animation-loop-count") Then SyncLoopButton m_SrcImage.ImgStorage.GetEntry_Long("animation-loop-count", 1)
    End If
    
    'If all frames have undefined frame times (e.g. none embedded a frame time in the layer name),
    ' default to a "fixed" frame time suggestion
    If m_FrameTimesUndefined And (Not m_SrcImage Is Nothing) Then btsFrameTimes.ListIndex = 0 Else btsFrameTimes.ListIndex = 1
    
End Sub

'Not required at present; may change if the exporter gains lossy optimization options
Private Sub cmdBar_RequestPreviewUpdate()
    'UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    
    'If a source image exists, synchronize output settings to match whatever its original animation
    ' settings were (if any)
    If (Not m_SrcImage Is Nothing) Then
        SyncLoopButton m_SrcImage.ImgStorage.GetEntry_Long("animation-loop-count", 1)
        
    'Otherwise, default to reasonable values
    Else
        SyncLoopButton 1
    End If
    
    'If all frames have undefined frame times (e.g. none embedded a frame time in the layer name),
    ' default to a "fixed" frame time suggestion
    If m_FrameTimesUndefined And (Not m_SrcImage Is Nothing) Then btsFrameTimes.ListIndex = 0 Else btsFrameTimes.ListIndex = 1
    
End Sub

Private Sub Form_Load()
    
    'Make sure our animation objects exist
    Set m_Thumbs = New pdSpriteSheet
    Set m_Timer = New pdTimerAnimation
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If (Not m_Timer Is Nothing) Then m_Timer.StopTimer
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_Timer = Nothing
    ReleaseFormTheming Me
End Sub

Private Sub SyncLoopButton(ByVal loopAmount As Long)
    If (loopAmount = 0) Then
        btsLoop.ListIndex = 1
    ElseIf (loopAmount >= 2) Then
        btsLoop.ListIndex = 2
        sldLoop.Value = loopAmount
    Else
        btsLoop.ListIndex = 0
    End If
    ReflowInterface
End Sub

Private Function GetExportParamString() As String

    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    'The loop setting is a little weird.  0 = loop infinitely, 1 = loop once, 2+ = loop that many times exactly
    If (btsLoop.ListIndex = 0) Then
        cParams.AddParam "animation-loop-count", 1
    ElseIf (btsLoop.ListIndex = 1) Then
        cParams.AddParam "animation-loop-count", 0
    Else
        cParams.AddParam "animation-loop-count", CLng(sldLoop.Value + 1)
    End If
    
    cParams.AddParam "use-fixed-frame-delay", (btsFrameTimes.ListIndex = 0)
    cParams.AddParam "frame-delay-default", sldFrameTime.Value
    cParams.AddParam "compression-level", sldCompression.Value
    
    GetExportParamString = cParams.GetParamString
    
End Function

'Load button icons and other various UI bits
Private Sub UpdateAgainstCurrentTheme()
    
    'Play and pause icons are generated at run-time, using the current UI accent color
    Dim btnIconSize As Long
    btnIconSize = btnPlay(0).GetWidth - Interface.FixDPI(4)
    
    Dim icoPlay As pdDIB
    Set icoPlay = Interface.GetRuntimeUIDIB(pdri_Play, btnIconSize)
    
    Dim icoPause As pdDIB
    Set icoPause = Interface.GetRuntimeUIDIB(pdri_Pause, btnIconSize)
    
    'Assign the icons
    btnPlay(0).AssignImage vbNullString, icoPlay
    btnPlay(0).AssignImage_Pressed vbNullString, icoPause
    
    'The 1x/repeat icons use prerendered graphics
    btnIconSize = btnIconSize - 4
    Dim tmpDIB As pdDIB
    If g_Resources.LoadImageResource("1x", tmpDIB, btnIconSize, btnIconSize, , False, g_Themer.GetGenericUIColor(UI_Accent)) Then btnPlay(1).AssignImage vbNullString, tmpDIB
    If g_Resources.LoadImageResource("infinity", tmpDIB, btnIconSize, btnIconSize, , False, g_Themer.GetGenericUIColor(UI_Accent)) Then btnPlay(1).AssignImage_Pressed vbNullString, tmpDIB
    
    'Add a special note to this particular 1x/repeat button, pointing out that it does
    ' *not* rely on the looping setting to the right.  (I have mixed feelings about the
    ' intuitiveness of this, but I feel like there needs to be *some* way to preview the
    ' animation as a loop without actually committing to it... idk, I may revisit.)
    Dim tText As String
    tText = g_Language.TranslateMessage("This button only affects the preview above.  The repeat setting on the right is what will be used by the exported image file.")
    btnPlay(1).AssignTooltip tText, UserControls.GetCommonTranslation(pduct_AnimationRepeatToggle)
    
End Sub

Private Sub m_Timer_DrawFrame(ByVal idxFrame As Long)
    
    If (m_FrameCount <= 0) Or (m_SrcImage Is Nothing) Then Exit Sub
    
    'Render the current frame
    RenderAnimationFrame
    
    'Synchronize the scrubber
    m_DoNotUpdate = True
    sldFrame.Value = idxFrame
    m_DoNotUpdate = False
    
End Sub

'Call at dialog initiation to produce a collection of animation thumbnails (and associated metadata,
' like frame delay times)
Private Sub UpdateAnimationSettings()
    
    If (m_SrcImage Is Nothing) Then Exit Sub

    'Suspend automatic control-based updates while we get everything synchronized
    m_DoNotUpdate = True
    
    'Load all animation frames.
    m_FrameCount = m_SrcImage.GetNumOfLayers
    If (m_FrameCount <= 0) Then Exit Sub
    ReDim m_Frames(0 To m_FrameCount - 1) As PD_AnimationFrame
    
    m_Thumbs.ResetCache
    m_Timer.NotifyFrameCount m_FrameCount
    
    sldFrame.Max = m_FrameCount - 1
    
    'In animation files, we currently assume all frames are the same size as the image itself,
    ' because this is how PD pre-processes them.  (This may change in the future.)
    Dim bWidth As Long, bHeight As Long
    bWidth = picPreview.GetWidth - 2
    bHeight = picPreview.GetHeight - 2
    
    'Figure out what size to use for the animation thumbnails
    Dim thumbSize As Long
    Dim thumbImageWidth As Long, thumbImageHeight As Long
    PDMath.ConvertAspectRatio m_SrcImage.Width, m_SrcImage.Height, bWidth, bHeight, thumbImageWidth, thumbImageHeight
    
    'Ensure the thumb isn't larger than the actual image
    If (thumbImageWidth > m_SrcImage.Width) Or (thumbImageHeight > m_SrcImage.Height) Then
        thumbImageWidth = m_SrcImage.Width
        thumbImageHeight = m_SrcImage.Height
    End If
    
    'Use the larger dimension to construct the thumb.  (For simplicity, thumbs are always square.)
    If (thumbImageWidth > thumbImageHeight) Then thumbSize = thumbImageWidth Else thumbSize = thumbImageHeight
    
    'Prepare our temporary animation buffer; we don't use it here, but it makes sense to initialize it
    ' to the required size now
    If (m_AniFrame Is Nothing) Then Set m_AniFrame = New pdDIB
    m_AniFrame.CreateBlank thumbSize, thumbSize, 32, 0, 0
    
    'Store the boundary rect of where the thumb will actually appear; we need this for rendering
    ' a transparency checkerboard
    With m_AniThumbBounds
        .Left = Int((thumbSize - thumbImageWidth) * 0.5 + 0.5)
        .Top = Int((thumbSize - thumbImageHeight) * 0.5 + 0.5)
        .Width = thumbImageWidth
        .Height = thumbImageHeight
    End With
    
    Dim numZeroFrameDelays As Long
    
    'Load all thumbnails
    Dim i As Long, tmpDIB As pdDIB
    For i = 0 To m_FrameCount - 1
        
        'Retrieve an updated thumbnail
        If (tmpDIB Is Nothing) Then Set tmpDIB = New pdDIB
        tmpDIB.CreateBlank thumbSize, thumbSize, 32, 0, 0
        
        m_Frames(i).afWidth = thumbSize
        m_Frames(i).afHeight = thumbSize
        
        m_SrcImage.GetLayerByIndex(i).RequestThumbnail_ImageCoords tmpDIB, m_SrcImage, thumbSize, False, VarPtr(m_AniThumbBounds)
        m_Frames(i).afThumbKey = m_Thumbs.AddImage(tmpDIB, Str$(i) & "|" & Str$(thumbSize))
        
        'Retrieve layer frame times and relay them to the animation object
        m_Frames(i).afFrameDelayMS = Animation.GetFrameTimeFromLayerName(m_SrcImage.GetLayerByIndex(i).GetLayerName(), 0)
        If (m_Frames(i).afFrameDelayMS = 0) Then numZeroFrameDelays = numZeroFrameDelays + 1
        
    Next i
    
    'If one or more valid frame time amounts were discovered, default to "pull frame times from
    ' layer names" - otherwise, default to a fixed delay for *all* frames.
    m_FrameTimesUndefined = (numZeroFrameDelays = m_FrameCount)
    
    'Relay frame times to the animator
    NotifyNewFrameTimes
    
    m_DoNotUpdate = False
    
    'Render the first frame of the animation
    RenderAnimationFrame
    
End Sub

Private Sub RelayAnimationSettings()
    m_Timer.NotifyFrameCount m_FrameCount
    NotifyNewFrameTimes
End Sub

'Render the current animation frame
Private Sub RenderAnimationFrame()
    
    If m_DoNotUpdate Then Exit Sub
    If (m_AniFrame Is Nothing) Then Exit Sub
    If (m_FrameCount <= 0) Or (m_SrcImage Is Nothing) Then Exit Sub
    
    Dim idxFrame As Long
    idxFrame = m_Timer.GetCurrentFrame()
    
    'Make sure the frame request is valid; if it isn't, exit immediately
    If (idxFrame >= 0) And (idxFrame < m_FrameCount) Then
        
        'Request the back buffer DC, and ask the support module to erase any existing rendering for us.
        m_AniFrame.ResetDIB 0
        
        'Paint a checkerboard background only over the relevant image region, followed by the frame itself
        With m_Frames(idxFrame)
            
            GDI_Plus.GDIPlusFillDIBRect_Pattern m_AniFrame, m_AniThumbBounds.Left, m_AniThumbBounds.Top, m_AniThumbBounds.Width, m_AniThumbBounds.Height, g_CheckerboardPattern, , False, True
            
            'Make sure we have the necessary image in the spritesheet cache
            If m_Thumbs.DoesImageExist(Str$(idxFrame) & "|" & Str$(.afWidth)) Then
                m_Thumbs.PaintCachedImage m_AniFrame.GetDIBDC, 0, 0, m_Frames(idxFrame).afThumbKey
            End If
            
        End With
        
        'Paint the final result to the screen, as relevant
        picPreview.CopyDIB m_AniFrame, False, True, True, True
        
    'If our frame counter is invalid, end all animations
    Else
        m_Timer.StopTimer
    End If
        
End Sub

Private Sub m_Timer_EndOfAnimation()

    m_DoNotUpdate = True
    If btnPlay(0).Value Then btnPlay(0).Value = False
    sldFrame.Value = m_Timer.GetCurrentFrame()
    m_DoNotUpdate = False
    
End Sub

Private Sub picPreview_WindowResizeDetected()
    ReflowInterface
    UpdateAnimationSettings
End Sub

Private Sub sldFrame_Change()
    If (Not m_DoNotUpdate) Then
        m_Timer.StopTimer
        m_Timer.SetCurrentFrame sldFrame.Value
    End If
End Sub

Private Sub ReflowInterface()
    
    Dim yPadding As Long, yPaddingTitle As Long
    yPadding = Interface.FixDPI(8)
    yPaddingTitle = Interface.FixDPI(12)
    
    Dim yOffset As Long
    yOffset = btsLoop.GetTop + btsLoop.GetHeight + yPadding
    
    sldLoop.Visible = (btsLoop.ListIndex = 2)
    If sldLoop.Visible Then
        sldLoop.SetTop yOffset
        yOffset = yOffset + sldLoop.GetHeight + yPaddingTitle
    Else
        yOffset = yOffset - yPadding + yPaddingTitle
    End If
    
    btsFrameTimes.SetTop yOffset
    yOffset = yOffset + btsFrameTimes.GetHeight + yPadding
    
    If (btsFrameTimes.ListIndex = 0) Then
        sldFrameTime.Caption = g_Language.TranslateMessage("frame time (in ms)")
    ElseIf (btsFrameTimes.ListIndex = 1) Then
        sldFrameTime.Caption = g_Language.TranslateMessage("frame time for undefined layers (in ms)")
    End If
    
    sldFrameTime.SetTop yOffset
    yOffset = yOffset + sldFrameTime.GetHeight + yPaddingTitle
    
    lblTitle.SetTop yOffset
    yOffset = yOffset + lblTitle.GetHeight + yPadding
    sldCompression.SetTop yOffset
    yOffset = yOffset + sldCompression.GetHeight + yPadding
    
End Sub

Private Sub sldFrameTime_Change()
    NotifyNewFrameTimes
End Sub

Private Sub NotifyNewFrameTimes()
    
    Dim useFixedTime As Boolean, fixedTimeMS As Long
    useFixedTime = (btsFrameTimes.ListIndex = 0)
    fixedTimeMS = sldFrameTime.Value
    
    Dim i As Long
    For i = 0 To m_FrameCount - 1
        If (m_Frames(i).afFrameDelayMS = 0) Or useFixedTime Then
            m_Timer.NotifyFrameTime fixedTimeMS, i
        Else
            m_Timer.NotifyFrameTime m_Frames(i).afFrameDelayMS, i
        End If
    Next i
    
End Sub
