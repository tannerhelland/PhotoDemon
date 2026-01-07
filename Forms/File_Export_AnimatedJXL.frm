VERSION 5.00
Begin VB.Form dialog_ExportAnimatedJXL 
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
   Icon            =   "File_Export_AnimatedJXL.frx":0000
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
   Begin PhotoDemon.pdSlider sldQuality 
      Height          =   735
      Left            =   6600
      TabIndex        =   8
      Top             =   4440
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1296
      Caption         =   "quality"
      FontSizeCaption =   10
      Max             =   100
      Value           =   90
      GradientColorRight=   1703935
      NotchPosition   =   2
      NotchValueCustom=   75
      DefaultValue    =   90
   End
   Begin PhotoDemon.pdSlider sldEffort 
      Height          =   735
      Left            =   6600
      TabIndex        =   9
      Top             =   5280
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1296
      Caption         =   "effort"
      FontSizeCaption =   10
      Min             =   1
      Max             =   9
      Value           =   7
      GradientColorRight=   1703935
      NotchPosition   =   2
      NotchValueCustom=   75
      DefaultValue    =   7
   End
End
Attribute VB_Name = "dialog_ExportAnimatedJXL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Animated JPEG XL export dialog
'Copyright 2019-2026 by Tanner Helland
'Created: 26/August/19
'Last updated: 13/October/23
'Last update: spin off JXL-specific dialog (using WebP as the base)
'
'In v9.2, PhotoDemon gained the ability to export animated JPEG XL files.  This dialog exposes relevant
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

'Final metadata XML packet, with all metadata settings defined as tag+value pairs.  May be unused for
' soem export dialogs.
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

'Because JPEG XL supports lossy compression, we need to estimate compressed frame results on-the-fly.
' This is challenging to do quickly and every optimization helps, including maintaining a persistent frame copy.
Private m_tmpAnimationFrame As pdDIB

'After many (failed) attempts to work with libjxl directly, stability issues prompted me to move everything
' out-of-process for this particular format.  This has unfortunate performance knock-on effects, which means
' we need to go to some extra lengths to make this dialog work okay-ish.
'
'As new animation frame previews are generated, this struct stores the results.  We only generate new frames
' as necessary.  This allows a running animation preview to slowly become more and more accurate as the user
' allows it to run.
'
'Because we cache temporary PNG files (used as the input to libjxl), we only have to produce new JXL output
' on settings changes.  However, all temp files must obviously be deleted when the dialog closes.
Private Type PD_JXLFramePreview
    pathToPNG As String
    settingsForPreviews As String
    jxlPreviewDIB As pdDIB
End Type

Private m_numFramePreviews As Long, m_jxlFrames() As PD_JXLFramePreview

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
    
    'Next, prepare various controls on the metadata panel (TODO pending ExifTool testing on JXL support)
    'mtdManager.SetParentImage m_SrcImage, PDIF_JXL
    
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
    m_FormatParamString = GetExportParamString()
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
    RenderAnimationFrame
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
    
    'libjxl-specific values follow
    sldQuality.Value = 90   'Default per libjxl
    sldEffort.Value = 7     'Default per libjxl
    
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
    
    If (m_numFramePreviews > 0) Then
        
        Dim i As Long
        For i = 0 To UBound(m_jxlFrames)
            If (LenB(m_jxlFrames(i).pathToPNG) <> 0) Then Files.FileDeleteIfExists m_jxlFrames(i).pathToPNG
        Next i
        
        m_numFramePreviews = 0
    
    End If
    
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
    
    'JPEG XL-specific settings follow
    cParams.AddParam "jxl-lossy-quality", sldQuality.Value
    cParams.AddParam "jxl-effort", sldEffort.Value
    
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
    
    'Prepare our temporary animation buffers; we don't use them here, but it makes sense to initialize it
    ' to the required size now (so the frame renderer doesn't have to worry about size validation)
    If (m_AniFrame Is Nothing) Then Set m_AniFrame = New pdDIB
    m_AniFrame.CreateBlank thumbSize, thumbSize, 32, 0, 0
    m_AniFrame.SetInitialAlphaPremultiplicationState True
    
    If (m_tmpAnimationFrame Is Nothing) Then Set m_tmpAnimationFrame = New pdDIB
    m_tmpAnimationFrame.CreateBlank thumbSize, thumbSize, 32, 0, 0
    m_tmpAnimationFrame.SetInitialAlphaPremultiplicationState True
    
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
    If (m_tmpAnimationFrame Is Nothing) Then Exit Sub
    If (m_FrameCount <= 0) Or (m_SrcImage Is Nothing) Then Exit Sub
    
    'Ensure our frame preview collection exists
    If (m_numFramePreviews <> m_FrameCount) Then
        m_numFramePreviews = m_FrameCount
        ReDim m_jxlFrames(0 To m_numFramePreviews - 1) As PD_JXLFramePreview
    End If
    
    Dim idxFrame As Long
    idxFrame = m_Timer.GetCurrentFrame()
    
    'Make sure the frame request is valid; if it isn't, exit immediately
    If (idxFrame >= 0) And (idxFrame < m_FrameCount) Then
        
        'Request the back buffer DC, and ask the support module to erase any existing rendering for us.
        m_AniFrame.ResetDIB 0
        
        'Paint a checkerboard background only over the relevant image region, followed by the frame itself
        With m_Frames(idxFrame)
            
            GDI_Plus.GDIPlusFillDIBRect_Pattern m_AniFrame, m_AniThumbBounds.Left, m_AniThumbBounds.Top, m_AniThumbBounds.Width, m_AniThumbBounds.Height, g_CheckerboardPattern, , False, True
            
            'Make sure we have the necessary image in the spritesheet cache, then paint the sprite to
            ' a *temporary image* (not the underlying frame DIB)
            If m_Thumbs.DoesImageExist(Str$(idxFrame) & "|" & Str$(.afWidth)) Then
                m_tmpAnimationFrame.ResetDIB 0
                m_Thumbs.PaintCachedImage m_tmpAnimationFrame.GetDIBDC, 0, 0, m_Frames(idxFrame).afThumbKey
            End If
            
        End With
        
        'Unique to JPEG XL is the need to generate lossy animation frames "on-the-fly".  There is (currently)
        ' no way to do this efficiently, so we must pay the full penalty of round-tripping each frame through
        ' a full pdDIB > PNG > JXL > PNG > pdDIB iteration (sigh).
        
        'Generate a specially crafted settings string
        Dim cParams As pdSerialize
        Set cParams = New pdSerialize
        cParams.AddParam "jxl-lossy-quality", sldQuality.Value, True, True
        cParams.AddParam "jxl-effort", 1&, True, True       'Always use minimal (fastest) effort for animation previews
        
        'See if the previous preview string is different (different JPEG-XL settings) or if the
        ' preview dimensions have changes (user resized window)
        Dim changeInSettings As Boolean, changeInSize As Boolean
        changeInSettings = (m_jxlFrames(idxFrame).settingsForPreviews <> cParams.GetParamString())
        
        If (Not m_jxlFrames(idxFrame).jxlPreviewDIB Is Nothing) Then
            changeInSize = (m_jxlFrames(idxFrame).jxlPreviewDIB.GetDIBWidth <> m_tmpAnimationFrame.GetDIBWidth) Or (m_jxlFrames(idxFrame).jxlPreviewDIB.GetDIBHeight <> m_tmpAnimationFrame.GetDIBHeight)
        Else
            changeInSize = True
        End If
        
        'If lossy compression is being used, generate a preview of the compression.  (Note that JPEG XL treats
        ' a compression value of 90 as "visually lossless", which is why we don't bother with lossy compression
        ' until quality drops *below* 90.)
        If (sldQuality.Value < 90) Then
            
            changeInSize = changeInSize Or (LenB(m_jxlFrames(idxFrame).pathToPNG) = 0)
            
            'Generate a temporary PNG for this frame, as necessary
            If changeInSize Then
                m_jxlFrames(idxFrame).pathToPNG = OS.UniqueTempFilename(customExtension:="png")
                Saving.QuickSaveDIBAsPNG m_jxlFrames(idxFrame).pathToPNG, m_tmpAnimationFrame, dontCompress:=True
            End If
            
            'Apply the compression as necessary (this may take awhile)
            If (changeInSize Or changeInSettings) Then
            
                Plugin_jxl.PreviewSingleFrameAsJXL m_jxlFrames(idxFrame).pathToPNG, m_tmpAnimationFrame, cParams.GetParamString()
                
                'Cache a local copy of the animation frame (so we don't have to generate it again unless settings change)
                If (m_jxlFrames(idxFrame).jxlPreviewDIB Is Nothing) Then Set m_jxlFrames(idxFrame).jxlPreviewDIB = New pdDIB
                m_jxlFrames(idxFrame).jxlPreviewDIB.CreateFromExistingDIB m_tmpAnimationFrame
                
                'Same for settings (this is how we detect when we need to generate a new preview)
                m_jxlFrames(idxFrame).settingsForPreviews = cParams.GetParamString()
                
            End If
            
        Else
            
            If (changeInSize Or changeInSettings) Then
                
                'If the preview string has changed, copy over the current frame as-is
                With m_jxlFrames(idxFrame)
                    Set .jxlPreviewDIB = New pdDIB
                    .jxlPreviewDIB.CreateFromExistingDIB m_tmpAnimationFrame
                    .settingsForPreviews = cParams.GetParamString()
                End With
                
            End If
            
        End If
        
        'Normally we would paint the previewed compression frame (in m_tmpAnimationFrame)
        ' onto the underlying frame object here, but for this format, we instead use a persistent
        ' collection of generated frames because they are so expensive to generate.
        m_jxlFrames(idxFrame).jxlPreviewDIB.AlphaBlendToDC m_AniFrame.GetDIBDC
        
        'Resources can be problematic on huge animations
        m_jxlFrames(idxFrame).jxlPreviewDIB.SuspendDIB autoKeepIfLarge:=False
        
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
    
    sldQuality.SetTop yOffset
    yOffset = yOffset + sldQuality.GetHeight + yPadding
    
    sldEffort.SetTop yOffset
    yOffset = yOffset + sldEffort.GetHeight + yPadding
    
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

Private Sub sldQuality_Change()
    RenderAnimationFrame
End Sub
