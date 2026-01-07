VERSION 5.00
Begin VB.Form FormAnimSpeed 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   " Playback speed"
   ClientHeight    =   6780
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   452
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   804
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdSlider sldSpeed 
      Height          =   855
      Left            =   6240
      TabIndex        =   6
      Top             =   1560
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   1508
      Caption         =   "speed modifier"
      Min             =   -16
      Max             =   16
      SigDigits       =   1
      GradientColorRight=   1703935
   End
   Begin PhotoDemon.pdButtonToolbox btnPlay 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   5520
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      DontHighlightDownState=   -1  'True
      StickyToggle    =   -1  'True
   End
   Begin PhotoDemon.pdSliderStandalone sldFrame 
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   5520
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   661
   End
   Begin PhotoDemon.pdPictureBox picPreview 
      Height          =   5295
      Left            =   120
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   9340
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   6030
      Width           =   12060
      _ExtentX        =   21273
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdButtonToolbox btnPlay 
      Height          =   375
      Index           =   1
      Left            =   5640
      TabIndex        =   3
      Top             =   5520
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      StickyToggle    =   -1  'True
   End
   Begin PhotoDemon.pdDropDown ddWhichFrame 
      Height          =   855
      Index           =   0
      Left            =   6240
      TabIndex        =   4
      Top             =   2520
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   1508
      Caption         =   "first frame for effect"
   End
   Begin PhotoDemon.pdDropDown ddWhichFrame 
      Height          =   855
      Index           =   1
      Left            =   6240
      TabIndex        =   5
      Top             =   3480
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   1508
      Caption         =   "last frame for effect"
   End
End
Attribute VB_Name = "FormAnimSpeed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Effect > Animation > Playback speed
'Copyright 2019-2026 by Tanner Helland
'Created: 26/August/19
'Last updated: 20/April/22
'Last update: split off from main Image > Animation dialog as a dedicated "effect"
'
'In v9.0, I added much better coverage of animated image formats to PhotoDemon.  This includes
' a few basic "effects" to permanently modify things like playback speed.
'
'Note that unlike static effects, animated effects use very different code for preview vs final
' execution.  This is necessary because pre-computing the effect for all frames is very energy
' intensive, so instead, we generate the preview "on the fly" and use totally different code for
' the final effect.  I am open to ideas for improving this, but remember - effects need to be
' preview-able in real-time on 20-year-old XP PCs.  It's a challenge.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'To avoid circular updates on animation state changes, we use this tracker
Private m_DoNotUpdate As Boolean

'A dedicated animation timer is used; it auto-corrects for frame time variations during rendering
Private WithEvents m_Timer As pdTimerAnimation
Attribute m_Timer.VB_VarHelpID = -1

'Thumb width/height is fixed according to the size of the preview window
Private m_ThumbWidth As Long, m_ThumbHeight As Long

'Animation updates are rendered to a temporary DIB, which is then forwarded to the preview window
Private m_AniFrame As pdDIB

Private m_Thumbs As pdSpriteSheet
Private m_Frames() As PD_AnimationFrame
Private m_FrameCount As Long
Private m_AniThumbBounds As RectF

'Apply an arbitrary background layer to other layers
Public Sub ApplyNewPlaybackSpeed(ByVal effectParams As String)
    
    SetProgBarMax PDImages.GetActiveImage.GetNumOfLayers()
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    'Retrieve parameters
    Dim frameRateModifier As Single
    frameRateModifier = cParams.GetSingle("speed-modifier", sldSpeed.Value, True)
    
    Dim idxFirstFrame As Long, idxLastFrame As Long
    idxFirstFrame = cParams.GetLong("idx-first-frame", 0, True)
    idxLastFrame = cParams.GetLong("idx-last-frame", PDImages.GetActiveImage.GetNumOfLayers() - 1, True)
    
    Dim i As Long
    For i = 0 To PDImages.GetActiveImage.GetNumOfLayers - 1
        
        ProgressBars.SetProgBarVal i + 1
        
        'Skip layers that are not in the target range
        If (i >= idxFirstFrame) And (i <= idxLastFrame) Then
            
            With PDImages.GetActiveImage.GetLayerByIndex(i)
                
                'Pull existing frame rate (if any)
                Dim curFrameTime As Single
                curFrameTime = .GetLayerFrameTimeInMS
                
                'How to handle null frame rate?  There's no good way, honestly.
                ' (We probably need a setting for this in the UI but I don't want to deal with this right now.)
                ' Today, let's just default to 30 fps (33.3 ms delay)
                If (curFrameTime = 0!) Then curFrameTime = 33.3!
                
                'Modify frame rate
                curFrameTime = GetModifiedFrameTime(curFrameTime, i, frameRateModifier, idxFirstFrame, idxLastFrame)
                PDImages.GetActiveImage.GetLayerByIndex(i).SetLayerFrameTimeInMS Int(curFrameTime + 0.5!)
                
                'After setting frame time, we need to synchronize each layer's name to its frame time.
                If (Animation.GetFrameTimeFromLayerName(.GetLayerName, 0) <> .GetLayerFrameTimeInMS) Then
                    .SetLayerName Animation.UpdateFrameTimeInLayerName(.GetLayerName, .GetLayerFrameTimeInMS())
                End If
                
            End With
            
            'Notify the parent image of the change
            PDImages.GetActiveImage.NotifyImageChanged UNDO_Layer_VectorSafe, i
            
        End If
        
    Next i
    
    'Notify the parent image of the final change
    PDImages.GetActiveImage.NotifyImageChanged UNDO_Image_VectorSafe
    
    'Redraw the screen and finalize the effect
    toolbar_Layers.NotifyLayerChange
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
    Message "Finished."
    ProgressBars.SetProgBarVal 0
    ProgressBars.ReleaseProgressBar
    
End Sub

Private Sub btnPlay_Click(Index As Integer, ByVal Shift As ShiftConstants)

    Select Case Index
    
        'Play/pause
        Case 0
            
            'When playing an animation, perform a failsafe frame check in case the user previously
            ' paused on the last frame, but has since enabled infinite looping playback.
            If btnPlay(Index).Value Then
                If (m_Timer.GetCurrentFrame() >= m_FrameCount - 1) Then m_Timer.SetCurrentFrame 0
                m_Timer.StartTimer
            Else
                m_Timer.StopTimer
            End If
                
        '1x/repeat
        Case 1
            m_Timer.SetRepeat btnPlay(Index).Value
    
    End Select

End Sub

Private Sub cmdBar_CancelClick()
    m_Timer.StopTimer
End Sub

Private Sub cmdBar_OKClick()
    
    'Halt any animations
    m_Timer.StopTimer
    
    'Process changes
    Process "Animation playback speed", , GetLocalParamString(), UNDO_Image
    
End Sub

Private Sub ddWhichFrame_Click(Index As Integer)
    NotifyNewFrameTimes
End Sub

Private Sub Form_Load()
    
    'Make sure our animation objects exist
    Set m_Thumbs = New pdSpriteSheet
    Set m_Timer = New pdTimerAnimation
    picPreview.RequestHighPerformanceRendering True
    
    'Populate the layer listboxes.  (The user can specify a sub-range of layers from these.)
    Dim i As Long
    If PDImages.IsImageActive() Then
        
        ddWhichFrame(0).SetAutomaticRedraws False
        ddWhichFrame(1).SetAutomaticRedraws False
        
        For i = 0 To PDImages.GetActiveImage.GetNumOfLayers - 1
            ddWhichFrame(0).AddItem PDImages.GetActiveImage.GetLayerByIndex(i).GetLayerName, i
            ddWhichFrame(1).AddItem PDImages.GetActiveImage.GetLayerByIndex(i).GetLayerName, i
        Next i
        
        ddWhichFrame(0).ListIndex = 0
        ddWhichFrame(1).ListIndex = ddWhichFrame(1).ListCount - 1
        
        cmdBar.RequestPresetNoLoad ddWhichFrame(0)
        cmdBar.RequestPresetNoLoad ddWhichFrame(1)
        
        ddWhichFrame(0).SetAutomaticRedraws True, True
        ddWhichFrame(1).SetAutomaticRedraws True, True
    
    End If
    
    'Default to infinite replays for this tool (since it's the only useful way to see frame rate changes)
    Me.btnPlay(1).Value = True
    m_Timer.SetRepeat True
    
    'Apply translations and visual themes
    Interface.ApplyThemeAndTranslations Me, True, True, picPreview.hWnd
    UpdateAgainstCurrentTheme
    
    'Update animation frames (so the user can preview them, obviously!)
    If PDImages.IsImageActive() Then UpdateAnimationSettings
    
    'Render the first frame of the animation
    RenderAnimationFrame
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If (Not m_Timer Is Nothing) Then m_Timer.StopTimer
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_Timer = Nothing
    ReleaseFormTheming Me
End Sub

Private Sub m_Timer_DrawFrame(ByVal idxFrame As Long)
    
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
    
    If (Not PDImages.IsImageActive()) Then Exit Sub
    
    'Suspend automatic control-based updates while we get everything synchronized
    m_DoNotUpdate = True
    
    'In animation files, we currently assume all frames are the same size as the image itself,
    ' because this is how PD pre-processes them.  (This may change in the future.)
    Dim bWidth As Long, bHeight As Long
    bWidth = picPreview.GetWidth - 2
    bHeight = picPreview.GetHeight - 2
    
    'Prepare our temporary animation buffer; we don't use it here, but it makes sense to initialize it
    ' to the required size now
    If (m_AniFrame Is Nothing) Then Set m_AniFrame = New pdDIB
    m_AniFrame.CreateBlank bWidth, bHeight, 32, 0, 0
    
    'Figure out what size to use for the animation thumbnails
    PDMath.ConvertAspectRatio PDImages.GetActiveImage.Width, PDImages.GetActiveImage.Height, bWidth, bHeight, m_ThumbWidth, m_ThumbHeight
    
    'Ensure the thumb isn't larger than the actual image
    If (m_ThumbWidth > PDImages.GetActiveImage.Width) Or (m_ThumbHeight > PDImages.GetActiveImage.Height) Then
        m_ThumbWidth = PDImages.GetActiveImage.Width
        m_ThumbHeight = PDImages.GetActiveImage.Height
    End If
    
    'If the thumb image width/height is the same as our current settings, we can keep our existing cache.
    If (m_ThumbWidth <> m_AniThumbBounds.Width) Or (m_ThumbHeight <> m_AniThumbBounds.Height) Or (m_FrameCount <> PDImages.GetActiveImage.GetNumOfLayers) Then
        
        'Load all animation frames.
        m_FrameCount = PDImages.GetActiveImage.GetNumOfLayers
        ReDim m_Frames(0 To m_FrameCount - 1) As PD_AnimationFrame
        
        m_Thumbs.ResetCache
        m_Timer.NotifyFrameCount m_FrameCount
        
        sldFrame.Max = m_FrameCount - 1
        
        'Store the boundary rect of where the thumb will actually appear; we need this for rendering
        ' a transparency checkerboard
        With m_AniThumbBounds
            .Left = 0
            .Top = 0
            .Width = m_ThumbWidth
            .Height = m_ThumbHeight
        End With
        
        'Before generating our preview images, we need to figure out how many frames we
        ' can fit on a shared spritesheet.  (We use sheets to cut down on resource usage;
        ' otherwise we may produce a horrifying number of GDI objects.)  The key here is
        ' that we don't want sheets to grow too large; if they're huge, they risk not
        ' having enough available memory to generate them at all.
        
        'For now, I use a (conservative?) upper limit of ~16mb per sheet (2x1920x1080x4)
        Dim sheetSizeLimit As Long
        sheetSizeLimit = 16777216
        
        Dim numFramesPerSheet As Long
        numFramesPerSheet = sheetSizeLimit / (m_ThumbWidth * m_ThumbHeight * 4)
        If (numFramesPerSheet < 2) Then numFramesPerSheet = 2
        m_Thumbs.SetMaxSpritesInColumn numFramesPerSheet
        
        'Load all thumbnails
        Dim i As Long, tmpDIB As pdDIB
        For i = 0 To m_FrameCount - 1
            
            'Retrieve an updated thumbnail
            If (tmpDIB Is Nothing) Then Set tmpDIB = New pdDIB
            tmpDIB.CreateBlank m_ThumbWidth, m_ThumbHeight, 32, 0, 0
            
            With m_Frames(i)
            
                .afWidth = m_ThumbWidth
                .afHeight = m_ThumbHeight
                
                PDImages.GetActiveImage.GetLayerByIndex(i).RequestThumbnail_ImageCoords tmpDIB, PDImages.GetActiveImage, PDMath.Max2Int(m_ThumbWidth, m_ThumbHeight), False, VarPtr(m_AniThumbBounds)
                .afThumbKey = m_Thumbs.AddImage(tmpDIB, Str$(i) & "|" & Str$(m_ThumbWidth))
                
                'Retrieve layer frame times and other metadata
                .afFrameDelayMS = PDImages.GetActiveImage.GetLayerByIndex(i).GetLayerFrameTimeInMS()
                .afFrameOpacity = PDImages.GetActiveImage.GetLayerByIndex(i).GetLayerOpacity()
                .afFrameBlendMode = PDImages.GetActiveImage.GetLayerByIndex(i).GetLayerBlendMode()
                .afFrameAlphaMode = PDImages.GetActiveImage.GetLayerByIndex(i).GetLayerAlphaMode()
                
            End With
            
        Next i
        
        'Relay frame times to the animator
        NotifyNewFrameTimes
        
    End If
        
    m_DoNotUpdate = False
    
    'Render the first frame of the animation
    RenderAnimationFrame
    
End Sub

'Render the current animation frame
Private Sub RenderAnimationFrame()
    
    'For performance reasons, we skip updates under a variety of circumstances
    If m_DoNotUpdate Then Exit Sub
    If (m_AniFrame Is Nothing) Then Exit Sub
    If Interface.GetDialogResizeFlag() Then
        picPreview.PaintText g_Language.TranslateMessage("waiting..."), 24
        Exit Sub
    End If
    
    Dim idxFrame As Long
    idxFrame = m_Timer.GetCurrentFrame()
    
    'We need to calculate x/y offsets relative to the current preview area
    Dim bWidth As Long, bHeight As Long
    bWidth = picPreview.GetWidth - 2
    bHeight = picPreview.GetHeight - 2
    
    Dim xOffset As Long, yOffset As Long
    xOffset = (bWidth - m_AniThumbBounds.Width) \ 2
    yOffset = (bHeight - m_AniThumbBounds.Height) \ 2
    
    'Make sure the frame request is valid; if it isn't, exit immediately
    If (idxFrame >= 0) And (idxFrame < m_FrameCount) Then
        
        'Request the back buffer DC, and ask the support module to erase any existing rendering for us.
        m_AniFrame.ResetDIB 0
        m_AniFrame.SetInitialAlphaPremultiplicationState True
        
        'Paint a stack consisting of: checkerboard background, then current frame.
        With m_Frames(idxFrame)
            
            'Checkerboard background
            GDI_Plus.GDIPlusFillDIBRect_Pattern m_AniFrame, xOffset, yOffset, m_AniThumbBounds.Width, m_AniThumbBounds.Height, g_CheckerboardPattern, , True, True
            
            'Now the current frame
            If m_Thumbs.DoesImageExist(Str$(idxFrame) & "|" & Str$(.afWidth)) Then
                m_Thumbs.PaintCachedImage m_AniFrame.GetDIBDC, xOffset, yOffset, m_Frames(idxFrame).afThumbKey, Int(m_Frames(idxFrame).afFrameOpacity * 2.55 + 0.5)
            End If
            
        End With
        
        'Paint the final result to the screen, as relevant
        picPreview.CopyDIB m_AniFrame, False, True, True, True
        
    'If our frame counter is invalid, end all animations
    Else
        m_Timer.StopTimer
    End If
    
    'Finally, update the slider tooltip to make it easier for the user to zero-in on a given frame
    If (Not g_Language Is Nothing) Then
        
        Dim numFrames As Long, curFrame As Long
        numFrames = m_FrameCount
        curFrame = idxFrame
        
        Dim totalTime As Long, curFrameTime As Long
        Dim i As Long
        For i = 0 To numFrames - 1
            totalTime = totalTime + Int(GetModifiedFrameTime(m_Frames(i).afFrameDelayMS, i, sldSpeed.Value, ddWhichFrame(0).ListIndex, ddWhichFrame(1).ListIndex) + 0.5!)
            curFrameTime = totalTime
        Next i
        
        Dim frameToolText As pdString
        Set frameToolText = New pdString
        frameToolText.Append g_Language.TranslateMessage("Current frame: %1 of %2", curFrame + 1, numFrames)
        frameToolText.Append ", "
        frameToolText.Append g_Language.TranslateMessage("%1 of %2", Strings.StrFromTimeInMS(curFrameTime, True), Strings.StrFromTimeInMS(totalTime, True))
        
        sldFrame.AssignTooltip frameToolText.ToString(), vbNullString, True
        
    End If
        
End Sub

Private Sub m_Timer_EndOfAnimation()
    m_DoNotUpdate = True
    If btnPlay(0).Value Then btnPlay(0).Value = False
    sldFrame.Value = m_Timer.GetCurrentFrame()
    m_DoNotUpdate = False
End Sub

'If the user clicks the preview window (for some reason), it'll trigger a redraw.
Private Sub picPreview_DrawMe(ByVal targetDC As Long, ByVal ctlWidth As Long, ByVal ctlHeight As Long)
    RenderAnimationFrame
End Sub

Private Sub picPreview_WindowResizeDetected()
    UpdateAnimationSettings
End Sub

Private Sub sldFrame_Change()
    If (Not m_DoNotUpdate) Then
        m_Timer.StopTimer
        m_Timer.SetCurrentFrame sldFrame.Value
    End If
End Sub

Private Sub NotifyNewFrameTimes()
    Dim i As Long
    For i = 0 To m_FrameCount - 1
        m_Timer.NotifyFrameTime Int(GetModifiedFrameTime(m_Frames(i).afFrameDelayMS, i, sldSpeed.Value, ddWhichFrame(0).ListIndex, ddWhichFrame(1).ListIndex) + 0.5!), i
    Next i
End Sub

Private Function GetLocalParamString() As String

    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    With cParams
        .AddParam "speed-modifier", sldSpeed.Value
        .AddParam "idx-first-frame", ddWhichFrame(0).ListIndex
        .AddParam "idx-last-frame", ddWhichFrame(1).ListIndex
    End With
    GetLocalParamString = cParams.GetParamString()
    
End Function

Private Function GetModifiedFrameTime(ByVal baseFrameTime As Single, ByVal idxFrame As Long, Optional ByVal speedModifier As Single = 0!, Optional ByVal idxFirstFrame As Long = -1, Optional ByVal idxLastFrame As Long = -1) As Long
    
    Dim frameRateModifier As Single
    If (speedModifier > 0!) Then
        frameRateModifier = 1! / (1! + speedModifier)
    ElseIf (speedModifier < 0!) Then
        frameRateModifier = Abs(speedModifier) + 1!
    Else
        frameRateModifier = 1!
    End If
    
    If (idxFirstFrame < 0) Then idxFirstFrame = ddWhichFrame(0).ListIndex
    If (idxLastFrame < 0) Then idxLastFrame = ddWhichFrame(1).ListIndex
    
    Dim idxMin As Long, idxMax As Long
    idxMin = PDMath.Min2Int(idxFirstFrame, idxLastFrame)
    idxMax = PDMath.Max2Int(idxFirstFrame, idxLastFrame)
    
    If (idxFrame >= idxMin) And (idxFrame <= idxMax) Then
        GetModifiedFrameTime = baseFrameTime * frameRateModifier
    Else
        GetModifiedFrameTime = baseFrameTime
    End If
    
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
    ' *not* rely on the neighboring looping setting.  (I have mixed feelings about the
    ' intuitiveness of this, but I feel like there needs to be *some* way to preview the
    ' animation as a loop without actually committing to it... idk, I may revisit.)
    btnPlay(1).AssignTooltip UserControls.GetCommonTranslation(pduct_AnimationRepeatToggle)
    
End Sub

Private Sub sldSpeed_Change()
    NotifyNewFrameTimes
End Sub
