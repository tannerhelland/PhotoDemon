VERSION 5.00
Begin VB.Form FormAnimBackground 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   " Animation background"
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
   Begin PhotoDemon.pdCheckBox chkDelete 
      Height          =   495
      Left            =   6360
      TabIndex        =   5
      Top             =   1080
      Width           =   5535
      _ExtentX        =   9975
      _ExtentY        =   873
      Caption         =   "after processing, delete this layer"
   End
   Begin PhotoDemon.pdDropDown ddLayer 
      Height          =   855
      Left            =   6240
      TabIndex        =   4
      Top             =   120
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   1508
      Caption         =   "background layer"
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
      TabIndex        =   6
      Top             =   1800
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   1508
      Caption         =   "first frame for effect"
   End
   Begin PhotoDemon.pdDropDown ddWhichFrame 
      Height          =   855
      Index           =   1
      Left            =   6240
      TabIndex        =   7
      Top             =   2760
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   1508
      Caption         =   "last frame for effect"
   End
End
Attribute VB_Name = "FormAnimBackground"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Effect > Animation > Add background/foreground
'Copyright 2019-2026 by Tanner Helland
'Created: 26/August/19
'Last updated: 31/March/22
'Last update: new options, including first/last frame selection, details on animation slider, and more
'
'In v9.0, PhotoDemon started gaining effects involving animated images.  This necessitated a bunch
' of new preview and UI code, since previewing effects in real-time is such an intensive process,
' and PD's existing solution was only ever designed for static images.
'
'This dialog served as the testbed for the first animation-related effect, and its code is now
' mirrored across other places in the project.
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

'To simplify my life, this dialog handles both background and foreground effects.
' (They're fundamentally the exact same effect - the only difference is whether we merge the target
' layer above or beneath each other frame in the animation.)
Private m_InBackgroundMode As Boolean

'To avoid circular updates on animation state changes, we use this tracker
Private m_DoNotUpdate As Boolean

'A dedicated animation timer is used; it auto-corrects for frame time variations during rendering
Private WithEvents m_Timer As pdTimerAnimation
Attribute m_Timer.VB_VarHelpID = -1

'Thumb width/height is fixed
Private m_ThumbWidth As Long, m_ThumbHeight As Long

Private m_Thumbs As pdSpriteSheet
Private m_Frames() As PD_AnimationFrame
Private m_FrameCount As Long
Private m_AniThumbBounds As RectF

'Animation updates are rendered to a temporary DIB, which is then forwarded to the preview window
Private m_AniFrame As pdDIB

'A copy of the current "background" layer (i.e. the user's requested background) is persistently
' maintained; depending on animation settings, it may be redrawn on each render, but efforts should
' be made to avoid reallocating it (as that's where performance will suffer most)
Private m_BackgroundFrame As pdDIB, m_BackgroundFrameIndex As Long

'To improve animation performance on larger images, a persistent blend DIB is used
Private m_BlendDIB As pdDIB

'For non-standard blend-modes, a compositor object is required
Private m_Compositor As pdCompositor

'Apply an arbitrary background layer to other layers
Public Sub ApplyAnimationBackground(ByVal effectParams As String)
    
    If m_InBackgroundMode Then
        Message "Applying background..."
    Else
        Message "Applying foreground..."
    End If
    
    SetProgBarMax PDImages.GetActiveImage.GetNumOfLayers()
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    'Retrieve parameters
    m_InBackgroundMode = cParams.GetBool("background-effect", True, True)
    
    Dim idxLayer As Long
    idxLayer = cParams.GetLong("target-layer-index", 0, True)
    
    Dim deleteLayerAfter As Boolean
    deleteLayerAfter = cParams.GetBool("delete-after", False, True)
    
    Dim idxFirstFrame As Long, idxLastFrame As Long
    idxFirstFrame = cParams.GetLong("idx-first-frame", 0)
    idxLastFrame = cParams.GetLong("idx-last-frame", PDImages.GetActiveImage.GetNumOfLayers() - 1)
    
    'Convert the target layer (background or foreground) to a null-padded layer, which is just a
    ' a layer at the same size as its parent image with no active transforms.
    ' (This makes it trivially mergeable.)
    Dim fxDIB As pdDIB
    PDImages.GetActiveImage.GetLayerByIndex(idxLayer).ConvertToNullPaddedLayer PDImages.GetActiveImage.Width, PDImages.GetActiveImage.Height, True
    Set fxDIB = PDImages.GetActiveImage.GetLayerByIndex(idxLayer).GetLayerDIB
    
    'We also need a scratch layer for compositing.
    Dim scratchDIB As pdDIB
    Set scratchDIB = New pdDIB
    scratchDIB.CreateBlank fxDIB.GetDIBWidth, fxDIB.GetDIBHeight, 32, 0, 0
    scratchDIB.SetInitialAlphaPremultiplicationState True
    
    'pdCompositor handles compositing duties.
    If (m_Compositor Is Nothing) Then Set m_Compositor = New pdCompositor
    
    Dim i As Long
    For i = 0 To PDImages.GetActiveImage.GetNumOfLayers - 1
        
        ProgressBars.SetProgBarVal i + 1
        
        'Skip the target layer (obviously)
        If (i <> idxLayer) Then
            
            'Skip layers that are not in the target range
            If (i >= idxFirstFrame) And (i <= idxLastFrame) Then
                
                'Convert this layer to a null-padded layer
                PDImages.GetActiveImage.GetLayerByIndex(i).ConvertToNullPaddedLayer PDImages.GetActiveImage.Width, PDImages.GetActiveImage.Height, True
                
                'Merge the transformed DIB with the effect DIB, c/o pdCompositor.
                ' (The order of merge is literally the only thing different between background and foreground mode!)
                scratchDIB.ResetDIB 0
                
                If m_InBackgroundMode Then
                    fxDIB.AlphaBlendToDC scratchDIB.GetDIBDC, Int(PDImages.GetActiveImage.GetLayerByIndex(idxLayer).GetLayerOpacity * 2.55 + 0.5)
                    m_Compositor.QuickMergeTwoDibsOfEqualSize scratchDIB, PDImages.GetActiveImage.GetLayerByIndex(i).GetLayerDIB, PDImages.GetActiveImage.GetLayerByIndex(i).GetLayerBlendMode(), PDImages.GetActiveImage.GetLayerByIndex(i).GetLayerOpacity(), PDImages.GetActiveImage.GetLayerByIndex(idxLayer).GetLayerAlphaMode(), PDImages.GetActiveImage.GetLayerByIndex(i).GetLayerAlphaMode()
                Else
                    PDImages.GetActiveImage.GetLayerByIndex(i).GetLayerDIB.AlphaBlendToDC scratchDIB.GetDIBDC, Int(PDImages.GetActiveImage.GetLayerByIndex(i).GetLayerOpacity() * 2.55 + 0.5)
                    m_Compositor.QuickMergeTwoDibsOfEqualSize scratchDIB, fxDIB, PDImages.GetActiveImage.GetLayerByIndex(idxLayer).GetLayerBlendMode, PDImages.GetActiveImage.GetLayerByIndex(idxLayer).GetLayerOpacity, PDImages.GetActiveImage.GetLayerByIndex(i).GetLayerAlphaMode(), PDImages.GetActiveImage.GetLayerByIndex(idxLayer).GetLayerAlphaMode()
                End If
                
                'Replace the layer with the newly composited image, then shrink the top layer to its smallest
                ' possible size (un-null-pad it)
                PDImages.GetActiveImage.GetLayerByIndex(i).GetLayerDIB.CreateFromExistingDIB scratchDIB
                PDImages.GetActiveImage.GetLayerByIndex(i).CropNullPaddedLayer
                
                'Notify the parent image of the change
                PDImages.GetActiveImage.NotifyImageChanged UNDO_Layer, i
                
            End If
                
        End If
        
    Next i
    
    'If the caller wants the target layer deleted, do that now
    If deleteLayerAfter Then PDImages.GetActiveImage.DeleteLayerByIndex idxLayer
    
    'Notify the parent image of the final change
    PDImages.GetActiveImage.NotifyImageChanged UNDO_Image
    
    'Redraw the screen and finalize the effect
    toolbar_Layers.NotifyLayerChange
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
    Message "Finished."
    ProgressBars.SetProgBarVal 0
    ProgressBars.ReleaseProgressBar
    
End Sub

'Pass TRUE for background mode; FALSE for foreground mode
Public Sub SetBackgroundMode(ByVal newMode As Boolean)
    
    m_InBackgroundMode = newMode
    
    'Apply translations and visual themes
    Interface.ApplyThemeAndTranslations Me, True, True, picPreview.hWnd
    UpdateAgainstCurrentTheme
    
    'Update animation frames (so the user can preview them, obviously!)
    If PDImages.IsImageActive() Then UpdateAnimationSettings
    
    'Render the first frame of the animation
    RenderAnimationFrame
    
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
    Process "Animation background", , GetLocalParamString(), UNDO_Image
    
End Sub

Private Sub ddLayer_Click()
    RenderAnimationFrame
End Sub

Private Sub ddWhichFrame_Click(Index As Integer)
    RenderAnimationFrame
End Sub

Private Sub Form_Load()
    
    'Make sure our animation objects exist
    Set m_Thumbs = New pdSpriteSheet
    Set m_Timer = New pdTimerAnimation
    picPreview.RequestHighPerformanceRendering True
    
    'Set some animation default values
    m_BackgroundFrameIndex = -1
    
    'Populate the layer listbox.  (The user will select their desired background layer from this box.)
    Dim i As Long
    If PDImages.IsImageActive() Then
        
        ddLayer.SetAutomaticRedraws False
        ddWhichFrame(0).SetAutomaticRedraws False
        ddWhichFrame(1).SetAutomaticRedraws False
        
        For i = 0 To PDImages.GetActiveImage.GetNumOfLayers - 1
            ddLayer.AddItem PDImages.GetActiveImage.GetLayerByIndex(i).GetLayerName, i
            ddWhichFrame(0).AddItem PDImages.GetActiveImage.GetLayerByIndex(i).GetLayerName, i
            ddWhichFrame(1).AddItem PDImages.GetActiveImage.GetLayerByIndex(i).GetLayerName, i
        Next i
        
        ddLayer.ListIndex = PDImages.GetActiveImage.GetActiveLayerIndex
        ddWhichFrame(0).ListIndex = 0
        ddWhichFrame(1).ListIndex = ddWhichFrame(1).ListCount - 1
        
        cmdBar.RequestPresetNoLoad ddLayer
        cmdBar.RequestPresetNoLoad ddWhichFrame(0)
        cmdBar.RequestPresetNoLoad ddWhichFrame(1)
        
        ddLayer.SetAutomaticRedraws True, True
        ddWhichFrame(0).SetAutomaticRedraws True, True
        ddWhichFrame(1).SetAutomaticRedraws True, True
    
    End If
    
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
        Set m_BackgroundFrame = Nothing
        
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
    
    'Recreate the background, as necessary
    Dim idxFrame As Long
    idxFrame = ddLayer.ListIndex
    
    If (m_BackgroundFrameIndex <> idxFrame) Or (m_BackgroundFrame Is Nothing) Then
        
        'Restore the correct frame time of the old background frame index
        If (m_BackgroundFrameIndex >= 0) Then m_Timer.NotifyFrameTime m_Frames(m_BackgroundFrameIndex).afFrameDelayMS, m_BackgroundFrameIndex
        
        'Ensure background DIB and temporary DIBs are allocated
        If (m_BackgroundFrame Is Nothing) Then Set m_BackgroundFrame = New pdDIB
        m_BackgroundFrame.CreateBlank m_ThumbWidth, m_ThumbHeight, 32, 0, 0
        m_BackgroundFrame.SetInitialAlphaPremultiplicationState True
        
        If (m_BlendDIB Is Nothing) Then Set m_BlendDIB = New pdDIB
        m_BlendDIB.CreateBlank m_AniFrame.GetDIBWidth, m_AniFrame.GetDIBHeight, 32, 0, 0
        m_BlendDIB.SetInitialAlphaPremultiplicationState True
        
        'Retrieve a copy of this layer at the current preview size
        PDImages.GetActiveImage.GetLayerByIndex(idxFrame).RequestThumbnail_ImageCoords m_BackgroundFrame, PDImages.GetActiveImage, PDMath.Max2Int(m_ThumbWidth, m_ThumbHeight), False, VarPtr(m_AniThumbBounds)
        
        'Update the frame pointer, and set the animation delay of the new background frame to 0
        m_BackgroundFrameIndex = idxFrame
        m_Timer.NotifyFrameTime 0, idxFrame
        
    End If
    
    'Now we start work on the current frame
    idxFrame = m_Timer.GetCurrentFrame()
    
    'If the scrubber points at the background frame, point at it the next layer in line (if possible).
    ' This fixes an issue where the desired background layer is also the first frame, and upon loading
    ' this dialog, the animation preview just shows the background frame.
    If (idxFrame = m_BackgroundFrameIndex) Then
        If (idxFrame < m_FrameCount - 1) Then
            idxFrame = idxFrame + 1
        ElseIf (idxFrame > 0) Then
            idxFrame = idxFrame - 1
        End If
    End If
    
    'We need to calculate x/y offsets relative to the current preview area
    Dim bWidth As Long, bHeight As Long
    bWidth = picPreview.GetWidth - 2
    bHeight = picPreview.GetHeight - 2
    
    Dim xOffset As Long, yOffset As Long
    xOffset = (bWidth - m_AniThumbBounds.Width) \ 2
    yOffset = (bHeight - m_AniThumbBounds.Height) \ 2
    
    'To support blend and alpha modes, we must use our internal compositor
    If (m_Compositor Is Nothing) Then Set m_Compositor = New pdCompositor
    
    Dim targetBlendMode As PD_BlendMode, targetOpacity As Long
    
    'Figure out blend mode and opacity of this operation; in background mode, these come from the current frame;
    ' in foreground mode, they come from the fixed foreground layer
    If m_InBackgroundMode Then
        targetBlendMode = m_Frames(idxFrame).afFrameBlendMode
    Else
        targetBlendMode = m_Frames(m_BackgroundFrameIndex).afFrameBlendMode
    End If
    
    targetOpacity = Int(m_Frames(m_BackgroundFrameIndex).afFrameOpacity * 2.55 + 0.5)
    
    'Alpha mode is a bit more complicated than blend mode, because we need to use non-standard paths if either
    ' the top *or* bottom layer have a non-standard setting.
    Dim alphaModeNormal As Boolean, bottomAlphaMode As PD_AlphaMode, topAlphaMode As PD_AlphaMode
    If m_InBackgroundMode Then
        bottomAlphaMode = m_Frames(m_BackgroundFrameIndex).afFrameAlphaMode
        topAlphaMode = m_Frames(idxFrame).afFrameAlphaMode
    Else
        bottomAlphaMode = m_Frames(idxFrame).afFrameAlphaMode
        topAlphaMode = m_Frames(m_BackgroundFrameIndex).afFrameAlphaMode
    End If
    
    alphaModeNormal = (bottomAlphaMode = AM_Normal) And (topAlphaMode = AM_Normal)
    
    'The background/foreground effect will be limited to the user's specified range.
    Dim firstFrameForEffect As Long, lastFrameForEffect As Long
    If (ddWhichFrame(0).ListIndex < ddWhichFrame(1).ListIndex) Then
        firstFrameForEffect = ddWhichFrame(0).ListIndex
        lastFrameForEffect = ddWhichFrame(1).ListIndex
    Else
        firstFrameForEffect = ddWhichFrame(1).ListIndex
        lastFrameForEffect = ddWhichFrame(0).ListIndex
    End If
    
    If (firstFrameForEffect < m_BackgroundFrameIndex) Then firstFrameForEffect = firstFrameForEffect - 1
    If (lastFrameForEffect < m_BackgroundFrameIndex) Then lastFrameForEffect = lastFrameForEffect - 1
    
    Dim thisFrameInEffectRange As Boolean
    thisFrameInEffectRange = (idxFrame >= firstFrameForEffect) And (idxFrame <= lastFrameForEffect)
    
    'Make sure the frame request is valid; if it isn't, exit immediately
    If (idxFrame >= 0) And (idxFrame < m_FrameCount) Then
        
        'Request the back buffer DC, and ask the support module to erase any existing rendering for us.
        m_AniFrame.ResetDIB 0
        m_AniFrame.SetInitialAlphaPremultiplicationState True
        
        'Paint a stack consisting of: checkerboard background, background layer, current frame.
        ' (In foreground mode, the order of "background layer" and "current frame" will be switched.)
        With m_Frames(idxFrame)
            
            'Normal blend mode allows us to just alpha-blend everything; this is faster than using
            ' a compositor, and we can do it directly atop a checkerboard background
            If (targetBlendMode = BM_Normal) And alphaModeNormal Then
                
                'Checkerboard background
                GDI_Plus.GDIPlusFillDIBRect_Pattern m_AniFrame, xOffset, yOffset, m_AniThumbBounds.Width, m_AniThumbBounds.Height, g_CheckerboardPattern, , True, True
                
                'When in background mode, paint the fixed background layer now
                If (m_InBackgroundMode And thisFrameInEffectRange) Then
                    m_BackgroundFrame.AlphaBlendToDC m_AniFrame.GetDIBDC, targetOpacity, xOffset, yOffset
                End If
                
                'Regardless of mode, paint the current frame directly into the animation buffer
                If m_Thumbs.DoesImageExist(Str$(idxFrame) & "|" & Str$(.afWidth)) Then
                    m_Thumbs.PaintCachedImage m_AniFrame.GetDIBDC, xOffset, yOffset, m_Frames(idxFrame).afThumbKey, Int(m_Frames(idxFrame).afFrameOpacity * 2.55 + 0.5)
                End If
                
                'If in foreground mode, we now need to paint the fixed top layer
                If ((Not m_InBackgroundMode) And thisFrameInEffectRange) Then
                    m_BackgroundFrame.AlphaBlendToDC m_AniFrame.GetDIBDC, targetOpacity, xOffset, yOffset
                End If
                
            'Non-standard blend-mode
            Else
                
                m_BlendDIB.ResetDIB 0
                m_BlendDIB.SetInitialAlphaPremultiplicationState True
                
                'We now need to copy the background layer (whatever it is) into the m_BlendMode DIB,
                ' at the background layer's expected opacity.
                If (m_InBackgroundMode And thisFrameInEffectRange) Then
                    m_BackgroundFrame.AlphaBlendToDC m_BlendDIB.GetDIBDC, targetOpacity, xOffset, yOffset
                Else
                    If m_Thumbs.DoesImageExist(Str$(idxFrame) & "|" & Str$(.afWidth)) Then
                        m_Thumbs.PaintCachedImage m_BlendDIB.GetDIBDC, xOffset, yOffset, m_Frames(idxFrame).afThumbKey, Int(m_Frames(idxFrame).afFrameOpacity * 2.55 + 0.5)
                    End If
                End If
                
                'm_BlendDIB now contains the background layer.  Generate a similar foreground layer,
                ' using m_AniFrame.
                If m_InBackgroundMode Then
                    If m_Thumbs.DoesImageExist(Str$(idxFrame) & "|" & Str$(.afWidth)) Then
                        m_Thumbs.PaintCachedImage m_AniFrame.GetDIBDC, xOffset, yOffset, m_Frames(idxFrame).afThumbKey, Int(m_Frames(idxFrame).afFrameOpacity * 2.55 + 0.5)
                    End If
                Else
                    If thisFrameInEffectRange Then m_BackgroundFrame.AlphaBlendToDC m_AniFrame.GetDIBDC, targetOpacity, xOffset, yOffset
                End If
                
                'NOTE: opacity has already been handled for all frames.
                
                'Use pdCompositor to blend the two layers using the expected blendmode and alpha modes
                If thisFrameInEffectRange Then
                    m_Compositor.QuickMergeTwoDibsOfEqualSize m_BlendDIB, m_AniFrame, targetBlendMode, 100#, bottomAlphaMode, topAlphaMode
                End If
                
                'Finally, replace the contents of the top layer with the expected checkerboard background,
                ' then merge the composited result atop that
                GDI_Plus.GDIPlusFillDIBRect_Pattern m_AniFrame, xOffset, yOffset, m_AniThumbBounds.Width, m_AniThumbBounds.Height, g_CheckerboardPattern, , True, True
                m_BlendDIB.AlphaBlendToDC m_AniFrame.GetDIBDC
                
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
            If (i <> m_BackgroundFrameIndex) Then
                totalTime = totalTime + m_Frames(i).afFrameDelayMS
                If (i < curFrame) Then curFrameTime = totalTime
            End If
        Next i
        
        If (curFrame > m_BackgroundFrameIndex) Then curFrame = curFrame - 1
        
        Dim frameToolText As pdString
        Set frameToolText = New pdString
        frameToolText.Append g_Language.TranslateMessage("Current frame: %1 of %2", curFrame + 1, numFrames - 1)
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
        m_Timer.NotifyFrameTime m_Frames(i).afFrameDelayMS, i
    Next i
End Sub

Private Function GetLocalParamString() As String

    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    With cParams
        .AddParam "background-effect", m_InBackgroundMode
        .AddParam "target-layer-index", ddLayer.ListIndex
        .AddParam "delete-after", chkDelete.Value
        .AddParam "idx-first-frame", ddWhichFrame(0).ListIndex
        .AddParam "idx-last-frame", ddWhichFrame(1).ListIndex
    End With
    GetLocalParamString = cParams.GetParamString()
    
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
    
    'Some captions also vary based on background/foreground mode
    If (Not g_Language Is Nothing) And (Not g_WindowManager Is Nothing) Then
        If m_InBackgroundMode Then
            g_WindowManager.SetWindowCaptionW Me.hWnd, " " & g_Language.TranslateMessage("Animation background")
            ddLayer.Caption = g_Language.TranslateMessage("background layer")
        Else
            g_WindowManager.SetWindowCaptionW Me.hWnd, " " & g_Language.TranslateMessage("Animation foreground")
            ddLayer.Caption = g_Language.TranslateMessage("foreground layer")
        End If
    End If
    
End Sub
