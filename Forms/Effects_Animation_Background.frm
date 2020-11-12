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
End
Attribute VB_Name = "FormAnimBackground"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Effect > Animation > Add background
'Copyright 2019-2020 by Tanner Helland
'Created: 26/August/19
'Last updated: 11/November/20
'Last update: spin off from the central Image > Animation dialog
'
'In v9.0, PhotoDemon started gaining effects involving animated images.  This necessitated a bunch
' of new preview and UI code, since previewing effects in real-time is such an intensive process,
' and PD's existing solution was only ever designed for static images.
'
'This dialog served as the testbed for the first animation-related effect, and its code is now
' mirrored across many other places in the project.
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

'Animation frames are stored in a spritesheet control, but to simplify display, we also cache a bunch
' of frame-related details.
Private Type PD_AnimationFrame
    
    'DIB parameters
    afThumbKey As Long
    afWidth As Long
    afHeight As Long
    
    'Metadata
    afFrameDelayOrig As Long
    
End Type

Private m_Thumbs As pdSpriteSheet
Private m_Frames() As PD_AnimationFrame
Private m_FrameCount As Long
Private m_AniThumbBounds As RectF

'Animation updates are rendered to a temporary DIB, which is then forwarded to the preview window
Private m_AniFrame As pdDIB

'Apply an arbitrary background layer to other layers
Public Sub ApplyAnimationBackground(ByVal effectParams As String)

End Sub

Private Sub btnPlay_Click(Index As Integer, ByVal Shift As ShiftConstants)

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

Private Sub cmdBar_CancelClick()
    m_Timer.StopTimer
End Sub

Private Sub cmdBar_OKClick()
    
    'Halt any animations
    m_Timer.StopTimer
    
    'Process changes
    Process "Animation background", , GetLocalParamString(), UNDO_Image
    
End Sub

Private Sub Form_Load()
    
    'Make sure our animation objects exist
    Set m_Thumbs = New pdSpriteSheet
    Set m_Timer = New pdTimerAnimation
    
    'Apply translations and visual themes
    Interface.ApplyThemeAndTranslations Me, True, True, picPreview.hWnd
    UpdateAgainstCurrentTheme
    
    'Update animation frames (so the user can preview them, obviously!)
    If PDImages.IsImageActive() Then UpdateAnimationSettings
    
    'Render the first frame of the animation
    RenderAnimationFrame
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Function GetLocalParamString() As String

    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    'TODO!
    
    GetLocalParamString = cParams.GetParamString
    
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
    Dim tText As String
    tText = g_Language.TranslateMessage("Toggle between 1x and repeating previews")
    btnPlay(1).AssignTooltip tText
    
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
    Dim thumbImageWidth As Long, thumbImageHeight As Long
    PDMath.ConvertAspectRatio PDImages.GetActiveImage.Width, PDImages.GetActiveImage.Height, bWidth, bHeight, thumbImageWidth, thumbImageHeight
    
    'Ensure the thumb isn't larger than the actual image
    If (thumbImageWidth > PDImages.GetActiveImage.Width) Or (thumbImageHeight > PDImages.GetActiveImage.Height) Then
        thumbImageWidth = PDImages.GetActiveImage.Width
        thumbImageHeight = PDImages.GetActiveImage.Height
    End If
    
    'If the thumb image width/height is the same as our current settings, we can keep our existing cache.
    If (thumbImageWidth <> m_AniThumbBounds.Width) Or (thumbImageHeight <> m_AniThumbBounds.Height) Or (m_FrameCount <> PDImages.GetActiveImage.GetNumOfLayers) Then
        
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
            .Width = thumbImageWidth
            .Height = thumbImageHeight
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
        numFramesPerSheet = sheetSizeLimit / (thumbImageWidth * thumbImageHeight * 4)
        If (numFramesPerSheet < 2) Then numFramesPerSheet = 2
        m_Thumbs.SetMaxSpritesInColumn numFramesPerSheet
        
        Dim numZeroFrameDelays As Long
        
        'Load all thumbnails
        Dim i As Long, tmpDIB As pdDIB
        For i = 0 To m_FrameCount - 1
            
            'Retrieve an updated thumbnail
            If (tmpDIB Is Nothing) Then Set tmpDIB = New pdDIB
            tmpDIB.CreateBlank thumbImageWidth, thumbImageHeight, 32, 0, 0
            
            m_Frames(i).afWidth = thumbImageWidth
            m_Frames(i).afHeight = thumbImageHeight
            
            PDImages.GetActiveImage.GetLayerByIndex(i).RequestThumbnail_ImageCoords tmpDIB, PDImages.GetActiveImage, PDMath.Max2Int(thumbImageWidth, thumbImageHeight), False, VarPtr(m_AniThumbBounds)
            m_Frames(i).afThumbKey = m_Thumbs.AddImage(tmpDIB, Str$(i) & "|" & Str$(thumbImageWidth))
            
            'Retrieve layer frame times and relay them to the animation object
            m_Frames(i).afFrameDelayOrig = PDImages.GetActiveImage.GetLayerByIndex(i).GetLayerFrameTimeInMS()
            If (m_Frames(i).afFrameDelayOrig = 0) Then numZeroFrameDelays = numZeroFrameDelays + 1
            
        Next i
        
        'Relay frame times to the animator
        NotifyNewFrameTimes
        
    End If
        
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
        
        'Paint a checkerboard background only over the relevant image region, followed by the frame itself
        With m_Frames(idxFrame)
            
            GDI_Plus.GDIPlusFillDIBRect_Pattern m_AniFrame, xOffset, yOffset, m_AniThumbBounds.Width, m_AniThumbBounds.Height, g_CheckerboardPattern, , True, True
            
            'Make sure we have the necessary image in the spritesheet cache
            If m_Thumbs.DoesImageExist(Str$(idxFrame) & "|" & Str$(.afWidth)) Then
                m_Thumbs.PaintCachedImage m_AniFrame.GetDIBDC, xOffset, yOffset, m_Frames(idxFrame).afThumbKey
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
        m_Timer.NotifyFrameTime m_Frames(i).afFrameDelayOrig, i
    Next i
End Sub
