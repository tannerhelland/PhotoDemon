VERSION 5.00
Begin VB.Form FormLevels 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Adjust Image Levels"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   195
   ClientWidth     =   12870
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   503
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   858
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdButton cmdAutoLevels 
      Height          =   600
      Left            =   120
      TabIndex        =   19
      Top             =   6000
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   873
      Caption         =   "set levels automatically"
   End
   Begin PhotoDemon.pdButtonToolbox cmdColorSelect 
      Height          =   375
      Index           =   0
      Left            =   7740
      TabIndex        =   17
      Top             =   3135
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      StickyToggle    =   -1  'True
   End
   Begin PhotoDemon.colorSelector csShadow 
      Height          =   375
      Left            =   7230
      TabIndex        =   13
      Top             =   3135
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      curColor        =   0
   End
   Begin VB.PictureBox picOutputArrows 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5760
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   473
      TabIndex        =   12
      Top             =   4590
      Width           =   7095
   End
   Begin VB.PictureBox picInputArrows 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5760
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   473
      TabIndex        =   11
      Top             =   2790
      Width           =   7095
   End
   Begin VB.PictureBox picHistogram 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   6000
      ScaleHeight     =   151
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   444
      TabIndex        =   10
      Top             =   480
      Width           =   6690
   End
   Begin VB.PictureBox picOutputGradient 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6000
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   444
      TabIndex        =   9
      Top             =   4200
      Width           =   6690
   End
   Begin PhotoDemon.textUpDown tudLevels 
      Height          =   345
      Index           =   0
      Left            =   6000
      TabIndex        =   4
      Top             =   3120
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   714
      Max             =   253
   End
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   6795
      Width           =   12870
      _ExtentX        =   22701
      _ExtentY        =   1323
      BackColor       =   14802140
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
      ColorSelection  =   -1  'True
   End
   Begin PhotoDemon.textUpDown tudLevels 
      Height          =   345
      Index           =   1
      Left            =   8760
      TabIndex        =   5
      Top             =   3120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   714
      Min             =   0.01
      Max             =   0.99
      SigDigits       =   2
      Value           =   0.5
   End
   Begin PhotoDemon.textUpDown tudLevels 
      Height          =   345
      Index           =   2
      Left            =   11490
      TabIndex        =   6
      Top             =   3120
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   714
      Min             =   2
      Max             =   255
      Value           =   255
   End
   Begin PhotoDemon.textUpDown tudLevels 
      Height          =   345
      Index           =   3
      Left            =   6000
      TabIndex        =   7
      Top             =   4920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   714
      Max             =   255
   End
   Begin PhotoDemon.textUpDown tudLevels 
      Height          =   345
      Index           =   4
      Left            =   11355
      TabIndex        =   8
      Top             =   4920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   714
      Max             =   255
      Value           =   255
   End
   Begin PhotoDemon.colorSelector csHighlight 
      Height          =   375
      Left            =   10920
      TabIndex        =   14
      Top             =   3135
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
   End
   Begin PhotoDemon.buttonStrip btsChannel 
      Height          =   600
      Left            =   6030
      TabIndex        =   15
      Top             =   6000
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   1058
   End
   Begin PhotoDemon.pdButtonToolbox cmdColorSelect 
      Height          =   375
      Index           =   1
      Left            =   10530
      TabIndex        =   18
      Top             =   3135
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      StickyToggle    =   -1  'True
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "channel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Index           =   0
      Left            =   6000
      TabIndex        =   16
      Top             =   5640
      Width           =   810
   End
   Begin VB.Label lblOutput 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "output levels"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   6000
      TabIndex        =   2
      Top             =   3840
      Width           =   1350
   End
   Begin VB.Label lblInput 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "input levels"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   6000
      TabIndex        =   1
      Top             =   120
      Width           =   1200
   End
End
Attribute VB_Name = "FormLevels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Image Levels
'Copyright 2006-2015 by Tanner Helland
'Created: 22/July/06
'Last updated: 07/September/15
'Last update: overhaul the underlying histogram UI code, using a new centralized histogram renderer
'
'This tool allows the user to adjust image levels.  Its behavior is based off Photoshop's Levels tool, and identical
' values entered into both programs should yield an identical image.
'
'Unfortunately, to perfectly mimic Photoshop's behavior, some fairly involved (i.e. incomprehensible) math is required.
' To mitigate the speed implications of such convoluted math, a number of look-up tables are used.  This makes the
' function quite fast, but at a hit to readability.  My apologies to anyone trying to understand how the function works.
'
'As of June '14, per-channel levels, set-by-color options, and Auto Levels are now supported.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Constants required for creating a gamma curve from .1 to 10
Private Const MAXGAMMA As Double = 1.8460498941512
Private Const MIDGAMMA As Double = 0.68377223398334
Private Const ROOT10 As Double = 3.16227766

'An image of the current image histogram is generated for each channel, then displayed as requested
Private hDIB() As pdDIB

'Copies of the "slider arrows" used to display and control input/output level manipulation
Private m_Arrows(0 To 2) As pdDIB

'For convenience, the dimensions and offsets of the UI arrows are stored in these variables.  Note that there are two
' extra offsets relative to the arrow DIBs themselves; this is because we render two copies of the black and white
' arrows to the screen, one each for input and output levels.
Private m_ArrowOffsets(0 To 4) As Long
Private m_ArrowWidth As Long, m_ArrowHalfWidth As Long
Private m_DstArrowBoxWidth As Long, m_DstArrowBoxOffset As Long

'Current channel ([0, 3] where 0 = red, 1 = green, 2 = blue, 3 = luminance)
Private m_curChannel As Long

'Because the user can now change levels independently for each of Red, Green, Blue, and Luminance, we must store all
' level values internally (rather than relying on the text up/down controls to do it for us).  Also, because the
' midtone values are floating-point, we declare the whole tracking array as Double-type (even though shadow, highlight,
' and output levels are integers).  The layout of this array is [channel, level adjustment].
Private m_LevelValues(0 To 3, 0 To 4) As Double

'Two special input classes are required; one each for the input and output arrow boxes
Private WithEvents cMouseEventsIn As pdInputMouse
Attribute cMouseEventsIn.VB_VarHelpID = -1
Private WithEvents cMouseEventsOut As pdInputMouse
Attribute cMouseEventsOut.VB_VarHelpID = -1

'If the user is using the mouse to slide nodes around, these values will be used to store the node's index
Private m_ActiveArrow As Long

'To prevent complicated interactions related to the max/min codependence of input shadow and highlight values, m_DisableMaxMinLimits
' can be used to disable automatic bounds-checking of input/output values.  Set this to TRUE when overwriting all on-screen level
' values with the ones stored in memory (e.g. when the user is changing the active channel, so the whole screen gets refreshed).
' When the new values have all been set, restore this to FALSE, then make a single call to fixScrollBars() to establish the new
' max/min bounds.
Private m_DisableMaxMinLimits As Boolean

'When a new channel is selected, refresh all text box values to match the new channel's stored values
Private Sub btsChannel_Click(ByVal buttonIndex As Long)

    m_curChannel = buttonIndex
    
    'Draw the relevant histogram onto the histogram box
    On Error GoTo ignoreChannelRender
    picHistogram.Picture = LoadPicture("")
    If Not hDIB(m_curChannel) Is Nothing Then hDIB(m_curChannel).alphaBlendToDC picHistogram.hDC
    picHistogram.Picture = picHistogram.Image
    
    'Update the text boxes to match the values for the selected channel
ignoreChannelRender:
    updateTextBoxes
    
    'Update the preview.  (The preview itself doesn't actually need to be redrawn, but that function is responsible for
    ' syncing the text box values with the arrow positions.)
    updatePreview

End Sub

'Auto levels wil calculate new levels for the user, using the getIdealLevelParamString function below
Private Sub cmdAutoLevels_Click()
    
    'Retrieve the ideal level param string
    Dim pString As String
    pString = getIdealLevelParamString(pdImages(g_CurrentImage).getActiveDIB)
    
    'Level value parsing is easily handled via PD's standard param string parser class
    Dim cParams As pdParamString
    Set cParams = New pdParamString
    cParams.setParamString pString
    
    Dim i As Long
    For i = 0 To 19
        m_LevelValues(i \ 5, i Mod 5) = cParams.GetDouble(i + 1)
    Next i

    'Update the text boxes to match the new values
    updateTextBoxes
    
    'Redraw the screen
    updatePreview
    
End Sub

'Returns the ideal param string for a given DIB.  "Auto levels" relies on this function to retrieve best values for a function.
' Note that PD's White Balance tool is effectively just an auto-levels function, with a variable "ignore percentage" that the
' user can set.  Similarly, the shadow/highlights tool allows for separate "ignore percentages" for shadows and highlights, but
' is otherwise effectively this same algorithm.
Public Function getIdealLevelParamString(ByRef srcDIB As pdDIB) As String

    'Create a local array and point it at the source DIB's pixel data
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepSafeArray tmpSA, srcDIB
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.getDIBWidth - 1
    finalY = srcDIB.getDIBHeight - 1
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = srcDIB.getDIBColorDepth \ 8
    
    'Color values
    Dim r As Long, g As Long, b As Long, l As Long
    
    'Maximum and minimum values, which will be detected by our initial histogram run
    Dim RMax As Byte, gMax As Byte, bMax As Byte, lMax As Byte
    Dim RMin As Byte, gMin As Byte, bMin As Byte, lMin As Byte
    RMax = 0: gMax = 0: bMax = 0: lMax = 0
    RMin = 255: gMin = 255: bMin = 255: lMin = 255
    
    'Calculate a percentage to ignore at either end.  Photoshop defaults to 0.5%, but the actual value might need to vary on
    ' a per-image basis... hard to know what the best approach is.
    Dim percentIgnore As Double
    percentIgnore = 0.005
    
    'Prepare histogram arrays
    Dim rCount(0 To 255) As Long, gCount(0 To 255) As Long, bCount(0 To 255) As Long, lCount(0 To 255) As Long
    For x = 0 To 255
        rCount(x) = 0
        gCount(x) = 0
        bCount(x) = 0
        lCount(x) = 0
    Next x
    
    'Build an image histogram
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        rCount(r) = rCount(r) + 1
        gCount(g) = gCount(g) + 1
        bCount(b) = bCount(b) + 1
        
        l = (213 * r + 715 * g + 72 * b) \ 1000
        lCount(l) = lCount(l) + 1
    Next y
    Next x
    
     'With the histogram complete, we can now figure out how to stretch the RGB channels. We do this by calculating a min/max
    ' ratio where the top and bottom 0.05% (or user-specified value) of pixels are ignored.
    
    Dim foundYet As Boolean
    foundYet = False
    
    Dim NumOfPixels As Long
    NumOfPixels = (finalX + 1) * (finalY + 1)
    
    Dim wbThreshold As Long
    wbThreshold = NumOfPixels * percentIgnore
    
    r = 0: g = 0: b = 0: l = 0
    
    Dim rTally As Long, gTally As Long, bTally As Long, lTally As Long
    rTally = 0: gTally = 0: bTally = 0: lTally = 0
    
    'Find minimum values of red, green, blue, and luminance
    Do
        If rCount(r) + rTally < wbThreshold Then
            r = r + 1
            rTally = rTally + rCount(r)
        Else
            RMin = r
            foundYet = True
        End If
    Loop While foundYet = False
        
    foundYet = False
        
    Do
        If gCount(g) + gTally < wbThreshold Then
            g = g + 1
            gTally = gTally + gCount(g)
        Else
            gMin = g
            foundYet = True
        End If
    Loop While foundYet = False
    
    foundYet = False
    
    Do
        If bCount(b) + bTally < wbThreshold Then
            b = b + 1
            bTally = bTally + bCount(b)
        Else
            bMin = b
            foundYet = True
        End If
    Loop While foundYet = False
    
    foundYet = False
    
    Do
        If lCount(l) + lTally < wbThreshold Then
            l = l + 1
            lTally = lTally + lCount(l)
        Else
            lMin = l
            foundYet = True
        End If
    Loop While foundYet = False
    
    'Now, find maximum values of red, green, blue, and luminance
    foundYet = False
    
    r = 255: g = 255: b = 255: l = 255
    rTally = 0: gTally = 0: bTally = 0: lTally = 0
    
    Do
        If rCount(r) + rTally < wbThreshold Then
            r = r - 1
            rTally = rTally + rCount(r)
        Else
            RMax = r
            foundYet = True
        End If
    Loop While foundYet = False
        
    foundYet = False
        
    Do
        If gCount(g) + gTally < wbThreshold Then
            g = g - 1
            gTally = gTally + gCount(g)
        Else
            gMax = g
            foundYet = True
        End If
    Loop While foundYet = False
    
    foundYet = False
    
    Do
        If bCount(b) + bTally < wbThreshold Then
            b = b - 1
            bTally = bTally + bCount(b)
        Else
            bMax = b
            foundYet = True
        End If
    Loop While foundYet = False
    
    foundYet = False
    
    Do
        If lCount(l) + lTally < wbThreshold Then
            l = l - 1
            lTally = lTally + lCount(l)
        Else
            lMax = l
            foundYet = True
        End If
    Loop While foundYet = False
    
    'We now have an idealized max/min for each of red, green, blue, and luminance.
    
    'One of the problems with auto-levels is that it can introduce nasty color casts to the image.  Consider an image of a Caucasian
    ' human face; generally speaking, red tends to be exposed fairly equally across skin tones, but green and blue are much more
    ' variable according to background elements.  A face against a bright blue sky will tend to have blue concentrated at the high
    ' end of the scale, so when we auto-level it, those blue levels will get spread across the full spectrum, introducing an
    ' unpleasant purplish-cast across the subject's skin.
    
    'To avoid this in PD (and to produce a really kick-ass Auto-Level result), we split the calculated auto-level adjustment equally
    ' between per-channel corrections and net luminance corrections.  This roughly maintains the existing color spread of the image,
    ' while removing any obviously bad results, and producing a consistently well-exposed final image.  It also serves to balance out
    ' color temperature in an elegant way, without subjecting photos to the standard over-cooled look of other auto-level tools.
    RMin = RMin \ 2
    gMin = gMin \ 2
    bMin = bMin \ 2
    lMin = lMin \ 2
    
    RMax = RMax + ((255 - RMax) \ 2)
    gMax = gMax + ((255 - gMax) \ 2)
    bMax = bMax + ((255 - bMax) \ 2)
    lMax = lMax + ((255 - lMax) \ 2)
    
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Return our assembled data in param-string compatible format
    getIdealLevelParamString = buildParams(RMin, 0.5, RMax, 0, 255, gMin, 0.5, gMax, 0, 255, bMin, 0.5, bMax, 0, 255, lMin, 0.5, lMax, 0, 255)

End Function

'Because the Levels dialog only uses one set of UI controls for all channels, we must manually write out preset data for each channel.
' This event will be raised whenever the command bar needs custom data from us.
Private Sub cmdBar_AddCustomPresetData()
    cmdBar.addPresetData "MultichannelLevelData", getLevelsParamString()
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Levels", , getLevelsParamString(), UNDO_LAYER
End Sub

'Randomize button (command bar)
Private Sub cmdBar_RandomizeClick()

    Randomize Timer

    Dim i As Long
    For i = 0 To 3
    
        'Set random shadow and highlight input levels
        m_LevelValues(i, 0) = Rnd * 125
        m_LevelValues(i, 2) = Rnd * 125 + 128
        
        'Set a random midtone value (range 0.01 - 0.99)
        m_LevelValues(i, 1) = Rnd
        If m_LevelValues(i, 1) < 0.01 Then m_LevelValues(i, 1) = 0.01
        If m_LevelValues(i, 1) > 0.99 Then m_LevelValues(i, 1) = 0.99
        
        'Set random output levels
        m_LevelValues(i, 3) = Rnd * 256
        m_LevelValues(i, 4) = Rnd * 256
    
    Next i
    
    'Update the text boxes to match the new values
    updateTextBoxes
    
    'Redraw the screen
    updatePreview

End Sub

'When a preset is loaded from file, we need to retrieve the custom levels information alongside it
Private Sub cmdBar_ReadCustomPresetData()
    
    'Retrieve a string containing all relevant layer information
    Dim tmpString As String
    tmpString = cmdBar.retrievePresetData("MultichannelLevelData")
    
    'Valid preset data was found
    If Len(tmpString) <> 0 Then
    
        'Level value parsing will be handled via PD's standard param string parser class
        Dim cParams As pdParamString
        Set cParams = New pdParamString
        cParams.setParamString tmpString
        
        Dim i As Long
        For i = 0 To 19
            m_LevelValues(i \ 5, i Mod 5) = cParams.GetDouble(i + 1)
        Next i
    
        'Update the text boxes to match the new values
        updateTextBoxes
        
        'Redraw the screen
        updatePreview
    
    'Valid preset data was *not* found, possibly because the user just upgraded from a past version of the Levels tool.
    ' Reset everything to default values
    Else
        Call cmdBar_ResetClick
    End If
    
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    
    'Make the RGB button pressed by default; this will be overridden by the user's last-used settings, if any exist
    m_curChannel = 3
    btsChannel.ListIndex = m_curChannel
        
    'Reset all values in our tracking array.
    Dim i As Long
    For i = 0 To 3
    
        'Input levels
        m_LevelValues(i, 0) = 0
        m_LevelValues(i, 1) = 0.5
        m_LevelValues(i, 2) = 255
        
        'Output levels
        m_LevelValues(i, 3) = 0
        m_LevelValues(i, 4) = 255
    
    Next i
    
    'Update the text boxes to match the new values
    updateTextBoxes
    
    'Redraw the screen
    updatePreview
    
End Sub

'Update all text box values to match the stored values of the current channel
Private Sub updateTextBoxes()

    cmdBar.markPreviewStatus False
    m_DisableMaxMinLimits = True
    
    'Set max/min values of the input shadow/highlight boxes to their max possible values.  This will prevent the current limits
    ' from affecting the new ones we are about to load.
    tudLevels(0).Max = 255
    tudLevels(2).Min = 0

    'Load the new values
    Dim i As Long
    For i = 0 To 4
        tudLevels(i) = m_LevelValues(m_curChannel, i)
    Next i
    
    'Update the text up/down max/min for shadow and highlight levels
    m_DisableMaxMinLimits = False
    FixScrollBars
    
    'Reinstate automatic preview updates
    cmdBar.markPreviewStatus True

End Sub

Private Sub cmdColorSelect_Click(Index As Integer)

    If cmdBar.previewsAllowed Then
    
        cmdBar.markPreviewStatus False
        
        'Toggle the other command button (as only one can be active at any time)
        If Index = 0 Then
            cmdColorSelect(1).Value = False
        Else
            cmdColorSelect(0).Value = False
        End If
        
        cmdBar.markPreviewStatus True
    
    End If

End Sub

Private Sub cMouseEventsIn_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)

    'Check the mouse position.  If it is over a slider, activate drag mode; otherwise, ignore the click.
    If (Button And pdLeftButton) <> 0 Then
        m_ActiveArrow = isCursorOverArrow(x, True)
    End If

End Sub

Private Sub cMouseEventsIn_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)

    'Left mouse button is down, and the user has a node selected
    If ((Button And pdLeftButton) <> 0) And (m_ActiveArrow >= 0) And (m_ActiveArrow <= 2) Then
    
        'Disable automatic preview updates
        cmdBar.markPreviewStatus False
        
        Dim newTUDValue As Double
        
        'Start by recalculating the x position relative to the histogram box
        Dim tmpX As Double
        tmpX = x - m_DstArrowBoxOffset
        tmpX = tmpX / m_DstArrowBoxWidth
        
        If tmpX < 0 Then tmpX = 0
        If tmpX > 1 Then tmpX = 1
        
        'Calculate a new value for the corresponding text box
        Select Case m_ActiveArrow
        
            'Shadow input node
            Case 0
                newTUDValue = tmpX * 255
                If newTUDValue > tudLevels(0).Max Then newTUDValue = tudLevels(0).Max
                tudLevels(0).Value = newTUDValue
            
            'Midtones input node
            Case 1
                newTUDValue = tmpX * 255
                newTUDValue = (newTUDValue - tudLevels(0).Value) / (tudLevels(2).Value - tudLevels(0).Value)
                If newTUDValue > tudLevels(1).Max Then
                    newTUDValue = tudLevels(1).Max
                ElseIf tmpX < tudLevels(1).Min Then
                    newTUDValue = tudLevels(1).Min
                End If
                tudLevels(1).Value = newTUDValue
                
            'Highlight input node
            Case 2
                newTUDValue = tmpX * 255
                If newTUDValue < tudLevels(2).Min Then newTUDValue = tudLevels(2).Min
                tudLevels(2).Value = newTUDValue
        
        End Select
        
        'Re-enable preview updates, and refresh the screen now
        cmdBar.markPreviewStatus True
        updatePreview
        
    'Left mouse button is not down
    Else
    
        'See if the cursor is over a slider.  If it is, change the cursor to a hand.
        If isCursorOverArrow(x, True) >= 0 Then
            cMouseEventsIn.setSystemCursor IDC_HAND
        Else
            cMouseEventsIn.setSystemCursor IDC_ARROW
        End If
        
    End If

End Sub

Private Sub cMouseEventsIn_MouseUpCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal ClickEventAlsoFiring As Boolean)
    m_ActiveArrow = -1
End Sub

Private Sub cMouseEventsOut_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)

    'Check the mouse position.  If it is over a slider, activate drag mode; otherwise, ignore the click.
    If (Button And pdLeftButton) <> 0 Then
        m_ActiveArrow = isCursorOverArrow(x, False)
    End If

End Sub

Private Sub cMouseEventsOut_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)

    'Left mouse button is down, and the user has a node selected
    If ((Button And pdLeftButton) <> 0) And (m_ActiveArrow >= 3) And (m_ActiveArrow <= 4) Then
    
        'Disable automatic preview updates
        cmdBar.markPreviewStatus False
        
        Dim newTUDValue As Double
        
        'Start by recalculating the x position relative to the histogram box
        Dim tmpX As Double
        tmpX = x - m_DstArrowBoxOffset
        tmpX = tmpX / m_DstArrowBoxWidth
        
        If tmpX < 0 Then tmpX = 0
        If tmpX > 1 Then tmpX = 1
        
        'Calculate a new value for the corresponding text box
        Select Case m_ActiveArrow
        
            'Black level node
            Case 3
                newTUDValue = tmpX * 255
                If newTUDValue > 255 Then
                    newTUDValue = 255
                ElseIf newTUDValue < 0 Then
                    newTUDValue = 0
                End If
                tudLevels(3).Value = newTUDValue
                
            'White level node
            Case 4
                newTUDValue = tmpX * 255
                If newTUDValue > 255 Then
                    newTUDValue = 255
                ElseIf newTUDValue < 0 Then
                    newTUDValue = 0
                End If
                tudLevels(4).Value = newTUDValue
        
        End Select
        
        'Re-enable preview updates, and refresh the screen now
        cmdBar.markPreviewStatus True
        updatePreview
        
    'Left mouse button is not down
    Else
    
        'See if the cursor is over a slider.  If it is, change the cursor to a hand.
        If isCursorOverArrow(x, False) >= 0 Then
            cMouseEventsOut.setSystemCursor IDC_HAND
        Else
            cMouseEventsOut.setSystemCursor IDC_ARROW
        End If
        
    End If

End Sub

Private Sub cMouseEventsOut_MouseUpCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal ClickEventAlsoFiring As Boolean)
    m_ActiveArrow = -1
End Sub

'For mouse events over the input or output box, this function can be used to determine if the cursor is over a slider arrow.
Private Function isCursorOverArrow(ByVal mouseX As Long, ByVal requestIsForInputArrows As Boolean) As Long

    Dim minDistance As Double, minDistanceIndex As Long
    minDistance = picInputArrows.ScaleWidth
    minDistanceIndex = -1
    
    Dim tmpDistance As Double
    
    'Because this function handles both input and output arrows, set array bounds accordingly
    Dim loopStart As Long, loopEnd As Long
    
    If requestIsForInputArrows Then
        loopStart = 0
        loopEnd = 2
    Else
        loopStart = 3
        loopEnd = 4
    End If
    
    Dim i As Long
    For i = loopStart To loopEnd
        tmpDistance = Abs(mouseX - (m_ArrowOffsets(i) + m_DstArrowBoxOffset))
        If tmpDistance < minDistance Then
            minDistance = tmpDistance
            minDistanceIndex = i
        End If
    Next i
    
    'The mouse must be within m_ArrowHalfWidth to even be counted.
    If minDistance < m_ArrowHalfWidth + 1 Then
        isCursorOverArrow = minDistanceIndex
    Else
        isCursorOverArrow = -1
    End If

End Function

'When the shadow or highlight color is changed by the user, update the Level parameters accordingly
Private Sub csHighlight_ColorChanged()

    If cmdBar.previewsAllowed Then
    
        'Disable automatic preview updates until our calculations are done.  (If we don't do this, we get infinite recursion from
        ' the updatePreview function attempting to set our color to match the new RGB values.)
        cmdBar.markPreviewStatus False
    
        Dim r As Long, g As Long, b As Long, l As Long
        r = ExtractR(csHighlight.Color)
        g = ExtractG(csHighlight.Color)
        b = ExtractB(csHighlight.Color)
        
        'Set the internal shadow colors to match these RGB values
        If r < m_LevelValues(0, 0) + 2 Then r = m_LevelValues(0, 0) + 2
        If g < m_LevelValues(1, 0) + 2 Then g = m_LevelValues(1, 0) + 2
        If b < m_LevelValues(2, 0) + 2 Then b = m_LevelValues(2, 0) + 2
        
        m_LevelValues(0, 2) = r
        m_LevelValues(1, 2) = g
        m_LevelValues(2, 2) = b
        
        l = (r + g + b) \ 3
        If l < m_LevelValues(3, 0) + 2 Then l = m_LevelValues(3, 0) + 2
        m_LevelValues(3, 2) = l
        
        'Update the active text box to match
        tudLevels(2) = m_LevelValues(m_curChannel, 2)
        
        'Re-enable automatic preview updates
        cmdBar.markPreviewStatus True
        
        'Redraw the preview
        updatePreview
        
    End If

End Sub

Private Sub csShadow_ColorChanged()

    If cmdBar.previewsAllowed Then
    
        cmdBar.markPreviewStatus False
    
        Dim r As Long, g As Long, b As Long, l As Long
        r = ExtractR(csShadow.Color)
        g = ExtractG(csShadow.Color)
        b = ExtractB(csShadow.Color)
        
        'Set the internal shadow colors to match these RGB values
        If r > m_LevelValues(0, 2) - 2 Then r = m_LevelValues(0, 2) - 2
        If g > m_LevelValues(1, 2) - 2 Then g = m_LevelValues(1, 2) - 2
        If b > m_LevelValues(2, 2) - 2 Then b = m_LevelValues(2, 2) - 2
        
        m_LevelValues(0, 0) = r
        m_LevelValues(1, 0) = g
        m_LevelValues(2, 0) = b
        
        l = (r + g + b) \ 3
        If l > m_LevelValues(3, 2) - 2 Then l = m_LevelValues(3, 2) - 2
        m_LevelValues(3, 0) = l
        
        'Update the active text box to match
        tudLevels(0) = m_LevelValues(m_curChannel, 0)
        
        cmdBar.markPreviewStatus True
        
        'Redraw the preview
        updatePreview
        
    End If

End Sub

Private Sub Form_Activate()
    
    'Note that the user is not currently interacting with a slider node
    m_ActiveArrow = -1
    
    'Fill the histogram arrays and prepare the overlay DIBs.  To conserve resources, this is only done once,
    ' when the dialog is first loaded.
    prepHistogramOverlays
        
    'Make the RGB button pressed by default; this will be overridden by the user's last-used settings, if any exist
    m_curChannel = 3
    btsChannel.ListIndex = m_curChannel
    
    m_DisableMaxMinLimits = False
    
    'Draw the default histogram onto the histogram box
    picHistogram.Picture = LoadPicture("")
    If Not hDIB(m_curChannel) Is Nothing Then hDIB(m_curChannel).alphaBlendToDC picHistogram.hDC
    picHistogram.Picture = picHistogram.Image
        
    'Load the arrow slider images from the resource file
    Dim i As Long
    For i = 0 To 2
        Set m_Arrows(i) = New pdDIB
    Next i
    
    loadResourceToDIB "LVL_ARROW_BLK", m_Arrows(0)
    loadResourceToDIB "LVL_ARROW_GRY", m_Arrows(1)
    loadResourceToDIB "LVL_ARROW_WHT", m_Arrows(2)
    
    'Store the arrow dimensions
    m_ArrowWidth = m_Arrows(0).getDIBWidth
    m_ArrowHalfWidth = m_ArrowWidth / 2
        
    'Calculate persistent width and offset values for the arrow interaction zones.  These must extend past the left and
    ' right borders of the desired area, so that the edges of the slider images are not cropped.
    m_DstArrowBoxWidth = picHistogram.ScaleWidth
    m_DstArrowBoxOffset = picHistogram.Left - picInputArrows.Left + 1
            
    'Render sample gradients for input/output levels
    Drawing.DrawGradient picOutputGradient, RGB(0, 0, 0), RGB(255, 255, 255), True
    
    'Draw a preview image
    cmdBar.markPreviewStatus True
    updatePreview

End Sub

Private Sub prepHistogramOverlays()
    
    'Even though we don't need log-based versions of the histogram data, the master function requires arrays for both.
    ' (TODO: fix this!  Most functions need one or the other; not both.)
    Dim hData() As Double
    Dim hDataLog() As Double
    Dim hMax() As Double
    Dim hMaxLog() As Double
    Dim hMaxPosition() As Byte
    
    'Gather histogram data for the current layer
    Histogram_Analysis.fillHistogramArrays hData, hDataLog, hMax, hMaxLog, hMaxPosition
    
    'Use that data to generate DIBs for the histogram data
    Histogram_Analysis.generateHistogramImages hData, hMax, hDIB, picHistogram.ScaleWidth, picHistogram.ScaleHeight
    
End Sub

Private Sub Form_Load()

    'Prevent automatic preview refreshes until we have finished initializing the dialog
    cmdBar.markPreviewStatus False
    
    'Populate the channel selector
    btsChannel.AddItem "red", 0
    btsChannel.AddItem "green", 1
    btsChannel.AddItem "blue", 2
    btsChannel.AddItem "RGB", 3
    
    btsChannel.AssignImageToItem 0, "", Interface.GetRuntimeUIDIB(PDRUID_CHANNEL_RED, 16, 2)
    btsChannel.AssignImageToItem 1, "", Interface.GetRuntimeUIDIB(PDRUID_CHANNEL_GREEN, 16, 2)
    btsChannel.AssignImageToItem 2, "", Interface.GetRuntimeUIDIB(PDRUID_CHANNEL_BLUE, 16, 2)
    btsChannel.AssignImageToItem 3, "", Interface.GetRuntimeUIDIB(PDRUID_CHANNEL_RGB, 24, 2)
    
    'Prepare the custom input handlers
    Set cMouseEventsIn = New pdInputMouse
    Set cMouseEventsOut = New pdInputMouse
    cMouseEventsIn.addInputTracker picInputArrows.hWnd, True, True, , True
    cMouseEventsOut.addInputTracker picOutputArrows.hWnd, True, True, , True
    
    'Add button images
    cmdColorSelect(0).AssignImage "EYE_DROPPER_GENERIC"
    cmdColorSelect(1).AssignImage "EYE_DROPPER_GENERIC"
    cmdColorSelect(0).AssignTooltip "When this button is active, you can set the shadow input level color by right-clicking a color in the preview window."
    cmdColorSelect(1).AssignTooltip "When this button is active, you can set the highlight input level color by right-clicking a color in the preview window."
    cmdColorSelect(0).Value = True
    
    'Apply translations and visual themes
    MakeFormPretty Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Draw an image based on user-adjusted input and output levels
Public Sub MapImageLevels(ByRef listOfLevels As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)

    If Not toPreview Then Message "Mapping new image levels..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    
    prepImageData tmpSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Color variables
    Dim r As Long, g As Long, b As Long
        
    'Look-up table for the midtone (gamma) leveled values
    Dim gValues(0 To 255) As Double
    
    'WARNING: This next chunk of code is a lot of messy math.  Don't worry too much
    ' if you can't make sense of it ;)
    
    'Fill the gamma table with appropriate gamma values (from 10 to .1, ranged quadratically)
    ' NOTE: This table is constant, and could theoretically be loaded from file instead of generated
    ' every time we run this function.
    Dim gStep As Double
    gStep = (MAXGAMMA + MIDGAMMA) / 127
    For x = 0 To 127
        gValues(x) = (CDbl(x) / 127) * MIDGAMMA
    Next x
    For x = 128 To 255
        gValues(x) = MIDGAMMA + (CDbl(x - 127) * gStep)
    Next x
    For x = 0 To 255
        gValues(x) = 1 / ((gValues(x) + 1 / ROOT10) ^ 2)
    Next x
    
    'Parse out individual level values into a master levels array
    Dim cParams As pdParamString
    Set cParams = New pdParamString
    cParams.setParamString listOfLevels
    
    Dim levelValues(0 To 3, 0 To 4) As Double
    
    Dim i As Long
    For i = 0 To 19
        levelValues(i \ 5, i Mod 5) = cParams.GetDouble(i + 1)
    Next i
    
    'Convert the midtone ratio into a byte (so we can access a look-up table with it)
    Dim bRatio(0 To 3) As Byte
    For i = 0 To 3
        bRatio(i) = CByte(levelValues(i, 1) * 255)
    Next i
    
    'Calculate a look-up table of gamma-corrected values based on the midtones scrollbar
    Dim gLevels(0 To 3, 0 To 255) As Byte
    Dim tmpGamma As Double
    
    For i = 0 To 3
        For x = 0 To 255
            tmpGamma = CDbl(x) / 255
            tmpGamma = tmpGamma ^ (1 / gValues(bRatio(i)))
            tmpGamma = tmpGamma * 255
            If tmpGamma > 255 Then
                tmpGamma = 255
            ElseIf tmpGamma < 0 Then
                tmpGamma = 0
            End If
            gLevels(i, x) = tmpGamma
        Next x
    Next i
    
    'Look-up table for the input leveled values
    Dim newLevels(0 To 3, 0 To 255) As Byte
    
    'Fill the look-up table with appropriately mapped input limits
    Dim pStep As Double
    
    For i = 0 To 3
    
        If (levelValues(i, 2) - levelValues(i, 0)) <> 0 Then
            pStep = 255 / (levelValues(i, 2) - levelValues(i, 0))
        Else
            pStep = 1
        End If
        
        For x = 0 To 255
            If x < levelValues(i, 0) Then
                newLevels(i, x) = 0
            ElseIf x > levelValues(i, 2) Then
                newLevels(i, x) = 255
            Else
                newLevels(i, x) = ByteMe(((CDbl(x) - levelValues(i, 0)) * pStep))
            End If
        Next x
        
    Next i
        
    'Now run all input-mapped values through our midtone-correction look-up
    For i = 0 To 3
        For x = 0 To 255
            newLevels(i, x) = gLevels(i, newLevels(i, x))
        Next x
    Next i
    
    'Last of all, remap all image values to match the user-specified output limits
    Dim oStep As Double
    
    For i = 0 To 3
    
        oStep = (levelValues(i, 4) - levelValues(i, 3)) / 255
    
        For x = 0 To 255
            newLevels(i, x) = ByteMe(levelValues(i, 3) + (CDbl(newLevels(i, x)) * oStep))
        Next x
    
    Next i
    
    'Now we can finally loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Get the source pixel color values
        r = newLevels(0, ImageData(QuickVal + 2, y))
        g = newLevels(1, ImageData(QuickVal + 1, y))
        b = newLevels(2, ImageData(QuickVal, y))
        
        'Assign new values looking the lookup table
        ImageData(QuickVal + 2, y) = newLevels(3, r)
        ImageData(QuickVal + 1, y) = newLevels(3, g)
        ImageData(QuickVal, y) = newLevels(3, b)
        
    Next y
        If Not toPreview Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x
            End If
        End If
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic

End Sub

'Used to convert Long-type variables to bytes (with proper [0,255] range)
Private Function ByteMe(ByVal bVal As Long) As Byte
    If bVal > 255 Then
        ByteMe = 255
    ElseIf bVal < 0 Then
        ByteMe = 0
    Else
        ByteMe = bVal
    End If
End Function

'Used to make sure the scroll bars have appropriate limits
Private Sub FixScrollBars()
    
    If Not m_DisableMaxMinLimits Then
    
        'The black tone input level is never allowed to be > the white tone input level.
        If tudLevels(0).Max <> tudLevels(2).Value - 2 Then tudLevels(0).Max = tudLevels(2).Value - 2
        
        ' Similarly, the white tone input level is never allowed to be < the black tone input level.
        If tudLevels(2).Min <> tudLevels(0).Value + 2 Then tudLevels(2).Min = tudLevels(0).Value + 2
        
    End If
    
End Sub

Private Sub updatePreview()
    
    If cmdBar.previewsAllowed And (Not m_Arrows(0) Is Nothing) Then
        
        cmdBar.markPreviewStatus False
        
        'Erase the picture boxes
        picInputArrows.Picture = LoadPicture("")
        picOutputArrows.Picture = LoadPicture("")
        
        'Synchronize the arrow offsets with the values of the corresponding text boxes
        m_ArrowOffsets(0) = (tudLevels(0).Value / 255) * m_DstArrowBoxWidth
        m_ArrowOffsets(2) = (tudLevels(2).Value / 255) * m_DstArrowBoxWidth
        
        m_ArrowOffsets(1) = tudLevels(1).Value * (m_ArrowOffsets(2) - m_ArrowOffsets(0)) + m_ArrowOffsets(0)
        
        m_ArrowOffsets(3) = (tudLevels(3).Value / 255) * m_DstArrowBoxWidth
        m_ArrowOffsets(4) = (tudLevels(4).Value / 255) * m_DstArrowBoxWidth
        
        'Render the arrows onto their respective picture boxes
        m_Arrows(0).alphaBlendToDC picInputArrows.hDC, 255, m_ArrowOffsets(0) - m_ArrowHalfWidth + m_DstArrowBoxOffset
        m_Arrows(1).alphaBlendToDC picInputArrows.hDC, 255, m_ArrowOffsets(1) - m_ArrowHalfWidth + m_DstArrowBoxOffset
        m_Arrows(2).alphaBlendToDC picInputArrows.hDC, 255, m_ArrowOffsets(2) - m_ArrowHalfWidth + m_DstArrowBoxOffset
        m_Arrows(0).alphaBlendToDC picOutputArrows.hDC, 255, m_ArrowOffsets(3) - m_ArrowHalfWidth + m_DstArrowBoxOffset
        m_Arrows(2).alphaBlendToDC picOutputArrows.hDC, 255, m_ArrowOffsets(4) - m_ArrowHalfWidth + m_DstArrowBoxOffset
        
        picInputArrows.Picture = picInputArrows.Image
        picInputArrows.Refresh
        picOutputArrows.Picture = picOutputArrows.Image
        picOutputArrows.Refresh
                
        'Update the shadow color box to match the new level values
        Dim r As Long, g As Long, b As Long, l As Long
        r = m_LevelValues(0, 0)
        g = m_LevelValues(1, 0)
        b = m_LevelValues(2, 0)
        
        l = (r + g + b) \ 3
        l = m_LevelValues(3, 0) - l
        
        r = ByteMe(r + l)
        g = ByteMe(g + l)
        b = ByteMe(b + l)
        
        csShadow.Color = RGB(r, g, b)
        
        'Repeat the above steps for the highlight box
        r = m_LevelValues(0, 2)
        g = m_LevelValues(1, 2)
        b = m_LevelValues(2, 2)
        
        l = (r + g + b) \ 3
        l = m_LevelValues(3, 2) - l
        
        r = ByteMe(r + l)
        g = ByteMe(g + l)
        b = ByteMe(b + l)
        
        csHighlight.Color = RGB(r, g, b)
        
        cmdBar.markPreviewStatus True
        
        'Actually render the levels effect
        MapImageLevels getLevelsParamString(), True, fxPreview
        
    End If
    
End Sub

Private Sub fxPreview_ColorSelected()

    'Assign the new color to the selected box
    If cmdColorSelect(0).Value Then
        csShadow.Color = fxPreview.SelectedColor
    Else
        csHighlight.Color = fxPreview.SelectedColor
    End If

End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

Private Sub tudLevels_Change(Index As Integer)
    
    'The shadow and highlight input levels limit each other's range; when they are changed, we need to update the max or min
    ' of the opposite control.
    If (Index = 0) Or (Index = 2) Then FixScrollBars
    
    'Store the changed value in our master levels array
    m_LevelValues(m_curChannel, Index) = tudLevels(Index)
    
    'Redraw the on-screen preview
    updatePreview
    
End Sub

'Convert all channel level values into a single list, built according to PD's internal string parameter format.
Private Function getLevelsParamString() As String

    Dim tmpString As String
    tmpString = ""
    
    Dim i As Long, j As Long
    For i = 0 To 3
    For j = 0 To 4
        tmpString = tmpString & m_LevelValues(i, j)
        If (i < 3) Or (j < 4) Then tmpString = tmpString & "|"
    Next j
    Next i
    
    getLevelsParamString = tmpString
    
End Function

