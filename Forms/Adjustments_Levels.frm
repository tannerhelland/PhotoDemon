VERSION 5.00
Begin VB.Form FormLevels 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Levels"
   ClientHeight    =   7545
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12870
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
   ScaleHeight     =   503
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   858
   Begin PhotoDemon.pdButton cmdAutoLevels 
      Height          =   600
      Left            =   120
      TabIndex        =   1
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
      TabIndex        =   2
      Top             =   3255
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      StickyToggle    =   -1  'True
   End
   Begin PhotoDemon.pdColorSelector csShadow 
      Height          =   375
      Left            =   7230
      TabIndex        =   11
      Top             =   3255
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      curColor        =   0
      ShowMainWindowColor=   0   'False
   End
   Begin PhotoDemon.pdPictureBoxInteractive picOutputArrows 
      Height          =   360
      Left            =   5760
      Top             =   4590
      Width           =   7095
      _ExtentX        =   0
      _ExtentY        =   0
   End
   Begin PhotoDemon.pdPictureBoxInteractive picInputArrows 
      Height          =   360
      Left            =   5760
      Top             =   2790
      Width           =   7095
      _ExtentX        =   0
      _ExtentY        =   0
   End
   Begin PhotoDemon.pdPictureBox picHistogram 
      Height          =   2295
      Left            =   6000
      Top             =   480
      Width           =   6690
      _ExtentX        =   0
      _ExtentY        =   0
   End
   Begin PhotoDemon.pdPictureBox picOutputGradient 
      Height          =   375
      Left            =   6000
      Top             =   4200
      Width           =   6690
      _ExtentX        =   0
      _ExtentY        =   0
   End
   Begin PhotoDemon.pdSpinner tudLevels 
      Height          =   345
      Index           =   0
      Left            =   6000
      TabIndex        =   4
      Top             =   3270
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   714
      Max             =   253
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   6795
      Width           =   12870
      _ExtentX        =   22701
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.pdSpinner tudLevels 
      Height          =   345
      Index           =   1
      Left            =   8760
      TabIndex        =   5
      Top             =   3270
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   714
      DefaultValue    =   0.5
      Min             =   0.01
      Max             =   0.99
      SigDigits       =   2
      Value           =   0.5
   End
   Begin PhotoDemon.pdSpinner tudLevels 
      Height          =   345
      Index           =   2
      Left            =   11490
      TabIndex        =   6
      Top             =   3270
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   714
      DefaultValue    =   255
      Min             =   2
      Max             =   255
      Value           =   255
   End
   Begin PhotoDemon.pdSpinner tudLevels 
      Height          =   345
      Index           =   3
      Left            =   6000
      TabIndex        =   7
      Top             =   5070
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   714
      Max             =   255
   End
   Begin PhotoDemon.pdSpinner tudLevels 
      Height          =   345
      Index           =   4
      Left            =   11355
      TabIndex        =   8
      Top             =   5070
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   714
      DefaultValue    =   255
      Max             =   255
      Value           =   255
   End
   Begin PhotoDemon.pdColorSelector csHighlight 
      Height          =   375
      Left            =   10920
      TabIndex        =   12
      Top             =   3255
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      ShowMainWindowColor=   0   'False
   End
   Begin PhotoDemon.pdButtonStrip btsChannel 
      Height          =   1080
      Left            =   6000
      TabIndex        =   9
      Top             =   5520
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   1905
      Caption         =   "channel"
   End
   Begin PhotoDemon.pdButtonToolbox cmdColorSelect 
      Height          =   375
      Index           =   1
      Left            =   10530
      TabIndex        =   10
      Top             =   3255
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      StickyToggle    =   -1  'True
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   1
      Left            =   6000
      Top             =   3840
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   503
      Caption         =   "output levels"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   0
      Left            =   6000
      Top             =   120
      Width           =   6720
      _ExtentX        =   11853
      _ExtentY        =   503
      Caption         =   "input levels"
      FontSize        =   12
      ForeColor       =   4210752
   End
End
Attribute VB_Name = "FormLevels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Image Levels
'Copyright 2006-2026 by Tanner Helland
'Created: 22/July/06
'Last updated: 15/October/20
'Last update: migrate remaining UI to elements to PD's internal toolkit
'
'This tool allows the user to adjust image levels.  Its behavior is based off Photoshop's Levels tool,
' and identical values entered into both programs should yield a roughly identical image.
'
'Unfortunately, to perfectly mimic Photoshop's behavior, some fairly involved (i.e. incomprehensible)
' math is required.  To mitigate the speed implications of such convoluted math, a lot of look-up
' tables are used.  This makes the function fast but somewhat unreadable.  My apologies to anyone
' who needs to understand how the function works... you may have your work cut out for you.
'
'As of June '14, per-channel levels, set-by-color options, and "Auto Levels" are now supported.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Constants required for creating a gamma curve from .1 to 10
Private Const MAXGAMMA As Double = 1.8460498941512
Private Const MIDGAMMA As Double = 0.68377223398334
Private Const ROOT10 As Double = 3.16227766

'Size of level bar "nodes" in the interactive UI.
Private Const LEVEL_NODE_WIDTH As Single = 12!
Private Const LEVEL_NODE_HEIGHT As Single = 14!

'An image of the current image histogram is generated for each channel, then displayed as requested
Private m_hDIB() As pdDIB

'For convenience, the dimensions and offsets of the UI arrows are stored in these variables.
' Note that there are a total of five offsets, matching the five sliders: three for input levels
' (including midtones), two for output levels.
Private m_ArrowOffsets(0 To 4) As Single
Private m_ArrowWidth As Single, m_ArrowHalfWidth As Single
Private m_DstArrowBoxWidth As Single, m_DstArrowBoxOffset As Single

'Current channel ([0, 3] where 0 = red, 1 = green, 2 = blue, 3 = luminance)
Private m_curChannel As Long

'Because the user can change levels independently for each of Red, Green, Blue, and Luminance,
' we must store all level values internally (rather than relying on the text up/down controls
' to do it for us).  Also, because the midtone values are floating-point, we declare the whole
' tracking array as Double-type (even though shadow, highlight, and output levels are integers).
' The layout of this array is [channel R/G/B/L, level adjustment].
Private m_LevelValues(0 To 3, 0 To 4) As Double

'Persistent DIBs are stored for the back buffers of the interactive picture boxes
Private m_InputDIB As pdDIB, m_OutputDIB As pdDIB

'When the user is interacting with input or output level nodes, these values are updated to match.
' (Note that the same [0, 4] indices are used to identify these nodes; also, these are set to -1
' when no node is active/hovered.)
Private m_ActiveArrow As Long, m_HoverArrow As Long

'Node UI render helpers
Private inactiveArrowFill As pd2DBrush, activeArrowFill As pd2DBrush
Private inactiveOutlinePen As pd2DPen, activeOutlinePen As pd2DPen

'To prevent complicated interactions related to the max/min codependence of input shadow and
' highlight values, m_DisableMaxMinLimits can be used to disable automatic bounds-checking of
' input/output values.  Set this to TRUE when overwriting all on-screen level values with the
' ones stored in memory (e.g. when the user is changing the active channel, so the whole screen
' gets refreshed). When the new values have all been set, restore this to FALSE, then make a
' single call to FixScrollBars() to establish the new max/min bounds.
Private m_DisableMaxMinLimits As Boolean

'When a new channel is selected, refresh all text box values to match the new channel's stored values
Private Sub btsChannel_Click(ByVal buttonIndex As Long)

    m_curChannel = buttonIndex
    
    'Draw the relevant histogram onto the histogram box
    picHistogram.RequestRedraw
    
    'Update the text boxes to match the values for the selected channel
    UpdateTextBoxes
    
    'Update the preview.  (The preview itself doesn't actually need to be redrawn, but that function is responsible for
    ' syncing the text box values with the arrow positions.)
    UpdatePreview

End Sub

'Auto levels wil calculate new levels for the user, using the getIdealLevelParamString function below
Private Sub cmdAutoLevels_Click()
    
    'Retrieve the ideal level param string
    Dim pString As String
    pString = GetIdealLevelParamString(PDImages.GetActiveImage.GetActiveDIB)
    
    'Level value parsing will be handled via PD's standard param string parser class
    FillLevelsFromParamString pString, m_LevelValues

    'Update the text boxes to match the new values
    UpdateTextBoxes
    
    'Redraw the screen
    UpdatePreview
    
End Sub

'Returns the ideal param string for a given DIB.  "Auto levels" relies on this function to retrieve best values for a function.
' Note that PD's White Balance tool is effectively just an auto-levels function, with a variable "ignore percentage" that the
' user can set.  Similarly, the shadow/highlights tool allows for separate "ignore percentages" for shadows and highlights, but
' is otherwise effectively this same algorithm.
Public Function GetIdealLevelParamString(ByRef srcDIB As pdDIB) As String

    'Create a local array and point it at the source DIB's pixel data
    Dim imageData() As Byte, tmpSA As SafeArray1D
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = (srcDIB.GetDIBWidth - 1) * 4
    finalY = srcDIB.GetDIBHeight - 1
    
    'Color values
    Dim r As Long, g As Long, b As Long, l As Long
    
    'Maximum and minimum values, which will be detected by our initial histogram run
    Dim rMax As Byte, gMax As Byte, bMax As Byte, lMax As Byte
    Dim rMin As Byte, gMin As Byte, bMin As Byte, lMin As Byte
    rMax = 0: gMax = 0: bMax = 0: lMax = 0
    rMin = 255: gMin = 255: bMin = 255: lMin = 255
    
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
    Dim numOfPixels As Long
    
    For y = initY To finalY
        srcDIB.WrapArrayAroundScanline imageData, tmpSA, y
    For x = initX To finalX Step 4
        
        'Ignore transparent pixels, as they don't provide meaningful RGB data
        If (imageData(x + 3) <> 0) Then
            
            b = imageData(x)
            g = imageData(x + 1)
            r = imageData(x + 2)
            
            bCount(b) = bCount(b) + 1
            gCount(g) = gCount(g) + 1
            rCount(r) = rCount(r) + 1
            
            l = (218 * r + 732 * g + 74 * b) \ 1024
            lCount(l) = lCount(l) + 1
            
            numOfPixels = numOfPixels + 1
            
        End If
        
    Next x
    Next y
    
    'Safely deallocate imageData()
    srcDIB.UnwrapArrayFromDIB imageData
    
     'With the histogram complete, we can now figure out how to stretch the RGB channels. We do this by calculating a min/max
    ' ratio where the top and bottom 0.05% (or user-specified value) of pixels are ignored.
    
    Dim foundYet As Boolean
    foundYet = False
    
    Dim wbThreshold As Long
    wbThreshold = numOfPixels * percentIgnore
    
    r = 0: g = 0: b = 0: l = 0
    
    Dim rTally As Long, gTally As Long, bTally As Long, lTally As Long
    rTally = 0: gTally = 0: bTally = 0: lTally = 0
    
    'Find minimum values of red, green, blue, and luminance
    Do
        If (rCount(r) + rTally < wbThreshold) Then
            r = r + 1
            rTally = rTally + rCount(r)
        Else
            rMin = r
            foundYet = True
        End If
    Loop While foundYet = False
        
    foundYet = False
        
    Do
        If (gCount(g) + gTally < wbThreshold) Then
            g = g + 1
            gTally = gTally + gCount(g)
        Else
            gMin = g
            foundYet = True
        End If
    Loop While foundYet = False
    
    foundYet = False
    
    Do
        If (bCount(b) + bTally < wbThreshold) Then
            b = b + 1
            bTally = bTally + bCount(b)
        Else
            bMin = b
            foundYet = True
        End If
    Loop While foundYet = False
    
    foundYet = False
    
    Do
        If (lCount(l) + lTally < wbThreshold) Then
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
        If (rCount(r) + rTally < wbThreshold) Then
            r = r - 1
            rTally = rTally + rCount(r)
        Else
            rMax = r
            foundYet = True
        End If
    Loop While foundYet = False
        
    foundYet = False
        
    Do
        If (gCount(g) + gTally < wbThreshold) Then
            g = g - 1
            gTally = gTally + gCount(g)
        Else
            gMax = g
            foundYet = True
        End If
    Loop While foundYet = False
    
    foundYet = False
    
    Do
        If (bCount(b) + bTally < wbThreshold) Then
            b = b - 1
            bTally = bTally + bCount(b)
        Else
            bMax = b
            foundYet = True
        End If
    Loop While foundYet = False
    
    foundYet = False
    
    Do
        If (lCount(l) + lTally < wbThreshold) Then
            l = l - 1
            lTally = lTally + lCount(l)
        Else
            lMax = l
            foundYet = True
        End If
    Loop While foundYet = False
    
    'We now have an idealized max/min for each of red, green, blue, and luminance.
    
    'One of the problems with auto-levels is that it can introduce nasty color casts
    ' to the image.  Consider an image of a Caucasian human face; generally speaking,
    ' red tend to be exposed fairly equally across skin tones, but green and blue are
    ' much more variable according to background elements.  A face against a bright
    ' blue sky will tend to have blue concentrated at the high end of the scale,
    ' so when we auto-level it, those blue levels get spread across the full spectrum,
    ' introducing an unpleasant purplish-cast to skin tones.
    
    'To avoid this in PD (and to produce a kick-ass Auto-Level result), we split the
    ' calculated auto-level adjustment equally between per-channel corrections and net
    ' luminance corrections.  This roughly maintains the existing color spread of the
    ' image, while removing any obviously bad results, and producing a consistently
    ' well-exposed final image.  It also serves to balance out color temperature in an
    ' elegant way, without subjecting photos to the standard over-cooled look of other
    ' auto-level tools.
    rMin = rMin \ 2
    gMin = gMin \ 2
    bMin = bMin \ 2
    lMin = lMin \ 2
    
    rMax = rMax + ((255 - rMax) \ 2)
    gMax = gMax + ((255 - gMax) \ 2)
    bMax = bMax + ((255 - bMax) \ 2)
    lMax = lMax + ((255 - lMax) \ 2)
    
    'Convert the calculated values to a valid paramstring equivalent
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "redinputmin", rMin
        .AddParam "redinputmid", 0.5
        .AddParam "redinputmax", rMax
        .AddParam "redoutputmin", 0
        .AddParam "redoutputmax", 255
        
        .AddParam "greeninputmin", gMin
        .AddParam "greeninputmid", 0.5
        .AddParam "greeninputmax", gMax
        .AddParam "greenoutputmin", 0
        .AddParam "greenoutputmax", 255
        
        .AddParam "blueinputmin", bMin
        .AddParam "blueinputmid", 0.5
        .AddParam "blueinputmax", bMax
        .AddParam "blueoutputmin", 0
        .AddParam "blueoutputmax", 255
        
        .AddParam "rgbinputmin", lMin
        .AddParam "rgbinputmid", 0.5
        .AddParam "rgbinputmax", lMax
        .AddParam "rgboutputmin", 0
        .AddParam "rgboutputmax", 255
    End With
    
    GetIdealLevelParamString = cParams.GetParamString()
    
End Function

'Because the Levels dialog only uses one set of UI controls for all channels, we must manually write out preset data for each channel.
' This event will be raised whenever the command bar needs custom data from us.
Private Sub cmdBar_AddCustomPresetData()
    cmdBar.AddPresetData "MultichannelLevelData", GetLevelsParamString()
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Levels", , GetLevelsParamString(), UNDO_Layer
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
        If (m_LevelValues(i, 1) < 0.01) Then m_LevelValues(i, 1) = 0.01
        If (m_LevelValues(i, 1) > 0.99) Then m_LevelValues(i, 1) = 0.99
        
        'Set random output levels
        m_LevelValues(i, 3) = Rnd * 256
        m_LevelValues(i, 4) = Rnd * 256
    
    Next i
    
    'Update the text boxes to match the new values
    UpdateTextBoxes
    
    'Redraw the screen
    UpdatePreview

End Sub

'When a preset is loaded from file, we need to retrieve the custom levels information alongside it
Private Sub cmdBar_ReadCustomPresetData()
    
    'Retrieve a string containing all relevant layer information
    Dim tmpString As String
    tmpString = cmdBar.RetrievePresetData("MultichannelLevelData")
    
    'Valid preset data was found
    If (LenB(tmpString) <> 0) Then
    
        'Level value parsing will be handled via PD's standard param string parser class
        FillLevelsFromParamString tmpString, m_LevelValues
        
        'Update the text boxes to match the new values
        UpdateTextBoxes
        
        'Redraw the screen
        UpdatePreview
    
    'Valid preset data was *not* found, possibly because the user just upgraded from a past version of the Levels tool.
    ' Reset everything to default values
    Else
        cmdBar_ResetClick
    End If
    
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
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
    UpdateTextBoxes
    
    'Redraw the screen
    UpdatePreview
    
End Sub

'Update all text box values to match the stored values of the current channel
Private Sub UpdateTextBoxes()

    cmdBar.SetPreviewStatus False
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
    cmdBar.SetPreviewStatus True

End Sub

Private Sub cmdColorSelect_Click(Index As Integer, ByVal Shift As ShiftConstants)
    
    pdFxPreview.AllowColorSelection = (cmdColorSelect(0).Value Or cmdColorSelect(1).Value)
    
    If cmdBar.PreviewsAllowed Then
    
        cmdBar.SetPreviewStatus False
        
        'Toggle the other command button (as only one can be active at any time)
        If (Index = 0) Then
            cmdColorSelect(1).Value = False
        Else
            cmdColorSelect(0).Value = False
        End If
        
        cmdBar.SetPreviewStatus True
    
    End If

End Sub

Private Sub Form_Activate()
    m_DisableMaxMinLimits = False
    FixScrollBars
End Sub

'For mouse events over the input or output box, this function can be used to determine if the cursor is over a slider node.
' To try and "optimize" arrow selection, distance is calculated to the centerpoint of each node, and the smallest distance
' is treated as the "best" match.
Private Function IsCursorOverArrow(ByVal mouseX As Long, ByVal requestIsForInputArrows As Boolean) As Long

    Dim minDistance As Single, minDistanceIndex As Long
    minDistance = picInputArrows.GetWidth
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
        tmpDistance = Abs(mouseX - m_ArrowOffsets(i))
        If (tmpDistance < minDistance) Then
            minDistance = tmpDistance
            minDistanceIndex = i
        End If
    Next i
    
    'The mouse must be within m_ArrowHalfWidth to even be counted.
    If (minDistance <= m_ArrowHalfWidth) Then
        IsCursorOverArrow = minDistanceIndex
    Else
        IsCursorOverArrow = -1
    End If

End Function

'When the shadow or highlight color is changed by the user, update the Level parameters accordingly
Private Sub csHighlight_ColorChanged()

    If cmdBar.PreviewsAllowed Then
    
        'Disable automatic preview updates until our calculations are done.  (If we don't do this, we get infinite recursion from
        ' the updatePreview function attempting to set our color to match the new RGB values.)
        cmdBar.SetPreviewStatus False
    
        Dim r As Long, g As Long, b As Long, l As Long
        r = Colors.ExtractRed(csHighlight.Color)
        g = Colors.ExtractGreen(csHighlight.Color)
        b = Colors.ExtractBlue(csHighlight.Color)
        
        'Set the internal shadow colors to match these RGB values
        If (r < m_LevelValues(0, 0) + 2) Then r = m_LevelValues(0, 0) + 2
        If (g < m_LevelValues(1, 0) + 2) Then g = m_LevelValues(1, 0) + 2
        If (b < m_LevelValues(2, 0) + 2) Then b = m_LevelValues(2, 0) + 2
        
        m_LevelValues(0, 2) = r
        m_LevelValues(1, 2) = g
        m_LevelValues(2, 2) = b
        
        l = (r + g + b) \ 3
        If (l < m_LevelValues(3, 0) + 2) Then l = m_LevelValues(3, 0) + 2
        m_LevelValues(3, 2) = l
        
        'Update the active text box to match
        tudLevels(2) = m_LevelValues(m_curChannel, 2)
        
        'Re-enable automatic preview updates
        cmdBar.SetPreviewStatus True
        
        'Redraw the preview
        UpdatePreview
        
    End If

End Sub

Private Sub csShadow_ColorChanged()

    If cmdBar.PreviewsAllowed Then
    
        cmdBar.SetPreviewStatus False
    
        Dim r As Long, g As Long, b As Long, l As Long
        r = Colors.ExtractRed(csShadow.Color)
        g = Colors.ExtractGreen(csShadow.Color)
        b = Colors.ExtractBlue(csShadow.Color)
        
        'Set the internal shadow colors to match these RGB values
        If (r > m_LevelValues(0, 2) - 2) Then r = m_LevelValues(0, 2) - 2
        If (g > m_LevelValues(1, 2) - 2) Then g = m_LevelValues(1, 2) - 2
        If (b > m_LevelValues(2, 2) - 2) Then b = m_LevelValues(2, 2) - 2
        
        m_LevelValues(0, 0) = r
        m_LevelValues(1, 0) = g
        m_LevelValues(2, 0) = b
        
        l = (r + g + b) \ 3
        If (l > m_LevelValues(3, 2) - 2) Then l = m_LevelValues(3, 2) - 2
        m_LevelValues(3, 0) = l
        
        'Update the active text box to match
        tudLevels(0) = m_LevelValues(m_curChannel, 0)
        
        cmdBar.SetPreviewStatus True
        
        'Redraw the preview
        UpdatePreview
        
    End If

End Sub

Private Sub PrepHistogramOverlays()
    
    'Even though we don't need log-based versions of the histogram data, the central function requires arrays for both.
    ' (TODO: fix this!  Most functions need one or the other; not both.)
    Dim hData() As Long, hDataLog() As Double
    Dim hMax() As Long, hMaxLog() As Double, hMaxPosition() As Byte
    
    'Gather histogram data for the current layer
    Histograms.FillHistogramArrays hData, hDataLog, hMax, hMaxLog, hMaxPosition, True
    
    'Use that data to generate DIBs for the histogram data
    Histograms.GenerateHistogramImages hData, hMax, m_hDIB, picHistogram.GetWidth, picHistogram.GetHeight, True
    
End Sub

Private Sub Form_Load()

    'Prevent automatic preview refreshes until we have finished initializing the dialog
    m_DisableMaxMinLimits = True
    cmdBar.SetPreviewStatus False
    
    'Populate the channel selector
    btsChannel.AddItem "red", 0
    btsChannel.AddItem "green", 1
    btsChannel.AddItem "blue", 2
    btsChannel.AddItem "RGB", 3
    
    Dim btnImageSize As Long, btnImageSizeGroup As Long
    btnImageSize = Interface.FixDPI(16)
    btnImageSizeGroup = Interface.FixDPI(24)
    btsChannel.AssignImageToItem 0, , Interface.GetRuntimeUIDIB(pdri_ChannelRed, btnImageSize, 2), btnImageSize, btnImageSize
    btsChannel.AssignImageToItem 1, , Interface.GetRuntimeUIDIB(pdri_ChannelGreen, btnImageSize, 2), btnImageSize, btnImageSize
    btsChannel.AssignImageToItem 2, , Interface.GetRuntimeUIDIB(pdri_ChannelBlue, btnImageSize, 2), btnImageSize, btnImageSize
    btsChannel.AssignImageToItem 3, , Interface.GetRuntimeUIDIB(pdri_ChannelRGB, btnImageSizeGroup, 2), btnImageSizeGroup, btnImageSizeGroup
    
    'Add button images
    Dim dropperSize As Long
    dropperSize = Interface.FixDPI(16)
    cmdColorSelect(0).AssignImage "generic_dropper", , dropperSize, dropperSize
    cmdColorSelect(1).AssignImage "generic_dropper", , dropperSize, dropperSize
    cmdColorSelect(0).AssignTooltip "When this button is active, you can set the shadow input level color by right-clicking a color in the preview window."
    cmdColorSelect(1).AssignTooltip "When this button is active, you can set the highlight input level color by right-clicking a color in the preview window."
    
    'Note that the user is not currently interacting with a slider node
    m_HoverArrow = -1
    m_ActiveArrow = -1
    
    'Prep a bunch of drawing objects related to rendering the interactive nodes
    Drawing2D.QuickCreateSolidBrush inactiveArrowFill, g_Themer.GetGenericUIColor(UI_Background)
    Drawing2D.QuickCreateSolidBrush activeArrowFill, g_Themer.GetGenericUIColor(UI_AccentLight)
    Drawing2D.QuickCreateSolidPen inactiveOutlinePen, 1#, g_Themer.GetGenericUIColor(UI_GrayDark)
    Drawing2D.QuickCreateSolidPen activeOutlinePen, 1#, g_Themer.GetGenericUIColor(UI_Accent)
    
    'Fill the histogram arrays and prepare the overlay DIBs.  To conserve resources, this is only done once,
    ' when the dialog is first loaded.
    PrepHistogramOverlays
        
    'Make the RGB button pressed by default; this will be overridden by the user's last-used settings, if any exist
    m_curChannel = 3
    btsChannel.ListIndex = m_curChannel
    
    'Draw the default histogram onto the histogram box
    picHistogram.RequestRedraw
    
    'Store the arrow dimensions
    m_ArrowWidth = LEVEL_NODE_WIDTH
    m_ArrowHalfWidth = m_ArrowWidth / 2
        
    'Calculate persistent width and offset values for the arrow interaction zones.  These must extend past the left and
    ' right borders of the desired area, so that the edges of the slider images are not cropped.
    m_DstArrowBoxWidth = picHistogram.GetWidth - 2
    m_DstArrowBoxOffset = picHistogram.GetLeft - picInputArrows.GetLeft + 1
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview

End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Draw an image based on user-adjusted input and output levels
Public Sub MapImageLevels(ByRef listOfLevels As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)

    If (Not toPreview) Then Message "Mapping new image levels..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim tmpSA As SafeArray2D, tmpSA1D As SafeArray1D
    EffectPrep.PrepImageData tmpSA, toPreview, dstPic
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left * 4
    initY = curDIBValues.Top
    finalX = curDIBValues.Right * 4
    finalY = curDIBValues.Bottom
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not toPreview) Then ProgressBars.SetProgBarMax finalY
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
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
        gValues(x) = 1# / ((gValues(x) + 1# / ROOT10) ^ 2)
    Next x
    
    'Parse out individual level values into a central levels array
    Dim levelValues(0 To 3, 0 To 4) As Double
    FillLevelsFromParamString listOfLevels, levelValues
    
    'Convert the midtone ratio into a byte (so we can access a look-up table with it)
    Dim i As Long, bRatio(0 To 3) As Byte
    For i = 0 To 3
        bRatio(i) = CByte(levelValues(i, 1) * 255)
    Next i
    
    'Calculate a look-up table of gamma-corrected values based on the midtones scrollbar
    Dim gLevels(0 To 3, 0 To 255) As Byte
    Dim tmpGamma As Double
    
    For i = 0 To 3
        For x = 0 To 255
            tmpGamma = CDbl(x) / 255#
            tmpGamma = tmpGamma ^ (1# / gValues(bRatio(i)))
            tmpGamma = tmpGamma * 255
            If (tmpGamma > 255) Then tmpGamma = 255
            If (tmpGamma < 0) Then tmpGamma = 0
            gLevels(i, x) = tmpGamma
        Next x
    Next i
    
    'Look-up table for the input leveled values
    Dim newLevels(0 To 255, 0 To 3) As Byte
    
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
                newLevels(x, i) = 0
            ElseIf x > levelValues(i, 2) Then
                newLevels(x, i) = 255
            Else
                newLevels(x, i) = ByteMe(((CDbl(x) - levelValues(i, 0)) * pStep))
            End If
        Next x
        
    Next i
    
    'Now run all input-mapped values through our midtone-correction look-up
    For i = 0 To 3
        For x = 0 To 255
            newLevels(x, i) = gLevels(i, newLevels(x, i))
        Next x
    Next i
    
    'Last of all, remap all image values to match the user-specified output limits
    Dim oStep As Double
    
    For i = 0 To 3
        oStep = (levelValues(i, 4) - levelValues(i, 3)) / 255
        For x = 0 To 255
            newLevels(x, i) = ByteMe(levelValues(i, 3) + (CDbl(newLevels(x, i)) * oStep))
        Next x
    Next i
    
    'Now we can finally loop through each pixel in the image, converting values as we go
    Dim imageData() As Byte
    For y = initY To finalY
        workingDIB.WrapArrayAroundScanline imageData, tmpSA1D, y
    For x = initX To finalX Step 4
        
        'Get the source pixel color values
        b = newLevels(imageData(x), 2)
        g = newLevels(imageData(x + 1), 1)
        r = newLevels(imageData(x + 2), 0)
        
        'Assign new values looking the lookup table
        imageData(x) = newLevels(b, 3)
        imageData(x + 1) = newLevels(g, 3)
        imageData(x + 2) = newLevels(r, 3)
        
    Next x
        If (Not toPreview) Then
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal y
            End If
        End If
    Next y
    
    'Safely deallocate imageData()
    workingDIB.UnwrapArrayFromDIB imageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic

End Sub

'Used to convert Long-type variables to bytes (with proper [0,255] range)
Private Function ByteMe(ByVal bVal As Long) As Long
    ByteMe = bVal
    If (bVal > 255) Then ByteMe = 255
    If (bVal < 0) Then ByteMe = 0
End Function

'Used to make sure the scroll bars have appropriate limits
Private Sub FixScrollBars()
    
    If (Not m_DisableMaxMinLimits) Then
    
        'The black tone input level is never allowed to be > the white tone input level.
        If (tudLevels(0).Max <> tudLevels(2).Value - 2) Then tudLevels(0).Max = tudLevels(2).Value - 2
        
        ' Similarly, the white tone input level is never allowed to be < the black tone input level.
        If (tudLevels(2).Min <> tudLevels(0).Value + 2) Then tudLevels(2).Min = tudLevels(0).Value + 2
        
    End If
    
End Sub

Private Sub UpdatePreview(Optional ByVal alsoUpdateEffect As Boolean = True)
    
    If cmdBar.PreviewsAllowed And PDMain.IsProgramRunning() Then
        
        cmdBar.SetPreviewStatus False
        
        'Initialize backbuffers for the interactive picture boxes
        If (m_InputDIB Is Nothing) Then Set m_InputDIB = New pdDIB
        m_InputDIB.CreateBlank picInputArrows.GetWidth, picInputArrows.GetHeight, 32, g_Themer.GetGenericUIColor(UI_Background), 255
        If (m_OutputDIB Is Nothing) Then Set m_OutputDIB = New pdDIB
        m_OutputDIB.CreateBlank picOutputArrows.GetWidth, picOutputArrows.GetHeight, 32, g_Themer.GetGenericUIColor(UI_Background), 255
        
        'Synchronize the arrow offsets with the values of the corresponding text boxes
        ' (input levels)
        m_ArrowOffsets(0) = (tudLevels(0).Value / 255) * m_DstArrowBoxWidth + m_DstArrowBoxOffset
        m_ArrowOffsets(2) = (tudLevels(2).Value / 255) * m_DstArrowBoxWidth + m_DstArrowBoxOffset
        m_ArrowOffsets(1) = tudLevels(1).Value * (m_ArrowOffsets(2) - m_ArrowOffsets(0)) + m_ArrowOffsets(0)
        
        ' (output levels)
        m_ArrowOffsets(3) = (tudLevels(3).Value / 255) * m_DstArrowBoxWidth + m_DstArrowBoxOffset
        m_ArrowOffsets(4) = (tudLevels(4).Value / 255) * m_DstArrowBoxWidth + m_DstArrowBoxOffset
        
        'Each level node is basically comprised of three parts:
        ' 1) An upward arrowhead pointing at the node's precise position
        ' 2) a colored block representing the node's type (e.g. "shadows" vs "midtones" vs "highlights")
        ' 3) An outline encompassing (1) and (2), which is colored based on the node's hover state
        
        'To simplify things, we assemble generic paths for (1) and (2), then simply translate and draw them for each individual node.
        Dim baseArrow As pd2DPath, baseBlock As pd2DPath
        Set baseArrow = New pd2DPath
        Set baseBlock = New pd2DPath
        
        'The base arrow is centered at 0, for convenience when translating
        Dim triangleHalfWidth As Single, triangleHeight As Single
        triangleHalfWidth = (LEVEL_NODE_WIDTH / 2)
        triangleHeight = (picInputArrows.GetHeight - LEVEL_NODE_HEIGHT) - 1
        baseArrow.AddTriangle -1 * triangleHalfWidth, triangleHeight, 0, 0, triangleHalfWidth, triangleHeight
        
        'Next up is the colored block, also centered horizontally around position 0
        baseBlock.AddRectangle_Relative -1 * LEVEL_NODE_WIDTH \ 2, triangleHeight, LEVEL_NODE_WIDTH, LEVEL_NODE_HEIGHT
        
        'We also want some duplicate nodes, to remove the need to reset our base node shapes between draws
        Dim tmpArrow As pd2DPath, tmpBlock As pd2DPath
        Set tmpArrow = New pd2DPath
        Set tmpBlock = New pd2DPath
        
        'Finally, some generic scale factors to simplify the process of positioning nodes (who store their positions on the range [0, 1])
        Dim hOffset As Single, hScaleFactor As Single
        hOffset = picHistogram.GetLeft - picInputArrows.GetLeft + 1
        hScaleFactor = (picHistogram.GetWidth - 3)
        
        '...and pen/fill objects for the actual rendering
        Dim blockFill As pd2DBrush
        Set blockFill = New pd2DBrush
        blockFill.SetBrushMode P2_BM_Solid
        blockFill.SetBrushOpacity 100#
        
        Dim cSurface As pd2DSurface, cBrush As pd2DBrush
        
        'Fill the target picture boxes with the current background color
        Drawing2D.QuickCreateSolidBrush cBrush, g_Themer.GetGenericUIColor(UI_Background)
        Drawing2D.QuickCreateSurfaceFromDIB cSurface, m_InputDIB, False
        PD2D.FillRectangleF cSurface, cBrush, 0, 0, picInputArrows.GetWidth, picInputArrows.GetHeight
        Drawing2D.QuickCreateSurfaceFromDIB cSurface, m_OutputDIB, False
        PD2D.FillRectangleF cSurface, cBrush, 0, 0, picOutputArrows.GetWidth, picOutputArrows.GetHeight
        
        cSurface.SetSurfaceAntialiasing P2_AA_HighQuality
        
        'Draw each node in turn
        Dim i As Long, targetColor As Long
        For i = 0 To 4
            
            'Copy the base shapes
            tmpArrow.CloneExistingPath baseArrow
            tmpBlock.CloneExistingPath baseBlock
            
            'Translate them to this node's position
            tmpArrow.TranslatePath m_ArrowOffsets(i), 0
            tmpBlock.TranslatePath m_ArrowOffsets(i), 0
            
            'The node's colored block is rendered the same regardless of hover
            If (i = 0) Then
                targetColor = RGB(0, 0, 0)
            ElseIf (i = 1) Then
                targetColor = RGB(127, 127, 127)
            ElseIf (i = 2) Then
                targetColor = RGB(255, 255, 255)
            ElseIf (i = 3) Then
                targetColor = RGB(0, 0, 0)
            ElseIf (i = 4) Then
                targetColor = RGB(255, 255, 255)
            End If
            
            'Make sure we target the right picture box!
            If (i < 3) Then
                Drawing2D.QuickCreateSurfaceFromDIB cSurface, m_InputDIB, True
            Else
                Drawing2D.QuickCreateSurfaceFromDIB cSurface, m_OutputDIB, True
            End If
            
            blockFill.SetBrushColor targetColor
            PD2D.FillPath cSurface, blockFill, tmpBlock
            
            'The node outline and arrow fill varies by hover/active state
            If ((i = m_ActiveArrow) Or (i = m_HoverArrow)) Then
                PD2D.DrawPath cSurface, activeOutlinePen, tmpBlock
                PD2D.FillPath cSurface, activeArrowFill, tmpArrow
                PD2D.DrawPath cSurface, activeOutlinePen, tmpArrow
            Else
                PD2D.DrawPath cSurface, inactiveOutlinePen, tmpBlock
                PD2D.FillPath cSurface, inactiveArrowFill, tmpArrow
                PD2D.DrawPath cSurface, inactiveOutlinePen, tmpArrow
            End If
            
        Next i
        
        Set cSurface = Nothing: Set cBrush = Nothing
        
        'Relay changes to the screen
        picInputArrows.RequestRedraw alsoUpdateEffect
        picOutputArrows.RequestRedraw alsoUpdateEffect
                
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
        
        cmdBar.SetPreviewStatus True
        
        'Actually render the levels effect
        If alsoUpdateEffect Then MapImageLevels GetLevelsParamString(), True, pdFxPreview
        
    End If
    
End Sub

Private Sub pdFxPreview_ColorSelected()

    'Assign the new color to the selected box
    If cmdColorSelect(0).Value Then
        csShadow.Color = pdFxPreview.SelectedColor
    Else
        csHighlight.Color = pdFxPreview.SelectedColor
    End If

End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Sub picHistogram_DrawMe(ByVal targetDC As Long, ByVal ctlWidth As Long, ByVal ctlHeight As Long)
    On Error GoTo IgnoreChannelRender
    If (Not m_hDIB(m_curChannel) Is Nothing) Then m_hDIB(m_curChannel).AlphaBlendToDC targetDC
IgnoreChannelRender:
End Sub

Private Sub picInputArrows_DrawMe(ByVal targetDC As Long, ByVal ctlWidth As Long, ByVal ctlHeight As Long)
    If (Not m_InputDIB Is Nothing) Then GDI.BitBltWrapper targetDC, 0, 0, ctlWidth, ctlHeight, m_InputDIB.GetDIBDC, 0, 0, vbSrcCopy
End Sub

Private Sub picInputArrows_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)

    'Check the mouse position.  If it is over a slider, activate drag mode; otherwise, ignore the click.
    If ((Button And pdLeftButton) <> 0) Then
        m_ActiveArrow = IsCursorOverArrow(x, True)
    End If

End Sub

Private Sub picInputArrows_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    If (m_HoverArrow >= 0) Then
        m_HoverArrow = -1
        UpdatePreview False
    End If
End Sub

Private Sub picInputArrows_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)

    'If the mouse is not down, check for a hovered node
    If ((Button And pdLeftButton) = 0) Then
        Dim hoverCheck As Long
        hoverCheck = IsCursorOverArrow(x, True)
        If (hoverCheck <> m_HoverArrow) Then
            m_HoverArrow = hoverCheck
            UpdatePreview False
        End If
    End If
    
    'Left mouse button is down, and the user has a node selected
    If (((Button And pdLeftButton) <> 0) And (m_ActiveArrow >= 0) And (m_ActiveArrow <= 2)) Then
    
        'Disable automatic preview updates
        cmdBar.SetPreviewStatus False
        
        Dim newTUDValue As Double
        
        'Start by recalculating the x position relative to the histogram box
        Dim tmpX As Double
        tmpX = x - m_DstArrowBoxOffset
        tmpX = tmpX / m_DstArrowBoxWidth
        
        If (tmpX < 0) Then tmpX = 0
        If (tmpX > 1) Then tmpX = 1
        
        'Calculate a new value for the corresponding text box
        Select Case m_ActiveArrow
        
            'Shadow input node
            Case 0
                newTUDValue = tmpX * 255
                If (newTUDValue > tudLevels(0).Max) Then newTUDValue = tudLevels(0).Max
                tudLevels(0).Value = newTUDValue
            
            'Midtones input node
            Case 1
                newTUDValue = tmpX * 255
                newTUDValue = (newTUDValue - tudLevels(0).Value) / (tudLevels(2).Value - tudLevels(0).Value)
                If (newTUDValue > tudLevels(1).Max) Then
                    newTUDValue = tudLevels(1).Max
                ElseIf (tmpX < tudLevels(1).Min) Then
                    newTUDValue = tudLevels(1).Min
                End If
                tudLevels(1).Value = newTUDValue
                
            'Highlight input node
            Case 2
                newTUDValue = tmpX * 255
                If (newTUDValue < tudLevels(2).Min) Then newTUDValue = tudLevels(2).Min
                tudLevels(2).Value = newTUDValue
        
        End Select
        
        'Re-enable preview updates, and refresh the screen now
        cmdBar.SetPreviewStatus True
        UpdatePreview
        
    'Left mouse button is not down
    Else
    
        'See if the cursor is over a slider.  If it is, change the cursor to a hand.
        If (IsCursorOverArrow(x, True) >= 0) Then
            picInputArrows.SetCursorCustom IDC_HAND
        Else
            picInputArrows.SetCursorCustom IDC_ARROW
        End If
        
    End If

End Sub

Private Sub picInputArrows_MouseUpCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal clickEventAlsoFiring As Boolean, ByVal timeStamp As Long)
    m_ActiveArrow = -1
End Sub

Private Sub picOutputArrows_DrawMe(ByVal targetDC As Long, ByVal ctlWidth As Long, ByVal ctlHeight As Long)
    If (Not m_OutputDIB Is Nothing) Then GDI.BitBltWrapper targetDC, 0, 0, ctlWidth, ctlHeight, m_OutputDIB.GetDIBDC, 0, 0, vbSrcCopy
End Sub

Private Sub picOutputArrows_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)

    'Check the mouse position.  If it is over a slider, activate drag mode; otherwise, ignore the click.
    If (Button And pdLeftButton) <> 0 Then
        m_ActiveArrow = IsCursorOverArrow(x, False)
    End If
    
End Sub

Private Sub picOutputArrows_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    If (m_HoverArrow >= 0) Then
        m_HoverArrow = -1
        UpdatePreview False
    End If
End Sub

Private Sub picOutputArrows_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)

    'If the mouse is not down, check for a hovered node
    If ((Button And pdLeftButton) = 0) Then
        Dim hoverCheck As Long
        hoverCheck = IsCursorOverArrow(x, False)
        If (hoverCheck <> m_HoverArrow) Then
            m_HoverArrow = hoverCheck
            UpdatePreview False
        End If
    End If
    
    'Left mouse button is down, and the user has a node selected
    If ((Button And pdLeftButton) <> 0) And (m_ActiveArrow >= 3) And (m_ActiveArrow <= 4) Then
    
        'Disable automatic preview updates
        cmdBar.SetPreviewStatus False
        
        Dim newTUDValue As Double
        
        'Start by recalculating the x position relative to the histogram box
        Dim tmpX As Double
        tmpX = x - m_DstArrowBoxOffset
        tmpX = tmpX / m_DstArrowBoxWidth
        
        If (tmpX < 0) Then tmpX = 0
        If (tmpX > 1) Then tmpX = 1
        
        'Calculate a new value for the corresponding text box
        Select Case m_ActiveArrow
        
            'Black level node
            Case 3
                newTUDValue = tmpX * 255
                If (newTUDValue > 255) Then
                    newTUDValue = 255
                ElseIf (newTUDValue < 0) Then
                    newTUDValue = 0
                End If
                tudLevels(3).Value = newTUDValue
                
            'White level node
            Case 4
                newTUDValue = tmpX * 255
                If (newTUDValue > 255) Then
                    newTUDValue = 255
                ElseIf (newTUDValue < 0) Then
                    newTUDValue = 0
                End If
                tudLevels(4).Value = newTUDValue
        
        End Select
        
        'Re-enable preview updates, and refresh the screen now
        cmdBar.SetPreviewStatus True
        UpdatePreview
        
    'Left mouse button is not down
    Else
    
        'See if the cursor is over a slider.  If it is, change the cursor to a hand.
        If (IsCursorOverArrow(x, False) >= 0) Then
            picOutputArrows.SetCursorCustom IDC_HAND
        Else
            picOutputArrows.SetCursorCustom IDC_ARROW
        End If
        
    End If

End Sub

Private Sub picOutputArrows_MouseUpCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal clickEventAlsoFiring As Boolean, ByVal timeStamp As Long)
    m_ActiveArrow = -1
End Sub

Private Sub picOutputGradient_DrawMe(ByVal targetDC As Long, ByVal ctlWidth As Long, ByVal ctlHeight As Long)

    'Render sample gradients for input/output levels
    Dim boundsRectF As RectF
    With boundsRectF
        .Left = 0
        .Top = 0
        .Height = picOutputGradient.GetHeight
        .Width = picOutputGradient.GetWidth
    End With
    
    Dim cSurface As pd2DSurface
    Drawing2D.QuickCreateSurfaceFromDC cSurface, targetDC, False
    
    Dim cBrush As pd2DBrush
    Drawing2D.QuickCreateTwoColorGradientBrush cBrush, boundsRectF, vbBlack, vbWhite
    PD2D.FillRectangleF_FromRectF cSurface, cBrush, boundsRectF
    
    Dim cPen As pd2DPen
    Drawing2D.QuickCreateSolidPen cPen, penColor:=g_Themer.GetGenericUIColor(UI_GrayNeutral)
    
    With boundsRectF
        PD2D.DrawRectangleF cSurface, cPen, .Left, .Top, .Width - 1!, .Height - 1!
    End With
    
End Sub

Private Sub tudLevels_Change(Index As Integer)
    
    'The shadow and highlight input levels limit each other's range; when they are changed, we need to update the max or min
    ' of the opposite control.
    If (Index = 0) Or (Index = 2) Then FixScrollBars
    
    'Store the changed value in our central levels array
    m_LevelValues(m_curChannel, Index) = tudLevels(Index)
    
    'Redraw the on-screen preview
    UpdatePreview
    
End Sub

'Convert all channel level values into a single list, built according to PD's internal string parameter format.
Private Function GetLevelsParamString() As String
    
    'Remember that the layout of our central tracking array is [channel R/G/B/L, level adjustment].
    ' Level adjustment values are, in order: input min, input mid, input max, output min, output max.
    ' Private m_LevelValues(0 To 3, 0 To 4) As Double
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    Dim i As Long, thisColorName As String
    For i = 0 To 3
    
        'Determine a name for this type of level adjustment
        If (i = 0) Then
            thisColorName = "red"
        ElseIf (i = 1) Then
            thisColorName = "green"
        ElseIf (i = 2) Then
            thisColorName = "blue"
        ElseIf (i = 3) Then
            thisColorName = "rgb"
        End If
        
        With cParams
            .AddParam thisColorName & "inputmin", m_LevelValues(i, 0)
            .AddParam thisColorName & "inputmid", m_LevelValues(i, 1)
            .AddParam thisColorName & "inputmax", m_LevelValues(i, 2)
            .AddParam thisColorName & "outputmin", m_LevelValues(i, 3)
            .AddParam thisColorName & "outputmax", m_LevelValues(i, 4)
        End With
        
    Next i
    
    GetLevelsParamString = cParams.GetParamString()
    
End Function

'Given an XML param string, fill the m_LevelValues() array with the stored param string values
Private Sub FillLevelsFromParamString(ByVal paramString As String, ByRef dstLevels() As Double)
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString paramString
    
    Dim i As Long, thisColorName As String
    For i = 0 To 3
    
        'Determine a name for this type of level adjustment
        If (i = 0) Then
            thisColorName = "red"
        ElseIf (i = 1) Then
            thisColorName = "green"
        ElseIf (i = 2) Then
            thisColorName = "blue"
        ElseIf (i = 3) Then
            thisColorName = "rgb"
        End If
        
        With cParams
            dstLevels(i, 0) = .GetDouble(thisColorName & "inputmin", 0#)
            dstLevels(i, 1) = .GetDouble(thisColorName & "inputmid", 0.5)
            dstLevels(i, 2) = .GetDouble(thisColorName & "inputmax", 255#)
            dstLevels(i, 3) = .GetDouble(thisColorName & "outputmin", 0#)
            dstLevels(i, 4) = .GetDouble(thisColorName & "outputmax", 255#)
        End With
        
    Next i
    
End Sub
