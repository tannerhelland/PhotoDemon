VERSION 5.00
Begin VB.Form FormLevels 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Adjust Image Levels"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   195
   ClientWidth     =   12180
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
   ScaleHeight     =   435
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   812
   ShowInTaskbar   =   0   'False
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
      ScaleWidth      =   425
      TabIndex        =   12
      Top             =   4710
      Width           =   6375
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
      ScaleWidth      =   425
      TabIndex        =   11
      Top             =   2790
      Width           =   6375
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
      ScaleWidth      =   391
      TabIndex        =   10
      Top             =   480
      Width           =   5895
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
      ScaleWidth      =   391
      TabIndex        =   9
      Top             =   4320
      Width           =   5895
   End
   Begin PhotoDemon.textUpDown tudShadows 
      Height          =   405
      Left            =   6000
      TabIndex        =   4
      Top             =   3120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   714
      Max             =   253
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5775
      Width           =   12180
      _ExtentX        =   21484
      _ExtentY        =   1323
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.textUpDown tudMidtones 
      Height          =   405
      Left            =   8400
      TabIndex        =   5
      Top             =   3120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   714
      Min             =   0.01
      Max             =   0.99
      SigDigits       =   2
      Value           =   0.5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.textUpDown tudHighlights 
      Height          =   405
      Left            =   10560
      TabIndex        =   6
      Top             =   3120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   714
      Min             =   2
      Max             =   255
      Value           =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.textUpDown tudBlackPoint 
      Height          =   405
      Left            =   6000
      TabIndex        =   7
      Top             =   5040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   714
      Max             =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.textUpDown tudWhitePoint 
      Height          =   405
      Left            =   10560
      TabIndex        =   8
      Top             =   5040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   714
      Max             =   255
      Value           =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Top             =   3960
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
'Copyright ©2006-2014 by Tanner Helland
'Created: 22/July/06
'Last updated: 06/June/14
'Last update: redesigned the entire dialog to better match Level dialogs in other photo editors.
'
'This tool allows the user to adjust image levels.  Its behavior is based off Photoshop's Levels tool, and identical
' values entered into both programs should yield an identical image.
'
'Unfortunately, to perfectly mimic Photoshop's behavior, some fairly involved (i.e. incomprehensible) math is required.
' To mitigate the speed implications of such convoluted math, a number of look-up tables are used.  This makes the
' function quite fast, but at a hit to readability.  My apologies to anyone trying to understand how the function works.
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

'These five arrays will hold histogram data for the current image.  They are filled when the form is activated, and
' not modified again unless the form is unloaded and reopened.
Private hData() As Double
Private hDataLog() As Double
Private hMax() As Double
Private hMaxLog() As Double
Private hMaxPosition() As Byte

'An image of the current image histogram is drawn once each for regular and logarithmic, then stored to these DIBs.
Private hDIB(0 To 3) As pdDIB, hLogDIB(0 To 3) As pdDIB

'Copies of the "slider arrows" used to display and control input/output level manipulation
Private m_Arrows(0 To 2) As pdDIB

'For convenience, the dimensions and offsets of the UI arrows are stored in these variables.  Note that there are two
' extra offsets relative to the arrow DIBs themselves; this is because we render two copies of the black and white
' arrows to the screen, one each for input and output levels.
Private m_ArrowOffsets(0 To 4) As Long
Private m_ArrowWidth As Long, m_ArrowHalfWidth As Long, m_ArrowHeight As Long
Private m_DstArrowBoxWidth As Long, m_DstArrowBoxOffset As Long

'Two special input classes are required; one each for the input and output arrow boxes
Private WithEvents cMouseEventsIn As pdInput
Attribute cMouseEventsIn.VB_VarHelpID = -1
Private WithEvents cMouseEventsOut As pdInput
Attribute cMouseEventsOut.VB_VarHelpID = -1

'If the user is using the mouse to slide nodes around, these values will be used to store the node's index
Private m_ActiveArrow As Long

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'OK button
Private Sub cmdBar_OKClick()
    Process "Levels", , buildParams(tudShadows.Value, tudMidtones.Value, tudHighlights.Value, tudBlackPoint.Value, tudWhitePoint.Value), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    
    'Set the output levels to (0-255)
    tudBlackPoint.Value = 0
    tudWhitePoint.Value = 255
    
    'Set the input levels to (0-255)
    tudShadows.Value = 0
    tudHighlights.Value = 255
    FixScrollBars
    
    'Set the midtone level to default (0.5)
    tudMidtones.Value = 0.5
    
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
                If newTUDValue > tudShadows.Max Then newTUDValue = tudShadows.Max
                tudShadows.Value = newTUDValue
            
            'Midtones input node
            Case 1
                newTUDValue = tmpX * 255
                newTUDValue = (newTUDValue - tudShadows.Value) / (tudHighlights.Value - tudShadows.Value)
                If newTUDValue > tudMidtones.Max Then
                    newTUDValue = tudShadows.Max
                ElseIf tmpX < tudMidtones.Min Then
                    newTUDValue = tudMidtones.Min
                End If
                tudMidtones.Value = newTUDValue
                
            'Highlight input node
            Case 2
                newTUDValue = tmpX * 255
                If newTUDValue < tudHighlights.Min Then newTUDValue = tudHighlights.Min
                tudHighlights.Value = newTUDValue
        
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
                tudBlackPoint.Value = newTUDValue
                
            'White level node
            Case 4
                newTUDValue = tmpX * 255
                If newTUDValue > 255 Then
                    newTUDValue = 255
                ElseIf newTUDValue < 0 Then
                    newTUDValue = 0
                End If
                tudWhitePoint.Value = newTUDValue
        
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

Private Sub Form_Activate()
    
    'Note that the user is not currently interacting with a slider node
    m_ActiveArrow = -1
    
    'Fill the histogram arrays and prepare the overlay DIBs.  To conserve resources, this is only done once,
    ' when the dialog is first loaded.
    prepHistogramOverlays
    
    'Draw the relevant histogram onto the histogram box
    Dim m_curChannel As Long
    m_curChannel = 3
    BitBlt picHistogram.hDC, 1, 0, hDIB(m_curChannel).getDIBWidth, hDIB(m_curChannel).getDIBHeight, hDIB(m_curChannel).getDIBDC, 0, 0, vbSrcCopy
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
    
    'Prepare the custom input handlers
    Set cMouseEventsIn = New pdInput
    Set cMouseEventsOut = New pdInput
    
    cMouseEventsIn.addInputTracker picInputArrows.hWnd, True, True, , True
    cMouseEventsOut.addInputTracker picOutputArrows.hWnd, True, True, , True
    
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
    'Render sample gradients for input/output levels
    Drawing.DrawGradient picOutputGradient, RGB(0, 0, 0), RGB(255, 255, 255), True
    
    'Draw a preview image
    cmdBar.markPreviewStatus True
    updatePreview

End Sub

Private Sub prepHistogramOverlays()

    'Gather histogram data for the current layer
    fillHistogramArrays hData, hDataLog, hMax, hMaxLog, hMaxPosition

    Dim yMax As Double
    Dim hLookupX() As Double
    Dim hColor As Long
    
    'Initialize the background histogram image DIBs
    Dim i As Long, j As Long
    
    For i = 0 To 3
    
        'Initialize this channel's DIB
        Set hDIB(i) = New pdDIB
        Set hLogDIB(i) = New pdDIB
        hDIB(i).createBlank picHistogram.ScaleWidth - 2, picHistogram.ScaleHeight
        hLogDIB(i).createFromExistingDIB hDIB(i)
        
        yMax = 0.9 * hDIB(i).getDIBHeight
        
        'Build a look-up table of x-positions for the histogram data
        ReDim hLookupX(0 To 255) As Double
        
        For j = 0 To 255
            hLookupX(j) = (CDbl(j) / 255) * hDIB(i).getDIBWidth
        Next j
        
        'The color of the histogram changes for each channel
        Select Case i
        
            'Red
            Case 0
                hColor = RGB(255, 60, 80)
            
            'Green
            Case 1
                hColor = RGB(60, 210, 80)
            
            'Blue
            Case 2
                hColor = RGB(60, 100, 255)
            
            'Luminance
            Case 3
                hColor = RGB(192, 192, 192)
        
        
        End Select
        
        'Render the histogram data to each DIB (one for regular, one for logarithmic)
        For j = 1 To 255
            GDIPlusDrawLineToDC hDIB(i).getDIBDC, hLookupX(j - 1), hDIB(i).getDIBHeight - (hData(i, j - 1) / hMax(i)) * yMax, hLookupX(j), hDIB(i).getDIBHeight - (hData(i, j) / hMax(i)) * yMax, hColor, 255
            GDIPlusDrawLineToDC hLogDIB(i).getDIBDC, hLookupX(j - 1), hDIB(i).getDIBHeight - (hDataLog(i, j - 1) / hMaxLog(i)) * yMax, hLookupX(j), hDIB(i).getDIBHeight - (hDataLog(i, j) / hMaxLog(i)) * yMax, hColor, 255
        Next j
        
        'Beneath each line, add an even lighter "filled" version of the line
        For j = 0 To 255
            GDIPlusDrawLineToDC hDIB(i).getDIBDC, hLookupX(j), hDIB(i).getDIBHeight - (hData(i, j) / hMax(i)) * yMax - 1, hLookupX(j), hDIB(i).getDIBHeight, hColor, 128
            GDIPlusDrawLineToDC hLogDIB(i).getDIBDC, hLookupX(j), hDIB(i).getDIBHeight - (hDataLog(i, j) / hMaxLog(i)) * yMax - 1, hLookupX(j), hDIB(i).getDIBHeight, hColor, 128
        Next j
    
    Next i

End Sub

Private Sub Form_Load()

    'Prevent automatic preview refreshes until we have finished initializing the dialog
    cmdBar.markPreviewStatus False

End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Draw an image based on user-adjusted input and output levels
Public Sub MapImageLevels(ByVal inLLimit As Long, ByVal inMLimit As Double, ByVal inRLimit As Long, ByVal outLLimit As Long, ByVal outRLimit As Long, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)

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
    
    'Convert the midtone ratio into a byte (so we can access a look-up table with it)
    Dim bRatio As Byte
    bRatio = CByte(inMLimit * 255)
    
    'Calculate a look-up table of gamma-corrected values based on the midtones scrollbar
    Dim gLevels(0 To 255) As Byte
    Dim tmpGamma As Double
    For x = 0 To 255
        tmpGamma = CDbl(x) / 255
        tmpGamma = tmpGamma ^ (1 / gValues(bRatio))
        tmpGamma = tmpGamma * 255
        If tmpGamma > 255 Then
            tmpGamma = 255
        ElseIf tmpGamma < 0 Then
            tmpGamma = 0
        End If
        gLevels(x) = tmpGamma
    Next x
    
    'Look-up table for the input leveled values
    Dim newLevels(0 To 255) As Byte
    
    'Fill the look-up table with appropriately mapped input limits
    Dim pStep As Double
    pStep = 255 / (CSng(inRLimit) - CSng(inLLimit))
    For x = 0 To 255
        If x < inLLimit Then
            newLevels(x) = 0
        ElseIf x > inRLimit Then
            newLevels(x) = 255
        Else
            newLevels(x) = ByteMe(((CSng(x) - CSng(inLLimit)) * pStep))
        End If
    Next x
    
    'Now run all input-mapped values through our midtone-correction look-up
    For x = 0 To 255
        newLevels(x) = gLevels(newLevels(x))
    Next x
    
    'Last of all, remap all image values to match the user-specified output limits
    Dim oStep As Double
    oStep = (CSng(outRLimit) - CSng(outLLimit)) / 255
    For x = 0 To 255
        newLevels(x) = ByteMe(CSng(outLLimit) + (CSng(newLevels(x)) * oStep))
    Next x
    
    'Now we can finally loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Assign new values looking the lookup table
        ImageData(QuickVal + 2, y) = newLevels(r)
        ImageData(QuickVal + 1, y) = newLevels(g)
        ImageData(QuickVal, y) = newLevels(b)
        
    Next y
        If toPreview = False Then
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
    
    'The black tone input level is never allowed to be > the white tone input level.
    If tudShadows.Max <> tudHighlights.Value - 2 Then tudShadows.Max = tudHighlights.Value - 2
    
    ' Similarly, the white tone input level is never allowed to be < the black tone input level.
    If tudHighlights.Min <> tudShadows.Value + 2 Then tudHighlights.Min = tudShadows.Value + 2
    
End Sub

Private Sub updatePreview()
    
    If cmdBar.previewsAllowed Then
        
        'Erase the picture boxes
        picInputArrows.Picture = LoadPicture("")
        picOutputArrows.Picture = LoadPicture("")
        
        'Synchronize the arrow offsets with the values of the corresponding text boxes
        m_ArrowOffsets(0) = (tudShadows.Value / 255) * m_DstArrowBoxWidth
        m_ArrowOffsets(2) = (tudHighlights.Value / 255) * m_DstArrowBoxWidth
        
        m_ArrowOffsets(1) = tudMidtones.Value * (m_ArrowOffsets(2) - m_ArrowOffsets(0)) + m_ArrowOffsets(0)
        
        m_ArrowOffsets(3) = (tudBlackPoint.Value / 255) * m_DstArrowBoxWidth
        m_ArrowOffsets(4) = (tudWhitePoint.Value / 255) * m_DstArrowBoxWidth
        
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
        
        'Actually render the levels effect
        MapImageLevels tudShadows.Value, tudMidtones.Value, tudHighlights.Value, tudBlackPoint.Value, tudWhitePoint.Value, True, fxPreview
        
    End If
    
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

Private Sub tudBlackPoint_Change()
    updatePreview
End Sub

Private Sub tudHighlights_Change()
    FixScrollBars
    updatePreview
End Sub

Private Sub tudMidtones_Change()
    updatePreview
End Sub

Private Sub tudShadows_Change()
    FixScrollBars
    updatePreview
End Sub

Private Sub tudWhitePoint_Change()
    updatePreview
End Sub
