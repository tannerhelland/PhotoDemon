VERSION 5.00
Begin VB.Form FormHistogram 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   " Histogram"
   ClientHeight    =   9045
   ClientLeft      =   120
   ClientTop       =   360
   ClientWidth     =   10590
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
   ScaleHeight     =   603
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   706
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdSlider sldSmoothing 
      Height          =   375
      Left            =   7200
      TabIndex        =   8
      Top             =   6480
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      Min             =   1
      Max             =   100
      Value           =   50
      DefaultValue    =   50
   End
   Begin PhotoDemon.pdCommandBarMini cmdBar 
      Align           =   2  'Align Bottom
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   8310
      Width           =   10590
      _ExtentX        =   18680
      _ExtentY        =   1296
   End
   Begin PhotoDemon.pdPictureBox picGradient 
      Height          =   300
      Left            =   120
      Top             =   4200
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   529
   End
   Begin PhotoDemon.pdPictureBox picH 
      Height          =   3975
      Left            =   120
      Top             =   120
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   7011
   End
   Begin PhotoDemon.pdCheckBox chkLog 
      Height          =   330
      Left            =   7200
      TabIndex        =   0
      Top             =   5640
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   582
      Caption         =   "use logarithmic values"
      Value           =   0   'False
   End
   Begin PhotoDemon.pdCheckBox chkChannel 
      Height          =   330
      Index           =   0
      Left            =   4560
      TabIndex        =   1
      Top             =   5160
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   582
      Caption         =   "red"
   End
   Begin PhotoDemon.pdCheckBox chkChannel 
      Height          =   330
      Index           =   1
      Left            =   4560
      TabIndex        =   2
      Top             =   5640
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   582
      Caption         =   "green"
   End
   Begin PhotoDemon.pdCheckBox chkChannel 
      Height          =   330
      Index           =   2
      Left            =   4560
      TabIndex        =   3
      Top             =   6120
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   582
      Caption         =   "blue"
   End
   Begin PhotoDemon.pdCheckBox chkChannel 
      Height          =   330
      Index           =   3
      Left            =   4560
      TabIndex        =   4
      Top             =   6600
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   582
      Caption         =   "luminance"
      Value           =   0   'False
   End
   Begin PhotoDemon.pdCheckBox chkFillCurve 
      Height          =   330
      Left            =   7200
      TabIndex        =   5
      Top             =   5160
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   582
      Caption         =   "fill histogram curves"
   End
   Begin PhotoDemon.pdLabel lblVisibleChannels 
      Height          =   285
      Left            =   4440
      Top             =   4680
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   503
      Caption         =   "visible channels"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   0
      Left            =   240
      Top             =   4680
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   503
      Caption         =   "statistics"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblMouseInstructions 
      Height          =   450
      Left            =   360
      Top             =   7680
      Width           =   9885
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "(Note: move the mouse over the histogram to calculate these values)"
      ForeColor       =   8421504
      Layout          =   1
      UseCustomForeColor=   -1  'True
   End
   Begin PhotoDemon.pdLabel lblDrawOptions 
      Height          =   285
      Left            =   7080
      Top             =   4680
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   503
      Caption         =   "rendering options"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblValue 
      Height          =   240
      Index           =   0
      Left            =   960
      Top             =   5880
      Width           =   390
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "level"
      ForeColor       =   4194304
   End
   Begin PhotoDemon.pdLabel lblMaxCount 
      Height          =   240
      Left            =   360
      Top             =   5520
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   423
      Caption         =   "maximum count:"
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblValue 
      Height          =   240
      Index           =   1
      Left            =   960
      Top             =   6240
      Width           =   285
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "red"
      ForeColor       =   192
      Layout          =   2
      UseCustomForeColor=   -1  'True
   End
   Begin PhotoDemon.pdLabel lblValueTitle 
      Height          =   240
      Index           =   1
      Left            =   360
      Top             =   6240
      Width           =   360
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "red:"
      ForeColor       =   4210752
      Layout          =   2
   End
   Begin PhotoDemon.pdLabel lblValueTitle 
      Height          =   240
      Index           =   0
      Left            =   360
      Top             =   5880
      Width           =   465
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "level:"
      ForeColor       =   4210752
      Layout          =   2
   End
   Begin PhotoDemon.pdLabel lblTotalPixels 
      Height          =   240
      Left            =   360
      Top             =   5160
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   423
      Caption         =   "total pixels:"
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblValueTitle 
      Height          =   240
      Index           =   2
      Left            =   360
      Top             =   6600
      Width           =   570
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "green:"
      ForeColor       =   4210752
      Layout          =   2
   End
   Begin PhotoDemon.pdLabel lblValue 
      Height          =   240
      Index           =   2
      Left            =   1080
      Top             =   6600
      Width           =   495
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "green"
      ForeColor       =   32768
      Layout          =   2
      UseCustomForeColor=   -1  'True
   End
   Begin PhotoDemon.pdLabel lblValueTitle 
      Height          =   240
      Index           =   3
      Left            =   360
      Top             =   6960
      Width           =   435
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "blue:"
      ForeColor       =   4210752
      Layout          =   2
   End
   Begin PhotoDemon.pdLabel lblValue 
      Height          =   240
      Index           =   3
      Left            =   1080
      Top             =   6960
      Width           =   360
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "blue"
      ForeColor       =   12582912
      Layout          =   2
      UseCustomForeColor=   -1  'True
   End
   Begin PhotoDemon.pdLabel lblValueTitle 
      Height          =   240
      Index           =   4
      Left            =   360
      Top             =   7320
      Width           =   945
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "luminance:"
      ForeColor       =   4210752
      Layout          =   2
   End
   Begin PhotoDemon.pdLabel lblValue 
      Height          =   240
      Index           =   4
      Left            =   1560
      Top             =   7320
      Width           =   870
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "luminance"
      ForeColor       =   -2147483640
      Layout          =   2
   End
   Begin PhotoDemon.pdCheckBox chkSmooth 
      Height          =   330
      Left            =   7200
      TabIndex        =   7
      Top             =   6120
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   582
      Caption         =   "use moving average"
   End
End
Attribute VB_Name = "FormHistogram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Histogram Handler
'Copyright 2001-2026 by Tanner Helland
'Created: 6/12/01
'Last updated: 28/February/25
'Last update: merge rendering code from Histograms.GenerateHistogramImages(), so this dialog uses an identical
'             rendering strategy to the Curves, Levels, etc dialogs
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Have we generated a histogram yet?
Private m_histogramGenerated As Boolean

Private Enum PD_HistogramChannel
    hc_Red = 0
    hc_Green = 1
    hc_Blue = 2
    hc_Luminance = 3
    [hc_NumOfChannels] = 4
End Enum

#If False Then
    Private Const hc_Red = 0, hc_Green = 1, hc_Blue = 2, hc_Luminance = 3, hc_NumOfChannels = 4
#End If

'Histogram data for each particular type (r/g/b/luminance)
Private m_hData() As Long
Private m_hDataLog() As Double

'Maximum histogram values (r/g/b/luminance)
Private m_channelMax() As Long
Private m_channelMaxLog() As Double
Private m_channelMaxPosition() As Byte

'To improve histogram render performance, we cache a number of translated strings; this saves us having to re-translate them
' every time the histogram is redrawn.
Private m_strTotalPixels As String
Private m_strMaxCount As String
Private m_strRed As String, m_strGreen As String, m_strBlue As String, m_strLuminance As String
Private m_strLevel As String

'Rendering surface for the histogram and the small gradient indicator beneath the histogram
Private m_HistogramImage As pdDIB, m_HistogramGradientImage As pdDIB

'Width/height padding for the histogram image itself; mirrors values from Histograms.GenerateHistogramImages()
Private Const HIST_WIDTH_PADDING As Single = 2!
Private Const HIST_HEIGHT_PADDING As Single = 8!

'When channels are enabled or disabled, redraw the histogram
Private Sub chkChannel_Click(Index As Integer)
    DrawHistogram
End Sub

Private Sub chkFillCurve_Click()
    DrawHistogram
End Sub

Private Sub chkLog_Click()
    DrawHistogram
End Sub

'[Re]draw the histogram image.  Must be called when UI elements change (e.g. the channel checkboxes).
Public Sub DrawHistogram(Optional ByVal refreshScreen As Boolean = True)
    
    'If histogram data hasn't been generated, exit
    If (Not m_histogramGenerated) Then Exit Sub
    
    'This dialog is resizable, so dimensions are calculated at run-time
    Dim imgWidth As Long, imgHeight As Long
    imgWidth = picH.GetWidth
    imgHeight = picH.GetHeight
    
    'Initialize (as necessary) and reset the backbuffer to full transparency
    If (m_HistogramImage Is Nothing) Then Set m_HistogramImage = New pdDIB
    If (m_HistogramImage.GetDIBWidth <> imgWidth) Or (m_HistogramImage.GetDIBHeight <> imgHeight) Then m_HistogramImage.CreateBlank picH.GetWidth, picH.GetHeight, 32
    m_HistogramImage.ResetDIB 0
    m_HistogramImage.SetInitialAlphaPremultiplicationState True
    
    'We want to overlay histogram layers using custom blend modes, which means we need both...
    ' 1) a temporary image buffer (as the source surface), and...
    ' 2) a compositor object (for blending)
    Dim tmpImage As pdDIB
    Set tmpImage = New pdDIB
    tmpImage.CreateBlank m_HistogramImage.GetDIBWidth, m_HistogramImage.GetDIBHeight, 32, 0, 0
    
    Dim cCompositor As pdCompositor
    Set cCompositor = New pdCompositor
    
    'We now need to calculate a max histogram value based on which RGB channels are enabled.
    ' (Unlike other dialogs, the user can freely turn specific channels on/off here.)
    Dim maxRGB As Long, maxLum As Long
    Dim maxRGBLog As Double, maxLumLog As Double
    Dim maxChannel As PD_HistogramChannel
    
    Dim i As Long, x As Long
    For i = hc_Red To hc_Blue
        If (chkChannel(i).Value And (m_channelMax(i) > maxRGB)) Then
            maxRGB = m_channelMax(i)
            maxRGBLog = m_channelMaxLog(i)
            maxChannel = i
        End If
    Next i
    
    maxLum = m_channelMax(hc_Luminance)
    maxLumLog = m_channelMaxLog(hc_Luminance)
    
    'dstHeight is used to determine the height of the maximum value in the histogram.
    ' We want it to be slightly shorter than the height of the picture box;
    ' this way the tallest histogram value fills the entire box.
    Dim dstWidth As Single, dstHeight As Single
    dstWidth = imgWidth - HIST_WIDTH_PADDING * 2
    dstHeight = imgHeight - HIST_HEIGHT_PADDING
    
    'pd2D will be used for rendering, so we simply need to construct a polyline to hand it.
    ' If the user wants us to *fill* the histogram, we will need to add corner points to the
    ' finished line to construct a filled shape - two extra points exist so that the left and right
    ' histogram points extend to the edge of the image (so 255 + 2), plus another 2 points for the
    ' bottom two corners (255 + 2 + 2.)
    Dim listOfPoints() As PointFloat
    ReDim listOfPoints(0 To 259) As PointFloat
    
    'Build a look-up table of x-positions for the histogram data;
    ' these are equally distributed across the width of the target image
    ' (with a little room on either side for padding).
    Dim hLookupX() As Double
    ReDim hLookupX(0 To 255) As Double
    For x = 0 To 255
        hLookupX(x) = (CSng(x + 1) / 257#) * CSng(imgWidth)
    Next x
    
    'Necessary pd2D rendering objects
    Dim cSurface As pd2DSurface, cPen As pd2DPen, cBrush As pd2DBrush
    
    'We'll need to draw up to four lines - one each for red, green, blue, and luminance,
    ' depending on what channels the user has enabled in the UI.
    Dim curMax As Long, curMaxLog As Double
    Dim hType As PD_HistogramChannel, targetColor As Long
    
    'Iterate channels and render each in turn
    For hType = 0 To hc_NumOfChannels - 1

        'Only draw this histogram channel if the user requests it
        If chkChannel(hType).Value Then
            
            'Reset the temporary surface (which only holds the rendering for *this* channel)
            tmpImage.ResetDIB 0
            tmpImage.SetInitialAlphaPremultiplicationState True

            'The type of histogram we're drawing will determine the color of the histogram
            'line - we'll make it match what we're drawing (red/green/blue/black)
            Select Case hType

                Case hc_Red
                    targetColor = g_Themer.GetGenericUIColor(UI_ChannelRed)
                Case hc_Green
                    targetColor = g_Themer.GetGenericUIColor(UI_ChannelGreen)
                Case hc_Blue
                    targetColor = g_Themer.GetGenericUIColor(UI_ChannelBlue)
                Case hc_Luminance
                    targetColor = g_Themer.GetGenericUIColor(UI_GrayDark)

            End Select

            'The luminance channel is a special case - it uses its own max values, so check for that here
            If (hType = hc_Luminance) Then
                curMax = maxLum
                curMaxLog = maxLumLog
            Else
                curMax = maxRGB
                curMaxLog = maxRGBLog
            End If
            
            'Failsafe only (should never trigger)
            If (curMax < 1) Then curMax = 1
            If (curMaxLog < 1#) Then curMaxLog = 1#
            
            'Iterate through the histogram and construct a matching on-screen point for each value
            For x = 0 To 255
                listOfPoints(x + 1).x = HIST_WIDTH_PADDING + (CSng(x) * dstWidth) / 255!
                If chkLog.Value Then
                    listOfPoints(x + 1).y = HIST_HEIGHT_PADDING + (dstHeight - (m_hDataLog(hType, x) * dstHeight) / curMaxLog)
                Else
                    listOfPoints(x + 1).y = HIST_HEIGHT_PADDING + (dstHeight - (m_hData(hType, x) * dstHeight) / curMax)
                End If
            Next x
            
            'Manually populate the first and last points in the collection (staged slightly off-screen)
            listOfPoints(0).x = 0!
            listOfPoints(0).y = listOfPoints(1).y
            listOfPoints(257).x = imgWidth
            listOfPoints(257).y = listOfPoints(256).y
            
            'Apply gentle smoothing to the line to improve its visual appearance.
            ' (This better matches behavior in other editors - and looks nicer - but the user can
            '  toggle it on/off depending on what they need from the histogram.)
            Dim numOfPoints As Long
            numOfPoints = 257
            If chkSmooth.Value Then
                PDMath.SmoothLineY listOfPoints, numOfPoints, sldSmoothing.Value / 100!
                PDMath.SimplifyLine listOfPoints, numOfPoints
            End If
            
            'Re-fill the first and last points; this helps the curvature of the render correctly point
            ' in the direction of either line endpoint.
            listOfPoints(0).x = 0!
            listOfPoints(0).y = listOfPoints(1).y
            listOfPoints(numOfPoints).x = imgWidth
            listOfPoints(numOfPoints).y = listOfPoints(numOfPoints - 1).y
            numOfPoints = numOfPoints + 1
            
            'Assemble a drawing surface
            Drawing2D.QuickCreateSurfaceFromDIB cSurface, tmpImage, True
            cSurface.SetSurfacePixelOffset P2_PO_Half
            
            'If the user wants the histogram filled, render the fill prior to stroking the outline.
            ' (Note that we don't fill the luminance curve, however, since it would just be gray
            '  and blow out the other colors beneath it.)
            If chkFillCurve.Value And (hType <> hc_Luminance) Then
                    
                'Connect either end of the polyline, because we can only fill a polygon
                listOfPoints(numOfPoints).x = imgWidth + 1
                listOfPoints(numOfPoints).y = imgHeight * 2
                listOfPoints(numOfPoints + 1).x = -1!
                listOfPoints(numOfPoints + 1).y = imgHeight
                numOfPoints = numOfPoints + 2

                'Construct a matching fill brush for the target color, then fill the region beneath the curve
                Drawing2D.QuickCreateSolidBrush cBrush, targetColor, 15!
                PD2D.FillPolygonF_FromPtF cSurface, cBrush, numOfPoints, VarPtr(listOfPoints(0)), True, 0.25
                Set cBrush = Nothing

            End If

            'Stroke the histogram outline, then free all rendering objects.
            ' (They will be re-created on the next pass with new attributes.)
            Drawing2D.QuickCreateSolidPen cPen, 1!, targetColor, 100!, P2_LJ_Round, P2_LC_Round
            PD2D.DrawLinesF_FromPtF cSurface, cPen, numOfPoints, VarPtr(listOfPoints(0)), True, 0.25

            Set cSurface = Nothing
            Set cPen = Nothing

            'Merge this temporary image onto the base image using the Overlay blend mode
            ' (which prevents the colors from becoming a gray mess).
            cCompositor.QuickMergeTwoDibsOfEqualSize m_HistogramImage, tmpImage, BM_Overlay
            
        End If

    Next hType
    
    'With all chosen colors rendered, we now want to merge the result onto the UI background color.
    tmpImage.ResetDIB 0
    tmpImage.FillWithColor g_Themer.GetGenericUIColor(UI_Background), 100!
    m_HistogramImage.AlphaBlendToDC tmpImage.GetDIBDC
    Set m_HistogramImage = tmpImage
    
    'Flip the assembled image to the screen
    If refreshScreen Then picH.CopyDIB m_HistogramImage, False, True, True
    
    'Last but not least, generate the statistics at the bottom of the form.
    
    'Total number of pixels
    lblTotalPixels.Caption = m_strTotalPixels & (PDImages.GetActiveImage.Width * PDImages.GetActiveImage.Height)
    
    'Assemble the string using a pdString object
    Dim cString As pdString
    Set cString = New pdString
    With cString
        
        'Maximum value; if a color channel is enabled, list that maximum (otherwise we'll use luminance)
        If (chkChannel(0).Value Or chkChannel(1).Value Or chkChannel(2).Value) Then
            
            'Grab the maximum channel value
            curMax = m_channelMax(maxChannel)
            .Append m_strMaxCount
            .Append CStr(curMax)
            
            'Also display the name of that channel
            Select Case maxChannel
                Case 0
                    .Append " (" & m_strRed
                Case 1
                    .Append " (" & m_strGreen
                Case 2
                    .Append " (" & m_strBlue
            End Select
            
            .Append ", " & m_strLevel
            .Append " " & m_channelMaxPosition(maxChannel)
            .Append ")"
            
        'If no colors are being rendered, default to max luminance
        ElseIf chkChannel(3).Value Then
            .Append m_strMaxCount
            .Append CStr(m_channelMax(3))
            .Append " (" & m_strLuminance
            .Append ", " & m_strLevel
            .Append " " & m_channelMaxPosition(3)
            .Append ")"
        End If
        
        lblMaxCount.Caption = .ToString
        
    End With
        
End Sub

Private Sub chkSmooth_Click()
    UpdateSmoothingVisibility
    DrawHistogram
End Sub

Private Sub UpdateSmoothingVisibility()
    sldSmoothing.Visible = chkSmooth.Value
End Sub

Private Sub Form_Activate()
    UpdateSmoothingVisibility
End Sub

Private Sub Form_Deactivate()
    m_histogramGenerated = False
End Sub

Private Sub Form_Load()
    
    'Initialize a blank histogram canvas and gradient canvas (for the helper gradient that appears beneath the
    ' histogram image)
    m_histogramGenerated = False
    Set m_HistogramImage = New pdDIB
    m_HistogramImage.CreateBlank picH.GetWidth, picH.GetHeight, 32
    
    Set m_HistogramGradientImage = New pdDIB
    m_HistogramGradientImage.CreateBlank picGradient.GetWidth, picGradient.GetHeight, 32
    
    'Apply visual themes and translations
    ApplyThemeAndTranslations Me
    UpdateSmoothingVisibility
    
    'Cache the translation for several dynamic strings; this is more efficient than retranslating them over and over
    m_strTotalPixels = g_Language.TranslateMessage("total pixels") & ": "
    m_strMaxCount = g_Language.TranslateMessage("max count") & ": "
    m_strRed = g_Language.TranslateMessage("red")
    m_strGreen = g_Language.TranslateMessage("green")
    m_strBlue = g_Language.TranslateMessage("blue")
    m_strLuminance = g_Language.TranslateMessage("luminance")
    m_strLevel = g_Language.TranslateMessage("level")
    
    'Some color values need to be custom-assigned based on the current theme
    lblValue(1).ForeColor = g_Themer.GetGenericUIColor(UI_ChannelRed)
    lblValue(2).ForeColor = g_Themer.GetGenericUIColor(UI_ChannelGreen)
    lblValue(3).ForeColor = g_Themer.GetGenericUIColor(UI_ChannelBlue)
    
    'Blank out the specific level labels populated by moving the mouse across the form.
    ' (Also, align the value labels with their (potentially translated) corresponding title labels.)
    Dim i As Long
    For i = 0 To lblValue.Count - 1
        lblValue(i).SetLeft lblValueTitle(i).GetLeft + lblValueTitle(i).GetWidth + Interface.FixDPI(8)
        lblValue(i).Caption = vbNullString
    Next i
    
    If (Not m_histogramGenerated) Then TallyHistogramValues
    DrawHistogram
    
End Sub

'If the form is resized, adjust all the controls to match
Private Sub Form_Resize()

    picH.SetWidth Me.ScaleWidth - picH.GetLeft - Interface.FixDPI(8)
    picGradient.SetWidth Me.ScaleWidth - picGradient.GetLeft - Interface.FixDPI(8)
    
    'Now draw a little gradient below the histogram window, to help orient the user
    DrawHistogramGradient RGB(0, 0, 0), RGB(255, 255, 255)
    
    'Only draw the histogram if the histogram data has been initialized
    ' (This is necessary because VB triggers the Resize event before the Activate event)
    If m_histogramGenerated Then DrawHistogram
    
End Sub

'UNLOAD form
Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
    Message "Finished."
End Sub

'We'll use this routine only to draw the gradient below the histogram window.  This code is old, but it works ;)
Private Sub DrawHistogramGradient(ByVal srcColor1 As Long, ByVal srcColor2 As Long)
    
    With m_HistogramGradientImage
        If (.GetDIBWidth <> picGradient.GetWidth) Or (.GetDIBHeight <> picGradient.GetHeight) Then .CreateBlank picGradient.GetWidth, picGradient.GetHeight, 32
    End With
    
    m_HistogramGradientImage.ResetDIB 0
    
    'Create a horizontal gradient brush
    Dim boundRect As RectF
    With boundRect
        .Left = 0
        .Top = 0
        .Width = picGradient.GetWidth
        .Height = picGradient.GetHeight
    End With
    
    Dim cBrush As pd2DBrush
    Drawing2D.QuickCreateTwoColorGradientBrush cBrush, boundRect, srcColor1, srcColor2
    
    'Create a surface for the destination picture box
    Dim cSurface As pd2DSurface
    Drawing2D.QuickCreateSurfaceFromDIB cSurface, m_HistogramGradientImage, False
    
    'Fill the picture box with our constructed gradient brush
    PD2D.FillRectangleF_FromRectF cSurface, cBrush, boundRect
    
    Set cSurface = Nothing
    Set cBrush = Nothing
    
    picGradient.CopyDIB m_HistogramGradientImage, True, True, True
    
End Sub

'Build the histogram tables.  This only needs to be called once, when the image is changed. It will generate all histogram
' data for all channels (including luminance, and all log variants).
Public Sub TallyHistogramValues()

    'Notify the user that the histogram is being generated
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    
    Message "Updating histogram..."
    
    'Blank the red, green, blue, and luminance count text boxes
    Dim i As Long
    For i = 0 To lblValue.Count - 1
        lblValue(i).Caption = vbNullString
    Next i
    
    'Use our new external function to fill the important histogram arrays
    Histograms.FillHistogramArrays m_hData, m_hDataLog, m_channelMax, m_channelMaxLog, m_channelMaxPosition
    
    m_histogramGenerated = True
    Message "Finished."

End Sub

Private Sub picH_DrawMe(ByVal targetDC As Long, ByVal ctlWidth As Long, ByVal ctlHeight As Long)
    DrawHistogram True
End Sub

'When the mouse moves over the histogram, display the level and count for the histogram
'entry at the x-value over which the mouse passes
Private Sub picH_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    
    Dim xCalc As Long
    xCalc = Int((CSng(x - HIST_WIDTH_PADDING) / CSng(picH.GetWidth - HIST_WIDTH_PADDING * 2)) * 255 + 0.5)
    If (xCalc < 0) Then xCalc = 0
    If (xCalc > 255) Then xCalc = 255
    
    lblValue(0).Caption = xCalc
    lblValue(1).Caption = m_hData(0, xCalc)
    lblValue(2).Caption = m_hData(1, xCalc)
    lblValue(3).Caption = m_hData(2, xCalc)
    lblValue(4).Caption = m_hData(3, xCalc)
    
End Sub

Private Sub sldSmoothing_Change()
    DrawHistogram
End Sub
