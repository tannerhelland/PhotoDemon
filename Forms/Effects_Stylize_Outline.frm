VERSION 5.00
Begin VB.Form FormOutlineEffect 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Outline"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12030
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
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   802
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdColorSelector csBackground 
      Height          =   975
      Left            =   5880
      TabIndex        =   5
      Top             =   4200
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   1720
      Caption         =   "background color"
   End
   Begin PhotoDemon.pdButtonStrip btsEdgeType 
      Height          =   975
      Left            =   5880
      TabIndex        =   4
      Top             =   2040
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   1720
      Caption         =   "edge type"
   End
   Begin PhotoDemon.pdPenSelector pnsOutline 
      Height          =   1335
      Left            =   5880
      TabIndex        =   3
      Top             =   480
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   2355
      Caption         =   "outline style"
   End
   Begin PhotoDemon.pdSlider sldThreshold 
      Height          =   735
      Left            =   5880
      TabIndex        =   2
      Top             =   3240
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   1296
      Caption         =   "threshold"
      Max             =   100
      SigDigits       =   1
      GradientColorRight=   1703935
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
      DisableZoomPan  =   -1  'True
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   12030
      _ExtentX        =   21220
      _ExtentY        =   1323
   End
End
Attribute VB_Name = "FormOutlineEffect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Image Outline Effect Dialog
'Copyright 2017-2017 by Tanner Helland
'Created: 05/January/17
'Last updated: 01/August/17
'Last update: fix potential OOB error on "alpha" edge mode
'
'I actually built this algorithm for internal purposes, because it's helpful to render outlines around various
' resource PNGs to ensure they stand out against variable background colors.  Since the effect works well, I
' thought users might find it helpful too.
'
'Future feature enhancements could let the user use something other than transparency as the threshold for
' determining the contour (e.g. some base color), or a more sophisticated interpretation that also handles
' interior holes (which are ignored at present).
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Find the outline boundary of an image and paint it with a variable pen stroke
Public Sub ApplyOutlineEffect(ByVal parameterList As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    'Parse out the parameter list.  Note that not all parameters are used by all modes.
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.SetParamString parameterList
    
    Dim edgeThreshold As Double, edgeType As Long, edgeColor As Long, edgeStyle As String
    edgeThreshold = cParams.GetDouble("edge-threshold", 0#)
    edgeType = cParams.GetLong("edge-type", 0)
    edgeColor = cParams.GetLong("background-color")
    edgeStyle = cParams.GetString("pen-style")
    
    'Passed sensitivity values are on the range [0, 100].  Normalize these to [0, 255] or [0, 1]
    ' depending on the edge detection method.
    If (edgeType = 0) Then
        edgeThreshold = edgeThreshold * 2.55
        If (edgeThreshold = 255) Then edgeThreshold = 254
    Else
        edgeThreshold = edgeThreshold * 0.01
        If (edgeThreshold >= 1#) Then edgeThreshold = 0.999
    End If
    
    If (Not toPreview) Then Message "Generating image outline..."
    
    'PD uses a generic edge-detection mechanism that operates on generic byte arrays.  This offers a
    ' number of performance benefits (because we don't have to worry about edges or complex edge descriptors),
    ' but it means we have to manually generate a "1bpp" array from the image data.
    
    'For now, transparency is the only way to define an image edge.  Use the passed threshold to generate
    ' a 1bpp array that we can pass to the edge detector.
    Dim srcImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    
    EffectPrep.PrepImageData tmpSA, toPreview, dstPic, , , CBool(edgeType = 0)
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(tmpSA), 4
    
    If (Not toPreview) Then
        SetProgBarMax 2
        SetProgBarVal 0
    End If
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    Dim iWidth As Long, iHeight As Long
    iWidth = finalX - initX
    iHeight = finalY - initY
    
    'To spare our edge detector from worrying about edge pixels (which slow down processing due to obnoxious
    ' nested If/Then statements), we declare our input array with a guaranteed list of non-edge pixels on
    ' all sides.
    Dim edgeData() As Byte
    ReDim edgeData(0 To iWidth, 0 To iHeight) As Byte
    
    Dim xOffset As Long, yOffset As Long
    xOffset = -initX
    yOffset = -initY
    
    'Most of the time, edges are calculated using the image's alpha channel.
    If (edgeType = 0) Then
        
        For y = initY To finalY
        For x = initX To finalX
            If (srcImageData(x * 4 + 3, y) > edgeThreshold) Then edgeData(x + xOffset, y + yOffset) = 255
        Next x
        Next y
    
    'If the user wants us to perform color-matching, this is more complicated.
    Else
    
        Dim targetR As Long, targetG As Long, targetB As Long
        targetR = Colors.ExtractRed(edgeColor)
        targetG = Colors.ExtractGreen(edgeColor)
        targetB = Colors.ExtractBlue(edgeColor)
        
        Dim r As Long, g As Long, b As Long
        Dim rgbDistance As Long, rgbMaxDistance As Double
        rgbMaxDistance = 1# / (255# * 3#)
        
        Dim quickX As Long
        For y = initY To finalY
        For x = initX To finalX
            quickX = x * 4
            b = srcImageData(quickX, y)
            g = srcImageData(quickX + 1, y)
            r = srcImageData(quickX + 2, y)
            
            'Perform a very "quick and dirty" color comparison
            rgbDistance = Abs(r - targetR) + Abs(g - targetG) + Abs(b - targetB)
            If ((rgbDistance * rgbMaxDistance) > edgeThreshold) Then edgeData(x + xOffset, y + yOffset) = 255
            
        Next x
        Next y
    
    End If
    
    'We no longer need direct access to pixel bits
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    
    If (Not toPreview) Then SetProgBarVal 1
    
    'With an edge array successfully assembled, prepare an edge detector
    Dim cEdges As pdEdgeDetector
    Set cEdges = New pdEdgeDetector
    
    'We now need to convert our "threshold map" into an "edge only" map.  The edge detection class can
    ' do this for us, using a minesweeper-style algorithm.
    Dim finalEdgeData() As Byte
    cEdges.MakeArrayEdgeSafe edgeData, finalEdgeData, iWidth, iHeight
    
    'Run the path analyzer
    Dim finalPath As pd2DPath
    cEdges.FindAllEdges finalPath, finalEdgeData, 1, 1, iWidth + 1, iHeight + 1, -xOffset - 1, -yOffset - 1
    
    If (Not toPreview) Then SetProgBarVal 2
    
    'If we generated edges using color comparisons, premultiply alpha now
    If (edgeType <> 0) Then workingDIB.SetAlphaPremultiplication True
    
    'Use pd2D to render the outline onto the image
    Dim cPainter As pd2DPainter, cSurface As pd2DSurface, cPen As pd2DPen
    Drawing2D.QuickCreatePainter cPainter
    Drawing2D.QuickCreateSurfaceFromDIB cSurface, workingDIB, True
    
    Set cPen = New pd2DPen
    cPen.SetPenPropertiesFromXML edgeStyle
    
    cPainter.DrawPath cSurface, cPen, finalPath
    Set cPen = Nothing: Set cSurface = Nothing: Set cPainter = Nothing
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic, True

End Sub

Private Sub btsEdgeType_Click(ByVal buttonIndex As Long)
    UpdateVisibleEdgeOptions
    UpdatePreview
End Sub

Private Sub UpdateVisibleEdgeOptions()
    csBackground.Visible = CBool(btsEdgeType.ListIndex = 1)
End Sub

Private Sub cmdBar_OKClick()
    Process "Outline", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub csBackground_ColorChanged()
    UpdatePreview
End Sub

Private Sub Form_Load()
    
    cmdBar.MarkPreviewStatus False
    
    btsEdgeType.AddItem "alpha", 0
    btsEdgeType.AddItem "color", 1
    btsEdgeType.ListIndex = 0
    UpdateVisibleEdgeOptions
    
    ApplyThemeAndTranslations Me
    
    cmdBar.MarkPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then Me.ApplyOutlineEffect GetLocalParamString(), True, pdFxPreview
End Sub

Private Function GetLocalParamString() As String
    GetLocalParamString = BuildParamList("edge-threshold", sldThreshold.Value, "edge-type", btsEdgeType.ListIndex, "background-color", csBackground.Color, "pen-style", pnsOutline.Pen)
End Function

Private Sub pnsOutline_PenChanged()
    UpdatePreview
End Sub

Private Sub sldThreshold_Change()
    UpdatePreview
End Sub
