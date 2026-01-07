VERSION 5.00
Begin VB.Form FormExposure 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Exposure"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11685
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
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   779
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdPictureBox picChart 
      Height          =   2415
      Left            =   8400
      Top             =   240
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   4260
   End
   Begin PhotoDemon.pdSlider sltExposure 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   2880
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1244
      Caption         =   "exposure compensation (stops)"
      Min             =   -5
      Max             =   5
      SigDigits       =   2
      SliderTrackStyle=   2
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.pdSlider sltOffset 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   3720
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1244
      Caption         =   "offset"
      Min             =   -1
      Max             =   1
      SigDigits       =   2
   End
   Begin PhotoDemon.pdSlider sltGamma 
      Height          =   705
      Left            =   6000
      TabIndex        =   4
      Top             =   4560
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1244
      Caption         =   "gamma"
      Min             =   0.01
      Max             =   2
      SigDigits       =   2
      Value           =   1
      NotchPosition   =   2
      NotchValueCustom=   1
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   1005
      Index           =   2
      Left            =   6000
      Top             =   1320
      Width           =   2220
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "new exposure curve:"
      FontSize        =   12
      ForeColor       =   4210752
      Layout          =   1
   End
End
Attribute VB_Name = "FormExposure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Exposure Dialog
'Copyright 2013-2026 by Tanner Helland
'Created: 13/July/13
'Last updated: 28/April/20
'Last update: overhaul live chart display to use pd2D; minor perf improvements
'
'Basic image exposure adjustment dialog.  Exposure is a complex topic in photography, and (obviously) the best way to
' adjust it is at image capture time.  This is because true exposure relies on a number of variables (see
' https://en.wikipedia.org/wiki/Exposure_%28photography%29) inherent in the scene itself, with a technical definition
' of "the accumulated physical quantity of visible light energy applied to a surface during a given exposure time."
' Once a set amount of light energy has been applied to a digital sensor and the resulting photo is captured, actual
' exposure can never fully be corrected or adjusted in post-production.
'
'That said, in the event that a poor choice is made at time of photography, certain approximate adjustments can be
' applied in post-production, with the understanding that missing shadows and highlights cannot be "magically"
' recreated out of thin air.  This is done by approximating an EV adjustment using a simple power-of-two formula.
' For more information on exposure compensation, see
' https://en.wikipedia.org/wiki/Exposure_value#Exposure_compensation_in_EV
'
'Also, I have mixed feelings about dumping brightness and gamma corrections on this dialog, but Photoshop does it,
' so we may as well, too.  (They can always be ignored if you just want "pure" exposure correction.)
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Exposure curve image
Private m_Graph As pdDIB

'Adjust an image's exposure.
' PRIMARY INPUT: exposureAdjust represents the number of stops to correct the image.  Each stop corresponds to a power-of-2
'                 increase (+values) or decrease (-values) in luminance.  Thus an EV of -1 will cut the amount of light in
'                 half, while an EV of +1 will double the amount of light.
Public Sub Exposure(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    If (Not toPreview) Then Message "Adjusting image exposure..."
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim exposureAdjust As Double, offsetAdjust As Double, gammaAdjust As Double
    exposureAdjust = cParams.GetDouble("exposure", 0#)
    offsetAdjust = cParams.GetDouble("offset", 0#)
    gammaAdjust = cParams.GetDouble("gamma", 1#)
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte, tmpSA As SafeArray2D, tmpSA1D As SafeArray1D
    EffectPrep.PrepImageData tmpSA, toPreview, dstPic
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not toPreview) Then ProgressBars.SetProgBarMax finalY
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'Exposure can be easily applied using a look-up table
    Dim gLookup(0 To 255) As Byte
    For x = 0 To 255
        gLookup(x) = GetCorrectedValue(x, 255, exposureAdjust, offsetAdjust, gammaAdjust)
    Next x
    
    'Loop through each pixel in the image, converting values as we go
    initX = initX * 4
    finalX = finalX * 4
    
    For y = initY To finalY
        workingDIB.WrapArrayAroundScanline imageData, tmpSA1D, y
    For x = initX To finalX Step 4
        
        'Apply a new value based on the lookup table
        imageData(x) = gLookup(imageData(x))
        imageData(x + 1) = gLookup(imageData(x + 1))
        imageData(x + 2) = gLookup(imageData(x + 2))
        
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

Private Function GetCorrectedValue(ByVal inputVal As Single, ByVal inputMax As Single, ByVal newExposure As Single, ByVal newOffset As Single, ByVal newGamma As Single) As Double
    
    Dim tmpCalculation As Double
    
    'Convert incoming value to the [0, 1] scale
    tmpCalculation = inputVal / inputMax
    
    'Apply exposure (simple power-of-two calculation)
    tmpCalculation = tmpCalculation * 2# ^ (newExposure)
    
    'Apply offset (brightness)
    tmpCalculation = tmpCalculation + newOffset
    
    'Apply gamma
    If (newGamma < 0.01) Then newGamma = 0.01
    If (tmpCalculation > 0#) Then tmpCalculation = tmpCalculation ^ (1# / newGamma)
    
    'Return to the original [0, inputMax] scale
    tmpCalculation = tmpCalculation * inputMax
    
    'Apply clipping
    If (tmpCalculation < 0#) Then tmpCalculation = 0#
    If (tmpCalculation > inputMax) Then tmpCalculation = inputMax
    
    GetCorrectedValue = tmpCalculation
    
End Function

Private Sub cmdBar_OKClick()
    Process "Exposure", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltGamma.Value = 1#
End Sub

Private Sub Form_Load()
    cmdBar.SetPreviewStatus False
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Redrawing a preview of the exposure effect also redraws the exposure curve (which isn't really a curve, but oh well)
Private Sub UpdatePreview()
    
    If cmdBar.PreviewsAllowed And sltExposure.IsValid And sltOffset.IsValid And sltGamma.IsValid And (Not g_Themer Is Nothing) Then
    
        Dim prevX As Double, prevY As Double
        Dim curX As Double, curY As Double
        Dim x As Long
        
        If (m_Graph Is Nothing) Then Set m_Graph = New pdDIB
        
        Dim xWidth As Long, yHeight As Long
        xWidth = picChart.GetWidth
        yHeight = picChart.GetHeight
        If (m_Graph.GetDIBWidth <> xWidth) Or (m_Graph.GetDIBHeight <> yHeight) Then m_Graph.CreateBlank xWidth, yHeight, 32, vbWhite, 255
        
        'pd2D handles rendering duties
        Dim cSurface As pd2DSurface
        Set cSurface = New pd2DSurface
        cSurface.WrapSurfaceAroundPDDIB m_Graph
        cSurface.SetSurfaceAntialiasing P2_AA_None
        cSurface.SetSurfacePixelOffset P2_PO_Normal
        
        'We first want to wipe the old chart, then draw a gray line across the diagonal for reference
        Dim cBrush As pd2DBrush
        Set cBrush = New pd2DBrush
        cBrush.SetBrushColor g_Themer.GetGenericUIColor(UI_Background)
        PD2D.FillRectangleI cSurface, cBrush, 0, 0, xWidth - 1, yHeight - 1
        
        Dim cPen As pd2DPen
        Set cPen = New pd2DPen
        cPen.SetPenWidth 1!
        cPen.SetPenColor g_Themer.GetGenericUIColor(UI_GrayNeutral)
        PD2D.DrawRectangleI cSurface, cPen, 0, 0, xWidth - 1, yHeight - 1
        
        cSurface.SetSurfaceAntialiasing P2_AA_HighQuality
        cSurface.SetSurfacePixelOffset P2_PO_Half
        PD2D.DrawLineI cSurface, cPen, 0, yHeight, xWidth, 0
        
        'Draw the corresponding exposure curve (line, actually) for this EV
        Dim expVal As Double, offsetVal As Double, gammaVal As Double, tmpVal As Double
        expVal = sltExposure
        offsetVal = sltOffset
        gammaVal = sltGamma
        
        cPen.SetPenColor g_Themer.GetGenericUIColor(UI_Accent)
        cPen.SetPenWidth 1.6!
        cPen.SetPenLineJoin P2_LJ_Round
        cPen.SetPenLineCap P2_LC_Round
        
        'Shrink the chart by two pixels (to account for borders); we will add 1 to
        ' all coordinates in the inner loop to ensure the chart is properly centered
        yHeight = yHeight - 2
        xWidth = xWidth - 2
        
        prevX = 0
        prevY = yHeight
        curX = 0
        curY = yHeight
        
        Dim listOfPoints() As PointFloat, numOfPoints As Long
        ReDim listOfPoints(0 To xWidth) As PointFloat
        
        For x = 0 To xWidth
            
            'Get the corrected, clamped exposure value
            tmpVal = GetCorrectedValue(x, xWidth, expVal, offsetVal, gammaVal)
            
            'Because the graph may not be square, we also need to multiply the returned value
            ' by the graph's aspect ratio
            tmpVal = tmpVal * (yHeight / xWidth)
            
            'Invert this final value, because screen coordinates are upside-down
            tmpVal = yHeight - tmpVal
            
            'Clip the current points to the boundaries of the image, then add it to the running list
            curY = tmpVal
            curX = x
            If (x = 0) Then prevY = curY
            If (curY > yHeight - 1) Then curY = yHeight - 1
            
            listOfPoints(numOfPoints).x = curX + 1
            listOfPoints(numOfPoints).y = curY + 1
            numOfPoints = numOfPoints + 1
            
            prevX = curX
            prevY = curY
            
        Next x
        
        'Draw the finished line
        PD2D.DrawLinesF_FromPtF cSurface, cPen, numOfPoints, VarPtr(listOfPoints(0)), False
        Set cSurface = Nothing
        picChart.RequestRedraw True
        
        'Finally, apply the exposure correction to the preview image
        Me.Exposure GetLocalParamString(), True, pdFxPreview
        
    End If
    
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Sub picChart_DrawMe(ByVal targetDC As Long, ByVal ctlWidth As Long, ByVal ctlHeight As Long)
    If (Not m_Graph Is Nothing) Then GDI.BitBltWrapper targetDC, 0, 0, ctlWidth, ctlHeight, m_Graph.GetDIBDC, 0, 0, vbSrcCopy
End Sub

'Update the preview whenever the combination slider/text control has its value changed
Private Sub sltExposure_Change()
    UpdatePreview
End Sub

Private Sub sltGamma_Change()
    UpdatePreview
End Sub

Private Sub sltOffset_Change()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "exposure", sltExposure
        .AddParam "offset", sltOffset
        .AddParam "gamma", sltGamma
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
