VERSION 5.00
Begin VB.Form FormDehaze 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Dehaze"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11550
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
   ScaleWidth      =   770
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   11550
      _ExtentX        =   20373
      _ExtentY        =   1323
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
   Begin PhotoDemon.pdSlider sltRadius 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   1920
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1244
      Caption         =   "quality"
      Min             =   1
      Max             =   100
      Value           =   5
      DefaultValue    =   5
   End
   Begin PhotoDemon.pdSlider sltStrength 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   3000
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1244
      Caption         =   "strength"
      Min             =   1
      Max             =   100
      SigDigits       =   1
      Value           =   20
      DefaultValue    =   20
   End
End
Attribute VB_Name = "FormDehaze"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Image Dehazing Tool
'Copyright 2021-2021 by Tanner Helland
'Created: 08/September/21
'Last updated: 08/September/21
'Last update: initial build
'
'Automatic image dehazing is an area of active study.  It is of special interest to automative manufacturers,
' who need access to fast, high-quality dehazing algorithms for automative safety system cameras.
'
'PD's implementation is a WIP
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Apply a "dehaze" filter to an image, including automatic estimation of atmospheric lighting.
Public Sub ApplyDehaze(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
        
    If (Not toPreview) Then Message "Dehazing image..."
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    'Parameters are TBD
    Dim fxQuality As Double, blendStrength As Double
    fxQuality = cParams.GetDouble("radius", 5#)
    blendStrength = cParams.GetDouble("strength", 20#)
    
    'Generate a preview copy of the image
    Dim dstSA As SafeArray2D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    'Dehazing works by first estimating atmostpheric light.  A number of approaches are available for this;
    ' in PD, we use the quad-tree method proposed here: https://www.hindawi.com/journals/mpe/2018/9241629/
    Dim atmValue As Single
    atmValue = EstimateAtmosphericLighting(workingDIB)
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic
        
End Sub

Private Function EstimateAtmosphericLighting(ByRef srcDIB As pdDIB) As Single

    'PD's atmospheric light estimator works similar to the quad-tree method proposed here:
    ' https://www.hindawi.com/journals/mpe/2018/9241629/
    ' My implementation of their algorithm is a novel one designed against VB6's particular quirks.
    
    'Start by generating a channel-minimum map of the full image (e.g the smallest value of RGB for
    ' each pixel, regardless of channel).
    Dim minVal() As Byte
    ReDim minVal(0 To srcDIB.GetDIBWidth - 1, 0 To srcDIB.GetDIBHeight - 1) As Byte
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.GetDIBWidth - 1
    finalY = srcDIB.GetDIBHeight - 1
    
    Dim r As Long, g As Long, b As Long, cMin As Long
    
    Dim srcSA As SafeArray1D, srcPixels() As Byte
    For y = 0 To finalY
        srcDIB.WrapArrayAroundScanline srcPixels, srcSA, y
    For x = 0 To finalX
        
        b = srcPixels(x * 4)
        g = srcPixels(x * 4 + 1)
        r = srcPixels(x * 4 + 2)
        
        If (b < g) Then
            If (b < r) Then cMin = b Else cMin = r
        Else
            If (g < r) Then cMin = g Else cMin = r
        End If
        
        minVal(x, y) = cMin
        
    Next x
    Next y
    
    srcDIB.UnwrapArrayFromDIB srcPixels
    
    'With minimum values calculated, we now need to recursively subdivide the image into quadrants.
    ' For each quadrant (in each pass), find the quad with the *largest average value*, then subdivide
    ' that quad and repeat the process until some minimum threshold size is reached.
    ' (In PD's implementation, we stop when *either* the width or height reaches the minimum size.
    ' This could easily be modified, below, if different termination behavior is desired.)
    Dim minSize As Long
    minSize = 16
    
    'Start by pre-setting a target rect
    Dim curRect As RectL
    With curRect
        .Left = 0
        .Top = 0
        .Right = finalX
        .Bottom = finalY
    End With
    
    'Ensure the initial rect is a valid size; if it *isn't*, immediately return the average value.
    If ((curRect.Bottom - curRect.Top + 1) <= minSize) Or ((curRect.Right - curRect.Left + 1) < minSize) Then
        EstimateAtmosphericLighting = FindMeanOfQuad(minVal, curRect)
        Exit Function
    End If
    
    'Child rects and their means
    Dim childRects(0 To 3) As RectL
    Dim childMean(0 To 3) As Single
    
    Dim maxMean As Single, curMean As Single, testMean As Single
    
    Do
        
        'TEST ONLY: draw a rectangle around the target rect
        Dim tmpPen As pd2DPen, tmpSurface As pd2DSurface
        Drawing2D.QuickCreateSurfaceFromDIB tmpSurface, srcDIB, False
        Drawing2D.QuickCreateSolidPen tmpPen, 1!, RGB(255, 0, 0)
        PD2D.DrawRectangleI_FromRectL tmpSurface, tmpPen, curRect
        
        'Always start by ensuring the current rectangle is large enough to sub-divide.  If it doesn't,
        ' we've reached the end of the function.  (Note that, by design, this check must *not* be
        ' triggered by the whole image, because we won't have calculated an average value for the rect yet!
        ' Instead, that is treated as a special case, above, for perf reasons.)
        If ((curRect.Bottom - curRect.Top + 1) <= minSize) Or ((curRect.Right - curRect.Left + 1) < minSize) Then
            EstimateAtmosphericLighting = curMean
            Exit Do
        End If
        
        'Subdivide the current area into 4 child areas.  Rects are calculated in the following order:
        ' 0 1
        ' 2 3
        With childRects(0)
            .Left = curRect.Left
            .Right = curRect.Left + (curRect.Right - curRect.Left) \ 2
            .Top = curRect.Top
            .Bottom = curRect.Top + (curRect.Bottom - curRect.Top) \ 2
        End With
        With childRects(1)
            .Left = childRects(0).Right + 1
            .Right = curRect.Right
            .Top = curRect.Top
            .Bottom = childRects(0).Bottom
        End With
        With childRects(2)
            .Left = curRect.Left
            .Right = childRects(0).Right
            .Top = childRects(0).Bottom + 1
            .Bottom = curRect.Bottom
        End With
        With childRects(3)
            .Left = childRects(1).Left
            .Right = curRect.Right
            .Top = childRects(2).Top
            .Bottom = curRect.Bottom
        End With
        
        'Always reset the current maximum mean to an impossible value
        curMean = 0!
        
        'Iterate all sub-rects and cache the largest mean value (and associated rect)
        Dim i As Long
        For i = 0 To 3
            testMean = FindMeanOfQuad(minVal, childRects(i))
            If (testMean > curMean) Then
                curMean = testMean
                curRect = childRects(i)
            End If
        Next i
        
    'Repeat the process on the new target rect!
    Loop

End Function

'Given a source byte array and a valid (and it MUST be valid) sub-rect, return the mean value of
' entries inside said rect.
Private Function FindMeanOfQuad(ByRef srcArray() As Byte, ByRef srcRect As RectL) As Single

    Dim x As Long, y As Long, curSum As Long
    For y = srcRect.Top To srcRect.Bottom
    For x = srcRect.Left To srcRect.Right
        curSum = curSum + srcArray(x, y)
    Next x
    Next y
    
    FindMeanOfQuad = CDbl(curSum) / CDbl((srcRect.Right - srcRect.Left + 1) * (srcRect.Bottom - srcRect.Top + 1))

End Function

'OK button
Private Sub cmdBar_OKClick()
    Process "Dehaze", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub Form_Load()

    'Disable previews while we initialize everything
    cmdBar.SetPreviewStatus False
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub sltRadius_Change()
    UpdatePreview
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then ApplyDehaze GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Sub sltStrength_Change()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "radius", sltRadius.Value
        .AddParam "strength", sltStrength.Value
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
