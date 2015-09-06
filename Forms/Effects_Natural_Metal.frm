VERSION 5.00
Begin VB.Form FormMetal 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Metal"
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
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   12030
      _ExtentX        =   21220
      _ExtentY        =   1323
      BackColor       =   14802140
   End
   Begin PhotoDemon.sliderTextCombo sltRadius 
      Height          =   720
      Left            =   6000
      TabIndex        =   2
      Top             =   1680
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "smoothness"
      Max             =   200
      SigDigits       =   1
      Value           =   20
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.sliderTextCombo sltDetail 
      Height          =   720
      Left            =   6000
      TabIndex        =   3
      Top             =   600
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "detail"
      Max             =   16
      Value           =   4
      NotchPosition   =   2
      NotchValueCustom=   4
   End
   Begin PhotoDemon.colorSelector csHighlight 
      Height          =   975
      Left            =   6000
      TabIndex        =   4
      Top             =   2760
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1720
      Caption         =   "highlight color"
      curColor        =   14737632
   End
   Begin PhotoDemon.colorSelector csShadow 
      Height          =   975
      Left            =   6000
      TabIndex        =   5
      Top             =   3960
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1720
      Caption         =   "shadow color"
      curColor        =   4210752
   End
End
Attribute VB_Name = "FormMetal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'"Metal" or "Chrome" Image effect
'Copyright 2002-2015 by Tanner Helland
'Created: sometime 2002
'Last updated: 04/April/15
'Last update: rewrite function from scratch
'
'PhotoDemon's "Metal" filter is the rough equivalent of "Chrome" in Photoshop.  Our implementation is relatively
' straightforward; a normalized graymap is created for the image, then remapped according to a sinusoidal-like
' lookup table (created using the pdFilterLUT class).
'
'The user currently has control over two parameters: "smoothness", which determines a pre-effect blur radius,
' and "detail" which controls the number of octaves in the lookup table.
'
'Still TODO: allow the user to set a highlight and shadow color, instead of using boring ol' gray
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Apply a metallic "shimmer" to an image
Public Sub ApplyMetalFilter(ByVal steelDetail As Long, ByVal steelSmoothness As Double, Optional ByVal shadowColor As Long = 0, Optional ByVal highlightColor As Long = vbWhite, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If Not toPreview Then Message "Pouring smoldering metal onto image..."
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    
    'If this is a preview, we need to adjust the smoothness (kernel radius) to match the size of the preview box
    If toPreview Then steelSmoothness = steelSmoothness * curDIBValues.previewModifier
    
    'Decompose the shadow and highlight colors into their individual color components
    Dim rShadow As Long, gShadow As Long, bShadow As Long
    Dim rHighlight As Long, gHighlight As Long, bHighlight As Long
    
    rShadow = ExtractR(shadowColor)
    gShadow = ExtractG(shadowColor)
    bShadow = ExtractB(shadowColor)
    
    rHighlight = ExtractR(highlightColor)
    gHighlight = ExtractG(highlightColor)
    bHighlight = ExtractB(highlightColor)
    
    'Retrieve a normalized luminance map of the current image
    Dim grayMap() As Byte
    DIB_Handler.getDIBGrayscaleMap workingDIB, grayMap, True
    
    'If the user specified a non-zero smoothness, apply it now
    If steelSmoothness > 0 Then Filters_ByteArray.GaussianBlur_IIR_ByteArray grayMap, workingDIB.getDIBWidth, workingDIB.getDIBHeight, steelSmoothness, 3
        
    'Re-normalize the data (this ends up not being necessary, but it could be exposed to the user in a future update)
    'Filters_ByteArray.normalizeByteArray grayMap, workingDIB.getDIBWidth, workingDIB.getDIBHeight
    
    'Next, we need to generate a sinusoidal octave lookup table for the graymap.  This causes the luminance of the map to
    ' vary evently between the number of detail points requested by the user.
    
    'Detail cannot be lower than 2, but it is presented to the user as [0, (arbitrary upper bound)], so add two to the total now
    steelDetail = steelDetail + 2
    
    'We will be using pdFilterLUT to generate corresponding RGB lookup tables, which means we need to use POINTFLOAT arrays
    Dim rCurve() As POINTFLOAT, gCurve() As POINTFLOAT, bCurve() As POINTFLOAT
    ReDim rCurve(0 To steelDetail) As POINTFLOAT
    ReDim gCurve(0 To steelDetail) As POINTFLOAT
    ReDim bCurve(0 To steelDetail) As POINTFLOAT
    
    'For all channels, X values are evenly distributed from 0 to 255
    Dim i As Long
    For i = 0 To steelDetail
        rCurve(i).x = CDbl(i / steelDetail) * 255
        gCurve(i).x = CDbl(i / steelDetail) * 255
        bCurve(i).x = CDbl(i / steelDetail) * 255
    Next i
    
    'Y values alternate between the shadow and highlight colors; these are calculated on a per-channel basis
    For i = 0 To steelDetail
        
        If i Mod 2 = 0 Then
            rCurve(i).y = rShadow
            gCurve(i).y = gShadow
            bCurve(i).y = bShadow
        Else
            rCurve(i).y = rHighlight
            gCurve(i).y = gHighlight
            bCurve(i).y = bHighlight
        End If
        
    Next i
    
    'Convert our point array into color curves
    Dim rLookup() As Byte, gLookUp() As Byte, bLookup() As Byte
    
    Dim cLUT As pdFilterLUT
    Set cLUT = New pdFilterLUT
    cLUT.fillLUT_Curve rLookup, rCurve
    cLUT.fillLUT_Curve gLookUp, gCurve
    cLUT.fillLUT_Curve bLookup, bCurve
        
    'We are now ready to apply the final curve to the image!
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(dstSA), 4
    
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
    
    Dim grayVal As Long
    
    'Apply the filter
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
        
        grayVal = grayMap(x, y)
        
        ImageData(QuickVal, y) = bLookup(grayVal)
        ImageData(QuickVal + 1, y) = gLookUp(grayVal)
        ImageData(QuickVal + 2, y) = rLookup(grayVal)
        
    Next y
        If (x And progBarCheck) = 0 Then
            If userPressedESC() Then Exit For
            SetProgBarVal x
        End If
    Next x
        
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingDIB
    finalizeImageData toPreview, dstPic
            
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Metal", , buildParams(sltDetail, sltRadius, csShadow.Color, csHighlight.Color), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltRadius.Value = 20
    sltDetail.Value = 4
    csShadow.Color = RGB(30, 30, 30)
    csHighlight.Color = RGB(230, 230, 230)
End Sub

Private Sub csHighlight_ColorChanged()
    updatePreview
End Sub

Private Sub csShadow_ColorChanged()
    updatePreview
End Sub

Private Sub Form_Activate()
    
    'Apply translations and visual themes
    makeFormPretty Me
    
    'Draw an initial preview of the effect
    updatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub updatePreview()
    If cmdBar.previewsAllowed Then ApplyMetalFilter sltDetail.Value, sltRadius.Value, csShadow.Color, csHighlight.Color, True, fxPreview
End Sub

Private Sub sltDetail_Change()
    updatePreview
End Sub

Private Sub sltRadius_Change()
    updatePreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub


