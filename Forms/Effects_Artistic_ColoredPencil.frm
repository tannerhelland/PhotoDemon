VERSION 5.00
Begin VB.Form FormPencil 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Colored pencil"
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
   Begin VB.ComboBox cboStyle 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   6000
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1740
      Width           =   5895
   End
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
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.sliderTextCombo sltRadius 
      Height          =   720
      Left            =   6000
      TabIndex        =   2
      Top             =   2520
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "tip radius"
      Min             =   1
      Max             =   100
      Value           =   5
   End
   Begin PhotoDemon.sliderTextCombo sltIntensity 
      Height          =   720
      Left            =   6000
      TabIndex        =   3
      Top             =   3600
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   1270
      Caption         =   "pressure"
      Min             =   -100
      Max             =   200
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "style"
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
      Index           =   2
      Left            =   6000
      TabIndex        =   4
      Top             =   1320
      Width           =   480
   End
End
Attribute VB_Name = "FormPencil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Pencil Sketch Image Effect
'Copyright 2001-2015 by Tanner Helland
'Created: sometime 2001
'Last updated: 23/July/14
'Last update: overhauled algorithm, gave tool its own dialog
'
'PhotoDemon has provided a pencil sketch tool for a long time, but despite going through many incarnations, it always
' used low-quality, "quick and dirty" approximations.
'
'In July '14, this changed, and the entire tool was rethought from the ground up.  A dialog is now provided, with options
' for pencil style, tip thickness, and stroke pressure.  This yields much more flexible results, and the use of PD's
' central compositor for overlaying various image copies keeps things nice and fast.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Apply a "colored pencil" effect to an image
'Inputs:
' 1) radius of the pencil tip (min 1, no real max - but processing speed obviously drops as the radius increases)
' 2) color intensity, which controls the vibrance applied to the resulting color
' 3) pencil style, a nebulous setting that controls blend mode and post-processing, among other items.  Current values include:
'    0 - normal
'    1 - luminous
'    2 - pastel
'    3 - graphite
Public Sub fxColoredPencil(ByVal penRadius As Long, ByVal colorIntensity As Double, ByVal pencilStyle As Long, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If Not toPreview Then Message "Sketching image with pencils..."
    
    'Reverse the intensity input; this way, positive values make the image more vibrant.  Negative values make it less vibrant.
    ' Note that the adjustment also varies by pencil style; typically it's used as a vibrance adjustment, but in some modes,
    ' we switch it out for gamma or contrast control.
    Select Case pencilStyle
    
        Case 0, 1
            colorIntensity = -0.01 * colorIntensity
            
        Case 2, 3
            colorIntensity = (302 - (colorIntensity + 101)) / 300
            
    End Select
    
    'More color variables - in this case, sums for each color component
    Dim r As Long, g As Long, b As Long, maxVal As Long
    Dim amtVal As Double, avgVal As Double
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    
    'Create a copy of the image.  "Colored pencil" requires a blurred image copy as part of the effect, and we maintain that
    ' copy seprate from the original (as the two must be blended as the final step of the filter).
    Dim blurDIB As pdDIB
    Set blurDIB = New pdDIB
    blurDIB.createFromExistingDIB workingDIB
    
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    'If this is a preview, we need to adjust the kernel radius to match the size of the preview box
    If toPreview Then
        penRadius = penRadius * curDIBValues.previewModifier
        
    'If this is not a preview, initialize the main program progress bar
    Else
        
        SetProgBarMax finalY * 3 + finalX * 5
        
        Dim progBarCheck As Long
        progBarCheck = findBestProgBarValue()
        
    End If
    
    If penRadius < 1 Then penRadius = 1
    
    'Start by creating the blurred DIB
    If CreateApproximateGaussianBlurDIB(penRadius, workingDIB, blurDIB, 3, toPreview, finalY * 3 + finalX * 5) Then
        
        Dim progBarOffset As Long
        progBarOffset = finalY * 3 + finalX * 3
        
        'Now that we have a gaussian DIB created in blurDIB, we can point arrays toward it and the source DIB
        Dim srcImageData() As Byte
        Dim srcSA As SAFEARRAY2D
        prepSafeArray srcSA, blurDIB
        CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
        'These values will help us access locations in the array more quickly.
        ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
        Dim QuickVal As Long, qvDepth As Long
        qvDepth = curDIBValues.BytesPerPixel
                
        'Build a look-up table of grayscale values (faster than calculating it manually for each pixel)
        Dim grayLookUp() As Byte
        ReDim grayLookUp(0 To 765) As Byte
        
        For x = 0 To 765
            grayLookUp(x) = x \ 3
        Next x
                
        'Invert the source DIB, and optionally, apply grayscale as well
        For x = initX To finalX
            QuickVal = x * qvDepth
        For y = initY To finalY
            
            r = srcImageData(QuickVal + 2, y)
            g = srcImageData(QuickVal + 1, y)
            b = srcImageData(QuickVal, y)
            
            'Normally, we invert the raw pixel data only...
            If pencilStyle <> 1 Then
                srcImageData(QuickVal + 2, y) = 255 - r
                srcImageData(QuickVal + 1, y) = 255 - g
                srcImageData(QuickVal, y) = 255 - b
            
            '...but for the "luminous" color mode, we also convert the image to grayscale
            Else
                g = grayLookUp(r + g + b)
                srcImageData(QuickVal + 2, y) = 255 - g
                srcImageData(QuickVal + 1, y) = 255 - g
                srcImageData(QuickVal, y) = 255 - g
            End If
            
        Next y
            If Not toPreview Then
                If (x And progBarCheck) = 0 Then
                    If userPressedESC() Then Exit For
                    SetProgBarVal progBarOffset + x
                End If
            End If
        Next x
            
        'Release our array copy
        CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
        
        'Apply premultiplication to the layers prior to compositing
        blurDIB.setAlphaPremultiplication True
        workingDIB.setAlphaPremultiplication True
        
        'A pdCompositor class will help us blend the invert+blur image back onto the original image
        Dim cComposite As pdCompositor
        Set cComposite = New pdCompositor
        
        'Composite our invert+blur image against the base layer (workingDIB) using the COLOR DODGE blend mode;
        ' this will emphasize areas where the layers differ, while ignoring areas where they're the same.
        Dim tmpLayerTop As pdLayer, tmpLayerBottom As pdLayer
        Set tmpLayerTop = New pdLayer
        Set tmpLayerBottom = New pdLayer
        
        tmpLayerTop.InitializeNewLayer PDL_IMAGE, , blurDIB
        tmpLayerBottom.InitializeNewLayer PDL_IMAGE, , workingDIB
        
        'Normally we use the "color dodge" blend mode, but for "pastel" mode, we switch to linear dodge for a
        ' lighter, softer appearance
        If pencilStyle <> 2 Then
            tmpLayerTop.setLayerBlendMode BL_COLORDODGE
        Else
            tmpLayerTop.setLayerBlendMode BL_LINEARDODGE
        End If
        tmpLayerTop.setLayerOpacity 100
        
        cComposite.mergeLayers tmpLayerTop, tmpLayerBottom, True
        
        'Copy the finished DIB from the bottom layer back into workingDIB
        workingDIB.createFromExistingDIB tmpLayerBottom.layerDIB
        
        'Remove premultiplied alpha
        workingDIB.setAlphaPremultiplication False
        
        'Release our temporary layers and DIBs
        Set blurDIB = Nothing
        Set tmpLayerTop = Nothing
        Set tmpLayerBottom = Nothing
        
        'Some modes requires post-production gamma correction.  Build a lookup table now.
        Dim gammaTable() As Byte
        ReDim gammaTable(0 To 255) As Byte
        
        If (pencilStyle = 2) Or (pencilStyle = 3) Then
        
            Dim tmpVal As Double
            
            For x = 0 To 255
                tmpVal = x / 255
                tmpVal = tmpVal ^ (1 / colorIntensity)
                tmpVal = tmpVal * 255
                
                If tmpVal > 255 Then tmpVal = 255
                If tmpVal < 0 Then tmpVal = 0
        
                gammaTable(x) = tmpVal
            Next x
        
        End If
        
        'Point our byte array at workingDIB, so we can apply a final vibrance pass using the specified color intensity
        prepSafeArray srcSA, workingDIB
        CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
        progBarOffset = finalY * 3 + finalX * 4
        
        'Adjust vibrance
        For x = initX To finalX
            QuickVal = x * qvDepth
        For y = initY To finalY
            
            r = srcImageData(QuickVal + 2, y)
            g = srcImageData(QuickVal + 1, y)
            b = srcImageData(QuickVal, y)
                
            'Calculate the gray value using different methods for each pencil style
            Select Case pencilStyle
            
                Case 0, 1
            
                    avgVal = grayLookUp(r + g + b)
                    maxVal = Max3Int(r, g, b)
                    
                    'Calculate a vibrance-adjusted average, using the gray as our base
                    amtVal = ((Abs(maxVal - avgVal) / 127) * colorIntensity)
                    
                    If r <> maxVal Then r = r + (maxVal - r) * amtVal
                    If g <> maxVal Then g = g + (maxVal - g) * amtVal
                    If b <> maxVal Then b = b + (maxVal - b) * amtVal
                    
                    'Clamp values to [0,255] range
                    If r < 0 Then r = 0
                    If r > 255 Then r = 255
                    If g < 0 Then g = 0
                    If g > 255 Then g = 255
                    If b < 0 Then b = 0
                    If b > 255 Then b = 255
                
                Case 2
                    r = gammaTable(r)
                    g = gammaTable(g)
                    b = gammaTable(b)
                    
                Case 3
                    r = gammaTable(grayLookUp(r + g + b))
                    g = r
                    b = r
                
            End Select
            
            srcImageData(QuickVal + 2, y) = r
            srcImageData(QuickVal + 1, y) = g
            srcImageData(QuickVal, y) = b
            
        Next y
            If Not toPreview Then
                If (x And progBarCheck) = 0 Then
                    If userPressedESC() Then Exit For
                    SetProgBarVal progBarOffset + x
                End If
            End If
        Next x
        
        'Release our array once more
        CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
        Erase srcImageData
        
    End If
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
    
End Sub

Private Sub cboStyle_Click()
    updatePreview
End Sub

Private Sub cmdBar_OKClick()
    Process "Colored pencil", , buildParams(sltRadius, sltIntensity, cboStyle.ListIndex), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltRadius.Value = 3
End Sub

Private Sub Form_Activate()
    
    'Apply translations and visual themes
    MakeFormPretty Me
        
    'Draw a preview of the effect
    cmdBar.markPreviewStatus True
    updatePreview
    
End Sub

Private Sub Form_Load()
    
    'Disable previews until the dialog is fully loaded
    cmdBar.markPreviewStatus False
    
    'Populate the style drop-down
    cboStyle.Clear
    cboStyle.AddItem "Normal", 0
    cboStyle.AddItem "Luminous", 1
    cboStyle.AddItem "Pastel", 2
    cboStyle.AddItem "Graphite", 3
    
    cboStyle.ListIndex = 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub sltIntensity_Change()
    updatePreview
End Sub

Private Sub sltRadius_Change()
    updatePreview
End Sub

'Render a new effect preview
Private Sub updatePreview()
    If cmdBar.previewsAllowed Then fxColoredPencil sltRadius, sltIntensity, cboStyle.ListIndex, True, fxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

