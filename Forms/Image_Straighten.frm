VERSION 5.00
Begin VB.Form FormStraighten 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Straighten image"
   ClientHeight    =   6540
   ClientLeft      =   -15
   ClientTop       =   225
   ClientWidth     =   12105
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
   ScaleWidth      =   807
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   12105
      _ExtentX        =   21352
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
      DisableZoomPan  =   -1  'True
   End
   Begin PhotoDemon.sliderTextCombo sltAngle 
      Height          =   720
      Left            =   6000
      TabIndex        =   2
      Top             =   2160
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "angle"
      Min             =   -15
      Max             =   15
      SigDigits       =   2
   End
End
Attribute VB_Name = "FormStraighten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Image Straightening Interface
'Copyright 2014-2015 by Tanner Helland
'Created: 11/May/14
'Last updated: 11/May/14
'Last update: initial build, based heavily off PD's existing Rotate dialog
'
'This tool allows the user to straighten an image at an arbitrary angle in 1/100 degree increments.  FreeImage is
' required for the tool to work, as this relies upon FreeImage to perform the rotation in a fast, efficient
' manner.  The corresponding menu entry for this tool is hidden unless FreeImage is found.  (To confuse matters
' further, GDI+ is used to enlarge the straightened image.)
'
'At present, the tool assumes that you want to straighten the image around its center.  I don't have plans to
' change this behavior.
'
'To straighten a layer instead of the entire image, use the Layer -> Orientation -> Straighten menu.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'This temporary DIB will be used for rendering the preview
Private smallDIB As pdDIB

'This dialog can be used to resize the full image, or a single layer.  The requested target will be stored here,
' and can be externally accessed by the ResizeTarget property.
Private m_StraightenTarget As PD_ACTION_TARGET

Public Property Let StraightenTarget(newTarget As PD_ACTION_TARGET)
    m_StraightenTarget = newTarget
End Property

Public Sub StraightenImage(ByVal rotationAngle As Double, Optional ByVal thingToRotate As PD_ACTION_TARGET = PD_AT_WHOLEIMAGE, Optional ByVal isPreview As Boolean = False)
        
    'If the image contains an active selection, disable it before transforming the canvas
    If (thingToRotate = PD_AT_WHOLEIMAGE) And pdImages(g_CurrentImage).selectionActive And (Not isPreview) Then
        pdImages(g_CurrentImage).selectionActive = False
        pdImages(g_CurrentImage).mainSelection.lockRelease
    End If

    'FreeImage uses positive values to indicate counter-clockwise rotation.  While mathematically correct, I find this
    ' unintuitive for casual users.  PD reverses the rotationAngle value so that POSITIVE values indicate CLOCKWISE rotation.
    rotationAngle = -rotationAngle

    Dim tmpDIB As pdDIB, finalDIB As pdDIB
    Set tmpDIB = New pdDIB
    Set finalDIB = New pdDIB

    'Double-check that FreeImage exists
    If g_ImageFormats.FreeImageEnabled Then
        
        'Rotation requires quite a few variables, including a number of handles for passing data back-and-forth with FreeImage.
        Dim fi_DIB As Long, returnDIB As Long
        Dim nWidth As Double, nHeight As Double
        
        'One of the FreeImage rotation variants requires an explicit center point; calculate one in advance.
        Dim cx As Double, cy As Double
        
        'To solve the problem of auto-cropping the straightened image, additional variables are required.
        Dim solveAngle As Double, len1 As Double, len2 As Double, scaleFactor As Double
        
        If isPreview Then
            cx = smallDIB.getDIBWidth / 2
            cy = smallDIB.getDIBHeight / 2
        Else
        
            Select Case thingToRotate
            
                Case PD_AT_WHOLEIMAGE
                    cx = pdImages(g_CurrentImage).Width / 2
                    cy = pdImages(g_CurrentImage).Height / 2
                    
                Case PD_AT_SINGLELAYER
                    cx = pdImages(g_CurrentImage).getActiveDIB.getDIBWidth / 2
                    cy = pdImages(g_CurrentImage).getActiveDIB.getDIBHeight / 2
                    
            End Select
                    
        End If
        
        Dim sourceCropWidth As Double, sourceCropHeight As Double
        
        'Normally, I like to use identical code for previews and actual effects.  However, rotating is completely different
        ' for previews (where we do a single rotation of the composited image) vs the full images (independently rotating
        ' each layer, with support functions to null-pad and crop layers as necessary).  As such, there is some code
        ' duplication here, but I believe it makes the code much more readable.
        
        If isPreview Then
            
            'Give FreeImage a handle to our temporary rotation image
            fi_DIB = FreeImage_CreateFromDC(smallDIB.getDIBDC)
            
            'Ask it to rotate the image
            returnDIB = FreeImage_RotateEx(fi_DIB, rotationAngle, 0, 0, cx, cy, True)
            
            'As a failsafe, check the returned width/height (they should be identical our original input)
            nWidth = FreeImage_GetWidth(returnDIB)
            nHeight = FreeImage_GetHeight(returnDIB)
                        
            'Create a blank DIB to receive the rotated image from FreeImage
            tmpDIB.createBlank nWidth, nHeight, 32
            
            'Ask FreeImage to premultiply the image's alpha data
            FreeImage_PreMultiplyWithAlpha returnDIB
            
            'Copy the bits from the FreeImage DIB to our DIB
            SetDIBitsToDevice tmpDIB.getDIBDC, 0, 0, nWidth, nHeight, 0, 0, 0, nHeight, ByVal FreeImage_GetBits(returnDIB), ByVal FreeImage_GetInfo(returnDIB), 0&
            
            'With the transfer complete, release the FreeImage DIB and unload the library
            If fi_DIB <> 0 Then FreeImage_UnloadEx fi_DIB
            If returnDIB <> 0 Then FreeImage_UnloadEx returnDIB
            
            'Next, we need to calculate a scaling factor for the image.  Straightening applies a sort of auto-crop
            ' to the image to remove empty corners; by solving a triangle equation using the image diagonal, we
            ' can calculate the scaling factor needed.  Thank you to this article for the helpful diagram:
            ' http://stackoverflow.com/questions/18865837/image-straightening-in-android
            ' (Note that the stackoverflow link does not work for the case of width > height, and the instructions
            '  provided for correcting that case are *wrong*!)
            If nWidth < nHeight Then
                solveAngle = Atn(nHeight / nWidth)
                len1 = cx / Cos(solveAngle - Abs(rotationAngle * PI_DIV_180))
            Else
                solveAngle = Atn(nWidth / nHeight)
                len1 = cy / Cos(solveAngle - Abs(rotationAngle * PI_DIV_180))
            End If
            
            len2 = Sqr(cx * cx + cy * cy)
            scaleFactor = len2 / len1
            
            'Using our new scalefactor, calculate a source image width and height
            sourceCropWidth = nWidth * (1 / scaleFactor)
            sourceCropHeight = nHeight * (1 / scaleFactor)
            
            'Prepare a final DIB to receive the resized image
            finalDIB.createBlank nWidth, nHeight, 32, 0
            
            'Use GDI+ to copy the relevant source rectangle into the final DIB
            GDIPlusResizeDIB finalDIB, 0, 0, nWidth, nHeight, tmpDIB, (nWidth - sourceCropWidth) / 2, (nHeight - sourceCropHeight) / 2, sourceCropWidth, sourceCropHeight, InterpolationModeHighQualityBicubic
            
            'For previews only, before rendering the final DIB to the screen, going some helpful
            ' guidelines to help the user confirm the accuracy of their straightening.
            Dim lineOffset As Double, lineStepX As Double, lineStepY As Double
            lineStepX = (nWidth - 1) / 4
            lineStepY = (nHeight - 1) / 4
            
            Dim j As Long
            For j = 0 To 4
                lineOffset = lineStepX * j
                GDIPlusDrawLineToDC finalDIB.getDIBDC, lineOffset, 0, lineOffset, nHeight, RGB(255, 255, 0), 192, 1
                GDIPlusDrawLineToDC finalDIB.getDIBDC, lineOffset + lineStepX / 2, 0, lineOffset + lineStepX / 2, nHeight, RGB(255, 255, 0), 80, 1
                lineOffset = lineStepY * j
                GDIPlusDrawLineToDC finalDIB.getDIBDC, 0, lineOffset, nWidth, lineOffset, RGB(255, 255, 0), 192, 1
                GDIPlusDrawLineToDC finalDIB.getDIBDC, 0, lineOffset + lineStepY / 2, nWidth, lineOffset + lineStepY / 2, RGB(255, 255, 0), 80, 1
            Next j
                        
            'Finally, render the preview and erase the temporary DIB to conserve memory
            finalDIB.renderToPictureBox fxPreview.getPreviewPic
            fxPreview.setFXImage finalDIB
            
            Set tmpDIB = Nothing
            Set finalDIB = Nothing
            
        Else
            
            'FreeImage doesn't raise progress events, but we can use the number of layers as
            ' a stand-in progress parameter.
            If thingToRotate = PD_AT_WHOLEIMAGE Then
                Message "Straightening image..."
                SetProgBarMax pdImages(g_CurrentImage).getNumOfLayers
            Else
                Message "Straightening layer..."
                SetProgBarMax 1
            End If
            
            'Iterate through each layer, rotating as we go
            Dim tmpLayerRef As pdLayer
            
            'If we are rotating the entire image, we must handle all layers in turn.  Otherwise, we can handle just
            ' the active layer.
            Dim lInit As Long, lFinal As Long
            
            Select Case thingToRotate
            
                Case PD_AT_WHOLEIMAGE
                    lInit = 0
                    lFinal = pdImages(g_CurrentImage).getNumOfLayers - 1
                
                Case PD_AT_SINGLELAYER
                    lInit = pdImages(g_CurrentImage).getActiveLayerIndex
                    lFinal = pdImages(g_CurrentImage).getActiveLayerIndex
            
            End Select
            
            Dim i As Long
            For i = lInit To lFinal
            
                If thingToRotate = PD_AT_WHOLEIMAGE Then SetProgBarVal i
            
                'Retrieve a pointer to the layer of interest
                Set tmpLayerRef = pdImages(g_CurrentImage).getLayerByIndex(i)
                
                'Remove premultiplied alpha, if any
                tmpLayerRef.layerDIB.setAlphaPremultiplication False
                
                'Null-pad the layer
                If thingToRotate = PD_AT_WHOLEIMAGE Then tmpLayerRef.convertToNullPaddedLayer pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height
                                
                'Give FreeImage a handle to the layer's pixel data
                fi_DIB = FreeImage_CreateFromDC(tmpLayerRef.layerDIB.getDIBDC)
            
                'Ask FreeImage to rotate the DIB
                returnDIB = FreeImage_RotateEx(fi_DIB, rotationAngle, 0, 0, cx, cy, True)
                
                'As a failsafe, check the returned width/height (they should be unchanged)
                nWidth = FreeImage_GetWidth(returnDIB)
                nHeight = FreeImage_GetHeight(returnDIB)
                
                'Resize the layer's DIB in preparation for the transfer
                tmpLayerRef.layerDIB.createBlank nWidth, nHeight, 32
                
                'Ask FreeImage to premultiply the image's alpha data
                FreeImage_PreMultiplyWithAlpha returnDIB
                
                'Copy the bits from the FreeImage DIB to our DIB
                SetDIBitsToDevice tmpLayerRef.layerDIB.getDIBDC, 0, 0, nWidth, nHeight, 0, 0, 0, nHeight, ByVal FreeImage_GetBits(returnDIB), ByVal FreeImage_GetInfo(returnDIB), 0&
                
                'With the transfer complete, release the FreeImage DIB and unload the library
                If returnDIB <> 0 Then FreeImage_UnloadEx returnDIB
                If fi_DIB <> 0 Then FreeImage_UnloadEx fi_DIB
                
                'Next, we need to calculate a scaling factor for the image.  Straightening applies a sort of auto-crop
                ' to the image to remove empty corners; by solving a triangle equation using the image diagonal, we
                ' can calculate the scaling factor needed.  Thank you to this article for the helpful diagram:
                ' http://stackoverflow.com/questions/18865837/image-straightening-in-android
                ' (Note that the stackoverflow link does not work for the case of width > height, and the instructions
                '  provided for correcting that case are *wrong*!)
                If nWidth < nHeight Then
                    solveAngle = Atn(nHeight / nWidth)
                    len1 = cx / Cos(solveAngle - Abs(rotationAngle * PI_DIV_180))
                Else
                    solveAngle = Atn(nWidth / nHeight)
                    len1 = cy / Cos(solveAngle - Abs(rotationAngle * PI_DIV_180))
                End If
                
                len2 = Sqr(cx * cx + cy * cy)
                scaleFactor = len2 / len1
                
                'Using our new scalefactor, calculate a source image width and height
                sourceCropWidth = nWidth * (1 / scaleFactor)
                sourceCropHeight = nHeight * (1 / scaleFactor)
                
                'Prepare a final DIB to receive the resized image
                finalDIB.createBlank nWidth, nHeight, 32, 0
                
                'Use GDI+ to copy the relevant source rectangle into the final DIB
                GDIPlusResizeDIB finalDIB, 0, 0, nWidth, nHeight, tmpLayerRef.layerDIB, (nWidth - sourceCropWidth) / 2, (nHeight - sourceCropHeight) / 2, sourceCropWidth, sourceCropHeight, InterpolationModeHighQualityBicubic
                
                'Copy the resized DIB into its parent layer
                tmpLayerRef.layerDIB.createFromExistingDIB finalDIB
                
                'If resizing the entire image, remove any null-padding now
                If thingToRotate = PD_AT_WHOLEIMAGE Then tmpLayerRef.cropNullPaddedLayer
                
                'Notify the parent of the change
                pdImages(g_CurrentImage).notifyImageChanged UNDO_LAYER, i
                                
            'Continue with the next layer
            Next i
            
            'All layers have been rotated successfully!
            
            'Update the image's size
            If thingToRotate = PD_AT_WHOLEIMAGE Then
                pdImages(g_CurrentImage).updateSize False, nWidth, nHeight
                DisplaySize pdImages(g_CurrentImage)
            End If
            
            'Fit the new image on-screen and redraw its viewport
            Viewport_Engine.Stage1_InitializeBuffer pdImages(g_CurrentImage), FormMain.mainCanvas(0)
            
            Message "Straighten complete."
            SetProgBarVal 0
            releaseProgressBar
        
        End If
        
    Else
        Message "Arbitrary rotation requires the FreeImage plugin, which could not be located.  Rotation canceled."
        PDMsgBox "The FreeImage plugin is required for image rotation.  Please go to Tools -> Options -> Updates and allow PhotoDemon to download core plugins.  Then restart the program.", vbApplicationModal + vbOKOnly + vbInformation, "FreeImage plugin missing"
    End If
        
End Sub

'OK button
Private Sub cmdBar_OKClick()

    Select Case m_StraightenTarget
    
        Case PD_AT_WHOLEIMAGE
            Process "Straighten image", , buildParams(sltAngle, m_StraightenTarget), UNDO_IMAGE
        
        Case PD_AT_SINGLELAYER
            Process "Straighten layer", , buildParams(sltAngle, m_StraightenTarget), UNDO_LAYER
    
    End Select
    
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub Form_Activate()
            
    'Set the dialog caption to match the current resize operation (resize image or resize single layer)
    Select Case m_StraightenTarget
        
        Case PD_AT_WHOLEIMAGE
            Me.Caption = g_Language.TranslateMessage("Straighten image")
        
        Case PD_AT_SINGLELAYER
            Me.Caption = g_Language.TranslateMessage("Straighten layer")
        
    End Select
    
    'During the preview stage, we want to rotate a smaller version of the image or active layer.  This increases
    ' the speed of previewing immensely (especially for large images, like 10+ megapixel photos)
    Set smallDIB = New pdDIB
    
    'Determine a new image size that preserves the current aspect ratio
    Dim srcWidth As Long, srcHeight As Long
    Dim dWidth As Long, dHeight As Long
    
    Select Case m_StraightenTarget
        
        Case PD_AT_WHOLEIMAGE
            srcWidth = pdImages(g_CurrentImage).Width
            srcHeight = pdImages(g_CurrentImage).Height
        
        Case PD_AT_SINGLELAYER
            srcWidth = pdImages(g_CurrentImage).getActiveLayer.getLayerWidth(False)
            srcHeight = pdImages(g_CurrentImage).getActiveLayer.getLayerHeight(False)
        
    End Select
    
    convertAspectRatio srcWidth, srcHeight, fxPreview.getPreviewWidth, fxPreview.getPreviewHeight, dWidth, dHeight
    
    'Create a new, smaller image at those dimensions
    If (dWidth < srcWidth) Or (dHeight < srcHeight) Then
        
        smallDIB.createBlank dWidth, dHeight, 32, 0
        
        Select Case m_StraightenTarget
        
            Case PD_AT_WHOLEIMAGE
                pdImages(g_CurrentImage).getCompositedRect smallDIB, 0, 0, dWidth, dHeight, 0, 0, pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height, InterpolationModeHighQualityBicubic, , CLC_Generic
            
            Case PD_AT_SINGLELAYER
                GDIPlusResizeDIB smallDIB, 0, 0, dWidth, dHeight, pdImages(g_CurrentImage).getActiveDIB, 0, 0, pdImages(g_CurrentImage).getActiveDIB.getDIBWidth, pdImages(g_CurrentImage).getActiveDIB.getDIBHeight, InterpolationModeHighQualityBicubic
            
        End Select
        
    'The source image or layer is tiny; just use the whole thing!
    Else
    
        Select Case m_StraightenTarget
        
            Case PD_AT_WHOLEIMAGE
                pdImages(g_CurrentImage).getCompositedImage smallDIB
            
            Case PD_AT_SINGLELAYER
                smallDIB.createFromExistingDIB pdImages(g_CurrentImage).getActiveDIB
            
        End Select
        
    End If
        
    'Remove premultiplied alpha from the small DIB copy
    smallDIB.setAlphaPremultiplication False
    
    'Give the preview object a copy of this image data so it can show it to the user if requested
    fxPreview.setOriginalImage smallDIB
    
    'Apply translations and visual themes
    MakeFormPretty Me
        
    'Render a preview
    cmdBar.markPreviewStatus True
    updatePreview
        
End Sub

Private Sub Form_Load()

    'Disable previewing until the dialog is fully initialized
    cmdBar.markPreviewStatus False
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Redraw the on-screen preview of the rotated image
Private Sub updatePreview()
    If cmdBar.previewsAllowed Then StraightenImage sltAngle, m_StraightenTarget, True
End Sub

Private Sub sltAngle_Change()
    updatePreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

