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
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   12105
      _ExtentX        =   21352
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
      DisableZoomPan  =   -1  'True
   End
   Begin PhotoDemon.pdSlider sltAngle 
      Height          =   705
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
'Copyright 2014-2018 by Tanner Helland
'Created: 11/May/14
'Last updated: 11/May/14
'Last update: initial build, based heavily off PD's existing Rotate dialog
'
'This tool allows the user to straighten an image at an arbitrary angle in 1/100 degree increments.
'At present, the tool assumes that you want to straighten the image around its center.  I don't have
' plans to change this behavior.
'
'To straighten a layer instead of the entire image, use the Layer -> Orientation -> Straighten menu.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
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

Public Sub StraightenImage(ByVal processParameters As String, Optional ByVal isPreview As Boolean = False)
    
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.SetParamString processParameters
    
    Dim rotationAngle As Double, thingToRotate As PD_ACTION_TARGET
    
    With cParams
        rotationAngle = .GetDouble("angle", 0#)
        thingToRotate = .GetLong("target", PD_AT_WHOLEIMAGE)
    End With
    
    'If the image contains an active selection, disable it before transforming the canvas
    If (thingToRotate = PD_AT_WHOLEIMAGE) And pdImages(g_CurrentImage).IsSelectionActive And (Not isPreview) Then
        pdImages(g_CurrentImage).SetSelectionActive False
        pdImages(g_CurrentImage).MainSelection.LockRelease
    End If

    'Many 2D libraries use positive values to indicate counter-clockwise rotation.  While mathematically correct, I find this
    ' unintuitive for casual users.  PD reverses the rotationAngle value so that POSITIVE values indicate CLOCKWISE rotation.
    rotationAngle = -rotationAngle

    Dim tmpDIB As pdDIB, finalDIB As pdDIB
    Set tmpDIB = New pdDIB
    Set finalDIB = New pdDIB
    
    Dim srcWidth As Double, srcHeight As Double
    
    'To solve the problem of auto-cropping the straightened image, additional variables are required.
    Dim solveAngle As Double, len1 As Double, len2 As Double, scaleFactor As Double
    
    'This function handles three different cases: previews (which use a pre-composited image, for performance),
    ' single layers (where the layer is rotated around its own center point), and the full image (where each layer
    ' is null-padded in turn, straightened, then un-null-padded).
    If isPreview Then
        srcWidth = smallDIB.GetDIBWidth
        srcHeight = smallDIB.GetDIBHeight
    Else
        Select Case thingToRotate
            Case PD_AT_WHOLEIMAGE
                srcWidth = pdImages(g_CurrentImage).Width
                srcHeight = pdImages(g_CurrentImage).Height
            Case PD_AT_SINGLELAYER
                srcWidth = pdImages(g_CurrentImage).GetActiveDIB.GetDIBWidth
                srcHeight = pdImages(g_CurrentImage).GetActiveDIB.GetDIBHeight
        End Select
    End If
    
    'We want to rotate and scale the image around its center point (instead of [0, 0])
    Dim cx As Double, cy As Double
    cx = srcWidth / 2
    cy = srcHeight / 2
    
    Dim cTransform As pd2DTransform
    Set cTransform = New pd2DTransform
    
    Dim rotatePoints() As PointFloat
    
    'Normally, I like to use identical code for previews and actual effects.  However, rotating is completely different
    ' for previews (where we do a single rotation of the composited image) vs the full images (independently rotating
    ' each layer, with support functions to null-pad and crop layers as necessary).  As such, there is some code
    ' duplication here, but I believe it makes the code much more readable.
    If isPreview Then
        
        'Start by calculating the corner points of the image, when rotated at the specified angle
        PDMath.FindCornersOfRotatedRect srcWidth, srcHeight, rotationAngle, rotatePoints
        
        'Next, we need to calculate a scaling factor for the image.  Straightening applies a sort of auto-crop
        ' to the image to remove empty corners; by solving a triangle equation using the image diagonal, we
        ' can calculate the scaling factor needed.  Thank you to this article for the helpful diagram:
        ' http://stackoverflow.com/questions/18865837/image-straightening-in-android
        ' (Note that the stackoverflow link does not work for the case of width > height, and the instructions
        '  provided for correcting that case are *wrong*!)
        If srcWidth < srcHeight Then
            solveAngle = Atn(srcHeight / srcWidth)
            len1 = cx / Cos(solveAngle - Abs(rotationAngle * PI_DIV_180))
        Else
            solveAngle = Atn(srcWidth / srcHeight)
            len1 = cy / Cos(solveAngle - Abs(rotationAngle * PI_DIV_180))
        End If
        
        len2 = Sqr(cx * cx + cy * cy)
        scaleFactor = len2 / len1
        
        'Apply that scalefactor to our calculated rotation points
        cTransform.ApplyScaling scaleFactor, scaleFactor, cx, cy
        cTransform.ApplyTransformToPointFs VarPtr(rotatePoints(0)), 4
        
        'Prepare a final DIB to receive the resized image
        finalDIB.CreateBlank srcWidth, srcHeight, 32, 0
        
        'Rotate the new image into place
        GDI_Plus.GDIPlus_PlgBlt finalDIB, rotatePoints, smallDIB, 0, 0, smallDIB.GetDIBWidth, smallDIB.GetDIBHeight, , , False
        
        'For previews only, before rendering the final DIB to the screen, going some helpful
        ' guidelines to help the user confirm the accuracy of their straightening.
        Dim lineOffset As Double, lineStepX As Double, lineStepY As Double
        lineStepX = (srcWidth - 1) / 4
        lineStepY = (srcHeight - 1) / 4
        
        Dim j As Long
        For j = 0 To 4
            lineOffset = lineStepX * j
            GDIPlusDrawLineToDC finalDIB.GetDIBDC, lineOffset, 0, lineOffset, srcHeight, RGB(255, 255, 0), 192, 1
            GDIPlusDrawLineToDC finalDIB.GetDIBDC, lineOffset + lineStepX / 2, 0, lineOffset + lineStepX / 2, srcHeight, RGB(255, 255, 0), 80, 1
            lineOffset = lineStepY * j
            GDIPlusDrawLineToDC finalDIB.GetDIBDC, 0, lineOffset, srcWidth, lineOffset, RGB(255, 255, 0), 192, 1
            GDIPlusDrawLineToDC finalDIB.GetDIBDC, 0, lineOffset + lineStepY / 2, srcWidth, lineOffset + lineStepY / 2, RGB(255, 255, 0), 80, 1
        Next j
                    
        'Finally, render the preview and erase the temporary DIB to conserve memory
        pdFxPreview.SetFXImage finalDIB
        
    'This is *not* a preview
    Else
            
        'When rotating the entire image, we can use the number of layers as a stand-in progress parameter.
        If (thingToRotate = PD_AT_WHOLEIMAGE) Then
            Message "Straightening image..."
            SetProgBarMax pdImages(g_CurrentImage).GetNumOfLayers
        Else
            Message "Straightening layer..."
            SetProgBarMax 1
        End If
        
        Dim tmpLayerRef As pdLayer
            
        'When rotating the entire image, we must handle all layers in turn.  Otherwise, we can handle just the active layer.
        Dim lInit As Long, lFinal As Long
        
        Select Case thingToRotate
            Case PD_AT_WHOLEIMAGE
                lInit = 0
                lFinal = pdImages(g_CurrentImage).GetNumOfLayers - 1
            Case PD_AT_SINGLELAYER
                lInit = pdImages(g_CurrentImage).GetActiveLayerIndex
                lFinal = pdImages(g_CurrentImage).GetActiveLayerIndex
        End Select
        
        Dim i As Long
        For i = lInit To lFinal
        
            If (thingToRotate = PD_AT_WHOLEIMAGE) Then SetProgBarVal i
        
            'Retrieve a pointer to the layer of interest
            Set tmpLayerRef = pdImages(g_CurrentImage).GetLayerByIndex(i)
            
            'Null-pad the layer
            If (thingToRotate = PD_AT_WHOLEIMAGE) Then tmpLayerRef.ConvertToNullPaddedLayer pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height
            
            'Calculating the corner points of the layer, when rotated at the specified angle.
            PDMath.FindCornersOfRotatedRect srcWidth, srcHeight, rotationAngle, rotatePoints
            
            'Next, we need to calculate a scaling factor for the image.  Straightening applies a sort of auto-crop
            ' to the image to remove empty corners; by solving a triangle equation using the image diagonal, we
            ' can calculate the scaling factor needed.  Thank you to this article for the helpful diagram:
            ' http://stackoverflow.com/questions/18865837/image-straightening-in-android
            ' (Note that the stackoverflow link does not work for the case of width > height, and the instructions
            '  provided for correcting that case are *wrong*!)
            If (srcWidth < srcHeight) Then
                solveAngle = Atn(srcHeight / srcWidth)
                len1 = cx / Cos(solveAngle - Abs(rotationAngle * PI_DIV_180))
            Else
                solveAngle = Atn(srcWidth / srcHeight)
                len1 = cy / Cos(solveAngle - Abs(rotationAngle * PI_DIV_180))
            End If
            
            len2 = Sqr(cx * cx + cy * cy)
            scaleFactor = len2 / len1
            
            'Apply that scalefactor to our calculated rotation points
            cTransform.Reset
            cTransform.ApplyScaling scaleFactor, scaleFactor, cx, cy
            cTransform.ApplyTransformToPointFs VarPtr(rotatePoints(0)), 4
            
            'Prepare a final DIB to receive the resized image
            If (finalDIB.GetDIBWidth <> CLng(srcWidth)) Or (finalDIB.GetDIBHeight <> CLng(srcHeight)) Then
                finalDIB.CreateBlank CLng(srcWidth), CLng(srcHeight), 32, 0
                finalDIB.SetInitialAlphaPremultiplicationState True
            Else
                finalDIB.ResetDIB 0
            End If
            
            'Rotate the new image into place
            GDI_Plus.GDIPlus_PlgBlt finalDIB, rotatePoints, tmpLayerRef.layerDIB, 0, 0, tmpLayerRef.layerDIB.GetDIBWidth, tmpLayerRef.layerDIB.GetDIBHeight, , , False
            
            'Copy the resized DIB into its parent layer
            tmpLayerRef.layerDIB.CreateFromExistingDIB finalDIB
            
            'If resizing the entire image, remove any null-padding now
            If (thingToRotate = PD_AT_WHOLEIMAGE) Then tmpLayerRef.CropNullPaddedLayer
            
            'Notify the parent of the change
            pdImages(g_CurrentImage).NotifyImageChanged UNDO_Layer, i
                            
        'Continue with the next layer
        Next i
        
        'All layers have been rotated successfully!
        
        'Update the image's size (not technically necessary, but this triggers some other backend notifications that are relevant)
        If thingToRotate = PD_AT_WHOLEIMAGE Then
            pdImages(g_CurrentImage).UpdateSize False, srcWidth, srcHeight
            DisplaySize pdImages(g_CurrentImage)
        End If
        
        'Fit the new image on-screen and redraw its viewport
        ViewportEngine.Stage1_InitializeBuffer pdImages(g_CurrentImage), FormMain.MainCanvas(0)
        
        Message "Straighten complete."
        SetProgBarVal 0
        ReleaseProgressBar
    
    End If
        
End Sub

'OK button
Private Sub cmdBar_OKClick()

    Select Case m_StraightenTarget
        Case PD_AT_WHOLEIMAGE
            Process "Straighten image", , GetLocalParamString(), UNDO_Image
        Case PD_AT_SINGLELAYER
            Process "Straighten layer", , GetLocalParamString(), UNDO_Layer
    End Select
    
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub Form_Load()

    'Disable previewing until the dialog is fully initialized
    cmdBar.MarkPreviewStatus False
    
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
            srcWidth = pdImages(g_CurrentImage).GetActiveLayer.GetLayerWidth(False)
            srcHeight = pdImages(g_CurrentImage).GetActiveLayer.GetLayerHeight(False)
        
    End Select
    
    ConvertAspectRatio srcWidth, srcHeight, pdFxPreview.GetPreviewWidth, pdFxPreview.GetPreviewHeight, dWidth, dHeight
    
    'Create a new, smaller image at those dimensions
    If (dWidth <= srcWidth) Or (dHeight <= srcHeight) Then
        
        smallDIB.CreateBlank dWidth, dHeight, 32, 0
        
        Select Case m_StraightenTarget
        
            Case PD_AT_WHOLEIMAGE
            
                Dim dstRectF As RectF, srcRectF As RectF
                With dstRectF
                    .Left = 0#
                    .Top = 0#
                    .Width = dWidth
                    .Height = dHeight
                End With
                
                With srcRectF
                    .Left = 0#
                    .Top = 0#
                    .Width = pdImages(g_CurrentImage).Width
                    .Height = pdImages(g_CurrentImage).Height
                End With
                
                pdImages(g_CurrentImage).GetCompositedRect smallDIB, dstRectF, srcRectF, GP_IM_HighQualityBicubic, , CLC_Generic
            
            Case PD_AT_SINGLELAYER
                GDIPlusResizeDIB smallDIB, 0, 0, dWidth, dHeight, pdImages(g_CurrentImage).GetActiveDIB, 0, 0, pdImages(g_CurrentImage).GetActiveDIB.GetDIBWidth, pdImages(g_CurrentImage).GetActiveDIB.GetDIBHeight, GP_IM_HighQualityBicubic
            
        End Select
        
    'The source image or layer is tiny; just use the whole thing!
    Else
    
        Select Case m_StraightenTarget
        
            Case PD_AT_WHOLEIMAGE
                pdImages(g_CurrentImage).GetCompositedImage smallDIB
            
            Case PD_AT_SINGLELAYER
                smallDIB.CreateFromExistingDIB pdImages(g_CurrentImage).GetActiveDIB
            
        End Select
        
    End If
    
    'Give the preview object a copy of this image data so it can show it to the user if requested
    pdFxPreview.SetOriginalImage smallDIB
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me
    cmdBar.MarkPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Redraw the on-screen preview of the rotated image
Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then Me.StraightenImage GetLocalParamString(), True
End Sub

Private Sub sltAngle_Change()
    UpdatePreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    
    With cParams
        .AddParam "angle", sltAngle.Value
        .AddParam "target", m_StraightenTarget
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
