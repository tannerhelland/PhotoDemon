VERSION 5.00
Begin VB.Form FormStraighten 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Straighten"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11655
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
   ScaleWidth      =   777
   Begin PhotoDemon.pdButtonStrip btsGrid 
      Height          =   975
      Left            =   6000
      TabIndex        =   3
      Top             =   3120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1720
      Caption         =   "grid"
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   11655
      _ExtentX        =   20558
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
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1244
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
'Copyright 2014-2026 by Tanner Helland
'Created: 11/May/14
'Last updated: 29/April/20
'Last update: allow user to toggle preview grid; use pd2D for grid rendering
'
'This tool allows the user to straighten an image at an arbitrary angle in 1/100 degree increments.
'At present, the tool assumes that you want to straighten the image around its center.  I don't have
' plans to change this behavior.
'
'To straighten a layer instead of the entire image, use the Layer -> Orientation -> Straighten menu.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'This temporary DIB will be used for rendering the preview
Private m_smallDIB As pdDIB

'This dialog can be used to resize the full image, or a single layer.  The requested target will be stored here,
' and can be externally accessed by the ResizeTarget property.
Private m_StraightenTarget As PD_ActionTarget

Public Property Let StraightenTarget(newTarget As PD_ActionTarget)
    m_StraightenTarget = newTarget
End Property

Public Sub StraightenImage(ByVal processParameters As String, Optional ByVal isPreview As Boolean = False)
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString processParameters
    
    Dim rotationAngle As Double, thingToRotate As PD_ActionTarget, showPreviewGrid As Boolean
    
    With cParams
        rotationAngle = .GetDouble("angle", 0#)
        thingToRotate = .GetLong("target", pdat_Image)
        showPreviewGrid = .GetBool("preview-grid", True, True)
    End With
    
    'If the image contains an active selection, disable it before transforming the canvas
    If (thingToRotate = pdat_Image) And PDImages.GetActiveImage.IsSelectionActive And (Not isPreview) Then
        PDImages.GetActiveImage.SetSelectionActive False
        PDImages.GetActiveImage.MainSelection.LockRelease
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
        srcWidth = m_smallDIB.GetDIBWidth
        srcHeight = m_smallDIB.GetDIBHeight
    Else
        Select Case thingToRotate
            Case pdat_Image
                srcWidth = PDImages.GetActiveImage.Width
                srcHeight = PDImages.GetActiveImage.Height
            Case pdat_SingleLayer
                srcWidth = PDImages.GetActiveImage.GetActiveDIB.GetDIBWidth
                srcHeight = PDImages.GetActiveImage.GetActiveDIB.GetDIBHeight
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
        GDI_Plus.GDIPlus_PlgBlt finalDIB, rotatePoints, m_smallDIB, 0, 0, m_smallDIB.GetDIBWidth, m_smallDIB.GetDIBHeight, , GP_IM_HighQualityBicubic, False, True
        
        If showPreviewGrid Then
            
            'For previews only, before rendering the final DIB to the screen, going some helpful
            ' guidelines to help the user confirm the accuracy of their straightening.
            Dim lineOffset As Double, lineStepX As Double, lineStepY As Double
            lineStepX = (srcWidth - 1) / 4
            lineStepY = (srcHeight - 1) / 4
            
            Dim cSurface As pd2DSurface
            Set cSurface = New pd2DSurface
            cSurface.WrapSurfaceAroundPDDIB finalDIB
            cSurface.SetSurfaceAntialiasing P2_AA_None
            cSurface.SetSurfacePixelOffset P2_PO_Normal
            
            Dim cPen As pd2DPen
            Set cPen = New pd2DPen
            cPen.SetPenWidth 1!
            cPen.SetPenColor RGB(255, 255, 0)
            
            Dim j As Long
            For j = 0 To 4
                lineOffset = lineStepX * j
                cPen.SetPenOpacity 75
                PD2D.DrawLineI cSurface, cPen, lineOffset, 0, lineOffset, srcHeight
                cPen.SetPenOpacity 33
                PD2D.DrawLineI cSurface, cPen, lineOffset + lineStepX / 2, 0, lineOffset + lineStepX / 2, srcHeight
                lineOffset = lineStepY * j
                cPen.SetPenOpacity 75
                PD2D.DrawLineI cSurface, cPen, 0, lineOffset, srcWidth, lineOffset
                cPen.SetPenOpacity 33
                PD2D.DrawLineI cSurface, cPen, 0, lineOffset + lineStepY / 2, srcWidth, lineOffset + lineStepY / 2
            Next j
            
            Set cSurface = Nothing
            
        End If
        
        'Finally, render the preview and erase the temporary DIB to conserve memory
        pdFxPreview.SetFXImage finalDIB
        
    'This is *not* a preview
    Else
            
        'When rotating the entire image, we can use the number of layers as a stand-in progress parameter.
        If (thingToRotate = pdat_Image) Then
            Message "Straightening image..."
            SetProgBarMax PDImages.GetActiveImage.GetNumOfLayers
        Else
            Message "Straightening layer..."
            SetProgBarMax 1
        End If
        
        Dim tmpLayerRef As pdLayer
            
        'When rotating the entire image, we must handle all layers in turn.  Otherwise, we can handle just the active layer.
        Dim lInit As Long, lFinal As Long
        
        Select Case thingToRotate
            Case pdat_Image
                lInit = 0
                lFinal = PDImages.GetActiveImage.GetNumOfLayers - 1
            Case pdat_SingleLayer
                lInit = PDImages.GetActiveImage.GetActiveLayerIndex
                lFinal = PDImages.GetActiveImage.GetActiveLayerIndex
        End Select
        
        Dim i As Long
        For i = lInit To lFinal
        
            If (thingToRotate = pdat_Image) Then SetProgBarVal i
        
            'Retrieve a pointer to the layer of interest
            Set tmpLayerRef = PDImages.GetActiveImage.GetLayerByIndex(i)
            
            'Null-pad the layer
            If (thingToRotate = pdat_Image) Then tmpLayerRef.ConvertToNullPaddedLayer PDImages.GetActiveImage.Width, PDImages.GetActiveImage.Height
            
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
            GDI_Plus.GDIPlus_PlgBlt finalDIB, rotatePoints, tmpLayerRef.GetLayerDIB, 0, 0, tmpLayerRef.GetLayerDIB.GetDIBWidth, tmpLayerRef.GetLayerDIB.GetDIBHeight, , , False, True
            
            'Copy the resized DIB into its parent layer
            tmpLayerRef.GetLayerDIB.CreateFromExistingDIB finalDIB
            
            'If resizing the entire image, remove any null-padding now
            If (thingToRotate = pdat_Image) Then tmpLayerRef.CropNullPaddedLayer
            
            'Notify the parent of the change
            PDImages.GetActiveImage.NotifyImageChanged UNDO_Layer, i
                            
        'Continue with the next layer
        Next i
        
        'All layers have been rotated successfully!
        
        'Update the image's size (not technically necessary, but this triggers some other backend notifications that are relevant)
        If (thingToRotate = pdat_Image) Then
            PDImages.GetActiveImage.UpdateSize False, srcWidth, srcHeight
            Interface.DisplaySize PDImages.GetActiveImage()
            Tools.NotifyImageSizeChanged
        End If
        
        'Fit the new image on-screen and redraw its viewport
        Viewport.Stage1_InitializeBuffer PDImages.GetActiveImage(), FormMain.MainCanvas(0)
        
        Message "Straighten complete."
        SetProgBarVal 0
        ReleaseProgressBar
    
    End If
        
End Sub

Private Sub btsGrid_Click(ByVal buttonIndex As Long)
    UpdatePreview
End Sub

'OK button
Private Sub cmdBar_OKClick()

    Select Case m_StraightenTarget
        Case pdat_Image
            Process "Straighten image", , GetLocalParamString(), UNDO_Image
        Case pdat_SingleLayer
            Process "Straighten layer", , GetLocalParamString(), UNDO_Layer
    End Select
    
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub Form_Load()

    'Disable previewing until the dialog is fully initialized
    cmdBar.SetPreviewStatus False
    
    'Set the dialog caption to match the current resize operation (resize image or resize single layer)
    Select Case m_StraightenTarget
        
        Case pdat_Image
            If (Not g_WindowManager Is Nothing) Then g_WindowManager.SetWindowCaptionW Me.hWnd, g_Language.TranslateMessage("Straighten image")
        
        Case pdat_SingleLayer
            If (Not g_WindowManager Is Nothing) Then g_WindowManager.SetWindowCaptionW Me.hWnd, g_Language.TranslateMessage("Straighten layer")
        
    End Select
    
    btsGrid.AddItem "on", 0
    btsGrid.AddItem "off", 1
    btsGrid.ListIndex = 0
    
    UpdatePreviewSource
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Generate a new source for the preview
Private Sub UpdatePreviewSource()

    'During the preview stage, we want to rotate a smaller version of the image or active layer.  This increases
    ' the speed of previewing immensely (especially for large images, like 10+ megapixel photos)
    Set m_smallDIB = New pdDIB
    
    'Determine a new image size that preserves the current aspect ratio
    Dim srcWidth As Long, srcHeight As Long
    Dim dWidth As Long, dHeight As Long
    
    Select Case m_StraightenTarget
        
        Case pdat_Image
            srcWidth = PDImages.GetActiveImage.Width
            srcHeight = PDImages.GetActiveImage.Height
        
        Case pdat_SingleLayer
            srcWidth = PDImages.GetActiveImage.GetActiveLayer.GetLayerWidth(False)
            srcHeight = PDImages.GetActiveImage.GetActiveLayer.GetLayerHeight(False)
        
    End Select
    
    ConvertAspectRatio srcWidth, srcHeight, pdFxPreview.GetPreviewWidth, pdFxPreview.GetPreviewHeight, dWidth, dHeight
    
    'Create a new, smaller image at those dimensions
    If (dWidth <= srcWidth) Or (dHeight <= srcHeight) Then
        
        m_smallDIB.CreateBlank dWidth, dHeight, 32, 0
        
        Select Case m_StraightenTarget
        
            Case pdat_Image
            
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
                    .Width = PDImages.GetActiveImage.Width
                    .Height = PDImages.GetActiveImage.Height
                End With
                
                PDImages.GetActiveImage.GetCompositedRect m_smallDIB, dstRectF, srcRectF, GP_IM_HighQualityBicubic, , CLC_Generic
            
            Case pdat_SingleLayer
                GDIPlusResizeDIB m_smallDIB, 0, 0, dWidth, dHeight, PDImages.GetActiveImage.GetActiveDIB, 0, 0, PDImages.GetActiveImage.GetActiveDIB.GetDIBWidth, PDImages.GetActiveImage.GetActiveDIB.GetDIBHeight, GP_IM_HighQualityBicubic
            
        End Select
        
    'The source image or layer is tiny; just use the whole thing!
    Else
    
        Select Case m_StraightenTarget
        
            Case pdat_Image
                PDImages.GetActiveImage.GetCompositedImage m_smallDIB
            
            Case pdat_SingleLayer
                m_smallDIB.CreateFromExistingDIB PDImages.GetActiveImage.GetActiveDIB
            
        End Select
        
    End If
    
    'Give the preview object a copy of this image data so it can show it to the user if requested
    pdFxPreview.SetOriginalImage m_smallDIB
    
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
    UpdatePreviewSource
    UpdatePreview
End Sub

'IMPORTANT NOTE: any changes made here need to be mirrored to the Tools_Measure module, specifically the
' "straighten image/layer using this angle" feature.
Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "angle", sltAngle.Value
        .AddParam "target", m_StraightenTarget
        .AddParam "preview-grid", (btsGrid.ListIndex = 0), True
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
