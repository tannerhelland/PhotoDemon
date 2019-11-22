VERSION 5.00
Begin VB.Form FormRotate 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Rotate Image"
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
   Begin PhotoDemon.pdColorSelector csBackground 
      Height          =   615
      Left            =   6120
      TabIndex        =   6
      Top             =   4920
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1085
   End
   Begin PhotoDemon.pdButtonStrip btsResize 
      Height          =   1095
      Left            =   6000
      TabIndex        =   3
      Top             =   1080
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1931
      Caption         =   "size"
   End
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
      TabIndex        =   2
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
      DisableZoomPan  =   -1  'True
   End
   Begin PhotoDemon.pdSlider sltAngle 
      Height          =   705
      Left            =   6000
      TabIndex        =   1
      Top             =   180
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "angle"
      Min             =   -360
      Max             =   360
      SigDigits       =   2
   End
   Begin PhotoDemon.pdButtonStrip btsResample 
      Height          =   1095
      Left            =   6000
      TabIndex        =   4
      Top             =   2400
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1931
      Caption         =   "quality"
   End
   Begin PhotoDemon.pdButtonStrip btsBackground 
      Height          =   1095
      Left            =   6000
      TabIndex        =   5
      Top             =   3720
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1931
      Caption         =   "border regions"
   End
End
Attribute VB_Name = "FormRotate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Image Rotation Interface
'Copyright 2012-2019 by Tanner Helland
'Created: 12/November/12
'Last updated: 06/June/16
'Last update: total overhaul to improve performance, quality, and feature set.  FreeImage is no longer involved.
'
'This tool allows the user to rotate an image at an arbitrary angle in 1/100 degree increments.  At present, it's assumed
' you want to rotate the image around its center.
'
'To rotate a layer instead of the entire image, use the Layer menu.  Rotation is also available in the
' Effect -> Distort menu, which behaves like a standard distort tool (with extra options related to distorting).
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
Private m_RotateTarget As PD_ACTION_TARGET

Public Property Let RotateTarget(newTarget As PD_ACTION_TARGET)
    m_RotateTarget = newTarget
End Property

Public Sub RotateArbitrary(ByVal rotationParameters As String, Optional ByVal isPreview As Boolean = False)
    
    'First, parse out individual XML parameters
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.SetParamString rotationParameters
    
    Dim thingToRotate As PD_ACTION_TARGET
    thingToRotate = cParams.GetLong("target", PD_AT_WHOLEIMAGE)
    
    Dim rotationAngle As Double
    rotationAngle = -1# * cParams.GetDouble("angle", 0#)
    
    Dim resizeToFit As Boolean
    resizeToFit = Strings.StringsEqual(cParams.GetString("style", "enlarge"), "enlarge", True)
    
    Dim rotationQuality As Long
    rotationQuality = cParams.GetLong("quality", 2)
    
    Dim rotationTransparent As Boolean, rotationBackColor As Long
    rotationTransparent = cParams.GetBool("transparentbackground", True)
    rotationBackColor = cParams.GetLong("backgroundcolor", vbWhite)
    
    Dim gdipRotationQuality As GP_InterpolationMode
    If (rotationQuality = 0) Then
        gdipRotationQuality = GP_IM_NearestNeighbor
    ElseIf (rotationQuality = 1) Then
        gdipRotationQuality = GP_IM_Bilinear
    Else
        gdipRotationQuality = GP_IM_Bicubic
    End If
    
    'If we're rotating an entire image, and a selection tool is active, disable the selection before rotating
    If (thingToRotate = PD_AT_WHOLEIMAGE) And PDImages.GetActiveImage.IsSelectionActive And (Not isPreview) Then
        PDImages.GetActiveImage.SetSelectionActive False
        PDImages.GetActiveImage.MainSelection.LockRelease
    End If
    
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
        
    'Normally, I like to use identical code for previews and actual effects.  However, rotating is completely different
    ' for previews (where we do a single rotation of the composited image) vs the full images (independently rotating
    ' each layer, with support functions to null-pad and crop layers as necessary).  As such, there is some code
    ' duplication here, but I believe it makes the code much more readable.
    If isPreview Then
        
        If resizeToFit Then
            GDI_Plus.GDIPlus_RotateDIBPlgStyle smallDIB, tmpDIB, rotationAngle, False, gdipRotationQuality, rotationTransparent, rotationBackColor
        Else
            tmpDIB.CreateBlank smallDIB.GetDIBWidth, smallDIB.GetDIBHeight, smallDIB.GetDIBColorDepth, 0, 0
            GDI_Plus.GDIPlus_RotateDIBPlgStyle smallDIB, tmpDIB, rotationAngle, True, gdipRotationQuality, rotationTransparent, rotationBackColor
        End If
        
        tmpDIB.SetInitialAlphaPremultiplicationState True
        pdFxPreview.SetFXImage tmpDIB
        
    Else
            
        'We don't currently use a progress callback for GDI+ events, but in this case, we can use the number of layers as
        ' a stand-in progress parameter.
        If (thingToRotate = PD_AT_WHOLEIMAGE) Then
            Message "Rotating image..."
            SetProgBarMax PDImages.GetActiveImage.GetNumOfLayers
        Else
            Message "Rotating layer..."
            SetProgBarMax 1
        End If
        
        'Iterate through each layer, rotating as we go
        Dim tmpLayerRef As pdLayer
        Dim origOffsetX As Long, origOffsetY As Long
        
        'If we are rotating the entire image, we must handle all layers in turn.  Otherwise, we can handle just
        ' the active layer.
        Dim lInit As Long, lFinal As Long
        
        Select Case thingToRotate
        
            Case PD_AT_WHOLEIMAGE
                lInit = 0
                lFinal = PDImages.GetActiveImage.GetNumOfLayers - 1
            
            Case PD_AT_SINGLELAYER
                lInit = PDImages.GetActiveImage.GetActiveLayerIndex
                lFinal = PDImages.GetActiveImage.GetActiveLayerIndex
        
        End Select
        
        Dim i As Long
        For i = lInit To lFinal
        
            If (thingToRotate = PD_AT_WHOLEIMAGE) Then SetProgBarVal i
        
            'Retrieve a pointer to the layer of interest
            Set tmpLayerRef = PDImages.GetActiveImage.GetLayerByIndex(i)
            
            'If we are only resizing a single layer, make a copy of the layer's current x/y offsets.  We will use these
            ' to re-center the layer after it has been resized.
            origOffsetX = tmpLayerRef.GetLayerOffsetX + (tmpLayerRef.GetLayerWidth(False) \ 2)
            origOffsetY = tmpLayerRef.GetLayerOffsetY + (tmpLayerRef.GetLayerHeight(False) \ 2)
            
            'Null-pad the layer
            If (thingToRotate = PD_AT_WHOLEIMAGE) Then tmpLayerRef.ConvertToNullPaddedLayer PDImages.GetActiveImage.Width, PDImages.GetActiveImage.Height
            
            'There are two ways to rotate an image - enlarging the canvas to receive the fully rotated copy, or
            ' leaving the image the same size and truncating corners.  These require two different FreeImage functions.
            If resizeToFit Then
                
                'If the user wants us to fill the border regions of the rotated image with color, we only obey their command
                ' for the base layer.  Layers atop the base layer can receive transparency in their border regions without trouble.
                If (thingToRotate = PD_AT_WHOLEIMAGE) And (Not rotationTransparent) And (i > lInit) Then
                    GDI_Plus.GDIPlus_RotateDIBPlgStyle tmpLayerRef.layerDIB, tmpDIB, rotationAngle, False, gdipRotationQuality, True
                Else
                    GDI_Plus.GDIPlus_RotateDIBPlgStyle tmpLayerRef.layerDIB, tmpDIB, rotationAngle, False, gdipRotationQuality, rotationTransparent, rotationBackColor
                End If
                
            Else
                If (tmpDIB.GetDIBWidth <> PDImages.GetActiveImage.Width) Or (tmpDIB.GetDIBHeight <> PDImages.GetActiveImage.Height) Then
                    tmpDIB.CreateBlank PDImages.GetActiveImage.Width, PDImages.GetActiveImage.Height, tmpLayerRef.layerDIB.GetDIBColorDepth, 0, 0
                Else
                    tmpDIB.ResetDIB 0
                End If
                
                If (thingToRotate = PD_AT_WHOLEIMAGE) And (Not rotationTransparent) And (i > lInit) Then
                    GDI_Plus.GDIPlus_RotateDIBPlgStyle tmpLayerRef.layerDIB, tmpDIB, rotationAngle, True, gdipRotationQuality, True
                Else
                    GDI_Plus.GDIPlus_RotateDIBPlgStyle tmpLayerRef.layerDIB, tmpDIB, rotationAngle, True, gdipRotationQuality, rotationTransparent, rotationBackColor
                End If
                
            End If
            
            'Copy the end result into the source layer
            tmpLayerRef.layerDIB.CreateFromExistingDIB tmpDIB
            
            'If resizing the entire image, remove any null-padding now
            If thingToRotate = PD_AT_WHOLEIMAGE Then
                tmpLayerRef.CropNullPaddedLayer
            
            'If resizing only a single layer, re-center it according to its old offset
            Else
                tmpLayerRef.SetLayerOffsetX origOffsetX - (tmpLayerRef.GetLayerWidth(False) \ 2)
                tmpLayerRef.SetLayerOffsetY origOffsetY - (tmpLayerRef.GetLayerHeight(False) \ 2)
            End If
            
            'Notify the parent of the change
            PDImages.GetActiveImage.NotifyImageChanged UNDO_Layer, i
            
        'Continue with the next layer
        Next i
        
        'All layers have been rotated successfully!
        
        'Update the image's size
        If (thingToRotate = PD_AT_WHOLEIMAGE) And resizeToFit Then
            Dim newWidth As Double, newHeight As Double
            PDMath.FindBoundarySizeOfRotatedRect PDImages.GetActiveImage.Width, PDImages.GetActiveImage.Height, rotationAngle, newWidth, newHeight, False
            PDImages.GetActiveImage.UpdateSize False, newWidth, newHeight
            DisplaySize PDImages.GetActiveImage()
        End If
        
        'Fit the new image on-screen and redraw its viewport
        ViewportEngine.Stage1_InitializeBuffer PDImages.GetActiveImage(), FormMain.MainCanvas(0)
        
        Message "Rotation complete."
        SetProgBarVal 0
        ReleaseProgressBar
        
    End If
    
End Sub

Private Sub btsBackground_Click(ByVal buttonIndex As Long)
    UpdatePreview
    UpdateBackgroundColorVisibility
End Sub

Private Sub UpdateBackgroundColorVisibility()
    csBackground.Visible = (btsBackground.ListIndex <> 0)
End Sub

Private Sub btsResample_Click(ByVal buttonIndex As Long)
    UpdatePreview
End Sub

Private Sub btsResize_Click(ByVal buttonIndex As Long)
    UpdatePreview
End Sub

'OK button
Private Sub cmdBar_OKClick()
    
    Select Case m_RotateTarget
    
        Case PD_AT_WHOLEIMAGE
            Process "Arbitrary image rotation", , GetFunctionParamString(), UNDO_Image
            
        Case PD_AT_SINGLELAYER
            Process "Arbitrary layer rotation", , GetFunctionParamString(), UNDO_Layer
            
    End Select
    
End Sub

Private Function GetFunctionParamString() As String
    
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    
    With cParams
        .AddParam "target", m_RotateTarget
        If (btsResize.ListIndex = 1) Then .AddParam "style", "enlarge" Else .AddParam "style", "fit"
        .AddParam "angle", sltAngle.Value
        .AddParam "quality", btsResample.ListIndex
        .AddParam "transparentbackground", (btsBackground.ListIndex = 0)
        .AddParam "backgroundcolor", csBackground.Color
    End With
    
    GetFunctionParamString = cParams.GetParamString
    
End Function

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    btsResize.ListIndex = 1
End Sub

Private Sub csBackground_ColorChanged()
    UpdatePreview
End Sub

Private Sub Form_Load()

    'Disable previewing until the dialog is fully initialized
    cmdBar.SetPreviewStatus False
    
    btsResize.AddItem "keep current", 0
    btsResize.AddItem "enlarge to fit", 1
    btsResize.ListIndex = 1
    
    btsResample.AddItem "nearest-neighbor", 0
    btsResample.AddItem "bilinear", 1
    btsResample.AddItem "bicubic", 2
    btsResample.ListIndex = 2
    
    btsBackground.AddItem "transparent", 0
    btsBackground.AddItem "fill with color", 1
    btsBackground.ListIndex = 0
    UpdateBackgroundColorVisibility
    
    'Set the dialog caption to match the current resize operation (resize image or resize single layer)
    Select Case m_RotateTarget
        
        Case PD_AT_WHOLEIMAGE
            Me.Caption = g_Language.TranslateMessage("Rotate image")
        
        Case PD_AT_SINGLELAYER
            Me.Caption = g_Language.TranslateMessage("Rotate layer")
        
    End Select
    
    'During the preview stage, we want to rotate a smaller version of the image or active layer.  This increases
    ' the speed of previewing immensely (especially for large images, like 10+ megapixel photos)
    Set smallDIB = New pdDIB
    
    'Determine a new image size that preserves the current aspect ratio
    Dim srcWidth As Long, srcHeight As Long
    Dim dWidth As Long, dHeight As Long
    
    Select Case m_RotateTarget
        
        Case PD_AT_WHOLEIMAGE
            srcWidth = PDImages.GetActiveImage.Width
            srcHeight = PDImages.GetActiveImage.Height
        
        Case PD_AT_SINGLELAYER
            srcWidth = PDImages.GetActiveImage.GetActiveLayer.GetLayerWidth(False)
            srcHeight = PDImages.GetActiveImage.GetActiveLayer.GetLayerHeight(False)
        
    End Select
    
    ConvertAspectRatio srcWidth, srcHeight, pdFxPreview.GetPreviewWidth, pdFxPreview.GetPreviewHeight, dWidth, dHeight
    
    'Create a new, smaller image at those dimensions
    If (dWidth < srcWidth) Or (dHeight < srcHeight) Then
        
        smallDIB.CreateBlank dWidth, dHeight, 32, 0
        
        Select Case m_RotateTarget
        
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
                    .Width = PDImages.GetActiveImage.Width
                    .Height = PDImages.GetActiveImage.Height
                End With
                
                PDImages.GetActiveImage.GetCompositedRect smallDIB, dstRectF, srcRectF, GP_IM_HighQualityBicubic, , CLC_Generic
            
            Case PD_AT_SINGLELAYER
                GDIPlusResizeDIB smallDIB, 0, 0, dWidth, dHeight, PDImages.GetActiveImage.GetActiveDIB, 0, 0, PDImages.GetActiveImage.GetActiveDIB.GetDIBWidth, PDImages.GetActiveImage.GetActiveDIB.GetDIBHeight, GP_IM_HighQualityBicubic
            
        End Select
        
    'The source image or layer is tiny; just use the whole thing!
    Else
    
        Select Case m_RotateTarget
        
            Case PD_AT_WHOLEIMAGE
                PDImages.GetActiveImage.GetCompositedImage smallDIB
            
            Case PD_AT_SINGLELAYER
                smallDIB.CreateFromExistingDIB PDImages.GetActiveImage.GetActiveDIB
            
        End Select
        
    End If
    
    'Give the preview object a copy of this image data so it can show it to the user if requested
    pdFxPreview.SetOriginalImage smallDIB
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Redraw the on-screen preview of the rotated image
Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then RotateArbitrary GetFunctionParamString(), True
End Sub

Private Sub sltAngle_Change()
    UpdatePreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub
