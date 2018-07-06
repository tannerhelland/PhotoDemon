VERSION 5.00
Begin VB.Form FormResizeContentAware 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Content-aware resize"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   9705
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
   ScaleHeight     =   331
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   647
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   4215
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   1323
      AutoloadLastPreset=   -1  'True
   End
   Begin PhotoDemon.pdResize ucResize 
      Height          =   2850
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   5027
   End
   Begin PhotoDemon.pdLabel lblFlatten 
      Height          =   645
      Left            =   240
      Top             =   3480
      Visible         =   0   'False
      Width           =   9330
      _ExtentX        =   0
      _ExtentY        =   0
      Alignment       =   2
      Caption         =   "Note: this operation will flatten the image before resizing it."
      ForeColor       =   4210752
      Layout          =   1
   End
End
Attribute VB_Name = "FormResizeContentAware"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Content-Aware Resize (e.g. "content-aware scale" in Photoshop, "liquid rescale" in GIMP) Dialog
'Copyright 2014-2018 by Tanner Helland
'Created: 06/January/14
'Last updated: 29/July/14
'Last update: fixed some 32bpp issues, added serpentine scanning for ideal treatment of images on uniform backgrounds
'
'Content-aware scaling is a very exciting addition to PhotoDemon 6.2.  (As a comparison, PhotoShop didn't gain this
' feature until CS4, so it's pretty cutting-edge stuff!)
'
'Normal scaling algorithms work by shrinking or enlarging all image pixels equally.  Such algorithms make no distinction
' between visually important pixels and visually unimportant ones.  Unfortunately, when the aspect ratio of an image is
' changed using such an algorithm, noticeable distortion will result, and the end result will typically be unpleasant.
'
'Content-aware scaling circumvents this by selectively removing the least visually important parts of an image
' (as determined by some type of per-pixel "energy" calculation).  By preferentially removing uninteresting pixels
' before interesting ones, important parts of an image can be preserved while uninteresting parts are removed.  The
' result is often a much more aesthetically pleasing image, even under severe aspect ratio changes.
'
'For reference, the original 2007 paper that first proposed this technique - called "seam carving" is available here:
' http://www.win.tue.nl/~wstahw/edu/2IV05/seamcarving.pdf
'
'I have written PhotoDemon's implementation from scratch, using the original paper as my primary resource.  Unfortunately,
' my current implementation is quite slow (though still faster than many other implementations!) on account of all the
' seam finding operations that must be performed.  I am investigating ways to further improve the algorithm's performance,
' but I remain worried that this task may prove a bit much for VB6.  We'll see.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'This dialog can be used to resize the full image, or a single layer.  The requested target will be stored here,
' and can be externally accessed by the ResizeTarget property.
Private m_ResizeTarget As PD_ACTION_TARGET

Public Property Let ResizeTarget(newTarget As PD_ACTION_TARGET)
    m_ResizeTarget = newTarget
End Property

'OK button
Private Sub cmdBar_OKClick()
    
    'Place all settings in an XML string
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    With cParams
        .AddParam "width", ucResize.ResizeWidth
        .AddParam "height", ucResize.ResizeHeight
        .AddParam "unit", ucResize.UnitOfMeasurement
        .AddParam "ppi", ucResize.ResizeDPIAsPPI
        If (m_ResizeTarget = PD_AT_SINGLELAYER) Then
            .AddParam "target", "layer"
        Else
            .AddParam "target", "image"
        End If
    End With
    
    If (m_ResizeTarget = PD_AT_WHOLEIMAGE) Then
        Process "Content-aware image resize", , cParams.GetParamString(), UNDO_Image
    ElseIf (m_ResizeTarget = PD_AT_SINGLELAYER) Then
        Process "Content-aware layer resize", , cParams.GetParamString(), UNDO_Layer
    End If

End Sub

'I'm not sure that randomize serves any purpose on this dialog, but as I don't have a way to hide that button at
' present, simply randomize the width/height to +/- the current image's width/height divided by two.
Private Sub cmdBar_RandomizeClick()
    
    ucResize.LockAspectRatio = False
    
    Select Case m_ResizeTarget
    
        Case PD_AT_WHOLEIMAGE
            ucResize.ResizeWidthInPixels = (pdImages(g_CurrentImage).Width / 2) + (Rnd * pdImages(g_CurrentImage).Width)
            ucResize.ResizeHeightInPixels = (pdImages(g_CurrentImage).Height / 2) + (Rnd * pdImages(g_CurrentImage).Height)
        
        Case PD_AT_SINGLELAYER
            ucResize.ResizeWidthInPixels = (pdImages(g_CurrentImage).GetActiveLayer.GetLayerWidth(False) / 2) + (Rnd * pdImages(g_CurrentImage).GetActiveLayer.GetLayerWidth(False))
            ucResize.ResizeHeightInPixels = (pdImages(g_CurrentImage).GetActiveLayer.GetLayerHeight(False) / 2) + (Rnd * pdImages(g_CurrentImage).GetActiveLayer.GetLayerHeight(False))
    
    End Select
    
End Sub

Private Sub cmdBar_ResetClick()

    ucResize.UnitOfMeasurement = mu_Pixels
    
    'Automatically set the width and height text boxes to match the relevant image or layer's current dimensions
    Select Case m_ResizeTarget
    
        Case PD_AT_WHOLEIMAGE
            ucResize.SetInitialDimensions pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height, pdImages(g_CurrentImage).GetDPI
        
        Case PD_AT_SINGLELAYER
            ucResize.SetInitialDimensions pdImages(g_CurrentImage).GetActiveLayer.GetLayerWidth(False), pdImages(g_CurrentImage).GetActiveLayer.GetLayerHeight(False), pdImages(g_CurrentImage).GetDPI
    
    End Select

    ucResize.LockAspectRatio = True
    
End Sub

Private Sub Form_Activate()

    'Set the dialog caption to match the current resize operation (resize image or resize single layer)
    Select Case m_ResizeTarget
        
        Case PD_AT_WHOLEIMAGE
            Me.Caption = g_Language.TranslateMessage("Content-aware image resize")
        
        Case PD_AT_SINGLELAYER
            Me.Caption = g_Language.TranslateMessage("Content-aware layer resize")
        
    End Select

    'Automatically set the width and height text boxes to match the image's current dimensions
    ucResize.UnitOfMeasurement = mu_Pixels
    
    Select Case m_ResizeTarget
        
        Case PD_AT_WHOLEIMAGE
            ucResize.SetInitialDimensions pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height, pdImages(g_CurrentImage).GetDPI
            
        Case PD_AT_SINGLELAYER
            ucResize.SetInitialDimensions pdImages(g_CurrentImage).GetActiveLayer.GetLayerWidth(False), pdImages(g_CurrentImage).GetActiveLayer.GetLayerHeight(False), pdImages(g_CurrentImage).GetDPI
        
    End Select
    
    'Warn about flattening if the entire image is being content-aware-resized
    lblFlatten.Visible = (m_ResizeTarget = PD_AT_WHOLEIMAGE)
    
    ucResize.LockAspectRatio = False

End Sub

'LOAD dialog
Private Sub Form_Load()
    
    'Automatically set the width and height text boxes to match the image's current dimensions.  (Note that we must
    ' do this again in the Activate step, as the last-used settings will automatically override these values.  However,
    ' if we do not also provide these values here, the resize control may attempt to set parameters while having
    ' a width/height/resolution of 0, which will cause divide-by-zero errors.)
    Select Case m_ResizeTarget
    
        Case PD_AT_WHOLEIMAGE
            ucResize.SetInitialDimensions pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height, pdImages(g_CurrentImage).GetDPI
            
        Case PD_AT_SINGLELAYER
            ucResize.SetInitialDimensions pdImages(g_CurrentImage).GetActiveLayer.GetLayerWidth(False), pdImages(g_CurrentImage).GetActiveLayer.GetLayerHeight(False), pdImages(g_CurrentImage).GetDPI
        
    End Select
    
    'If the current image has more than one layer, warn the user that this action will flatten the image.
    If pdImages(g_CurrentImage).GetNumOfLayers > 1 Then
        lblFlatten.Visible = True
    Else
        lblFlatten.Visible = False
    End If
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Small wrapper for the seam carve function
Public Sub SmartResizeImage(ByVal xmlParams As String)
    
    'Parse incoming parameters from the XML string
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.SetParamString xmlParams
    
    Dim imgWidth As Long, imgHeight As Long, imgResizeUnit As PD_MeasurementUnit, imgDPI As Double
    Dim thingToResize As PD_ACTION_TARGET
    
    With cParams
        imgWidth = .GetDouble("width")
        imgHeight = .GetDouble("height")
        imgResizeUnit = .GetLong("unit", mu_Pixels)
        imgDPI = .GetDouble("ppi", 96)
        thingToResize = .GetLong("target", PD_AT_WHOLEIMAGE)
    End With
    
    'If the entire image is being resized, some extra preparation is required
    If (thingToResize = PD_AT_WHOLEIMAGE) Then
        
        'If a selection is active, remove it now
        If pdImages(g_CurrentImage).IsSelectionActive Then
            pdImages(g_CurrentImage).SetSelectionActive False
            pdImages(g_CurrentImage).MainSelection.LockRelease
        End If
                   
        'Flatten the image; note that we route this through the central processor, so that a proper Undo/Redo entry
        ' is created.  (This is especially important if the user presses ESC to cancel the seam-carving step.)
        Process "Flatten image", , , UNDO_Image
        
    End If

    'Create a temporary DIB, which will be passed to the master SeamCarveDIB function
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    tmpDIB.CreateFromExistingDIB pdImages(g_CurrentImage).GetActiveDIB
    If tmpDIB.GetAlphaPremultiplication Then tmpDIB.SetAlphaPremultiplication False
    
    'In past versions of the software, we could assume the passed measurements were always in pixels,
    ' but that is no longer the case!  Using the supplied "unit of measurement", convert the passed
    ' width and height values to pixel measurements.
    imgWidth = ConvertOtherUnitToPixels(imgResizeUnit, imgWidth, imgDPI, pdImages(g_CurrentImage).GetActiveLayer.GetLayerWidth(False))
    imgHeight = ConvertOtherUnitToPixels(imgResizeUnit, imgHeight, imgDPI, pdImages(g_CurrentImage).GetActiveLayer.GetLayerHeight(False))
    
    'Pass the temporary DIB to the master seam carve function
    If Me.SeamCarveDIB(tmpDIB, imgWidth, imgHeight) Then
        
        'Premultiply alpha
        If (Not tmpDIB.GetAlphaPremultiplication) Then tmpDIB.SetAlphaPremultiplication True
        
        'Copy the newly resized DIB back into its parent image
        pdImages(g_CurrentImage).GetActiveLayer.layerDIB.CreateFromExistingDIB tmpDIB
        Set tmpDIB = Nothing
        
        'Notify the parent of the change
        pdImages(g_CurrentImage).NotifyImageChanged UNDO_Layer, pdImages(g_CurrentImage).GetActiveLayerIndex
        
        'Update the main image's size and DPI values as necessary
        If thingToResize = PD_AT_WHOLEIMAGE Then
            pdImages(g_CurrentImage).UpdateSize False, imgWidth, imgHeight
            pdImages(g_CurrentImage).SetDPI imgDPI, imgDPI
            DisplaySize pdImages(g_CurrentImage)
        End If
        
        'Fit the new image on-screen and redraw its viewport
        ViewportEngine.Stage1_InitializeBuffer pdImages(g_CurrentImage), FormMain.MainCanvas(0)
        
    'Failsafe check for seam carving failure; this should never trigger
    Else
        pdImages(g_CurrentImage).UndoManager.RestoreUndoData
        Interface.NotifyImageChanged g_CurrentImage
    End If
    
    Message "Finished."

End Sub

'Resize a DIB via seam carving ("content-aware resize" in Photoshop, or "liquid rescale" in GIMP).
Public Function SeamCarveDIB(ByRef srcDIB As pdDIB, ByVal imgWidth As Long, ByVal imgHeight As Long) As Boolean

    'For more information on how seam-carving works, visit http://en.wikipedia.org/wiki/Seam_carving
    
    Message "Initializing content-aware resize engine..."
    
    'Before starting on seam carving, we must first generate an "energy map" for the image.  This can be done many ways,
    ' but since PD has a nice artistic contour algorithm already available, let's use that.
    Dim energyDIB As pdDIB
    Set energyDIB = New pdDIB
    energyDIB.CreateFromExistingDIB srcDIB
    
    GrayscaleDIB energyDIB, True
    CreateContourDIB True, srcDIB, energyDIB, True
    
    'Create a seam carver class, which will handle the technical details of the carve
    Dim seamCarver As pdSeamCarving
    Set seamCarver = New pdSeamCarving
    
    'Give the seam carving class a copy of our source and energy images
    seamCarver.SetSourceImage srcDIB
    seamCarver.SetEnergyImage energyDIB
    
    'We no longer need a copy of the energy image, so release it.
    Set energyDIB = Nothing
    
    Message "Applying content-aware resize..."
    
    'This initial seam-carving algorithm is not particularly well-implemented, but that's okay.  It's a starting point!
    seamCarver.StartSeamCarve imgWidth, imgHeight
    
    'Release the progress bar
    ReleaseProgressBar
    
    'Check for user cancellation; if none occurred, copy the seam-carved image into place
    If (Not g_cancelCurrentAction) Then
        srcDIB.CreateFromExistingDIB seamCarver.GetCarvedImage()
        SeamCarveDIB = True
    Else
        SeamCarveDIB = False
    End If
    
End Function

