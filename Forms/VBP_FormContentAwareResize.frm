VERSION 5.00
Begin VB.Form FormResizeContentAware 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Content-aware resize"
   ClientHeight    =   4275
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
   ScaleHeight     =   285
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   647
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   3525
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   1323
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoloadLastPreset=   -1  'True
   End
   Begin PhotoDemon.smartResize ucResize 
      Height          =   2850
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   5027
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblSize 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "new size:"
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
      Left            =   480
      TabIndex        =   2
      Top             =   120
      Width           =   990
   End
End
Attribute VB_Name = "FormResizeContentAware"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Content-Aware Resize(e.g. "content-aware scale" in Photoshop, "liquid rescale" in GIMP) Dialog
'Copyright ©2013-2014 by Tanner Helland
'Created: 06/January/14
'Last updated: 08/January/14
'Last update: finished work on enlarging support
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
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Used for maintaining ratios when the check box is clicked
Private wRatio As Double, hRatio As Double
Dim allowedToUpdateWidth As Boolean, allowedToUpdateHeight As Boolean

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

Private Sub cmdBar_ExtraValidations()
    If Not ucResize.IsValid(True) Then cmdBar.validationFailed
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Content-aware resize", , buildParams(ucResize.imgWidth, ucResize.imgHeight, ucResize.unitOfMeasurement, ucResize.imgDPIAsPPI)
End Sub

'I'm not sure that randomize serves any purpose on this dialog, but as I don't have a way to hide that button at
' present, simply randomize the width/height to +/- the current image's width/height divided by two.
Private Sub cmdBar_RandomizeClick()
    
    ucResize.lockAspectRatio = False
    ucResize.imgWidthInPixels = (pdImages(g_CurrentImage).Width / 2) + (Rnd * pdImages(g_CurrentImage).Width)
    ucResize.imgHeightInPixels = (pdImages(g_CurrentImage).Height / 2) + (Rnd * pdImages(g_CurrentImage).Height)
    
End Sub

Private Sub cmdBar_ResetClick()

    'Automatically set the width and height text boxes to match the image's current dimensions
    ucResize.unitOfMeasurement = MU_PIXELS
    ucResize.setInitialDimensions pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height, pdImages(g_CurrentImage).getDPI
    ucResize.lockAspectRatio = True
    
End Sub

'LOAD dialog
Private Sub Form_Load()
    
    'Automatically set the width and height text boxes to match the image's current dimensions
    ucResize.setInitialDimensions pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height, pdImages(g_CurrentImage).getDPI
    
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Small wrapper for the seam carve function
Public Sub SmartResizeImage(ByVal iWidth As Long, ByVal iHeight As Long, Optional ByVal unitOfMeasurement As MeasurementUnit = MU_PIXELS, Optional ByVal iDPI As Long)

    'TODO: make this function work with layers.  It may be best to restrict the function to layers, and warn of
    '       flattening if used at the Image level.

    'Create a temporary DIB, which will be passed to the master SeamCarveDIB function
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    tmpDIB.createFromExistingDIB pdImages(g_CurrentImage).getActiveDIB
    
    'In past versions of the software, we could assume the passed measurements were always in pixels,
    ' but that is no longer the case!  Using the supplied "unit of measurement", convert the passed
    ' width and height values to pixel measurements.
    iWidth = convertOtherUnitToPixels(unitOfMeasurement, iWidth, iDPI, pdImages(g_CurrentImage).Width)
    iHeight = convertOtherUnitToPixels(unitOfMeasurement, iHeight, iDPI, pdImages(g_CurrentImage).Height)
    
    'Pass the temporary DIB to the master seam carve function
    SeamCarveDIB tmpDIB, iWidth, iHeight
    
    'Copy the newly resized DIB back into its parent image
    'pdImages(g_CurrentImage).mainDIB.createFromExistingDIB tmpDIB
    Set tmpDIB = Nothing
    
    'Update the main image's size and DPI values
    pdImages(g_CurrentImage).updateSize
    pdImages(g_CurrentImage).setDPI iDPI, iDPI
    DisplaySize pdImages(g_CurrentImage)
    
    'Fit the new image on-screen and redraw its viewport
    PrepareViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0), "Content-aware resize"
    
    Message "Finished."

End Sub

'Resize a DIB via seam carving ("content-aware resize" in Photoshop, or "liquid rescale" in GIMP).
Public Function SeamCarveDIB(ByRef srcDIB As pdDIB, ByVal iWidth As Long, ByVal iHeight As Long) As Boolean

    'For more information on how seam-carving works, visit http://en.wikipedia.org/wiki/Seam_carving

    'If the image contains an active selection, disable it before doing any transformation work
    If pdImages(g_CurrentImage).selectionActive Then
        pdImages(g_CurrentImage).selectionActive = False
        pdImages(g_CurrentImage).mainSelection.lockRelease
    End If
    
    Message "Initializing content-aware resize engine..."
    
    'Before starting on seam carving, we must first generate an "energy map" for the image.  This can be done many ways,
    ' but since PD has a nice artistic contour algorithm already available, let's use that.
    Dim energyDIB As pdDIB
    Set energyDIB = New pdDIB
    energyDIB.createFromExistingDIB srcDIB
    CreateContourDIB True, srcDIB, energyDIB, True
    
    'Create a seam carver class, which will handle the technical details of the carve
    Dim seamCarver As pdSeamCarving
    Set seamCarver = New pdSeamCarving
    
    'Give the seam carving class a copy of our source and energy images
    seamCarver.setSourceImage srcDIB
    seamCarver.setEnergyImage energyDIB
    
    'We no longer need a copy of the energy image, so release it.
    Set energyDIB = Nothing
    
    Message "Applying content-aware resize..."
    
    'This initial seam-carving algorithm is not particularly well-implemented, but that's okay.  It's a starting point!
    seamCarver.startSeamCarve iWidth, iHeight
    
    'Release the progress bar
    releaseProgressBar
    
    'Check for user cancellation; if none occurred, copy the seam-carved image into place
    If Not cancelCurrentAction Then srcDIB.createFromExistingDIB seamCarver.getCarvedImage()
    
End Function
