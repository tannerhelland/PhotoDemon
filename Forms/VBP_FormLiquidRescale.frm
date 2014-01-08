VERSION 5.00
Begin VB.Form FormLiquidResize 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Liquid resize (content-aware scaling)"
   ClientHeight    =   3420
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
   ScaleHeight     =   228
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   647
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   2670
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
   Begin PhotoDemon.textUpDown tudWidth 
      Height          =   405
      Left            =   4320
      TabIndex        =   1
      Top             =   705
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   714
      Max             =   32767
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.textUpDown tudHeight 
      Height          =   405
      Left            =   4320
      TabIndex        =   2
      Top             =   1335
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   714
      Max             =   32767
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblWarning 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "WARNING! This tool is currently under heavy development.  It may not work as intended."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1815
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   2775
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblTitle 
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
      Index           =   0
      Left            =   3120
      TabIndex        =   8
      Top             =   240
      Width           =   990
   End
   Begin VB.Label lblAspectRatio 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "new aspect ratio will be"
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
      Left            =   3495
      TabIndex        =   7
      Top             =   1950
      Width           =   2490
   End
   Begin VB.Label lblHeightUnit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "pixels"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   5610
      TabIndex        =   6
      Top             =   1365
      Width           =   855
   End
   Begin VB.Label lblWidthUnit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "pixels"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   480
      Left            =   5610
      TabIndex        =   5
      Top             =   735
      Width           =   855
   End
   Begin VB.Label lblHeight 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "height:"
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
      Left            =   3480
      TabIndex        =   4
      Top             =   1365
      Width           =   750
   End
   Begin VB.Label lblWidth 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "width:"
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
      Left            =   3480
      TabIndex        =   3
      Top             =   735
      Width           =   675
   End
End
Attribute VB_Name = "FormLiquidResize"
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

'OK button
Private Sub cmdBar_OKClick()
    Process "Content-aware resize", , buildParams(tudWidth, tudHeight)
End Sub

'I'm not sure that randomize serves any purpose on this dialog, but as I don't have a way to hide that button at
' present, simply randomize the width/height to +/- the current image's width/height divided by two.
Private Sub cmdBar_RandomizeClick()
    tudWidth = (pdImages(g_CurrentImage).Width / 2) + (Rnd * pdImages(g_CurrentImage).Width)
    tudHeight = (pdImages(g_CurrentImage).Height / 2) + (Rnd * pdImages(g_CurrentImage).Height)
End Sub

Private Sub cmdBar_ResetClick()

    'Automatically set the width and height text boxes to match the image's current dimensions
    tudWidth.Value = pdImages(g_CurrentImage).Width
    tudHeight.Value = pdImages(g_CurrentImage).Height
    
End Sub

'Upon form activation, determine the ratio between the width and height of the image
Private Sub Form_Activate()
    
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
        
End Sub

'Certain actions are done at LOAD time instead of ACTIVATE time to minimize visible flickering
Private Sub Form_Load()
    
    'To prevent aspect ratio changes to one box resulting in recursion-type changes to the other, we only
    ' allow one box at a time to be updated.
    allowedToUpdateWidth = True
    allowedToUpdateHeight = True
    
    'Establish ratios
    wRatio = pdImages(g_CurrentImage).Width / pdImages(g_CurrentImage).Height
    hRatio = pdImages(g_CurrentImage).Height / pdImages(g_CurrentImage).Width
    
    'Automatically set the width and height text boxes to match the image's current dimensions
    tudWidth.Value = pdImages(g_CurrentImage).Width
    tudHeight.Value = pdImages(g_CurrentImage).Height
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Small wrapper for the seam carve function
Public Sub SmartResizeImage(ByVal iWidth As Long, ByVal iHeight As Long)

    'Create a temporary layer, which will be passed to the master SeamCarveLayer function
    Dim tmpLayer As pdLayer
    Set tmpLayer = New pdLayer
    tmpLayer.createFromExistingLayer pdImages(g_CurrentImage).getActiveLayer
    
    'Pass the temporary layer to the master seam carve function
    SeamCarveLayer tmpLayer, iWidth, iHeight
    
    'Copy the newly resized layer back into its parent image
    pdImages(g_CurrentImage).mainLayer.createFromExistingLayer tmpLayer
    Set tmpLayer = Nothing
    
    'Update the main image's size values
    pdImages(g_CurrentImage).updateSize
    DisplaySize pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height
    
    'Fit the new image on-screen and redraw its viewport
    PrepareViewport pdImages(g_CurrentImage).containingForm, "Content-aware resize"
    
    Message "Finished."

End Sub

'Resize a layer via seam carving ("content-aware resize" in Photoshop, or "liquid rescale" in GIMP).
Public Function SeamCarveLayer(ByRef srcLayer As pdLayer, ByVal iWidth As Long, ByVal iHeight As Long) As Boolean

    'For more information on how seam-carving works, visit http://en.wikipedia.org/wiki/Seam_carving

    'If the image contains an active selection, disable it before doing any transformation work
    If pdImages(g_CurrentImage).selectionActive Then
        pdImages(g_CurrentImage).selectionActive = False
        pdImages(g_CurrentImage).mainSelection.lockRelease
    End If
    
    Message "Initializing content-aware resize engine..."
    
    'Before starting on seam carving, we must first generate an "energy map" for the image.  This can be done many ways,
    ' but since PD has a nice artistic contour algorithm already available, let's use that.
    Dim energyLayer As pdLayer
    Set energyLayer = New pdLayer
    energyLayer.createFromExistingLayer srcLayer
    CreateContourLayer True, srcLayer, energyLayer, True
    
    'Create a seam carver class, which will handle the technical details of the carve
    Dim seamCarver As pdSeamCarving
    Set seamCarver = New pdSeamCarving
    
    'Give the seam carving class a copy of our source and energy images
    seamCarver.setSourceImage srcLayer
    seamCarver.setEnergyImage energyLayer
    
    'We no longer need a copy of the energy image, so release it.
    Set energyLayer = Nothing
    
    Message "Applying content-aware resize..."
    
    'This initial seam-carving algorithm is not particularly well-implemented, but that's okay.  It's a starting point!
    seamCarver.startSeamCarve iWidth, iHeight
    
    'TESTING ONLY!!!  Replace srcLayer with an energy map of the seam carve function
    srcLayer.createFromExistingLayer seamCarver.getCarvedImage()
    
End Function

'PhotoDemon now displays an approximate aspect ratio for the selected values.  This can be helpful when
' trying to select new width/height values for a specific application with a set aspect ratio (e.g. 16:9 screens).
Private Sub updateAspectRatio()

    'This sub may be called before all on-screen controls have been filled.  To prevent overflow errors, check for
    ' DIV-BY-0 in advance.
    If tudHeight = 0 Then Exit Sub

    Dim wholeNumber As Double, Numerator As Double, Denominator As Double
    
    If tudWidth.IsValid And tudHeight.IsValid Then
        convertToFraction tudWidth / tudHeight, wholeNumber, Numerator, Denominator, 4, 99.9
        
        'Aspect ratios are typically given in terms of base 10 if possible, so change values like 8:5 to 16:10
        If CLng(Denominator) = 5 Then
            Numerator = Numerator * 2
            Denominator = Denominator * 2
        End If
        
        lblAspectRatio.Caption = g_Language.TranslateMessage("new aspect ratio will be %1:%2", Numerator, Denominator)
    End If

End Sub

'If "Lock Image Aspect Ratio" is selected, these two routines keep all values in sync
Private Sub tudHeight_Change()
    updateAspectRatio
End Sub

Private Sub tudWidth_Change()
    updateAspectRatio
End Sub
