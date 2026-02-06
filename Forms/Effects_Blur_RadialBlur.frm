VERSION 5.00
Begin VB.Form FormRadialBlur 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Radial blur"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12030
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
   ScaleWidth      =   802
   Begin PhotoDemon.pdButtonStrip btsRender 
      Height          =   1095
      Left            =   6000
      TabIndex        =   4
      Top             =   3360
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1931
      Caption         =   "render emphasis"
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   12030
      _ExtentX        =   21220
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
   Begin PhotoDemon.pdSlider sltRadius 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   1680
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "angle"
      Min             =   1
      Max             =   360
      SigDigits       =   1
      Value           =   1
      DefaultValue    =   1
   End
   Begin PhotoDemon.pdCheckBox chkSymmetry 
      Height          =   300
      Left            =   6120
      TabIndex        =   3
      Top             =   2760
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   582
      Caption         =   "blur symmetrically"
   End
End
Attribute VB_Name = "FormRadialBlur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Radial Blur Tool
'Copyright 2013-2026 by Tanner Helland
'Created: 26/August/13
'Last updated: 27/July/17
'Last update: performance improvements, migrate to XML params
'
'To my knowledge, this tool is the first of its kind in VB6 - a radial blur tool that supports variable angles,
' and capable of operating in real-time.  This function is mostly just a wrapper to PD's horizontal blur and
' polar coordinate conversion functions; they do all the heavy lifting, as you can see from the code below.
'
'Performance is pretty good, all things considered, but be careful in the IDE. I STRONGLY recommend compiling
' the project before applying any actions at a large radius.
'
'All source code in this file is licensed under a modified BSD license. This means you may use the code in your own
' projects IF you provide attribution. For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Apply radial blur to an image
'Inputs: angle of the blur, and whether it should be symmetrical (e.g. equal in +/- angle amounts)
Public Sub RadialBlurFilter(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    If (Not toPreview) Then Message "Applying radial blur..."
        
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim bRadius As Double, blurSymmetrically As Boolean, useBilinear As Boolean
    
    With cParams
        bRadius = .GetDouble("radius", sltRadius.Value)
        blurSymmetrically = .GetBool("symmetry", False)
        useBilinear = .GetBool("bilinear", True)
    End With

    'Create a local array and point it at the pixel data of the current image
    Dim dstSA As SafeArray2D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic, , , True
    
    'Create a second local array. This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent blurred pixel values from spreading across the image as we go.)
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.CreateFromExistingDIB workingDIB
    
    'By dividing blur radius by 360 (its maximum value), we can use it as a fractional amount to determine the strength of our horizontal blur.
    Dim actualBlurSize As Long
    actualBlurSize = (bRadius / 360#) * curDIBValues.Width
    If (actualBlurSize < 1) Then actualBlurSize = 1
    
    Dim finalX As Long, finalY As Long
    finalX = workingDIB.GetDIBWidth
    finalY = workingDIB.GetDIBHeight
    
    'Because this function actually wraps three functions, calculating the progress bar maximum is a bit convoluted
    Dim newProgBarMax As Long
    newProgBarMax = finalY * 3
    
    'Start by converting the image to polar coordinates, using a specific set of actions to maximize quality
    If CreatePolarCoordDIB(1, 100, pdeo_Clamp, useBilinear, srcDIB, workingDIB, toPreview, newProgBarMax) Then
    
        'We now need to do something a little unconventional.  When converting to polar coordinates, the line running from
        ' the top-center of the image to the center point ends up being separated onto the full left and right sides of the
        ' polar coordinate image.  Because PD's box blur does not wrap around image edges (for performance reasons), this line
        ' doesn't get blurred properly, and when we convert back to rectangular coordinates, it forms a visible abberation
        ' running from the top-center of the image to the center point.  To prevent this, we must create a temporary copy of
        ' the image that is larger (by the width of the blur radius) on both sides.  We then place the polar coord image in the
        ' center of this larger image, then copy the relevant edge pixels onto either side of it.  When the blur is complete,
        ' we copy back just this center portion before converting from polar to rect coords for the final time.  This results
        ' in a proper blur.  (Hope you caught all that!  :p)
        
        'Start by calculating the temporary image's size and offset
        Dim srcWidth As Long, srcHeight As Long
        srcWidth = workingDIB.GetDIBWidth
        srcHeight = workingDIB.GetDIBHeight
        
        Dim dstWidth As Long
        dstWidth = srcWidth + actualBlurSize * 2
        
        Dim dstX As Long
        dstX = (dstWidth - srcWidth) \ 2
        
        'Create a temporary DIB to hold the blurred image
        Dim tmpDIB As pdDIB
        Set tmpDIB = New pdDIB
        tmpDIB.CreateBlank dstWidth, srcHeight, workingDIB.GetDIBColorDepth
        
        'Bitblt the original image onto the center of the temporary canvas
        GDI.BitBltWrapper tmpDIB.GetDIBDC, dstX, 0, srcWidth, srcHeight, workingDIB.GetDIBDC, 0, 0, vbSrcCopy
        
        'Apply two more blts - each of these will mirror an edge section of the source image
        GDI.BitBltWrapper tmpDIB.GetDIBDC, 0, 0, dstX, srcHeight, workingDIB.GetDIBDC, srcWidth - dstX, 0, vbSrcCopy
        GDI.BitBltWrapper tmpDIB.GetDIBDC, dstX + srcWidth, 0, dstX, srcHeight, workingDIB.GetDIBDC, 0, 0, vbSrcCopy
        
        'Change the srcDIB to be the same size as this working DIB, so it can receive the fully blurred image
        srcDIB.CreateBlank tmpDIB.GetDIBWidth, tmpDIB.GetDIBHeight, workingDIB.GetDIBColorDepth
    
        'Now we can apply the box blur to the temporary DIB, using the blur radius supplied by the user
        Dim leftRadius As Long
        If blurSymmetrically Then leftRadius = actualBlurSize Else leftRadius = 0
        
        If CreateHorizontalBlurDIB(leftRadius, actualBlurSize, tmpDIB, srcDIB, toPreview, newProgBarMax, finalY) Then
        
            'Copy the blurred results of the source DIB back into the temporary DIB
            tmpDIB.CreateFromExistingDIB srcDIB
            
            'Resize the source DIB to match the original image
            srcDIB.CreateBlank workingDIB.GetDIBWidth, workingDIB.GetDIBHeight, workingDIB.GetDIBColorDepth
            
            'Copy the correct chunk of the temporary DIB into the source DIB
            GDI.BitBltWrapper srcDIB.GetDIBDC, 0, 0, srcWidth, srcHeight, tmpDIB.GetDIBDC, dstX, 0, vbSrcCopy
            tmpDIB.EraseDIB
            
            'Finally, convert back to rectangular coordinates, using the opposite parameters of the first conversion
            CreatePolarCoordDIB 0, 100, pdeo_Clamp, useBilinear, srcDIB, workingDIB, toPreview, newProgBarMax, finalY * 2
            
        End If
        
        Set tmpDIB = Nothing
        
    End If
    
    Set srcDIB = Nothing
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingDIB
    EffectPrep.FinalizeImageData toPreview, dstPic, True
    
End Sub

Private Sub btsRender_Click(ByVal buttonIndex As Long)
    UpdatePreview
End Sub

Private Sub chkSymmetry_Click()
    UpdatePreview
End Sub

Private Sub cmdBar_OKClick()
    Process "Radial blur", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub Form_Load()
    
    'Disable previews until the form is fully initialized
    cmdBar.SetPreviewStatus False
    
    btsRender.AddItem "speed", 0
    btsRender.AddItem "accuracy", 1
    btsRender.ListIndex = 1
    
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

'Render a new effect preview
Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then RadialBlurFilter GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "radius", sltRadius.Value
        .AddParam "symmetry", chkSymmetry.Value
        .AddParam "bilinear", (btsRender.ListIndex = 1)
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
