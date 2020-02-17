VERSION 5.00
Begin VB.Form FormComicBook 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Comic book"
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
   Begin PhotoDemon.pdButtonStrip btsStrength 
      Height          =   1095
      Left            =   6000
      TabIndex        =   4
      Top             =   3600
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1931
      Caption         =   "brush smoothing"
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
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
   End
   Begin PhotoDemon.pdSlider sltInk 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   1680
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "ink"
      Max             =   100
      Value           =   50
      DefaultValue    =   50
   End
   Begin PhotoDemon.pdSlider sltColor 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   2640
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   1270
      Caption         =   "brush size"
      Max             =   50
      Value           =   5
      DefaultValue    =   5
   End
End
Attribute VB_Name = "FormComicBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Comic Book Image Effect
'Copyright 2013-2020 by Tanner Helland
'Created: 02/Feb/13 (ish... I didn't write it down, alas)
'Last updated: 23/May/16
'Last update: optimize function a bit
'
'PhotoDemon has provided a "comic book" effect for a long time, but despite going through many incarnations, it always
' used low-quality, "quick and dirty" approximations.
'
'In July '14, this changed, and the entire tool was rethought from the ground up.  A dialog is now provided, with
' various user-settable options.  This yields much more flexible results, and the use of PD's central compositor for
' overlaying intermediate image copies keeps things nice and fast.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Apply a "comic book" effect to an image
'Inputs:
' 1) strength of the inking
' 2) color smudging, which controls the radius of the median effect applied to the base image
Public Sub fxComicBook(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    'Parse out individual parameters
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.SetParamString effectParams
    
    Dim inkOpacity As Long, colorSmudge As Long, colorStrength As Long
    
    With cParams
        inkOpacity = .GetLong("ink", sltInk.Value)
        colorSmudge = .GetLong("size", sltColor.Value)
        colorStrength = .GetLong("smoothing", btsStrength.ListIndex)
    End With
    
    If (Not toPreview) Then Message "Animating image (stage %1 of %2)...", 1, 3 + colorStrength
    
    'Initiate PhotoDemon's central image handler
    Dim dstSA As SafeArray2D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic
    
    'During a preview, the smudge radius must be reduced to match the preview size
    If toPreview Then colorSmudge = colorSmudge * curDIBValues.previewModifier
    
    'During a preview, ink opacity is artificially reduced to give a better idea of how the final image will appear
    If toPreview Then inkOpacity = inkOpacity * curDIBValues.previewModifier
    
    'If this is not a preview, calculate a new maximum progress bar value.  This changes depending on the number of
    ' iterations we must run to obtain a proper colored image.
    Dim numOfSteps As Long, newProgBarMax As Long
    
    If (inkOpacity > 0) Then
        numOfSteps = 2 + colorStrength
    Else
        numOfSteps = 1 + colorStrength
    End If
    
    newProgBarMax = workingDIB.GetDIBWidth * numOfSteps + ((colorStrength + 1) * workingDIB.GetDIBWidth)
    If (Not toPreview) Then SetProgBarMax newProgBarMax
    
    'If the user wants the image inked, we're actually going to generate a contour map now, before applying any coloring.
    ' This gives us more interesting lines to work with.
    If (inkOpacity > 0) Then
            
        If (Not toPreview) Then Message "Animating image (stage %1 of %2)...", 1, numOfSteps
        
        'Create two copies of the working image.  This filter overlays an inked image over a color-smudged version, and we need to
        ' handle these separately until the final step.
        Dim inkDIB As pdDIB
        Set inkDIB = New pdDIB
        inkDIB.CreateFromExistingDIB workingDIB
        Filters_Layers.CreateContourDIB True, workingDIB, inkDIB, toPreview, newProgBarMax, 0
        
        'Apply premultiplication to the DIB prior to compositing
        inkDIB.SetAlphaPremultiplication True
        
    End If
    
    'We now need to obtain the underlying color-smudged version of the source image
    If (colorSmudge > 0) Then
        
        'Use PD's excellent bilateral smoothing function to handle color smudging.
        Dim i As Long
        For i = 0 To colorStrength
            
            If (Not toPreview) Then
                If (numOfSteps > 1) Then
                    If (inkOpacity > 0) Then
                        Message "Animating image (stage %1 of %2)...", i + 2, numOfSteps
                    Else
                        Message "Animating image (stage %1 of %2)...", i + 1, numOfSteps
                    End If
                Else
                    Message "Animating image..."
                End If
            End If
            
            CreateBilateralDIB workingDIB, colorSmudge, 100, 10, toPreview, newProgBarMax, workingDIB.GetDIBWidth * (i * 2 + 1)
            
        Next i
        
    End If
    
    'Return the image to the premultiplied alpha space
    workingDIB.SetAlphaPremultiplication True
    
    'If the caller doesn't want us to ink the image, we're all done!
    If (inkOpacity > 0) Then
        
        'With an ink image and color image now available, we can composite the two into a single "comic book" image
        ' via PD's central compositor.
        Dim cComposite As pdCompositor
        Set cComposite = New pdCompositor
        cComposite.QuickMergeTwoDibsOfEqualSize workingDIB, inkDIB, BM_Difference, inkOpacity
        
    End If
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingDIB
    EffectPrep.FinalizeImageData toPreview, dstPic, True
    
End Sub

Private Sub btsStrength_Click(ByVal buttonIndex As Long)
    UpdatePreview
End Sub

Private Sub cmdBar_OKClick()
    Process "Comic book", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub Form_Load()
    
    'Disable previews until the dialog is fully loaded
    cmdBar.SetPreviewStatus False
    
    'Populate the button strip
    btsStrength.AddItem "low", 0
    btsStrength.AddItem "medium", 1
    btsStrength.AddItem "high", 2
    btsStrength.ListIndex = 0
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Render a new effect preview
Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then Me.fxComicBook GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Sub sltColor_Change()
    UpdatePreview
End Sub

Private Sub sltInk_Change()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    
    With cParams
        .AddParam "ink", sltInk.Value
        .AddParam "size", sltColor.Value
        .AddParam "smoothing", btsStrength.ListIndex
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
