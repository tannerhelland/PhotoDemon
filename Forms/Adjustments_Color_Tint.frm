VERSION 5.00
Begin VB.Form FormTint 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Tint"
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
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
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
   Begin PhotoDemon.sliderTextCombo sltTint 
      CausesValidation=   0   'False
      Height          =   720
      Left            =   6000
      TabIndex        =   2
      Top             =   2400
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "tint"
      Min             =   -100
      Max             =   100
      SliderTrackStyle=   3
      GradientColorLeft=   15102446
      GradientColorRight=   8253041
      GradientColorMiddle=   16777215
   End
End
Attribute VB_Name = "FormTint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Tint Dialog
'Copyright 2014-2015 by Tanner Helland
'Created: 03/July/14
'Last updated: 03/July/14
'Last update: initial build
'
'Tint is a simple adjustment along the magenta/green axis of the image.  While of limited use in a
' separate dialog like this, PhotoDemon sticks to convention by providing it as a "quick-fix" non-destructive
' action, which also means that it needs to exist as a dedicated menu entry.
'
'The formula used here is more nuanced than the "quick fix" version.  This tool will attempt to preserve image
' luminance, by compensating for the loss (or gain) of green light via adjustments to the red and blue channels.
' This provides a better end result, but note that it *will* differ from a matching adjustment via the
' tint quick-fix slider.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Change the tint of an image
' INPUT: tint adjustment, [-100, 100], 0 = no change
Public Sub adjustTint(ByVal tintAdjustment As Long, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If Not toPreview Then Message "Re-tinting image..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    
    prepImageData tmpSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Color and grayscale variables
    Dim r As Long, g As Long, b As Long
    Dim h As Double, s As Double, v As Double, origV As Double
    
    'Build a look-up table of tint values.  (Tint only affects the green channel)
    Dim gLookUp() As Long
    ReDim gLookUp(0 To 255) As Long
    For x = 0 To 255
        g = x + (tintAdjustment \ 2)
        If g > 255 Then g = 255
        If g < 0 Then g = 0
        gLookUp(x) = g
    Next x
    
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
        
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Calculate luminance
        origV = getLuminance(r, g, b) / 255
        
        'Convert the re-tinted colors to HSL
        tRGBToHSL r, gLookUp(g), b, h, s, v
        
        'Convert back to RGB
        tHSLToRGB h, s, origV, r, g, b
        
        'Assign new values
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
        
    Next y
        If Not toPreview Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x
            End If
        End If
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic

End Sub

Private Sub cmdBar_OKClick()
    Process "Tint", , buildParams(sltTint), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub Form_Activate()
    
    'Apply translations and visual themes
    MakeFormPretty Me
    
    'Draw a preview of the effect
    updatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

Private Sub sltTint_Change()
    updatePreview
End Sub

Private Sub updatePreview()
    If cmdBar.previewsAllowed Then adjustTint sltTint, True, fxPreview
End Sub

