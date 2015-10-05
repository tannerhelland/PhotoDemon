VERSION 5.00
Begin VB.Form FormBlackLight 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Black light"
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
   Begin PhotoDemon.sliderTextCombo sltIntensity 
      Height          =   720
      Left            =   6000
      TabIndex        =   2
      Top             =   2280
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   1270
      Caption         =   "intensity"
      Min             =   1
      SigDigits       =   2
      Value           =   2
   End
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
End
Attribute VB_Name = "FormBlackLight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Blacklight Form
'Copyright 2001-2015 by Tanner Helland
'Created: some time 2001
'Last updated: 01/October/13
'Last update: use a floating-point slider for more precise results
'
'I found this effect on accident, and it has gradually become one of my favorite effects.
' Visually stunning on many photographs.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Perform a blacklight filter
'Input: strength of the filter (min 1, no real max - but above 7 it becomes increasingly blown-out)
Public Sub fxBlackLight(Optional ByVal Weight As Double = 2#, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If Not toPreview Then Message "Illuminating image with imaginary blacklight..."
    
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
    Dim grayVal As Byte
    
    'Build a look-up table of grayscale values (faster than calculating it manually for each pixel)
    Dim grayLookUp(0 To 765) As Byte
    For x = 0 To 765
        grayLookUp(x) = x \ 3
    Next x
    
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Calculate the gray value using the look-up table
        grayVal = grayLookUp(r + g + b)
        
        'Perform the blacklight conversion
        r = Abs(r - grayVal) * Weight
        g = Abs(g - grayVal) * Weight
        b = Abs(b - grayVal) * Weight
        
        If r > 255 Then r = 255
        If g > 255 Then g = 255
        If b > 255 Then b = 255
        
        'Assign that gray value to each color channel
        ImageData(QuickVal, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal + 2, y) = b
        
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
    Process "Black light", , buildParams(sltIntensity), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltIntensity.Value = 2#
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

'Update the preview whenever the combination slider/text control has its value changed
Private Sub sltIntensity_Change()
    updatePreview
End Sub

Private Sub updatePreview()
    If cmdBar.previewsAllowed Then fxBlackLight sltIntensity, True, fxPreview
End Sub
