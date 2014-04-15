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
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   12105
      _ExtentX        =   21352
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
   End
   Begin PhotoDemon.smartOptionButton optRotate 
      Height          =   360
      Index           =   0
      Left            =   6120
      TabIndex        =   4
      Top             =   3330
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   635
      Caption         =   "adjust size to fit rotated image"
      Value           =   -1  'True
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
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
      DisableZoomPan  =   -1  'True
   End
   Begin PhotoDemon.smartOptionButton optRotate 
      Height          =   360
      Index           =   1
      Left            =   6120
      TabIndex        =   5
      Top             =   3720
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   635
      Caption         =   "keep image at its present size"
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
   Begin PhotoDemon.sliderTextCombo sltAngle 
      Height          =   495
      Left            =   6000
      TabIndex        =   6
      Top             =   2280
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Min             =   -360
      Max             =   360
      SigDigits       =   2
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
   Begin VB.Label lblRotatedCanvas 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "rotated image size:"
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
      Left            =   6000
      TabIndex        =   2
      Top             =   2880
      Width           =   2025
   End
   Begin VB.Label lblAmount 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "rotation angle:"
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
      Left            =   6000
      TabIndex        =   1
      Top             =   1920
      Width           =   1560
   End
End
Attribute VB_Name = "FormRotate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Image Rotation Interface
'Copyright ©2012-2014 by Tanner Helland
'Created: 12/November/12
'Last updated: 14/April/14
'Last update: rotate now works with layers!
'
'This tool allows the user to rotate an image at an arbitrary angle in 1/100 degree increments.  FreeImage is
' required for the tool to work, as this relies upon FreeImage to perform the rotation in a fast, efficient
' manner.  The corresponding menu entry for this tool is hidden unless FreeImage is found.  (I could add a
' GDI+ fallback as well, but it's waaaay down my list of priorities.)
'
'At present, the tool assumes that you want to rotate the image around its center.
'
'To rotate a layer instead of the entire image, use the Layer menu.  Rotation is also available in the
' Effect -> Distort menu, which can provide cool artistic effect when combined with selections.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'This temporary DIB will be used for rendering the preview
Dim smallDIB As pdDIB

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

Public Sub RotateArbitrary(ByVal canvasResize As Long, ByVal rotationAngle As Double, Optional ByVal isPreview As Boolean = False)
    
    'If the image contains an active selection, disable it before transforming the canvas
    If pdImages(g_CurrentImage).selectionActive Then
        pdImages(g_CurrentImage).selectionActive = False
        pdImages(g_CurrentImage).mainSelection.lockRelease
    End If

    'FreeImage uses positive values to indicate counter-clockwise rotation.  While mathematically correct, I find this
    ' unintuitive for casual users.  PD reverses the rotationAngle value so that POSITIVE values indicate CLOCKWISE rotation.
    rotationAngle = -rotationAngle

    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB

    'Double-check that FreeImage exists
    If g_ImageFormats.FreeImageEnabled Then
        
        'Rotation requires quite a few variables, including a number of handles for passing data back-and-forth with FreeImage.
        Dim fi_DIB As Long, returnDIB As Long
        Dim nWidth As Long, nHeight As Long
        
        'One of the FreeImage rotation variants requires an explicit center point; calculate one in advance.
        Dim cx As Double, cy As Double
        
        If isPreview Then
            cx = smallDIB.getDIBWidth / 2
            cy = smallDIB.getDIBHeight / 2
        Else
            cx = pdImages(g_CurrentImage).Width / 2
            cy = pdImages(g_CurrentImage).Height / 2
        End If
        
        
        'Normally, I like to use identical code for previews and actual effects.  However, rotating is completely different
        ' for previews (where we do a single rotation of the composited image) vs the full images (independently rotating
        ' each layer, with support functions to null-pad and crop layers as necessary).  As such, there is some code
        ' duplication here, but I believe it makes the code much more readable.
        
        If isPreview Then
            
            'Give FreeImage a handle to our temporary rotation image
            fi_DIB = FreeImage_CreateFromDC(smallDIB.getDIBDC)
            
            'There are two ways to rotate an image - enlarging the canvas to receive the fully rotated copy, or
            ' leaving the image the same size and truncating corners.  These require two different FreeImage functions.
            Select Case canvasResize
            
                'Resize the canvas to accept the new image
                Case 0
                    returnDIB = FreeImage_Rotate(fi_DIB, rotationAngle, RGB(255, 255, 255))
                    
                    nWidth = FreeImage_GetWidth(returnDIB)
                    nHeight = FreeImage_GetHeight(returnDIB)
                    
                'Leave the canvas the same size
                Case 1
                    returnDIB = FreeImage_RotateEx(fi_DIB, rotationAngle, 0, 0, cx, cy, True)
                    
                    nWidth = FreeImage_GetWidth(returnDIB)
                    nHeight = FreeImage_GetHeight(returnDIB)
            
            End Select
            
            'Create a blank DIB to receive the rotated image from FreeImage
            tmpDIB.createBlank nWidth, nHeight, 32
            
            'Ask FreeImage to premultiply the image's alpha data
            FreeImage_PreMultiplyWithAlpha returnDIB
            
            'Copy the bits from the FreeImage DIB to our DIB
            SetDIBitsToDevice tmpDIB.getDIBDC, 0, 0, nWidth, nHeight, 0, 0, 0, nHeight, ByVal FreeImage_GetBits(returnDIB), ByVal FreeImage_GetInfo(returnDIB), 0&
                
            'Finally, render the preview and erase the temporary DIB to conserve memory
            tmpDIB.renderToPictureBox fxPreview.getPreviewPic
            fxPreview.setFXImage tmpDIB
            
            Set tmpDIB = Nothing
            
        Else
        
            Message "Rotating image..."
            
            'FreeImage doesn't raise progress events, but we can use the number of layers as
            ' a stand-in progress parameter.
            SetProgBarMax pdImages(g_CurrentImage).getNumOfLayers
            
            'Iterate through each layer, rotating as we go
            Dim tmpLayerRef As pdLayer
            
            Dim i As Long
            For i = 0 To pdImages(g_CurrentImage).getNumOfLayers - 1
            
                SetProgBarVal i
            
                'Retrieve a pointer to the layer of interest
                Set tmpLayerRef = pdImages(g_CurrentImage).getLayerByIndex(i)
                
                'Remove premultiplied alpha, if any
                tmpLayerRef.layerDIB.fixPremultipliedAlpha False
                
                'Null-pad the layer
                tmpLayerRef.convertToNullPaddedLayer pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height
                
                'Give FreeImage a handle to the layer's pixel data
                fi_DIB = FreeImage_CreateFromDC(tmpLayerRef.layerDIB.getDIBDC)
            
                'There are two ways to rotate an image - enlarging the canvas to receive the fully rotated copy, or
                ' leaving the image the same size and truncating corners.  These require two different FreeImage functions.
                Select Case canvasResize
                
                    'Resize the canvas to accept the new image
                    Case 0
                        returnDIB = FreeImage_Rotate(fi_DIB, rotationAngle, RGB(255, 255, 255))
                        
                        nWidth = FreeImage_GetWidth(returnDIB)
                        nHeight = FreeImage_GetHeight(returnDIB)
                        
                    'Leave the canvas the same size
                    Case 1
                        returnDIB = FreeImage_RotateEx(fi_DIB, rotationAngle, 0, 0, cx, cy, True)
                        
                        nWidth = FreeImage_GetWidth(returnDIB)
                        nHeight = FreeImage_GetHeight(returnDIB)
                
                End Select
                
                'Resize the layer's DIB in preparation for the transfer
                tmpLayerRef.layerDIB.createBlank nWidth, nHeight, 32
                
                'Ask FreeImage to premultiply the image's alpha data
                FreeImage_PreMultiplyWithAlpha returnDIB
                
                'Copy the bits from the FreeImage DIB to our DIB
                SetDIBitsToDevice tmpLayerRef.layerDIB.getDIBDC, 0, 0, nWidth, nHeight, 0, 0, 0, nHeight, ByVal FreeImage_GetBits(returnDIB), ByVal FreeImage_GetInfo(returnDIB), 0&
                
                'Remove any null-padding
                tmpLayerRef.cropNullPaddedLayer
                
            'Continue with the next layer
            Next i
            
            'All layers have been rotated successfully!
            
            'Update the image's size
            pdImages(g_CurrentImage).updateSize False, nWidth, nHeight
            DisplaySize pdImages(g_CurrentImage)
            
            'Fit the new image on-screen and redraw it
            FitImageToViewport
            
            Message "Rotation complete."
            SetProgBarVal 0
        
        End If
        
        'With the transfer complete, release the FreeImage DIB and unload the library
        If returnDIB <> 0 Then FreeImage_UnloadEx returnDIB
        If fi_DIB <> 0 Then FreeImage_UnloadEx fi_DIB
        
    Else
        Message "Arbitrary rotation requires the FreeImage plugin, which could not be located.  Rotation canceled."
        pdMsgBox "The FreeImage plugin is required for image rotation.  Please go to Tools -> Options -> Updates and allow PhotoDemon to download core plugins.  Then restart the program.", vbApplicationModal + vbOKOnly + vbInformation, "FreeImage plugin missing"
    End If
        
End Sub

'OK button
Private Sub cmdBar_OKClick()
    If optRotate(0) Then
        Process "Arbitrary rotation", , buildParams(0, sltAngle)
    Else
        Process "Arbitrary rotation", , buildParams(1, sltAngle)
    End If
    
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub Form_Activate()
            
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
        
    'Render a preview
    cmdBar.markPreviewStatus True
    updatePreview
        
End Sub

Private Sub Form_Load()

    'Disable previewing until the dialog is fully initialized
    cmdBar.markPreviewStatus False
    
    'During the preview stage, we want to rotate a smaller version of the image.  This increases the speed of
    ' previewing immensely (especially for large images, like 10+ megapixel photos)
    Set smallDIB = New pdDIB
            
    'Determine a new image size that preserves the current aspect ratio
    Dim dWidth As Long, dHeight As Long
    convertAspectRatio pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height, fxPreview.getPreviewWidth, fxPreview.getPreviewHeight, dWidth, dHeight
    
    'Create a new, smaller image at those dimensions
    If (dWidth < pdImages(g_CurrentImage).Width) Or (dHeight < pdImages(g_CurrentImage).Height) Then
        
        Dim tmpDIB As pdDIB
        Set tmpDIB = New pdDIB
        pdImages(g_CurrentImage).getCompositedImage tmpDIB
        
        smallDIB.createFromExistingDIB tmpDIB, dWidth, dHeight, True
        
        Set tmpDIB = Nothing
        
    Else
        pdImages(g_CurrentImage).getCompositedImage smallDIB
    End If
        
    'Remove premultiplied alpha from the small DIB copy
    smallDIB.fixPremultipliedAlpha False
    
    'Give the preview object a copy of this image data so it can show it to the user if requested
    fxPreview.setOriginalImage smallDIB
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub OptRotate_Click(Index As Integer)
    updatePreview
End Sub

'Redraw the on-screen preview of the rotated image
Private Sub updatePreview()
    
    If cmdBar.previewsAllowed Then
        If optRotate(0).Value Then
            RotateArbitrary 0, sltAngle, True
        Else
            RotateArbitrary 1, sltAngle, True
        End If
    End If

End Sub

Private Sub sltAngle_Change()
    updatePreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

