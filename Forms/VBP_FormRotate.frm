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
   StartUpPosition =   1  'CenterOwner
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   6
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
      TabIndex        =   3
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
      TabIndex        =   2
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.smartOptionButton optRotate 
      Height          =   360
      Index           =   1
      Left            =   6120
      TabIndex        =   4
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
      TabIndex        =   5
      Top             =   2280
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Min             =   -360
      Max             =   360
      SigDigits       =   1
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
      TabIndex        =   1
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
      TabIndex        =   0
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
'Copyright ©2012-2013 by Tanner Helland
'Created: 12/November/12
'Last updated: 24/August/13
'Last update: add command bar
'
'This tool allows the user to rotate an image at an arbitrary angle in 1/10 degree increments.  FreeImage is required
' for the tool to work, as this relies upon FreeImage to perform the rotation in a fast, efficient manner.  The
' corresponding menu entry for this tool is hidden unless FreeImage is found.
'
'At present, the tool assumes that you want to rotate the image around its center.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

'Required for copying rotated image data from a FreeImage object
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As Any, ByVal wUsage As Long) As Long

'This temporary layer will be used for rendering the preview
Dim smallLayer As pdLayer

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

Public Sub RotateArbitrary(ByVal canvasResize As Long, ByVal rotationAngle As Double, Optional ByVal isPreview As Boolean = False)

    'If the image contains an active selection, disable it before transforming the canvas
    If pdImages(CurrentImage).selectionActive Then
        pdImages(CurrentImage).selectionActive = False
        pdImages(CurrentImage).mainSelection.lockRelease
        metaToggle tSelection, False
    End If

    'FreeImage uses positive values to indicate counter-clockwise rotation.  While mathematically correct, I find this
    ' unintuitive for casual users.  PD reverses the rotationAngle value so that POSITIVE values indicate CLOCKWISE rotation.
    rotationAngle = -rotationAngle

    'Double-check that FreeImage exists
    If g_ImageFormats.FreeImageEnabled Then
    
        'If a selection is active, remove it.  (This is not the most elegant solution, but it isn't fixable until masked
        ' selections are implemented.)
        If pdImages(CurrentImage).selectionActive Then
            pdImages(CurrentImage).selectionActive = False
            pdImages(CurrentImage).mainSelection.lockRelease
            metaToggle tSelection, False
        End If
        
        If isPreview = False Then Message "Rotating image (this may take a few seconds)..."
        
        'Load the FreeImage dll into memory
        Dim hLib As Long
        hLib = LoadLibrary(g_PluginPath & "FreeImage.dll")
                
        'Convert our current layer to a FreeImage-type DIB
        Dim fi_DIB As Long
                
        If isPreview Then
            fi_DIB = FreeImage_CreateFromDC(smallLayer.getLayerDC)
        Else
            fi_DIB = FreeImage_CreateFromDC(pdImages(CurrentImage).mainLayer.getLayerDC)
        End If
        
        'Use that handle to request an image resize
        If fi_DIB <> 0 Then
        
            Dim returnDIB As Long
            
            Dim nWidth As Long, nHeight As Long
            
            'There are currently two ways to resize an image - enlarging the canvas to receive the new image, or
            ' leaving the image the same size.  These require two different FreeImage functions.
            Select Case canvasResize
            
                'Resize the canvas to accept the new image
                Case 0
                    returnDIB = FreeImage_Rotate(fi_DIB, rotationAngle, RGB(255, 255, 255))
                    
                    nWidth = FreeImage_GetWidth(returnDIB)
                    nHeight = FreeImage_GetHeight(returnDIB)
                    
                'Leave the canvas the same size
                Case 1
                
                    'This call requires an explicit center-point.  Calculate one accordingly.
                    Dim cx As Double, cy As Double
                    
                    If isPreview Then
                        cx = smallLayer.getLayerWidth / 2
                        cy = smallLayer.getLayerHeight / 2
                    Else
                        cx = pdImages(CurrentImage).Width / 2
                        cy = pdImages(CurrentImage).Height / 2
                    End If
                    
                    returnDIB = FreeImage_RotateEx(fi_DIB, rotationAngle, 0, 0, cx, cy, True)
                    
                    nWidth = FreeImage_GetWidth(returnDIB)
                    nHeight = FreeImage_GetHeight(returnDIB)
            
            End Select
            
            'If this is only a preview, use a temporary object to hold the rotated image
            If isPreview Then
            
                Dim tmpLayer As pdLayer
                Set tmpLayer = New pdLayer
                tmpLayer.createBlank nWidth, nHeight, pdImages(CurrentImage).mainLayer.getLayerColorDepth
            
                'Copy the bits from the FreeImage DIB to our DIB
                SetDIBitsToDevice tmpLayer.getLayerDC, 0, 0, nWidth, nHeight, 0, 0, 0, nHeight, ByVal FreeImage_GetBits(returnDIB), ByVal FreeImage_GetInfo(returnDIB), 0&
                
                'If tmpLayer.getLayerColorDepth = 32 Then tmpLayer.compositeBackgroundColor
                
                'Finally, render the preview and erase the temporary layer to conserve memory
                DrawPreviewImage fxPreview.getPreviewPic, True, tmpLayer
                fxPreview.setFXImage tmpLayer
                
                tmpLayer.eraseLayer
                Set tmpLayer = Nothing
            
            Else
                
                'Resize the image's main layer in preparation for the transfer
                pdImages(CurrentImage).mainLayer.createBlank nWidth, nHeight, pdImages(CurrentImage).mainLayer.getLayerColorDepth
                
                'Copy the bits from the FreeImage DIB to our DIB
                SetDIBitsToDevice pdImages(CurrentImage).mainLayer.getLayerDC, 0, 0, nWidth, nHeight, 0, 0, 0, nHeight, ByVal FreeImage_GetBits(returnDIB), ByVal FreeImage_GetInfo(returnDIB), 0&
                  
                'Update the size variables
                pdImages(CurrentImage).updateSize
                DisplaySize pdImages(CurrentImage).Width, pdImages(CurrentImage).Height
            
                'Fit the new image on-screen and redraw it
                FitImageToViewport
                FitWindowToImage
            
            End If
            
            'With the transfer complete, release the FreeImage DIB and unload the library
            If returnDIB <> 0 Then FreeImage_UnloadEx returnDIB
            If fi_DIB <> 0 Then FreeImage_UnloadEx fi_DIB
            FreeLibrary hLib
            
        Else
            FreeLibrary hLib
        End If
        
        If isPreview = False Then Message "Rotation complete."
        
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
    Set smallLayer = New pdLayer
            
    'Determine a new image size that preserves the current aspect ratio
    Dim dWidth As Long, dHeight As Long
    convertAspectRatio pdImages(CurrentImage).Width, pdImages(CurrentImage).Height, fxPreview.getPreviewWidth, fxPreview.getPreviewHeight, dWidth, dHeight
            
    'Create a new, smaller image at those dimensions
    If (dWidth < pdImages(CurrentImage).Width) Or (dHeight < pdImages(CurrentImage).Height) Then
        smallLayer.createFromExistingLayer pdImages(CurrentImage).mainLayer, dWidth, dHeight, True
    Else
        smallLayer.createFromExistingLayer pdImages(CurrentImage).mainLayer
    End If
        
    'Give the preview object a copy of this image data so it can show it to the user if requested
    fxPreview.setOriginalImage smallLayer
    
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
