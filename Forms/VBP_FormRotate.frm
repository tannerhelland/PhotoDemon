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
   Begin PhotoDemon.smartOptionButton optRotate 
      Height          =   360
      Index           =   0
      Left            =   6120
      TabIndex        =   8
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
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   9120
      TabIndex        =   0
      Top             =   5910
      Width           =   1365
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   10590
      TabIndex        =   1
      Top             =   5910
      Width           =   1365
   End
   Begin VB.TextBox txtAngle 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   11040
      MaxLength       =   6
      TabIndex        =   4
      Text            =   "0.0"
      Top             =   2280
      Width           =   735
   End
   Begin VB.HScrollBar hsAngle 
      Height          =   255
      LargeChange     =   10
      Left            =   6120
      Max             =   1800
      Min             =   -1800
      TabIndex        =   3
      Top             =   2340
      Width           =   4815
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.smartOptionButton optRotate 
      Height          =   360
      Index           =   1
      Left            =   6120
      TabIndex        =   9
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
   Begin VB.Label lblBackground 
      Height          =   855
      Left            =   0
      TabIndex        =   6
      Top             =   5760
      Width           =   12135
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
      TabIndex        =   5
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
      TabIndex        =   2
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
'Copyright ©2000-2013 by Tanner Helland
'Created: 12/November/12
'Last updated: 12/November/12
'Last update: initial build
'
'This tool allows the user to rotate an image at an arbitrary angle in 1/10 degree increments.  FreeImage is required
' for the tool to work, as this relies upon FreeImage to perform the rotation in a fast, efficient manner.  The
' corresponding menu entry for this tool is hidden unless FreeImage is found.
'
'At present, the tool assumes that you want to rotate the image around its center.
'
'***************************************************************************

Option Explicit

'Required for copying rotated image data from a FreeImage object
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As Any, ByVal wUsage As Long) As Long

'This temporary layer will be used for rendering the preview
Dim smallLayer As pdLayer

'Use this to prevent the text box and scroll bar from updating each other in an endless loop
Dim userChange As Boolean

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub cmdOK_Click()

    'Before rendering anything, check to make sure the text boxes have valid input
    If Not EntryValid(txtAngle, -180, 180, True, True) Then
        AutoSelectText txtAngle
        Exit Sub
    End If

    Me.Visible = False
    
    'Based on the user's selection, submit the proper processor request
    If optRotate(0) Then
        Process FreeRotate, 0, CSng(hsAngle / 10)
    Else
        Process FreeRotate, 1, CSng(hsAngle / 10)
    End If
    
    Unload Me
    
End Sub

Public Sub RotateArbitrary(ByVal canvasResize As Long, ByVal rotationAngle As Double, Optional ByVal isPreview As Boolean = False)

    'FreeImage uses positive values to indicate counter-clockwise rotation.  I disagree with this interpretation.  Thus, reverse
    ' the rotationAngle value so that POSITIVE values indicate CLOCKWISE rotation.
    rotationAngle = -rotationAngle

    'Double-check that FreeImage exists
    If g_ImageFormats.FreeImageEnabled Then
    
        'If a selection is active, remove it.  (This is not the most elegant solution, but it isn't fixable until masked
        ' selections are implemented.)
        If pdImages(CurrentImage).selectionActive Then
            pdImages(CurrentImage).selectionActive = False
            tInit tSelection, False
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

Private Sub Form_Activate()
    
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
        
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
    
    userChange = True
    
    'Render a preview
    updatePreview
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Keep the scroll bar and the text box values in sync
Private Sub hsAngle_Change()
    If userChange = True Then
        txtAngle.Text = Format(CSng(hsAngle.Value) / 10, "##0.0")
        txtAngle.Refresh
    End If
    updatePreview
End Sub

Private Sub hsAngle_Scroll()
    txtAngle.Text = Format(CSng(hsAngle.Value) / 10, "##0.0")
    txtAngle.Refresh
    updatePreview
End Sub

Private Sub OptRotate_Click(Index As Integer)
    updatePreview
End Sub

Private Sub txtAngle_GotFocus()
    AutoSelectText txtAngle
End Sub

Private Sub txtAngle_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtAngle, True, True
    If EntryValid(txtAngle, hsAngle.Min / 10, hsAngle.Max / 10, False, False) Then
        userChange = False
        hsAngle.Value = Val(txtAngle) * 10
        userChange = True
    End If
End Sub

'Redraw the on-screen preview of the rotated image
Private Sub updatePreview()

    If optRotate(0).Value Then
        RotateArbitrary 0, CDbl(hsAngle / 10), True
    Else
        RotateArbitrary 1, CDbl(hsAngle / 10), True
    End If

End Sub
