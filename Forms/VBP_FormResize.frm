VERSION 5.00
Begin VB.Form FormResize 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Resize Image"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   7020
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
   ScaleHeight     =   348
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   468
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin PhotoDemon.smartCheckBox chkNames 
      Height          =   480
      Left            =   600
      TabIndex        =   12
      Top             =   3720
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   847
      Caption         =   "show technical names"
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
   Begin VB.CommandButton CmdResize 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   4050
      TabIndex        =   0
      Top             =   4590
      Width           =   1365
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   5520
      TabIndex        =   1
      Top             =   4590
      Width           =   1365
   End
   Begin VB.ComboBox cboResample 
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
      Left            =   600
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3240
      Width           =   5535
   End
   Begin PhotoDemon.smartCheckBox chkRatio 
      Height          =   480
      Left            =   4200
      TabIndex        =   4
      Top             =   975
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   847
      Caption         =   "lock aspect ratio"
      Value           =   1
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
   Begin PhotoDemon.textUpDown tudWidth 
      Height          =   405
      Left            =   1440
      TabIndex        =   2
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
      Left            =   1440
      TabIndex        =   3
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
      Left            =   240
      TabIndex        =   14
      Top             =   240
      Width           =   990
   End
   Begin VB.Label lblAspectRatio 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "selected aspect ratio is"
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
      Left            =   615
      TabIndex        =   13
      Top             =   1950
      Width           =   2370
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   272
      X2              =   272
      Y1              =   57
      Y2              =   106
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   272
      X2              =   232
      Y1              =   56
      Y2              =   56
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   232
      X2              =   272
      Y1              =   105
      Y2              =   105
   End
   Begin VB.Label lblBackground 
      Height          =   855
      Left            =   -2280
      TabIndex        =   11
      Top             =   4440
      Width           =   9975
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
      Left            =   2730
      TabIndex        =   10
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
      Left            =   2730
      TabIndex        =   9
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
      Left            =   600
      TabIndex        =   8
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
      Left            =   600
      TabIndex        =   7
      Top             =   735
      Width           =   675
   End
   Begin VB.Label lblResample 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "resize quality:"
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
      Left            =   240
      TabIndex        =   6
      Top             =   2760
      Width           =   1470
   End
End
Attribute VB_Name = "FormResize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Image Size Handler
'Copyright ©2001-2013 by Tanner Helland
'Created: 6/12/01
'Last updated: 13/June/13
'Last update: redesigned the form to incorporate features originally meant for "Smart Resize"
'
'Handles all image-size related functions.  Currently supports nearest-neighbor and halftone resampling
' (via the API; not 100% accurate but faster than doing it manually), bilinear resampling via pure VB, and
' a number of more advanced resampling techniques via FreeImage.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

Private Type resampleAlgorithm
    Name As String
    ProgramID As Long
End Type

Dim resampleTypes() As resampleAlgorithm
Dim numResamples As Long
Dim lastSelectedResample As Long

'Used for maintaining ratios when the check box is clicked
Private wRatio As Double, hRatio As Double

Dim allowedToUpdateWidth As Boolean, allowedToUpdateHeight As Boolean

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

Private Sub addResample(ByVal rName As String, ByVal rID As Long)
    resampleTypes(numResamples).Name = rName
    resampleTypes(numResamples).ProgramID = rID
    numResamples = numResamples + 1
    ReDim Preserve resampleTypes(0 To numResamples) As resampleAlgorithm
End Sub

'Display all available resample algorithms in the combo box (contingent on the "show technical names" check box as well)
Private Sub refillResampleBox(Optional ByVal isFirstTime As Boolean = False)

    ReDim resampleTypes(0) As resampleAlgorithm
    numResamples = 0

    'Use friendly names
    If Not CBool(chkNames) Then
        
        lblResample.Caption = g_Language.TranslateMessage("resize quality:")
        
        'FreeImage is required for best output.  Without it, only a small number of resample algorithms are implemented.
        If g_ImageFormats.FreeImageEnabled Then
            addResample g_Language.TranslateMessage("Best general purpose"), RESIZE_BICUBIC_CATMULL
            addResample g_Language.TranslateMessage("Best for photographs"), RESIZE_LANCZOS
            addResample g_Language.TranslateMessage("Best for text and illustration"), RESIZE_BICUBIC_MITCHELL
            addResample g_Language.TranslateMessage("Fastest"), RESIZE_NORMAL
        Else
            addResample g_Language.TranslateMessage("Best for photographs"), RESIZE_BILINEAR
            addResample g_Language.TranslateMessage("Best for text and illustration"), RESIZE_HALFTONE
            addResample g_Language.TranslateMessage("Fastest"), RESIZE_NORMAL
        End If
    
    'Use technical names
    Else
        
        lblResample.Caption = g_Language.TranslateMessage("resampling algorithm:")
        
        'Prepare a list of available resample algorithms
        addResample g_Language.TranslateMessage("Nearest Neighbor"), RESIZE_NORMAL
        addResample g_Language.TranslateMessage("Halftone"), RESIZE_HALFTONE
        addResample g_Language.TranslateMessage("Bilinear"), RESIZE_BILINEAR
        
        'If the FreeImage library is available, add additional resize options to the combo box
        If g_ImageFormats.FreeImageEnabled Then
            addResample g_Language.TranslateMessage("B-Spline"), RESIZE_BSPLINE
            addResample g_Language.TranslateMessage("Bicubic (Mitchell and Netravali)"), RESIZE_BICUBIC_MITCHELL
            addResample g_Language.TranslateMessage("Bicubic (Catmull-Rom)"), RESIZE_BICUBIC_CATMULL
            addResample g_Language.TranslateMessage("Sinc (Lanczos3)"), RESIZE_LANCZOS
        End If
                
    End If
    
    'Populate the combo box
    cboResample.Clear
    Dim I As Long
    For I = 0 To numResamples - 1
        cboResample.AddItem resampleTypes(I).Name, I
    Next I
    
    'If this is the first time we are filling the combo box, provide an intelligent default setting
    If isFirstTime Then
    
        'Friendly names
        If Not CBool(chkNames) Then
            cboResample.ListIndex = 0
            
        'Technical names
        Else
        
            'FreeImage enabled
            If g_ImageFormats.FreeImageEnabled Then
                cboResample.ListIndex = 5
                
            'FreeImage not enabled
            Else
                cboResample.ListIndex = 2
            End If
        
        End If
    
    'If this is not the first time we are creating a list of resample methods, re-select whatever method the
    ' user had previously selected (if available; otherwise, redirect them to the best general-purpose algorithm)
    Else
    
        Dim targetResampleMethod As Long
        targetResampleMethod = lastSelectedResample
        
        'Some technical options are not available under friendly names, so redirect them to something similar
        If CBool(chkNames) Then
            Select Case lastSelectedResample
                Case 1 To 3
                    targetResampleMethod = RESIZE_BICUBIC_CATMULL
            End Select
        End If
        
        'Find the matching resample method in the new combo box
        For I = 0 To cboResample.ListCount - 1
            If resampleTypes(I).ProgramID = targetResampleMethod Then
                cboResample.ListIndex = I
                Exit For
            End If
        Next I
    
    End If

End Sub

Private Sub cboResample_Click()
    lastSelectedResample = resampleTypes(cboResample.ListIndex).ProgramID
End Sub

'New to v5.6, PhotoDemon gives the user friendly resample names by default.  They can toggle these off at their liking.
Private Sub chkNames_Click()
    refillResampleBox
End Sub

'If the ratio button is checked, update the height box to reflect the image's current aspect ratio
Private Sub ChkRatio_Click()
    If CBool(chkRatio.Value) Then tudHeight = Int((tudWidth * hRatio) + 0.5)
End Sub

'Perform a resize operation
Private Sub CmdResize_Click()
    
    'Before resizing anything, check to make sure the textboxes have valid input
    If tudWidth.IsValid And tudHeight.IsValid Then
        
        Me.Visible = False
        Process "Resize", , buildParams(tudWidth, tudHeight, resampleTypes(cboResample.ListIndex).ProgramID)
        Unload Me
        
    End If
    
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

'Upon form activation, determine the ratio between the width and height of the image
Private Sub Form_Activate()
    
    'To prevent aspect ratio changes to one box resulting in recursion-type changes to the other, we only
    ' allow one box at a time to be updated.
    allowedToUpdateWidth = True
    allowedToUpdateHeight = True
    
    'Establish ratios
    wRatio = pdImages(CurrentImage).Width / pdImages(CurrentImage).Height
    hRatio = pdImages(CurrentImage).Height / pdImages(CurrentImage).Width

    'Populate the number of available resampling algorithms
    refillResampleBox True
    
    'Automatically set the width and height text boxes to match the image's current dimensions
    tudWidth.Value = pdImages(CurrentImage).Width
    tudHeight.Value = pdImages(CurrentImage).Height
    
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Resize an image using the FreeImage library.  Very fast.
Private Sub FreeImageResize(ByVal iWidth As Long, iHeight As Long, ByVal interpolationMethod As Long)
    
    'Double-check that FreeImage exists
    If g_ImageFormats.FreeImageEnabled Then
    
        'Load the FreeImage dll into memory
        Dim hLib As Long
        hLib = LoadLibrary(g_PluginPath & "FreeImage.dll")
        
        Message "Resampling image using the FreeImage plugin..."
        
        'Convert our current layer to a FreeImage-type DIB
        Dim fi_DIB As Long
        fi_DIB = FreeImage_CreateFromDC(pdImages(CurrentImage).mainLayer.getLayerDC)
        
        'Use that handle to request an image resize
        If fi_DIB <> 0 Then
            
            Dim returnDIB As Long
            returnDIB = FreeImage_RescaleByPixel(fi_DIB, iWidth, iHeight, True, interpolationMethod)
            
            'Resize our main layer in preparation for the transfer
            pdImages(CurrentImage).mainLayer.createBlank iWidth, iHeight, pdImages(CurrentImage).mainLayer.getLayerColorDepth
            
            'Copy the bits from the FreeImage DIB to our DIB
            SetDIBitsToDevice pdImages(CurrentImage).mainLayer.getLayerDC, 0, 0, iWidth, iHeight, 0, 0, 0, iHeight, ByVal FreeImage_GetBits(returnDIB), ByVal FreeImage_GetInfo(returnDIB), 0&
     
            'With the transfer complete, release the FreeImage DIB and unload the library
            If returnDIB <> 0 Then FreeImage_UnloadEx returnDIB
            FreeLibrary hLib
     
            'Update the size variables
            pdImages(CurrentImage).updateSize
            DisplaySize pdImages(CurrentImage).Width, pdImages(CurrentImage).Height
        
            'Fit the new image on-screen and redraw it
            FitImageToViewport
            FitWindowToImage
            
        Else
            FreeLibrary hLib
        End If
        
    End If
    
End Sub

'Resize an image using any one of several resampling algorithms.  (Some algorithms are provided by FreeImage.)
Public Sub ResizeImage(ByVal iWidth As Long, ByVal iHeight As Long, ByVal iMethod As Byte)

    'If the image contains an active selection, automatically resize it to match the new image.
    Dim selActive As Boolean
    Dim tsLeft As Double, tsTop As Double, tsWidth As Double, tsHeight As Double
    
    If pdImages(CurrentImage).selectionActive Then
        selActive = True
        
        'Remember all the current selection values
        tsLeft = pdImages(CurrentImage).mainSelection.boundLeft
        tsTop = pdImages(CurrentImage).mainSelection.boundTop
        tsWidth = pdImages(CurrentImage).mainSelection.boundWidth
        tsHeight = pdImages(CurrentImage).mainSelection.boundHeight
        
        'Deactivate the current selection
        pdImages(CurrentImage).selectionActive = False
        tInit tSelection, False
        
        'Note the ratio between the original width/height values and the new ones
        wRatio = iWidth / pdImages(CurrentImage).Width
        hRatio = iHeight / pdImages(CurrentImage).Height
        
    Else
        selActive = False
    End If

    'Because most resize methods require a temporary layer, create one here
    Dim tmpLayer As pdLayer
    Set tmpLayer = New pdLayer

    Select Case iMethod

        'Nearest neighbor...
        Case RESIZE_NORMAL
        
            Message "Resizing image..."
            
            'Copy the current layer into this temporary layer at the new size
            tmpLayer.createFromExistingLayer pdImages(CurrentImage).mainLayer, iWidth, iHeight, False
            
            'Now copy the resized image back into the main layer
            pdImages(CurrentImage).mainLayer.createFromExistingLayer tmpLayer
            
            'Update the size to match
            pdImages(CurrentImage).updateSize
            DisplaySize pdImages(CurrentImage).Width, pdImages(CurrentImage).Height
            
            'Fit the new image on-screen and redraw it
            FitOnScreen
            
        'Halftone resampling... I'm not sure what to actually call it, but since it's based off the
        ' StretchBlt mode Microsoft calls "halftone," I'm sticking with that
        Case RESIZE_HALFTONE
            
            Message "Resizing image..."
            
            'Copy the current layer into this temporary layer at the new size
            tmpLayer.createFromExistingLayer pdImages(CurrentImage).mainLayer, iWidth, iHeight, True
                
            'Now copy the resized image back into the main layer
            pdImages(CurrentImage).mainLayer.createFromExistingLayer tmpLayer
            
            'Update the size to match
            pdImages(CurrentImage).updateSize
            DisplaySize pdImages(CurrentImage).Width, pdImages(CurrentImage).Height
            
            'Fit the new image on-screen and redraw it
            FitOnScreen
            
        'True bilinear sampling
        Case RESIZE_BILINEAR
        
            'If FreeImage is enabled, use their bilinear filter.  Similar results, much faster.
            If g_ImageFormats.FreeImageEnabled Then
            
                FreeImageResize iWidth, iHeight, FILTER_BILINEAR
            
            'If FreeImage is not enabled, we have to do the resample ourselves.
            Else
            
                Message "Resampling image..."
        
                'Create a local array and point it at the pixel data of the current image
                Dim srcImageData() As Byte
                Dim srcSA As SAFEARRAY2D
                prepImageData srcSA
                CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
                'Resize the temporary layer to the target size, and point a second local array at it
                tmpLayer.createBlank iWidth, iHeight, pdImages(CurrentImage).mainLayer.getLayerColorDepth
                
                Dim dstImageData() As Byte
                Dim dstSA As SAFEARRAY2D
                
                prepSafeArray dstSA, tmpLayer
                CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
                
                'These values will help us access locations in the array more quickly.
                ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
                Dim qvDepth As Long
                qvDepth = tmpLayer.getLayerColorDepth \ 8
                
                'Create a filter support class, which will aid with edge handling and interpolation
                Dim fSupport As pdFilterSupport
                Set fSupport = New pdFilterSupport
                fSupport.setDistortParameters qvDepth, EDGE_CLAMP, True, curLayerValues.MaxX, curLayerValues.MaxY
    
                'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
                ' based on the size of the area to be processed.
                Dim progBarCheck As Long
                SetProgBarMax iWidth
                progBarCheck = findBestProgBarValue()
            
                'Resampling requires many variables
                
                'Scaled ratios between the old x and y values and the new ones
                Dim xScale As Double, yScale As Double
                xScale = (pdImages(CurrentImage).Width - 1) / iWidth
                yScale = (pdImages(CurrentImage).Height - 1) / iHeight
                            
                'Coordinate variables for source and destination
                Dim x As Long, y As Long
                Dim srcX As Double, srcY As Double
                            
                For x = 0 To iWidth - 1
                    
                    'Generate the x calculation variables
                    srcX = x * xScale
                    
                    'Draw each pixel in the new image
                    For y = 0 To iHeight - 1
                        
                        'Generate the y calculation variables
                        srcY = y * yScale
                        
                        'The lovely .setPixels routine will handle edge detection and interpolation for us as necessary
                        fSupport.setPixels x, y, srcX, srcY, srcImageData, dstImageData
                                            
                    Next y
                
                    If (x And progBarCheck) = 0 Then SetProgBarVal x
                    
                Next x
            
                'Now copy the resized image back into the main layer
                pdImages(CurrentImage).mainLayer.createFromExistingLayer tmpLayer
                
                'With our work complete, point both ImageData() arrays away from their DIBs and deallocate them
                CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
                Erase srcImageData
                
                CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
                Erase dstImageData
                
                'Update the size variables
                pdImages(CurrentImage).updateSize
                DisplaySize pdImages(CurrentImage).Width, pdImages(CurrentImage).Height
            
                SetProgBarVal 0
            
                'Fit the new image on-screen and redraw it
                FitOnScreen
                
            End If
        
        Case RESIZE_BSPLINE
            FreeImageResize iWidth, iHeight, FILTER_BSPLINE
            
        Case RESIZE_BICUBIC_MITCHELL
            FreeImageResize iWidth, iHeight, FILTER_BICUBIC
            
        Case RESIZE_BICUBIC_CATMULL
            FreeImageResize iWidth, iHeight, FILTER_CATMULLROM
        
        Case RESIZE_LANCZOS
            FreeImageResize iWidth, iHeight, FILTER_LANCZOS3
            
    End Select
    
    'Release our temporary layer
    Set tmpLayer = Nothing
    
    'If the image had a selection, recreate it - but make it match the new image size
    If selActive Then
        
        'Populate the selection text boxes (which are now invisible)
        FormMain.tudSelLeft(0) = Int(tsLeft * wRatio)
        FormMain.tudSelTop(0) = Int(tsTop * hRatio)
        FormMain.tudSelWidth(0) = Int(tsWidth * wRatio)
        FormMain.tudSelHeight(0) = Int(tsHeight * hRatio)
        
        'Reactivate the current selection with the new values
        tInit tSelection, True
        pdImages(CurrentImage).mainSelection.updateViaTextBox 0
        pdImages(CurrentImage).selectionActive = True
        
        'Redraw the image
        RenderViewport pdImages(CurrentImage).containingForm
        
    End If
    
    Message "Finished."
    
End Sub

'PhotoDemon now displays an approximate aspect ratio for the selected values.  This can be helpful when
' trying to select new width/height values for a specific application with a set aspect ratio (e.g. 16:9 screens).
Private Sub updateAspectRatio()

    Dim wholeNumber As Double, Numerator As Double, Denominator As Double
    
    If tudWidth.IsValid And tudHeight.IsValid Then
        ConvertToFraction tudWidth / tudHeight, wholeNumber, Numerator, Denominator, 4, 99.9
        'Numerator = (wholeNumber * Denominator) + Numerator
        
        'Aspect ratios are typically given in terms of base 10 if possible, so change something like 8:5 to 16:10
        If CLng(Denominator) = 5 Then
            Numerator = Numerator * 2
            Denominator = Denominator * 2
        End If
        
        lblAspectRatio.Caption = g_Language.TranslateMessage("selected aspect ratio is %1:%2", Numerator, Denominator)
    End If

End Sub

'If "Lock Image Aspect Ratio" is selected, these two routines keep all values in sync
Private Sub tudHeight_Change()
    If CBool(chkRatio) And allowedToUpdateWidth Then
        allowedToUpdateHeight = False
        tudWidth = Int((tudHeight * wRatio) + 0.5)
        allowedToUpdateHeight = True
    End If
    
    updateAspectRatio
    
End Sub

Private Sub tudWidth_Change()
    If CBool(chkRatio) And allowedToUpdateHeight Then
        allowedToUpdateWidth = False
        tudHeight = Int((tudWidth * hRatio) + 0.5)
        allowedToUpdateWidth = True
    End If
    
    updateAspectRatio
    
End Sub
