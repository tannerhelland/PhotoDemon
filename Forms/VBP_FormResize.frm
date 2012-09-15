VERSION 5.00
Begin VB.Form FormResize 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   " Resize Image"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   4005
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
   ScaleHeight     =   277
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   267
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.VScrollBar VSHeight 
      Height          =   420
      Left            =   2430
      Max             =   32766
      Min             =   1
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1140
      Value           =   15000
      Width           =   270
   End
   Begin VB.VScrollBar VSWidth 
      Height          =   420
      Left            =   2430
      Max             =   32766
      Min             =   1
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   540
      Value           =   15000
      Width           =   270
   End
   Begin VB.TextBox TxtHeight 
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
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Text            =   "N/A"
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox TxtWidth 
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
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Text            =   "N/A"
      Top             =   600
      Width           =   855
   End
   Begin VB.ComboBox cboResample 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2760
      Width           =   3495
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   3600
      Width           =   1125
   End
   Begin VB.CheckBox ChkRatio 
      Appearance      =   0  'Flat
      Caption         =   "Preserve size ratio"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   1920
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CommandButton CmdResize 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   3600
      Width           =   1125
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "pixels"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   2850
      TabIndex        =   10
      Top             =   1260
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "pixels"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   2850
      TabIndex        =   9
      Top             =   645
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Height:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Left            =   840
      TabIndex        =   8
      Top             =   1230
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Width:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Left            =   840
      TabIndex        =   7
      Top             =   645
      Width           =   555
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Resample method:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2520
      Width           =   2895
   End
End
Attribute VB_Name = "FormResize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Image Size Handler
'Copyright ©2000-2012 by Tanner Helland
'Created: 6/12/01
'Last updated: 10/September/12
'Last update: rewrote all resize functions against the new layer class
'
'Handles all image-size related functions.  Currently supports standard resizing and halftone resampling
' (via the API; not 100% accurate but faster than doing it in VB code) and bilinear resampling via pure VB)
'
'***************************************************************************

Option Explicit

'Resampling declarations
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal DX As Long, ByVal DY As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As Any, ByVal wUsage As Long) As Long

'Used to prevent the scroll bars from getting stuck in update loops
Dim updateWidthBar As Boolean, updateHeightBar As Boolean

'Used for maintaining ratios when the check box is clicked
Dim WRatio As Double, HRatio As Double

'If the ratio button is checked, then update the height box to match
Private Sub ChkRatio_Click()
    If ChkRatio.Value = vbChecked Then UpdateHeightBox
End Sub

'Resize an image using bicubic, bilinear, or nearest neighbor resampling
Public Sub ResizeImage(ByVal iWidth As Long, ByVal iHeight As Long, ByVal iMethod As Byte)

    'Because most resize methods require a temporary layer, create one here
    Dim tmpLayer As pdLayer
    Set tmpLayer = New pdLayer

    'Nearest neighbor...
    If iMethod = RESIZE_NORMAL Then
    
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
        
    'Halftone resampling... I'm not sure what to actually call it, but since it seems to be based off the
    ' StretchBlt mode Microsoft calls "halftone," I'm running with it
    ElseIf iMethod = RESIZE_HALFTONE Then
        
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
    ElseIf iMethod = RESIZE_BILINEAR Then
    
        'If FreeImage is enabled, use their bilinear filter.  Similar results, much faster.
        If FreeImageEnabled Then
        
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
            Dim QuickVal As Long, qvDepth As Long
            qvDepth = tmpLayer.getLayerColorDepth \ 8
            
            'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
            ' based on the size of the area to be processed.
            Dim progBarCheck As Long
            SetProgBarMax iWidth
            progBarCheck = findBestProgBarValue()
        
            'Resampling requires a ton of variables
        
            'Scaled ratios between the old x and y values and the new ones
            Dim xScale As Single, yScale As Single
            
            'Coordinate variables for source and destination
            Dim x As Long, y As Long
            Dim dstX As Single, dstY As Single
            
            'Interpolated X and Y values
            Dim IntrplX As Integer, IntrplY As Integer
            
            'Calculation variables
            Dim CalcX As Single, CalcY As Single, invCalcX As Single, invCalcY As Single
            
            'Red and green values we'll use to interpolate the new pixel
            Dim r As Long, r1 As Single, r2 As Single, r3 As Single, r4 As Single
            Dim g As Long, g1 As Single, g2 As Single, g3 As Single, g4 As Single
            Dim b As Long, b1 As Single, b2 As Single, b3 As Single, b4 As Single
            
            'Interpolated red, green, and blue
            Dim ir1 As Long, ig1 As Long, ib1 As Long
            Dim ir2 As Long, ig2 As Long, ib2 As Long
            
            'Shortcut variables for x positions
            Dim QuickX As Long, QuickXInt As Long, QuickXIntRight As Long
            
            'Get the ratio between the old image and the new one
            xScale = (pdImages(CurrentImage).Width - 1) / iWidth
            yScale = (pdImages(CurrentImage).Height - 1) / iHeight
            
            For x = 0 To iWidth - 1
                
                'Generate the x calculation variables
                dstX = x * xScale
                IntrplX = Int(dstX)
                CalcX = dstX - IntrplX
                invCalcX = 1 - CalcX
                
                QuickX = x * qvDepth
                QuickXInt = IntrplX * qvDepth
                QuickXIntRight = (IntrplX + 1) * qvDepth
    
                'Draw each pixel in the new image
                For y = 0 To iHeight - 1
                    
                    'Generate the y calculation variables
                    dstY = y * yScale
                    IntrplY = Int(dstY)
                    CalcY = dstY - IntrplY
                    invCalcY = 1 - CalcY
                    
                    'Get the 4 pixels around the interpolated one
                    r1 = srcImageData(QuickXInt + 2, IntrplY)
                    g1 = srcImageData(QuickXInt + 1, IntrplY)
                    b1 = srcImageData(QuickXInt, IntrplY)
                    
                    r2 = srcImageData(QuickXIntRight + 2, IntrplY)
                    g2 = srcImageData(QuickXIntRight + 1, IntrplY)
                    b2 = srcImageData(QuickXIntRight, IntrplY)
                    
                    r3 = srcImageData(QuickXInt + 2, IntrplY + 1)
                    g3 = srcImageData(QuickXInt + 1, IntrplY + 1)
                    b3 = srcImageData(QuickXInt, IntrplY + 1)
        
                    r4 = srcImageData(QuickXIntRight + 2, IntrplY + 1)
                    g4 = srcImageData(QuickXIntRight + 1, IntrplY + 1)
                    b4 = srcImageData(QuickXIntRight, IntrplY + 1)
        
                    'Interpolate the R,G,B values in the Y direction
                    ir1 = r1 * invCalcY + r3 * CalcY
                    ig1 = g1 * invCalcY + g3 * CalcY
                    ib1 = b1 * invCalcY + b3 * CalcY
                    ir2 = r2 * invCalcY + r4 * CalcY
                    ig2 = g2 * invCalcY + g4 * CalcY
                    ib2 = b2 * invCalcY + b4 * CalcY
                    
                    'Intepolate the R,G,B values in the X direction
                    r = ir1 * invCalcX + ir2 * CalcX
                    g = ig1 * invCalcX + ig2 * CalcX
                    b = ib1 * invCalcX + ib2 * CalcX
                    
                    'Make sure that the values are in the acceptable range
                    If r > 255 Then
                        r = 255
                    ElseIf r < 0 Then
                        r = 0
                    End If
                    If g > 255 Then
                        g = 255
                    ElseIf g < 0 Then
                        g = 0
                    End If
                    If b > 255 Then
                        b = 255
                    ElseIf b < 0 Then
                        b = 0
                    End If
                    
                    'Set this pixel onto the destination image
                    dstImageData(QuickX + 2, y) = r
                    dstImageData(QuickX + 1, y) = g
                    dstImageData(QuickX, y) = b
                
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
    
    ElseIf iMethod = RESIZE_BSPLINE Then
        FreeImageResize iWidth, iHeight, FILTER_BSPLINE
        
    ElseIf iMethod = RESIZE_BICUBIC_MITCHELL Then
        FreeImageResize iWidth, iHeight, FILTER_BICUBIC
        
    ElseIf iMethod = RESIZE_BICUBIC_CATMULL Then
        FreeImageResize iWidth, iHeight, FILTER_CATMULLROM
    
    ElseIf iMethod = RESIZE_LANCZOS Then
        FreeImageResize iWidth, iHeight, FILTER_LANCZOS3
        
    End If
    
    'Release our temporary layer
    Set tmpLayer = Nothing
    
    Message "Finished."
    
End Sub

'Perform a resize operation
Private Sub CmdResize_Click()
    
    'Before resizing anything, check to make sure the textboxes have valid input
    If Not EntryValid(TxtWidth, 1, 32767, True, True) Then
        AutoSelectText TxtWidth
        Exit Sub
    End If
    If Not EntryValid(TxtHeight, 1, 32767, True, True) Then
        AutoSelectText TxtHeight
        Exit Sub
    End If
    
    Me.Visible = False
    
    'Resample based on the combo box entry...
    Select Case cboResample.ListIndex
        Case 0
            Process ImageSize, val(TxtWidth), val(TxtHeight), RESIZE_NORMAL
        Case 1
            Process ImageSize, val(TxtWidth), val(TxtHeight), RESIZE_HALFTONE
        Case 2
            Process ImageSize, val(TxtWidth), val(TxtHeight), RESIZE_BILINEAR
        Case 3
            Process ImageSize, val(TxtWidth), val(TxtHeight), RESIZE_BSPLINE
        Case 4
            Process ImageSize, val(TxtWidth), val(TxtHeight), RESIZE_BICUBIC_MITCHELL
        Case 5
            Process ImageSize, val(TxtWidth), val(TxtHeight), RESIZE_BICUBIC_CATMULL
        Case 6
            Process ImageSize, val(TxtWidth), val(TxtHeight), RESIZE_LANCZOS
    End Select
    
    Unload Me
    
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

'Upon form load, determine the ratio between the width and height of the image
Private Sub Form_Load()
    
    'Add one to the displayed width and height, since we store them -1 for loops
    TxtWidth.Text = pdImages(CurrentImage).Width
    TxtHeight.Text = pdImages(CurrentImage).Height
    
    'Make the scroll bars match the text boxes
    updateWidthBar = False
    updateHeightBar = False
    VSWidth.Value = Abs(32767 - CInt(TxtWidth))
    VSHeight.Value = Abs(32767 - CInt(TxtHeight))
    updateWidthBar = True
    updateHeightBar = True
    
    'Establish ratios
    WRatio = pdImages(CurrentImage).Width / pdImages(CurrentImage).Height
    HRatio = pdImages(CurrentImage).Height / pdImages(CurrentImage).Width

    'Load up the combo box
    cboResample.AddItem "Nearest Neighbor", 0
    cboResample.AddItem "Halftone", 1
    cboResample.AddItem "Bilinear", 2
    cboResample.ListIndex = 2
    
    'If the FreeImage library is available, add additional resize options to the combo box
    If FreeImageEnabled = True Then
        cboResample.AddItem "B-Spline", 3
        cboResample.AddItem "Bicubic (Mitchell and Netravali)", 4
        cboResample.AddItem "Bicubic (Catmull-Rom)", 5
        cboResample.AddItem "Sinc (Lanczos3)", 6
        cboResample.ListIndex = 5
    End If
    
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
    
End Sub

'*************************************************************************************
'If "Preserve Size Ratio" is selected, this set of routines handles the preservation

Private Sub txtHeight_GotFocus()
    AutoSelectText TxtHeight
End Sub

Private Sub txtHeight_KeyUp(KeyCode As Integer, Shift As Integer)
    ChangeToHeight
End Sub

Private Sub TxtHeight_LostFocus()
    ChangeToHeight
End Sub

Private Sub txtWidth_GotFocus()
    AutoSelectText TxtWidth
End Sub

Private Sub TxtWidth_KeyUp(KeyCode As Integer, Shift As Integer)
    ChangeToWidth
End Sub

Private Sub TxtWidth_LostFocus()
    ChangeToWidth
End Sub

Private Sub UpdateHeightBox()
    updateHeightBar = False
    TxtHeight = Int((CDbl(val(TxtWidth)) * HRatio) + 0.5)
    VSHeight.Value = Abs(32767 - val(TxtHeight))
    updateHeightBar = True
End Sub

Private Sub UpdateWidthBox()
    updateWidthBar = False
    TxtWidth = Int((CDbl(val(TxtHeight)) * WRatio) + 0.5)
    VSWidth.Value = Abs(32767 - val(TxtWidth))
    updateWidthBar = True
End Sub
'*************************************************************************************

'Resize an image using the FreeImage library.  Very fast.
Private Sub FreeImageResize(ByVal iWidth As Long, iHeight As Long, ByVal interpolationMethod As Long)
    
    'Double-check that FreeImage exists
    If FreeImageEnabled Then
    
        'Load the FreeImage dll into memory
        Dim hLib As Long
        hLib = LoadLibrary(PluginPath & "FreeImage.dll")
        
        Message "Resampling image using the FreeImage library..."
        
        'Convert our current layer to a FreeImage-type DIB
        Dim fi_DIB As Long
        fi_DIB = FreeImage_CreateFromDC(pdImages(CurrentImage).mainLayer.getLayerDC)
        
        'Use that handle to save the image to GIF format, with required 8bpp (256 color) conversion
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
            FitOnScreen
            
        End If
        
    End If
    
End Sub

'Because the scrollbars work backward (up means up and down means down) we have to reverse its input
' relative to the associated text box
Private Sub VSHeight_Change()
    If updateHeightBar = True Then
        TxtHeight = Abs(32767 - CStr(VSHeight.Value))
        ChangeToHeight
    End If
End Sub

Private Sub VSWidth_Change()
    If updateWidthBar = True Then
        TxtWidth = Abs(32767 - CStr(VSWidth.Value))
        ChangeToWidth
    End If
End Sub

Private Sub ChangeToWidth()
    If EntryValid(TxtWidth, 1, 32767, False, True) Then
        updateWidthBar = False
        VSWidth.Value = Abs(32767 - CInt(TxtWidth))
        updateWidthBar = True
        If ChkRatio.Value = vbChecked Then
            UpdateHeightBox
        End If
    End If
End Sub

Private Sub ChangeToHeight()
    If EntryValid(TxtHeight, 1, 32767, False, True) Then
        updateHeightBar = False
        VSHeight.Value = Abs(32767 - CInt(TxtHeight))
        updateHeightBar = True
        If ChkRatio.Value = vbChecked Then
            UpdateWidthBox
        End If
    End If
End Sub
