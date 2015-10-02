VERSION 5.00
Begin VB.Form FormMosaic 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Mosaic"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12090
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
   ScaleHeight     =   434
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   806
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5760
      Width           =   12090
      _ExtentX        =   21325
      _ExtentY        =   1323
      BackColor       =   14802140
   End
   Begin PhotoDemon.smartCheckBox chkUnison 
      Height          =   330
      Left            =   6120
      TabIndex        =   4
      Top             =   4200
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   582
      Caption         =   "synchronize block size"
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.sliderTextCombo sltWidth 
      Height          =   720
      Left            =   6000
      TabIndex        =   2
      Top             =   2280
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "block width"
      Min             =   1
      Max             =   64
      Value           =   2
   End
   Begin PhotoDemon.sliderTextCombo sltHeight 
      Height          =   720
      Left            =   6000
      TabIndex        =   3
      Top             =   3240
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "block height"
      Min             =   1
      Max             =   64
      Value           =   2
   End
   Begin PhotoDemon.sliderTextCombo sltAngle 
      Height          =   720
      Left            =   6000
      TabIndex        =   1
      Top             =   1320
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "angle"
      Max             =   359.9
      SigDigits       =   1
   End
End
Attribute VB_Name = "FormMosaic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Pixelate/Mosaic filter interface
'Copyright 2000-2015 by Tanner Helland
'Created: 08/May/00
'Last updated: 02/October/15
'Last update: add support for variable angles
'
'Form for handling all the pixellation image transform code.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private Sub chkUnison_Click()
    If CBool(chkUnison) Then syncScrollBars True
    updatePreview
End Sub

'Apply a pixelate effect (sometimes called "mosaic") to an image
' Inputs: width and height of the desired pixelation tiles (in pixels), optional preview settings
Public Sub MosaicFilter(ByVal BlockSizeX As Long, ByVal BlockSizeY As Long, ByVal blockAngle As Double, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If Not toPreview Then Message "Applying mosaic..."
        
    'Grab a copy of the relevant pixel data from PD's main image data handler
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic, , , True
    
    'Make a note of the original image's size; we need this so we can restore the image to its original angle after
    ' the pixelation is complete.
    Dim origWidth As Long, origHeight As Long
    origWidth = workingDIB.getDIBWidth
    origHeight = workingDIB.getDIBHeight
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent already-mosaic'ed pixels from affecting the results of later pixels.)
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    
    'If an angle has been specified, we need to pre-rotate the image to match.
    If blockAngle <> 0 Then
        GDI_Plus.GDIPlus_GetRotatedClampedDIB workingDIB, srcDIB, blockAngle
        workingDIB.createFromExistingDIB srcDIB
    Else
        srcDIB.createFromExistingDIB workingDIB
    End If
    
    'Only now can we safely point arrays at their DIBs, as the DIBs will not be recreated again.  Note that we reverse
    ' the order of the source and destination DIBs if an angle is active; this spares us from having to perform an
    ' extra BitBlt after the operation is complete.
    
    If blockAngle = 0 Then
        prepSafeArray dstSA, workingDIB
        prepSafeArray srcSA, srcDIB
    Else
        prepSafeArray dstSA, srcDIB
        prepSafeArray srcSA, workingDIB
    End If
    
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = workingDIB.getDIBWidth - 1
    finalY = workingDIB.getDIBHeight - 1
    
    'If this is a preview, we need to adjust the mosaic values to match the size of the preview box
    If toPreview Then
        BlockSizeX = BlockSizeX * curDIBValues.previewModifier
        BlockSizeY = BlockSizeY * curDIBValues.previewModifier
        If BlockSizeX < 1 Then BlockSizeX = 1
        If BlockSizeY < 1 Then BlockSizeY = 1
    End If
    
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'Calculate how many mosaic tiles will fit on the current image's size
    Dim xLoop As Long, yLoop As Long
    xLoop = initX + Int(workingDIB.getDIBWidth \ BlockSizeX) + 1
    yLoop = initY + Int(workingDIB.getDIBHeight \ BlockSizeY) + 1
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not toPreview Then
        SetProgBarMax xLoop
        progBarCheck = findBestProgBarValue()
    End If
    
    'A number of other variables are required for the nested For..Next loops
    Dim dstXLoop As Long, dstYLoop As Long
    Dim initXLoop As Long, initYLoop As Long
    Dim i As Long, j As Long
    
    'We also need to count how many pixels must be averaged in each mosaic tile
    Dim NumOfPixels As Long
    
    'Finally, individual colors also need to be tracked
    Dim r As Long, g As Long, b As Long, a As Long
    
    'Loop through each pixel in the image, diffusing as we go
    For x = initX To xLoop
        QuickVal = x * qvDepth
    For y = initY To yLoop
        
        'This sub loop is to gather all of the data for the current mosaic tile
        initXLoop = x * BlockSizeX
        initYLoop = y * BlockSizeY
        dstXLoop = (x + 1) * BlockSizeX - 1
        dstYLoop = (y + 1) * BlockSizeY - 1
        
        For i = initXLoop To dstXLoop
            QuickVal = i * qvDepth
        For j = initYLoop To dstYLoop
        
            'If this particular pixel is off of the image, don't bother counting it
            If i > finalX Or j > finalY Then GoTo NextPixelatePixel1
            
            'Total up all the red, green, and blue values for the pixels within this
            'mosiac tile
            r = r + srcImageData(QuickVal + 2, j)
            g = g + srcImageData(QuickVal + 1, j)
            b = b + srcImageData(QuickVal, j)
            If qvDepth = 4 Then a = a + srcImageData(QuickVal + 3, j)
            
            'Count this as a valid pixel
            NumOfPixels = NumOfPixels + 1
            
NextPixelatePixel1:
        
        Next j
        Next i
        
        'If this tile is completely off of the image, don't worry about it and go to the next one
        If NumOfPixels = 0 Then GoTo NextPixelatePixel3
        
        'Take the average red, green, and blue values of all the pixels within this tile
        r = r \ NumOfPixels
        g = g \ NumOfPixels
        b = b \ NumOfPixels
        If qvDepth = 4 Then a = a \ NumOfPixels
        
        'Now run a loop through the same pixels you just analyzed, only this time you're gonna
        'draw the averaged color over the top of them
        For i = initXLoop To dstXLoop
            QuickVal = i * qvDepth
        For j = initYLoop To dstYLoop
        
            'Same thing as above - if it's off the image, ignore it
            If i > finalX Or j > finalY Then GoTo NextPixelatePixel2
            
            'Set the pixel
            If qvDepth = 4 Then dstImageData(QuickVal + 3, j) = a
            dstImageData(QuickVal + 2, j) = r
            dstImageData(QuickVal + 1, j) = g
            dstImageData(QuickVal, j) = b
            
NextPixelatePixel2:

        Next j
        Next i

NextPixelatePixel3:

        'Clear all the variables and go to the next pixel
        r = 0
        g = 0
        b = 0
        a = 0
        NumOfPixels = 0
        
    Next y
        If Not toPreview Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x
            End If
        End If
    Next x
    
    'With our work complete, point both ImageData() arrays away from their DIBs and deallocate them
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    Erase srcImageData
    
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    Erase dstImageData
    
    'If rotation was applied, restore the image to its original orientation.
    If blockAngle <> 0 Then
        workingDIB.createBlank origWidth, origHeight, srcDIB.getDIBColorDepth, 0, 0
        GDI_Plus.GDIPlus_RotateDIBPlgStyle srcDIB, workingDIB, -blockAngle, True
    End If
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic, True
    
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Mosaic", , buildParams(sltWidth.Value, sltHeight.Value, sltAngle.Value), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltAngle.Value = 0
    sltWidth.Value = 2
    sltHeight.Value = 2
End Sub

Private Sub Form_Activate()
    
    'Apply translations and visual themes
    MakeFormPretty Me
    
    'Request a preview
    cmdBar.markPreviewStatus True
    updatePreview
    
End Sub

Private Sub Form_Load()
    
    'Disable previews until the dialog is fully initialized
    cmdBar.markPreviewStatus False
    
    'Note the current image's width and height, which will be needed to adjust the preview effect
    If pdImages(g_CurrentImage).selectionActive Then
        sltWidth.Max = pdImages(g_CurrentImage).mainSelection.boundWidth
        sltHeight.Max = pdImages(g_CurrentImage).mainSelection.boundHeight
    Else
        sltWidth.Max = pdImages(g_CurrentImage).Width
        sltHeight.Max = pdImages(g_CurrentImage).Height
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Keep the two scroll bars in sync.  Some extra work has to be done to makes sure scrollbar max values aren't exceeded.
Private Sub syncScrollBars(ByVal srcHorizontal As Boolean)
    
    If sltWidth.Value = sltHeight.Value Then Exit Sub
    
    Dim tmpVal As Long
    
    If srcHorizontal Then
        tmpVal = sltWidth.Value
        If tmpVal < sltHeight.Max Then sltHeight.Value = sltWidth.Value Else sltHeight.Value = sltHeight.Max
    Else
        tmpVal = sltHeight.Value
        If tmpVal < sltWidth.Max Then sltWidth.Value = sltHeight.Value Else sltWidth.Value = sltWidth.Max
    End If
    
End Sub

'Redraw the effect preview
Private Sub updatePreview()
    If cmdBar.previewsAllowed Then MosaicFilter sltWidth.Value, sltHeight.Value, sltAngle.Value, True, fxPreview
End Sub

Private Sub sltAngle_Change()
    updatePreview
End Sub

Private Sub sltHeight_Change()
    If CBool(chkUnison) Then syncScrollBars False
    updatePreview
End Sub

Private Sub sltWidth_Change()
    If CBool(chkUnison) Then syncScrollBars True
    updatePreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

