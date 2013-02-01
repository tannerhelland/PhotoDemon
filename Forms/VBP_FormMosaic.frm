VERSION 5.00
Begin VB.Form FormMosaic 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Mosaic Options"
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
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin PhotoDemon.smartCheckBox chkUnison 
      Height          =   480
      Left            =   6120
      TabIndex        =   10
      Top             =   3600
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   847
      Caption         =   "keep both dimensions in sync"
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
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   9120
      TabIndex        =   0
      Top             =   5880
      Width           =   1365
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   10590
      TabIndex        =   1
      Top             =   5880
      Width           =   1365
   End
   Begin VB.HScrollBar hsHeight 
      Height          =   255
      Left            =   6120
      Max             =   64
      Min             =   1
      TabIndex        =   5
      Top             =   3000
      Value           =   2
      Width           =   4935
   End
   Begin VB.HScrollBar hsWidth 
      Height          =   255
      Left            =   6120
      Max             =   64
      Min             =   1
      TabIndex        =   3
      Top             =   2040
      Value           =   2
      Width           =   4935
   End
   Begin VB.TextBox txtHeight 
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
      Left            =   11160
      TabIndex        =   4
      Text            =   "2"
      Top             =   2940
      Width           =   615
   End
   Begin VB.TextBox txtWidth 
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
      Left            =   11160
      TabIndex        =   2
      Text            =   "2"
      Top             =   1980
      Width           =   615
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin VB.Label lblBackground 
      Height          =   855
      Left            =   0
      TabIndex        =   8
      Top             =   5730
      Width           =   12135
   End
   Begin VB.Label lblWidth 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "block width:"
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
      TabIndex        =   7
      Top             =   1680
      Width           =   1305
   End
   Begin VB.Label lblHeight 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "block height:"
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
      TabIndex        =   6
      Top             =   2640
      Width           =   1380
   End
End
Attribute VB_Name = "FormMosaic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Mosaic filter interface
'Copyright ©2000-2013 by Tanner Helland
'Created: 8/5/00
'Last updated: 10/September/12
'Last update: fixed many problems; rewrote code against new layer class, fixed code to work with mosaic sizes of 1,
'              changed BlockSize to be as large as the image if desired (previous limit was 255).
'
'Form for handling all the mosaic image transform code.
'
'***************************************************************************

Option Explicit

'When previewing, we need to modify the strength to be representative of the final filter.  This means dividing by the
' original image width in order to establish the right ratio.
Dim iWidth As Long, iHeight As Long

Dim userChange As Boolean

Private Sub chkUnison_Click()
    userChange = False
    If CBool(chkUnison) Then hsHeight.Value = hsWidth.Value
    userChange = True
    updatePreview
End Sub

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub cmdOK_Click()
    If EntryValid(txtWidth, hsWidth.Min, hsWidth.Max) Then
        If EntryValid(txtHeight, hsHeight.Min, hsHeight.Max) Then
            Me.Visible = False
            Process Mosaic, hsWidth.Value, hsHeight.Value
            Unload Me
        Else
            AutoSelectText txtHeight
        End If
    Else
        AutoSelectText txtWidth
    End If
End Sub

'Apply a mosaic effect (sometimes called "pixelize") to an image
' Inputs: width and height of the desired mosaic tiles (in pixels), optional preview settings
Public Sub MosaicFilter(ByVal BlockSizeX As Long, ByVal BlockSizeY As Long, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If toPreview = False Then Message "Repainting image in mosaic style..."
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent already-mosaic'ed pixels from affecting the results of later pixels.)
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    
    Dim srcLayer As pdLayer
    Set srcLayer = New pdLayer
    srcLayer.createFromExistingLayer workingLayer
    
    prepSafeArray srcSA, srcLayer
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
    
    'If this is a preview, we need to adjust the xDiffuse and yDiffuse values to match the size of the preview box
    If toPreview Then
        BlockSizeX = (BlockSizeX / iWidth) * curLayerValues.Width
        BlockSizeY = (BlockSizeY / iHeight) * curLayerValues.Height
        If BlockSizeX = 0 Then BlockSizeX = 1
        If BlockSizeY = 0 Then BlockSizeY = 1
    End If
    
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    'Calculate how many mosaic tiles will fit on the current image's size
    Dim xLoop As Long, yLoop As Long
    xLoop = initX + Int(workingLayer.getLayerWidth \ BlockSizeX) + 1
    yLoop = initY + Int(workingLayer.getLayerHeight \ BlockSizeY) + 1
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    SetProgBarMax xLoop
    progBarCheck = findBestProgBarValue()
    
    'A number of other variables are required for the nested For..Next loops
    Dim dstXLoop As Long, dstYLoop As Long
    Dim initXLoop As Long, initYLoop As Long
    Dim i As Long, j As Long
    
    'We also need to count how many pixels must be averaged in each mosaic tile
    Dim NumOfPixels As Long
    
    'Finally, individual colors also need to be tracked
    Dim r As Long, g As Long, b As Long
    
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
            If i > finalX Or j > finalY Then GoTo NextMosiacPixel1
            
            'Total up all the red, green, and blue values for the pixels within this
            'mosiac tile
            r = r + srcImageData(QuickVal + 2, j)
            g = g + srcImageData(QuickVal + 1, j)
            b = b + srcImageData(QuickVal, j)
            
            'Count this as a valid pixel
            NumOfPixels = NumOfPixels + 1
            
NextMosiacPixel1:
        
        Next j
        Next i
        
        'If this tile is completely off of the image, don't worry about it and go to the next one
        If NumOfPixels = 0 Then GoTo NextMosaicPixel3
        
        'Take the average red, green, and blue values of all the pixles within this tile
        r = r \ NumOfPixels
        g = g \ NumOfPixels
        b = b \ NumOfPixels
        
        'Now run a loop through the same pixels you just analyzed, only this time you're gonna
        'draw the averaged color over the top of them
        For i = initXLoop To dstXLoop
            QuickVal = i * qvDepth
        For j = initYLoop To dstYLoop
        
            'Same thing as above - if it's off the image, ignore it
            If i > finalX Or j > finalY Then GoTo NextMosiacPixel2
            
            'Set the pixel
            dstImageData(QuickVal + 2, j) = r
            dstImageData(QuickVal + 1, j) = g
            dstImageData(QuickVal, j) = b
            
NextMosiacPixel2:

        Next j
        Next i

NextMosaicPixel3:

        'Clear all the variables and go to the next pixel
        r = 0
        g = 0
        b = 0
        NumOfPixels = 0
        
    Next y
        If toPreview = False Then
            If (x And progBarCheck) = 0 Then SetProgBarVal x
        End If
    Next x
    
    'With our work complete, point both ImageData() arrays away from their DIBs and deallocate them
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    Erase srcImageData
    
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    Erase dstImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
    
End Sub

Private Sub Form_Activate()

    userChange = False

    'Note the current image's width and height, which will be needed to adjust the preview effect
    iWidth = pdImages(CurrentImage).Width
    iHeight = pdImages(CurrentImage).Height
        
    hsWidth.Max = pdImages(CurrentImage).Width
    hsHeight.Max = pdImages(CurrentImage).Height
    
    userChange = True
    
    MosaicFilter hsWidth.Value, hsHeight.Value, True, fxPreview
    
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub hsHeight_Change()
    userChange = False
    copyToTextBoxI txtHeight, hsHeight.Value
    If CBool(chkUnison) Then syncScrollBars False
    userChange = True
    updatePreview
End Sub

Private Sub hsWidth_Change()
    userChange = False
    copyToTextBoxI txtWidth, hsWidth.Value
    If CBool(chkUnison) Then syncScrollBars True
    userChange = True
    updatePreview
End Sub

Private Sub hsHeight_Scroll()
    userChange = False
    copyToTextBoxI txtHeight, hsHeight.Value
    If CBool(chkUnison) Then syncScrollBars False
    userChange = True
    updatePreview
End Sub

Private Sub hsWidth_Scroll()
    userChange = False
    copyToTextBoxI txtWidth, hsWidth.Value
    If CBool(chkUnison) Then syncScrollBars True
    userChange = True
    updatePreview
End Sub

Private Sub txtHeight_KeyUp(KeyCode As Integer, Shift As Integer)
    userChange = False
    textValidate txtHeight
    If EntryValid(txtHeight, hsHeight.Min, hsHeight.Max, False, False) Then hsHeight.Value = Val(txtHeight)
    userChange = True
    updatePreview
End Sub

Private Sub txtHeight_GotFocus()
    AutoSelectText txtHeight
End Sub

Private Sub txtWidth_KeyUp(KeyCode As Integer, Shift As Integer)
    userChange = False
    textValidate txtWidth
    If EntryValid(txtWidth, hsWidth.Min, hsWidth.Max, False, False) Then hsWidth.Value = Val(txtWidth)
    userChange = True
    updatePreview
End Sub

Private Sub txtWidth_GotFocus()
    AutoSelectText txtWidth
End Sub

'Keep the two scroll bars in sync.  Some extra work has to be done to makes sure scrollbar max values aren't exceeded.
Private Sub syncScrollBars(ByVal srcHorizontal As Boolean)
    
    If hsWidth.Value = hsHeight.Value Then Exit Sub
    
    Dim tmpVal As Long
    
    If srcHorizontal Then
        tmpVal = hsWidth.Value
        If tmpVal < hsHeight.Max Then hsHeight.Value = hsWidth.Value Else hsHeight.Value = hsHeight.Max
    Else
        tmpVal = hsHeight.Value
        If tmpVal < hsWidth.Max Then hsWidth.Value = hsHeight.Value Else hsWidth.Value = hsWidth.Max
    End If
    
End Sub

'Redraw the effect preview
Private Sub updatePreview()
    MosaicFilter hsWidth.Value, hsHeight.Value, True, fxPreview
End Sub
