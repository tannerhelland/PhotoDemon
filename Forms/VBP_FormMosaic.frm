VERSION 5.00
Begin VB.Form FormMosaic 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Mosaic Options"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6270
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
   ScaleHeight     =   420
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   418
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.PictureBox picPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2730
      Left            =   120
      ScaleHeight     =   180
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   191
      TabIndex        =   9
      Top             =   120
      Width           =   2895
   End
   Begin VB.PictureBox picEffect 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2730
      Left            =   3240
      ScaleHeight     =   180
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   191
      TabIndex        =   8
      Top             =   120
      Width           =   2895
   End
   Begin VB.HScrollBar hsHeight 
      Height          =   255
      Left            =   360
      Max             =   64
      Min             =   1
      TabIndex        =   5
      Top             =   4800
      Value           =   2
      Width           =   4935
   End
   Begin VB.HScrollBar hsWidth 
      Height          =   255
      Left            =   360
      Max             =   64
      Min             =   1
      TabIndex        =   3
      Top             =   3840
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
      Left            =   5400
      TabIndex        =   4
      Text            =   "2"
      Top             =   4740
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
      Left            =   5400
      TabIndex        =   2
      Text            =   "2"
      Top             =   3780
      Width           =   615
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   555
      Left            =   4800
      TabIndex        =   1
      Top             =   5640
      Width           =   1245
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   555
      Left            =   3480
      TabIndex        =   0
      Top             =   5640
      Width           =   1245
   End
   Begin VB.Label lblAfter 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "after"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   3360
      TabIndex        =   11
      Top             =   2880
      Width           =   360
   End
   Begin VB.Label lblBefore 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "before"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   2880
      Width           =   480
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
      Left            =   240
      TabIndex        =   7
      Top             =   3480
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
      Left            =   240
      TabIndex        =   6
      Top             =   4440
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

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub cmdOK_Click()
    If EntryValid(TxtWidth, hsWidth.Min, hsWidth.Max) Then
        If EntryValid(TxtHeight, hsHeight.Min, hsHeight.Max) Then
            Me.Visible = False
            Process Mosaic, hsWidth.Value, hsHeight.Value
            Unload Me
        Else
            AutoSelectText TxtHeight
        End If
    Else
        AutoSelectText TxtWidth
    End If
End Sub

'Apply a mosaic effect (sometimes called "pixelize") to an image
' Inputs: width and height of the desired mosaic tiles (in pixels), optional preview settings
Public Sub MosaicFilter(ByVal BlockSizeX As Long, ByVal BlockSizeY As Long, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As PictureBox)
    
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
    Dim R As Long, g As Long, b As Long
    
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
            R = R + srcImageData(QuickVal + 2, j)
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
        R = R \ NumOfPixels
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
            dstImageData(QuickVal + 2, j) = R
            dstImageData(QuickVal + 1, j) = g
            dstImageData(QuickVal, j) = b
            
NextMosiacPixel2:

        Next j
        Next i

NextMosaicPixel3:

        'Clear all the variables and go to the next pixel
        R = 0
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

    'Note the current image's width and height, which will be needed to adjust the preview effect
    iWidth = pdImages(CurrentImage).Width
    iHeight = pdImages(CurrentImage).Height
    
    'Draw a preview of the current image in the left picture box
    DrawPreviewImage picPreview
    
    hsWidth.Max = pdImages(CurrentImage).Width
    hsHeight.Max = pdImages(CurrentImage).Height
    MosaicFilter hsWidth.Value, hsHeight.Value, True, picEffect
    
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
    
End Sub

Private Sub hsHeight_Change()
    copyToTextBoxI TxtHeight, hsHeight.Value
    MosaicFilter hsWidth.Value, hsHeight.Value, True, picEffect
End Sub

Private Sub hsWidth_Change()
    copyToTextBoxI TxtWidth, hsWidth.Value
    MosaicFilter hsWidth.Value, hsHeight.Value, True, picEffect
End Sub

Private Sub hsHeight_Scroll()
    copyToTextBoxI TxtHeight, hsHeight.Value
    MosaicFilter hsWidth.Value, hsHeight.Value, True, picEffect
End Sub

Private Sub hsWidth_Scroll()
    copyToTextBoxI TxtWidth, hsWidth.Value
    MosaicFilter hsWidth.Value, hsHeight.Value, True, picEffect
End Sub

Private Sub txtHeight_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate TxtHeight
    If EntryValid(TxtHeight, hsHeight.Min, hsHeight.Max, False, False) Then hsHeight.Value = Val(TxtHeight)
End Sub

Private Sub txtHeight_GotFocus()
    AutoSelectText TxtHeight
End Sub

Private Sub txtWidth_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate TxtWidth
    If EntryValid(TxtWidth, hsWidth.Min, hsWidth.Max, False, False) Then hsWidth.Value = Val(TxtWidth)
End Sub

Private Sub txtWidth_GotFocus()
    AutoSelectText TxtWidth
End Sub

