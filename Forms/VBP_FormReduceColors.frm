VERSION 5.00
Begin VB.Form FormReduceColors 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Reduce Image Colors"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6480
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
   ScaleHeight     =   554
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   432
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2730
      Left            =   240
      ScaleHeight     =   180
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   191
      TabIndex        =   21
      Top             =   120
      Width           =   2895
   End
   Begin VB.PictureBox picEffect 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2730
      Left            =   3360
      ScaleHeight     =   180
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   191
      TabIndex        =   20
      Top             =   120
      Width           =   2895
   End
   Begin VB.HScrollBar hsBlue 
      Height          =   255
      Left            =   2520
      Max             =   64
      Min             =   2
      TabIndex        =   10
      Top             =   6195
      Value           =   6
      Width           =   3015
   End
   Begin VB.HScrollBar hsGreen 
      Height          =   255
      Left            =   2520
      Max             =   64
      Min             =   2
      TabIndex        =   8
      Top             =   5805
      Value           =   7
      Width           =   3015
   End
   Begin VB.HScrollBar hsRed 
      Height          =   255
      Left            =   2520
      Max             =   64
      Min             =   2
      TabIndex        =   6
      Top             =   5415
      Value           =   6
      Width           =   3015
   End
   Begin VB.OptionButton OptQuant 
      Appearance      =   0  'Flat
      Caption         =   "PhotoDemon Advanced (manual)"
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
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   4
      Top             =   4440
      Width           =   3495
   End
   Begin VB.TextBox TxtB 
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
      Left            =   5640
      MaxLength       =   3
      TabIndex        =   9
      Text            =   "6"
      Top             =   6150
      Width           =   615
   End
   Begin VB.TextBox TxtG 
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
      Left            =   5640
      MaxLength       =   3
      TabIndex        =   7
      Text            =   "7"
      Top             =   5760
      Width           =   615
   End
   Begin VB.TextBox TxtR 
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
      Left            =   5640
      MaxLength       =   3
      TabIndex        =   5
      Text            =   "6"
      Top             =   5370
      Width           =   615
   End
   Begin VB.CheckBox chkColorDither 
      Appearance      =   0  'Flat
      Caption         =   " Use error diffusion dithering"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   840
      TabIndex        =   11
      Top             =   7110
      Value           =   1  'Checked
      Width           =   2655
   End
   Begin VB.CheckBox chkSmartColors 
      Appearance      =   0  'Flat
      Caption         =   " Use realistic coloring"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3600
      TabIndex        =   12
      Top             =   7110
      Value           =   1  'Checked
      Width           =   2535
   End
   Begin VB.OptionButton OptQuant 
      Appearance      =   0  'Flat
      Caption         =   "Xiaolin Wu (automatic)"
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
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   3720
      Value           =   -1  'True
      Width           =   2775
   End
   Begin VB.OptionButton OptQuant 
      Appearance      =   0  'Flat
      Caption         =   "NeuQuant by Anthony Dekker (automatic)"
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
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   3
      Top             =   4080
      Width           =   4095
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   3720
      TabIndex        =   0
      Top             =   7680
      Width           =   1245
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   5040
      TabIndex        =   1
      Top             =   7680
      Width           =   1245
   End
   Begin VB.Label lblWarning 
      Caption         =   "Note: some options on this page have been disabled because the FreeImage plugin could not be found."
      ForeColor       =   &H000000C0&
      Height          =   1095
      Left            =   4470
      TabIndex        =   19
      Top             =   3600
      Visible         =   0   'False
      Width           =   1815
      WordWrap        =   -1  'True
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
      Left            =   3480
      TabIndex        =   23
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
      Left            =   360
      TabIndex        =   22
      Top             =   2880
      Width           =   480
   End
   Begin VB.Label lblOptions 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "advanced quantization options:"
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
      TabIndex        =   18
      Top             =   5040
      Width           =   3300
   End
   Begin VB.Label lblBlue 
      BackStyle       =   0  'Transparent
      Caption         =   "Possible blue values (2-64):"
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   6225
      Width           =   2175
   End
   Begin VB.Label lblGreen 
      BackStyle       =   0  'Transparent
      Caption         =   "Possible green values (2-64):"
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   5835
      Width           =   2535
   End
   Begin VB.Label lblRed 
      BackStyle       =   0  'Transparent
      Caption         =   "Possible red values (2-64):"
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   5445
      Width           =   2175
   End
   Begin VB.Label lblMaxColors 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "These parameters allow for a maximum of 252 colors in the quantized image."
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   360
      TabIndex        =   14
      Top             =   6675
      Width           =   5505
   End
   Begin VB.Label lblQuantMethod 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "quantization method:"
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
      Height          =   405
      Left            =   240
      TabIndex        =   13
      Top             =   3360
      Width           =   2265
   End
End
Attribute VB_Name = "FormReduceColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Color Reduction Form
'Copyright ©2000-2013 by Tanner Helland
'Created: 4/October/00
'Last updated: 11/September/12
'Last update: Rewrote all reduction algorithms against the new layer class and added previewing
'
'In the original incarnation of PhotoDemon, this was a central part of the project. I have since not used
' it as much (since the project has become almost entirely centered around 24-bit imaging), but the code
' is solid and the feature set is large.
'
'***************************************************************************

Option Explicit

'SetDIBitsToDevice is used to interact with the FreeImage DLL
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As Any, ByVal wUsage As Long) As Long

Private Sub ChkColorDither_Click()
    updateReductionPreview
End Sub

Private Sub chkSmartColors_Click()
    updateReductionPreview
End Sub

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub cmdOK_Click()

    'Check to see which method the user has requested
    
    'Xiaolin Wu
    If OptQuant(0).Value = True Then
        FormReduceColors.Visible = False
        Process ReduceColors, REDUCECOLORS_AUTO, FIQ_WUQUANT
        Unload Me
    
    'NeuQuant
    ElseIf OptQuant(1).Value = True Then
        FormReduceColors.Visible = False
        Process ReduceColors, REDUCECOLORS_AUTO, FIQ_NNQUANT
        Unload Me
    
    'Manual
    Else
    
        'Before reducing anything, check to make sure the textboxes have valid input
        If Not EntryValid(TxtR, hsRed.Min, hsRed.Max, True, True) Then
            AutoSelectText TxtR
            Exit Sub
        End If
        If Not EntryValid(TxtG, hsGreen.Min, hsGreen.Max, True, True) Then
            AutoSelectText TxtG
            Exit Sub
        End If
        If Not EntryValid(TxtB, hsBlue.Min, hsBlue.Max, True, True) Then
            AutoSelectText TxtB
            Exit Sub
        End If
        
        Me.Visible = False
        
        'Do the appropriate method of color reduction
        If chkColorDither.Value = vbUnchecked Then
            If chkSmartColors.Value = vbUnchecked Then
                Process ReduceColors, REDUCECOLORS_MANUAL, TxtR, TxtG, TxtB, False
            Else
                Process ReduceColors, REDUCECOLORS_MANUAL, TxtR, TxtG, TxtB, True
            End If
        Else
            If chkSmartColors.Value = vbUnchecked Then
                Process ReduceColors, REDUCECOLORS_MANUAL_ERRORDIFFUSION, TxtR, TxtG, TxtB, False
            Else
                Process ReduceColors, REDUCECOLORS_MANUAL_ERRORDIFFUSION, TxtR, TxtG, TxtB, True
            End If
        End If
        
        Unload Me
        
    End If
    
End Sub

Private Sub Form_Activate()
    
    'Only allow AutoReduction stuff if the FreeImage dll was found.
    If g_ImageFormats.FreeImageEnabled = False Then
        OptQuant(0).Enabled = False
        OptQuant(1).Enabled = False
        OptQuant(2).Value = True
        DisplayManualOptions True
        lblWarning.Visible = True
    Else
        OptQuant(0).Value = True
        DisplayManualOptions False
    End If
    
    'Draw the previews
    DrawPreviewImage picPreview
    updateReductionPreview
    
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Enable/disable the manual settings depending on which option button has been selected
Private Sub OptQuant_Click(Index As Integer)
    If OptQuant(2).Value = True Then DisplayManualOptions True Else DisplayManualOptions False
    updateReductionPreview
End Sub

'The large chunk of subs that follow serve to keep the text box and scroll bar values in lock-step
Private Sub hsRed_Change()
    copyToTextBoxI TxtR, hsRed.Value
    updateColorLabel
End Sub

Private Sub hsGreen_Change()
    copyToTextBoxI TxtG, hsGreen.Value
    updateColorLabel
End Sub

Private Sub hsBlue_Change()
    copyToTextBoxI TxtB, hsBlue.Value
    updateColorLabel
End Sub

Private Sub hsRed_Scroll()
    copyToTextBoxI TxtR, hsRed.Value
    updateColorLabel
End Sub

Private Sub hsGreen_Scroll()
    copyToTextBoxI TxtG, hsGreen.Value
    updateColorLabel
End Sub

Private Sub hsBlue_Scroll()
    copyToTextBoxI TxtB, hsBlue.Value
    updateColorLabel
End Sub

Private Sub TxtB_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate TxtB
    If EntryValid(TxtB, hsBlue.Min, hsBlue.Max, False, False) Then hsBlue.Value = Val(TxtB)
End Sub

Private Sub TxtB_GotFocus()
    AutoSelectText TxtB
End Sub

Private Sub TxtG_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate TxtG
    If EntryValid(TxtG, hsGreen.Min, hsGreen.Max, False, False) Then hsGreen.Value = Val(TxtG)
End Sub

Private Sub TxtG_GotFocus()
    AutoSelectText TxtG
End Sub

Private Sub TxtR_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate TxtR
    If EntryValid(TxtR, hsRed.Min, hsRed.Max, False, False) Then hsRed.Value = Val(TxtR)
End Sub

Private Sub TxtR_GotFocus()
    AutoSelectText TxtR
End Sub

'This lets the user know the max number of colors that the current set of quantization parameters will allow for
Private Sub updateColorLabel()
    If EntryValid(TxtR, hsRed.Min, hsRed.Max, False, False) And EntryValid(TxtG, hsGreen.Min, hsGreen.Max, False, False) And EntryValid(TxtB, hsBlue.Min, hsBlue.Max, False, False) Then
        lblMaxColors = "These parameters allow for a maximum of " & Val(TxtR) * Val(TxtG) * Val(TxtB) & " colors in the quantized image."
        updateReductionPreview
    Else
        lblMaxColors = "Color count could not be calculated due to invalid text box values."
    End If
End Sub

'Enable/disable the manual options depending on which quantization method has been selected
Private Sub DisplayManualOptions(ByVal toDisplay As Boolean)
    If toDisplay = False Then
        lblOptions.ForeColor = RGB(160, 160, 160)
        lblRed.ForeColor = RGB(160, 160, 160)
        lblGreen.ForeColor = RGB(160, 160, 160)
        lblBlue.ForeColor = RGB(160, 160, 160)
        lblMaxColors.ForeColor = RGB(160, 160, 160)
        TxtR.Enabled = False
        TxtG.Enabled = False
        TxtB.Enabled = False
        hsRed.Enabled = False
        hsGreen.Enabled = False
        hsBlue.Enabled = False
        chkColorDither.Enabled = False
        chkSmartColors.Enabled = False
    Else
        lblOptions.ForeColor = &H400000
        lblRed.ForeColor = &H400000
        lblGreen.ForeColor = &H400000
        lblBlue.ForeColor = &H400000
        lblMaxColors.ForeColor = &H400000
        TxtR.Enabled = True
        TxtG.Enabled = True
        TxtB.Enabled = True
        hsRed.Enabled = True
        hsGreen.Enabled = True
        hsBlue.Enabled = True
        chkColorDither.Enabled = True
        chkSmartColors.Enabled = True
    End If
End Sub

'Automatic 8-bit color reduction via the FreeImage DLL.
Public Sub ReduceImageColors_Auto(ByVal qMethod As Long, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As PictureBox)

    'If a selection is active, remove it.  (This is not the most elegant solution, but we can fix it at a later date.)
    If pdImages(CurrentImage).selectionActive Then
        pdImages(CurrentImage).selectionActive = False
        tInit tSelection, False
    End If

    'If this is a preview, we want to perform the color reduction on a temporary image
    If toPreview Then
        Dim tmpSA As SAFEARRAY2D
        prepImageData tmpSA, toPreview, dstPic
    End If

    'Make sure we found the FreeImage plug-in when the program was loaded
    If g_ImageFormats.FreeImageEnabled Then
    
        'Load the FreeImage dll into memory
        Dim hLib As Long
        hLib = LoadLibrary(g_PluginPath & "FreeImage.dll")
        
        If toPreview = False Then Message "Quantizing image using the FreeImage library..."
        
        'Convert our current layer to a FreeImage-type DIB
        Dim fi_DIB As Long
        
        If toPreview Then
            If workingLayer.getLayerColorDepth = 32 Then workingLayer.compositeBackgroundColor 255, 255, 255
            fi_DIB = FreeImage_CreateFromDC(workingLayer.getLayerDC)
        Else
            If pdImages(CurrentImage).mainLayer.getLayerColorDepth = 32 Then pdImages(CurrentImage).mainLayer.compositeBackgroundColor 255, 255, 255
            fi_DIB = FreeImage_CreateFromDC(pdImages(CurrentImage).mainLayer.getLayerDC)
        End If
        
        'Use that handle to save the image to GIF format, with required 8bpp (256 color) conversion
        If fi_DIB <> 0 Then
            
            Dim returnDIB As Long
            
            returnDIB = FreeImage_ColorQuantizeEx(fi_DIB, qMethod, True)
            
            'If this is a preview, render it to the temporary layer.  Otherwise, use the current main layer.
            If toPreview Then
                workingLayer.createBlank workingLayer.getLayerWidth, workingLayer.getLayerHeight, 24
                SetDIBitsToDevice workingLayer.getLayerDC, 0, 0, workingLayer.getLayerWidth, workingLayer.getLayerHeight, 0, 0, 0, workingLayer.getLayerHeight, ByVal FreeImage_GetBits(returnDIB), ByVal FreeImage_GetInfo(returnDIB), 0&
            Else
                pdImages(CurrentImage).mainLayer.createBlank pdImages(CurrentImage).Width, pdImages(CurrentImage).Height, 24
                SetDIBitsToDevice pdImages(CurrentImage).mainLayer.getLayerDC, 0, 0, pdImages(CurrentImage).Width, pdImages(CurrentImage).Height, 0, 0, 0, pdImages(CurrentImage).Height, ByVal FreeImage_GetBits(returnDIB), ByVal FreeImage_GetInfo(returnDIB), 0&
            End If
            
            'With the transfer complete, release the FreeImage DIB and unload the library
            If returnDIB <> 0 Then FreeImage_UnloadEx returnDIB
            FreeLibrary hLib
     
            'If this is a preview, draw the new image to the picture box and exit.  Otherwise, render the new main image layer.
            If toPreview Then
                finalizeImageData toPreview, dstPic
            Else
                ScrollViewport FormMain.ActiveForm
                Message "Image successfully quantized to 256 unique colors. "
            End If
            
        End If
        
    Else
        MsgBox "The FreeImage interface plug-ins (FreeImage.dll) was marked as missing or disabled upon program initialization." & vbCrLf & vbCrLf & "To enable support for this feature, please copy the FreeImage.dll file into the plug-in directory and reload " & PROGRAMNAME & ".", vbExclamation + vbOKOnly + vbApplicationModal, PROGRAMNAME & " FreeImage Interface Error"
        Exit Sub
    End If
    
End Sub

'Bit RGB color reduction (no error diffusion)
Public Sub ReduceImageColors_BitRGB(ByVal rValue As Byte, ByVal gValue As Byte, ByVal bValue As Byte, Optional ByVal smartColors As Boolean = False, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As PictureBox)

    If toPreview = False Then Message "Applying manual RGB modifications to image..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    
    prepImageData tmpSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Color variables
    Dim r As Long, g As Long, b As Long
    Dim mR As Single, mG As Single, mB As Single
    Dim cR As Long, cG As Long, cb As Long
    
    'New code for so-called "Intelligent Coloring"
    Dim rLookup() As Long
    Dim gLookup() As Long
    Dim bLookup() As Long
    Dim countLookup() As Long
    
    ReDim rLookup(0 To rValue, 0 To gValue, 0 To bValue) As Long
    ReDim gLookup(0 To rValue, 0 To gValue, 0 To bValue) As Long
    ReDim bLookup(0 To rValue, 0 To gValue, 0 To bValue) As Long
    ReDim countLookup(0 To rValue, 0 To gValue, 0 To bValue) As Long
    
    'Prepare inputted variables for the requisite maths
    rValue = rValue - 1
    gValue = gValue - 1
    bValue = bValue - 1
    mR = (256 / rValue)
    mG = (256 / gValue)
    mB = (256 / bValue)
    
    'Finally, prepare conversion look-up tables (which will make the actual color reduction much faster)
    Dim rQuick(0 To 255) As Byte, gQuick(0 To 255) As Byte, bQuick(0 To 255) As Byte
    For x = 0 To 255
        rQuick(x) = Int((x / mR) + 0.5)
        gQuick(x) = Int((x / mG) + 0.5)
        bQuick(x) = Int((x / mB) + 0.5)
    Next x
    
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Truncate R, G, and B values (posterize-style) into discreet increments.  0.5 is added for rounding purposes.
        cR = rQuick(r)
        cG = gQuick(g)
        cb = bQuick(b)
        
        'If we're doing Intelligent Coloring, place color values into a look-up table
        If smartColors = True Then
            rLookup(cR, cG, cb) = rLookup(cR, cG, cb) + r
            gLookup(cR, cG, cb) = gLookup(cR, cG, cb) + g
            bLookup(cR, cG, cb) = bLookup(cR, cG, cb) + b
            'Also, keep track of how many colors fall into this bucket (so we can later determine an average color)
            countLookup(cR, cG, cb) = countLookup(cR, cG, cb) + 1
        End If
        
        'Multiply the now-discretely divided R, G, and B values to (0-255) equivalents
        cR = cR * mR
        cG = cG * mG
        cb = cb * mB
        
        If cR > 255 Then cR = 255
        If cR < 0 Then cR = 0
        If cG > 255 Then cG = 255
        If cG < 0 Then cG = 0
        If cb > 255 Then cb = 255
        If cb < 0 Then cb = 0
        
        'If we are not doing Intelligent Coloring, assign the colors now (to avoid having to do another loop at the end)
        If smartColors = False Then
            ImageData(QuickVal + 2, y) = cR
            ImageData(QuickVal + 1, y) = cG
            ImageData(QuickVal, y) = cb
        End If
        
    Next y
        If toPreview = False Then
            If (x And progBarCheck) = 0 Then SetProgBarVal x
        End If
    Next x
    
    'Intelligent Coloring requires extra work.  Perform a second loop through the image, replacing values with their
    ' computed counterparts.
    If smartColors Then
    
        If toPreview = False Then
            SetProgBarVal getProgBarMax
            Message "Applying intelligent coloring..."
        End If
        
        'Find average colors based on color counts
        For r = 0 To rValue
        For g = 0 To gValue
        For b = 0 To bValue
            If countLookup(r, g, b) <> 0 Then
                rLookup(r, g, b) = Int(Int(rLookup(r, g, b)) / Int(countLookup(r, g, b)))
                gLookup(r, g, b) = Int(Int(gLookup(r, g, b)) / Int(countLookup(r, g, b)))
                bLookup(r, g, b) = Int(Int(bLookup(r, g, b)) / Int(countLookup(r, g, b)))
                If rLookup(r, g, b) > 255 Then rLookup(r, g, b) = 255
                If gLookup(r, g, b) > 255 Then gLookup(r, g, b) = 255
                If bLookup(r, g, b) > 255 Then bLookup(r, g, b) = 255
            End If
        Next b
        Next g
        Next r
        
        'Assign average colors back into the picture
        For x = initX To finalX
            QuickVal = x * qvDepth
        For y = initY To finalY
        
            r = ImageData(QuickVal + 2, y)
            g = ImageData(QuickVal + 1, y)
            b = ImageData(QuickVal, y)
            
            cR = rQuick(r)
            cG = gQuick(g)
            cb = bQuick(b)
            
            ImageData(QuickVal + 2, y) = rLookup(cR, cG, cb)
            ImageData(QuickVal + 1, y) = gLookup(cR, cG, cb)
            ImageData(QuickVal, y) = bLookup(cR, cG, cb)
            
        Next y
        Next x
        
    End If
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
    
End Sub

'Error Diffusion dithering to x# shades of color per component
Public Sub ReduceImageColors_BitRGB_ErrorDif(ByVal rValue As Byte, ByVal gValue As Byte, ByVal bValue As Byte, Optional ByVal smartColors As Boolean = False, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As PictureBox)
    
    If toPreview = False Then Message "Applying manual RGB modifications to image..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    
    prepImageData tmpSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    SetProgBarMax finalY
    progBarCheck = findBestProgBarValue()
    
    'Color variables
    Dim r As Long, g As Long, b As Long
    Dim cR As Long, cG As Long, cb As Long
    Dim iR As Long, iG As Long, iB As Long
    Dim mR As Single, mG As Single, mB As Single
    Dim eR As Single, eG As Single, eB As Single
    
    'New code for so-called "Intelligent Coloring"
    Dim rLookup() As Long
    Dim gLookup() As Long
    Dim bLookup() As Long
    Dim countLookup() As Long
    
    ReDim rLookup(0 To rValue, 0 To gValue, 0 To bValue) As Long
    ReDim gLookup(0 To rValue, 0 To gValue, 0 To bValue) As Long
    ReDim bLookup(0 To rValue, 0 To gValue, 0 To bValue) As Long
    ReDim countLookup(0 To rValue, 0 To gValue, 0 To bValue) As Long
    
    'Prepare inputted variables for the requisite maths
    rValue = rValue - 1
    gValue = gValue - 1
    bValue = bValue - 1
    mR = (256 / rValue)
    mG = (256 / gValue)
    mB = (256 / bValue)
    
    'Finally, prepare conversion look-up tables (which will make the actual color reduction much faster)
    Dim rQuick(0 To 255) As Byte, gQuick(0 To 255) As Byte, bQuick(0 To 255) As Byte
    For x = 0 To 255
        rQuick(x) = Int((x / mR) + 0.5)
        gQuick(x) = Int((x / mG) + 0.5)
        bQuick(x) = Int((x / mB) + 0.5)
    Next x
    
    'Loop through each pixel in the image, converting values as we go
    For y = initY To finalY
    For x = initX To finalX
        
        QuickVal = x * qvDepth
    
        'Get the source pixel color values
        iR = ImageData(QuickVal + 2, y)
        iG = ImageData(QuickVal + 1, y)
        iB = ImageData(QuickVal, y)
        
        r = iR + eR
        g = iG + eG
        b = iB + eB
        
        If r > 255 Then r = 255
        If g > 255 Then g = 255
        If b > 255 Then b = 255
        If r < 0 Then r = 0
        If g < 0 Then g = 0
        If b < 0 Then b = 0
        
        'Truncate R, G, and B values (posterize-style) into discreet increments.  0.5 is added for rounding purposes.
        cR = rQuick(r)
        cG = gQuick(g)
        cb = bQuick(b)
        
        'If we're doing Intelligent Coloring, place color values into a look-up table
        If smartColors = True Then
            rLookup(cR, cG, cb) = rLookup(cR, cG, cb) + r
            gLookup(cR, cG, cb) = gLookup(cR, cG, cb) + g
            bLookup(cR, cG, cb) = bLookup(cR, cG, cb) + b
            'Also, keep track of how many colors fall into this bucket (so we can later determine an average color)
            countLookup(cR, cG, cb) = countLookup(cR, cG, cb) + 1
        End If
        
        'Multiply the now-discretely divided R, G, and B values to (0-255) equivalents
        cR = cR * mR
        cG = cG * mG
        cb = cb * mB
        
        'Calculate error
        eR = iR - cR
        eG = iG - cG
        eB = iB - cb
        
        'Diffuse the error further (in a grid pattern) to prevent undesirable lining effects
        If (x + y) And 3 <> 0 Then
            eR = eR \ 2
            eG = eG \ 2
            eB = eB \ 2
        End If
        
        If cR > 255 Then cR = 255
        If cR < 0 Then cR = 0
        If cG > 255 Then cG = 255
        If cG < 0 Then cG = 0
        If cb > 255 Then cb = 255
        If cb < 0 Then cb = 0
        
        'If we are not doing Intelligent Coloring, assign the colors now (to avoid having to do another loop at the end)
        If smartColors = False Then
            ImageData(QuickVal + 2, y) = cR
            ImageData(QuickVal + 1, y) = cG
            ImageData(QuickVal, y) = cb
        End If
        
    Next x
        
        'At the end of each line, reset our error-tracking values
        eR = 0
        eG = 0
        eB = 0
        
        If toPreview = False Then
            If (y And progBarCheck) = 0 Then SetProgBarVal y
        End If
    Next y
    
    'Intelligent Coloring requires extra work.  Perform a second loop through the image, replacing values with their
    ' computed counterparts.
    If smartColors Then
        
        If toPreview = False Then
            SetProgBarVal getProgBarMax
            Message "Applying intelligent coloring..."
        End If
        
        'Find average colors based on color counts
        For r = 0 To rValue
        For g = 0 To gValue
        For b = 0 To bValue
            If countLookup(r, g, b) <> 0 Then
                rLookup(r, g, b) = Int(Int(rLookup(r, g, b)) / Int(countLookup(r, g, b)))
                gLookup(r, g, b) = Int(Int(gLookup(r, g, b)) / Int(countLookup(r, g, b)))
                bLookup(r, g, b) = Int(Int(bLookup(r, g, b)) / Int(countLookup(r, g, b)))
                If rLookup(r, g, b) > 255 Then rLookup(r, g, b) = 255
                If gLookup(r, g, b) > 255 Then gLookup(r, g, b) = 255
                If bLookup(r, g, b) > 255 Then bLookup(r, g, b) = 255
            End If
        Next b
        Next g
        Next r
        
        'Assign average colors back into the picture
        For y = initY To finalY
        For x = initX To finalX
            
            QuickVal = x * qvDepth
        
            iR = ImageData(QuickVal + 2, y)
            iG = ImageData(QuickVal + 1, y)
            iB = ImageData(QuickVal, y)
            
            r = iR + eR
            g = iG + eG
            b = iB + eB
            
            If r > 255 Then r = 255
            If g > 255 Then g = 255
            If b > 255 Then b = 255
            If r < 0 Then r = 0
            If g < 0 Then g = 0
            If b < 0 Then b = 0
            
            cR = rQuick(r)
            cG = gQuick(g)
            cb = bQuick(b)
            
            ImageData(QuickVal + 2, y) = rLookup(cR, cG, cb)
            ImageData(QuickVal + 1, y) = gLookup(cR, cG, cb)
            ImageData(QuickVal, y) = bLookup(cR, cG, cb)
            
            'Calculate the error for this pixel
            cR = cR * mR
            cG = cG * mG
            cb = cb * mB
        
            eR = iR - cR
            eG = iG - cG
            eB = iB - cb
            
            'Diffuse the error further (in a grid pattern) to prevent undesirable lining effects
            If (x + y) And 3 <> 0 Then
                eR = eR \ 2
                eG = eG \ 2
                eB = eB \ 2
            End If
            
        Next x
        
            'At the end of each line, reset our error-tracking values
            eR = 0
            eG = 0
            eB = 0
        
        Next y
        
    End If
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
    
End Sub

'Use this sub to update the on-screen preview
Private Sub updateReductionPreview()
    If OptQuant(0).Value = True Then
        ReduceImageColors_Auto FIQ_WUQUANT, True, picEffect
    ElseIf OptQuant(1).Value = True Then
        ReduceImageColors_Auto FIQ_NNQUANT, True, picEffect
    Else
        If EntryValid(TxtR, hsRed.Min, hsRed.Max, False, False) And EntryValid(TxtG, hsGreen.Min, hsGreen.Max, False, False) And EntryValid(TxtB, hsBlue.Min, hsBlue.Max, False, False) Then
            If chkColorDither.Value = vbUnchecked Then
                If chkSmartColors.Value = vbUnchecked Then
                    ReduceImageColors_BitRGB TxtR, TxtG, TxtB, False, True, picEffect
                Else
                    ReduceImageColors_BitRGB TxtR, TxtG, TxtB, True, True, picEffect
                End If
            Else
                If chkSmartColors.Value = vbUnchecked Then
                    ReduceImageColors_BitRGB_ErrorDif TxtR, TxtG, TxtB, False, True, picEffect
                Else
                    ReduceImageColors_BitRGB_ErrorDif TxtR, TxtG, TxtB, True, True, picEffect
                End If
            End If
        End If
    End If
End Sub

