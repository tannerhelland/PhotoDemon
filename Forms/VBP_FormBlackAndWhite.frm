VERSION 5.00
Begin VB.Form FormBlackAndWhite 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Black/White Color Conversion"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6255
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
   ScaleHeight     =   482
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   417
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboDither 
      BackColor       =   &H00FFFFFF&
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
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   5280
      Width           =   4935
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
      TabIndex        =   7
      Top             =   120
      Width           =   2895
   End
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
      TabIndex        =   6
      Top             =   120
      Width           =   2895
   End
   Begin VB.HScrollBar hsThreshold 
      Height          =   255
      Left            =   360
      Max             =   254
      Min             =   1
      TabIndex        =   1
      Top             =   3840
      Value           =   128
      Width           =   4935
   End
   Begin VB.TextBox txtThreshold 
      Alignment       =   2  'Center
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
      Height          =   315
      Left            =   5400
      TabIndex        =   0
      Text            =   "128"
      Top             =   3810
      Width           =   540
   End
   Begin VB.CheckBox chkAutoThreshold 
      Appearance      =   0  'Flat
      Caption         =   "Automatically choose the best threshold for this image/dithering combination"
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   4320
      Width           =   5895
   End
   Begin VB.CommandButton CmdCancel 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   6390
      Width           =   1125
   End
   Begin VB.CommandButton CmdOK 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      MaskColor       =   &H00000000&
      TabIndex        =   3
      Top             =   6390
      Width           =   1125
   End
   Begin VB.Label lblDither 
      BackStyle       =   0  'Transparent
      Caption         =   "Dithering method:"
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
      TabIndex        =   10
      Top             =   4920
      Width           =   2415
   End
   Begin VB.Label lblBWDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   $"VBP_FormBlackAndWhite.frx":0000
      ForeColor       =   &H00404040&
      Height          =   1095
      Left            =   240
      TabIndex        =   9
      Top             =   6120
      Width           =   3135
   End
   Begin VB.Label lblBeforeandAfter 
      BackStyle       =   0  'Transparent
      Caption         =   "  Before                                                           After"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Width           =   3975
   End
   Begin VB.Label lblHeader 
      BackStyle       =   0  'Transparent
      Caption         =   "Threshold:"
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
      Top             =   3480
      Width           =   2415
   End
End
Attribute VB_Name = "FormBlackAndWhite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Black/White Color Reduction Form
'Copyright ©2000-2012 by Tanner Helland
'Created: some time 2002
'Last updated: 11/September/12
'Last update: full rewrite
'
'The meat of this form is in the module with the same name...look there for
' real algorithm info.
'
'***************************************************************************

Option Explicit

Private Sub cboDither_Click()
    If CBool(chkAutoThreshold.Value) Then txtThreshold.Text = calculateOptimalThreshold(cboDither.ListIndex)
    masterBlackWhiteConversion txtThreshold, cboDither.ListIndex, , , True, picEffect
End Sub

'When the auto threshold button is clicked, disable the scroll bar and text box and calculate the optimal value immediately
Private Sub chkAutoThreshold_Click()
    
    If CBool(chkAutoThreshold.Value) Then
        hsThreshold.Enabled = False
        txtThreshold.Enabled = False
        txtThreshold.Text = calculateOptimalThreshold(cboDither.ListIndex)
    Else
        hsThreshold.Enabled = True
        txtThreshold.Enabled = True
    End If
    
    masterBlackWhiteConversion txtThreshold, cboDither.ListIndex, , , True, picEffect
    
End Sub

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub CmdOK_Click()

    'Checking the threshold value to make sure it's valid
    If EntryValid(txtThreshold, hsThreshold.Min, hsThreshold.Max) = False Then
        AutoSelectText txtThreshold
        Exit Sub
    End If
    
    Me.Visible = False
    
    Process BWMaster, txtThreshold, cboDither.ListIndex, 0, &HFFFFFF
    
    Unload Me
    
End Sub

'LOAD form
Private Sub Form_Load()
    
    'Populate the dither combobox
    cboDither.AddItem "None", 0
    cboDither.AddItem "Ordered (Bayer 4x4)", 1
    cboDither.AddItem "Ordered (Bayer 8x8)", 2
    'cboDither.AddItem "Genuine Floyd-Steinberg", 3
    cboDither.ListIndex = 2
    DoEvents
    
    'Draw the previews
    DrawPreviewImage picPreview
    masterBlackWhiteConversion txtThreshold, cboDither.ListIndex, , , True, picEffect
    
    'Assign the system hand cursor to all relevant objects
    setHandCursorForAll Me
    
End Sub

Private Sub hsThreshold_Change()
    txtThreshold.Text = hsThreshold.Value
End Sub

Private Sub hsThreshold_Scroll()
    chkAutoThreshold.Value = vbUnchecked
    txtThreshold.Text = hsThreshold.Value
End Sub

Private Sub txtThreshold_Change()
    If EntryValid(txtThreshold, hsThreshold.Min, hsThreshold.Max, False, False) Then
        hsThreshold.Value = val(txtThreshold)
        If txtThreshold.Enabled Then masterBlackWhiteConversion txtThreshold, cboDither.ListIndex, , , True, picEffect
    End If
End Sub

Private Sub txtThreshold_GotFocus()
    AutoSelectText txtThreshold
End Sub

'Calculate the optimal threshold for the current image
Private Function calculateOptimalThreshold(ByVal DitherMethod As Long) As Long

    'Create a local array and point it at the pixel data of the image
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    
    prepImageData tmpSA
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
    
    'Color variables
    Dim r As Long, g As Long, b As Long
    
    'Histogram tables
    Dim lLookup(0 To 255)
    Dim pLuminance As Long
    Dim NumOfPixels As Long
        
    'Loop through each pixel in the image, tallying values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        pLuminance = getLuminance(r, g, b)
        
        'Store this value in the histogram
        lLookup(pLuminance) = lLookup(pLuminance) + 1
        
        'Increment the pixel count
        NumOfPixels = NumOfPixels + 1
        
    Next y
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    workingLayer.eraseLayer
    Set workingLayer = Nothing
    
    'Divide the number of pixels by two
    NumOfPixels = NumOfPixels \ 2
            
    Dim pixelCount As Long
    pixelCount = 0
    x = 0
            
    'Loop through the histogram table until we have moved past half the pixels in the image
    Do
        pixelCount = pixelCount + lLookup(x)
        x = x + 1
    Loop While pixelCount < NumOfPixels
    
    'X now equals the value where half the image will be black and half will be white.  Use that to determine an optimal
    ' threshold based on the selected dithering mechanism.
    Select Case DitherMethod
    
        'No dither
        Case 0
            calculateOptimalThreshold = x
            
        'Bayer 4x4
        Case 1
            calculateOptimalThreshold = ((255 - x) + (127 * 3)) \ 4
        
        'Bayer 8x8
        Case 2
            calculateOptimalThreshold = (x + (127 * 7)) \ 8
    
    End Select
    
End Function

'Convert an image to black and white (1-bit image)
Public Sub masterBlackWhiteConversion(ByVal cThreshold As Long, Optional ByVal DitherMethod As Long = 0, Optional ByVal lowColor As Long = &H0, Optional ByVal highColor As Long = &HFFFFFF, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As PictureBox)

    If toPreview = False Then Message "Converting image to two colors..."
    
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
    
    'Low and high color values
    Dim lowR As Long, lowG As Long, lowB As Long
    Dim highR As Long, highG As Long, highB As Long
    
    lowR = ExtractR(lowColor)
    lowG = ExtractG(lowColor)
    lowB = ExtractB(lowColor)
    
    highR = ExtractR(highColor)
    highG = ExtractG(highColor)
    highB = ExtractB(highColor)
    
    'Calculationg color variables (including luminance)
    Dim r As Long, g As Long, b As Long, l As Long
    Dim lReference As Byte
    Dim xModQuick As Long
    Dim DitherTable() As Byte
    
    'Process the image based on the dither method requested
    Select Case DitherMethod
        
        'No dither, so just perform a quick and dirty threshold calculation
        Case 0
    
            For x = initX To finalX
                QuickVal = x * qvDepth
            For y = initY To finalY
        
                'Get the source pixel color values
                r = ImageData(QuickVal + 2, y)
                g = ImageData(QuickVal + 1, y)
                b = ImageData(QuickVal, y)
                
                'Convert those to a luminance value
                l = getLuminance(r, g, b)
            
                'Check the luminance against the threshold, and set new values accordingly
                If l >= cThreshold Then
                    ImageData(QuickVal + 2, y) = highR
                    ImageData(QuickVal + 1, y) = highG
                    ImageData(QuickVal, y) = highB
                Else
                    ImageData(QuickVal + 2, y) = lowR
                    ImageData(QuickVal + 1, y) = lowG
                    ImageData(QuickVal, y) = lowB
                End If
                
            Next y
                If toPreview = False Then
                    If (x And progBarCheck) = 0 Then SetProgBarVal x
                End If
            Next x
            
            
        'Ordered dither (Bayer 4x4).  Unfortunately, this routine requires a unique set of code owing to its specialized
        ' implementation. Coefficients derived from http://en.wikipedia.org/wiki/Ordered_dithering
        ' Note also that ordered dithers ignore the set threshold value.  This is by design.
        Case 1
        
            'First, prepare a Bayer dither table
            ReDim DitherTable(0 To 3, 0 To 3) As Byte
            
            DitherTable(0, 0) = 1
            DitherTable(0, 1) = 9
            DitherTable(0, 2) = 3
            DitherTable(0, 3) = 11
            
            DitherTable(1, 0) = 13
            DitherTable(1, 1) = 5
            DitherTable(1, 2) = 15
            DitherTable(1, 3) = 7
            
            DitherTable(2, 0) = 4
            DitherTable(2, 1) = 12
            DitherTable(2, 2) = 2
            DitherTable(2, 3) = 10
            
            DitherTable(3, 0) = 16
            DitherTable(3, 1) = 8
            DitherTable(3, 2) = 14
            DitherTable(3, 3) = 6
    
            'Convert the dither entries to 255-based values
            For x = 0 To 3
            For y = 0 To 3
                DitherTable(x, y) = DitherTable(x, y) * 16 - 1
            Next y
            Next x
            
            cThreshold = cThreshold * 2

            'Now loop through the image, using the dither values as our threshold
            For x = initX To finalX
                QuickVal = x * qvDepth
                xModQuick = x And 3
            For y = initY To finalY
        
                'Get the source pixel color values
                r = ImageData(QuickVal + 2, y)
                g = ImageData(QuickVal + 1, y)
                b = ImageData(QuickVal, y)
                
                'Convert those to a luminance value and add the value of the dither table
                l = getLuminance(r, g, b) + DitherTable(xModQuick, y And 3)
            
                'Check THAT value against the threshold, and set new values accordingly
                If l >= cThreshold Then
                    ImageData(QuickVal + 2, y) = highR
                    ImageData(QuickVal + 1, y) = highG
                    ImageData(QuickVal, y) = highB
                Else
                    ImageData(QuickVal + 2, y) = lowR
                    ImageData(QuickVal + 1, y) = lowG
                    ImageData(QuickVal, y) = lowB
                End If
                
            Next y
                If toPreview = False Then
                    If (x And progBarCheck) = 0 Then SetProgBarVal x
                End If
            Next x

        'Ordered dither (Bayer 8x8).  Unfortunately, this routine requires a unique set of code owing to its specialized
        ' implementation. Coefficients derived from http://en.wikipedia.org/wiki/Ordered_dithering
        ' Note also that ordered dithers ignore the set threshold value.  This is by design.
        Case 2
        
            'First, prepare a Bayer dither table
            ReDim DitherTable(0 To 7, 0 To 7) As Byte
            
            DitherTable(0, 0) = 1
            DitherTable(0, 1) = 49
            DitherTable(0, 2) = 13
            DitherTable(0, 3) = 61
            DitherTable(0, 4) = 4
            DitherTable(0, 5) = 52
            DitherTable(0, 6) = 16
            DitherTable(0, 7) = 64
            
            DitherTable(1, 0) = 33
            DitherTable(1, 1) = 17
            DitherTable(1, 2) = 45
            DitherTable(1, 3) = 29
            DitherTable(1, 4) = 36
            DitherTable(1, 5) = 20
            DitherTable(1, 6) = 48
            DitherTable(1, 7) = 32
            
            DitherTable(2, 0) = 9
            DitherTable(2, 1) = 57
            DitherTable(2, 2) = 5
            DitherTable(2, 3) = 53
            DitherTable(2, 4) = 12
            DitherTable(2, 5) = 60
            DitherTable(2, 6) = 8
            DitherTable(2, 7) = 56
            
            DitherTable(3, 0) = 41
            DitherTable(3, 1) = 25
            DitherTable(3, 2) = 37
            DitherTable(3, 3) = 21
            DitherTable(3, 4) = 44
            DitherTable(3, 5) = 28
            DitherTable(3, 6) = 40
            DitherTable(3, 7) = 24
            
            DitherTable(4, 0) = 3
            DitherTable(4, 1) = 51
            DitherTable(4, 2) = 15
            DitherTable(4, 3) = 63
            DitherTable(4, 4) = 2
            DitherTable(4, 5) = 50
            DitherTable(4, 6) = 14
            DitherTable(4, 7) = 62
            
            DitherTable(5, 0) = 35
            DitherTable(5, 1) = 19
            DitherTable(5, 2) = 47
            DitherTable(5, 3) = 31
            DitherTable(5, 4) = 34
            DitherTable(5, 5) = 18
            DitherTable(5, 6) = 46
            DitherTable(5, 7) = 30
    
            DitherTable(6, 0) = 11
            DitherTable(6, 1) = 59
            DitherTable(6, 2) = 7
            DitherTable(6, 3) = 55
            DitherTable(6, 4) = 10
            DitherTable(6, 5) = 58
            DitherTable(6, 6) = 6
            DitherTable(6, 7) = 54
            
            DitherTable(7, 0) = 43
            DitherTable(7, 1) = 27
            DitherTable(7, 2) = 39
            DitherTable(7, 3) = 23
            DitherTable(7, 4) = 42
            DitherTable(7, 5) = 26
            DitherTable(7, 6) = 38
            DitherTable(7, 7) = 22
            
            'Convert the dither entries to 255-based values
            For x = 0 To 7
            For y = 0 To 7
                DitherTable(x, y) = DitherTable(x, y) * 4 - 1
            Next y
            Next x

            cThreshold = cThreshold * 2

            'Now loop through the image, using the dither values as our threshold
            For x = initX To finalX
                QuickVal = x * qvDepth
                xModQuick = x And 7
            For y = initY To finalY
        
                'Get the source pixel color values
                r = ImageData(QuickVal + 2, y)
                g = ImageData(QuickVal + 1, y)
                b = ImageData(QuickVal, y)
                
                'Convert those to a luminance value and add the value of the dither table
                l = getLuminance(r, g, b) + DitherTable(xModQuick, y And 7)
            
                'Check THAT value against the threshold, and set new values accordingly
                If l >= cThreshold Then
                    ImageData(QuickVal + 2, y) = highR
                    ImageData(QuickVal + 1, y) = highG
                    ImageData(QuickVal, y) = highB
                Else
                    ImageData(QuickVal + 2, y) = lowR
                    ImageData(QuickVal + 1, y) = lowG
                    ImageData(QuickVal, y) = lowB
                End If
                
            Next y
                If toPreview = False Then
                    If (x And progBarCheck) = 0 Then SetProgBarVal x
                End If
            Next x
            
    End Select
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
   


    
    

End Sub

Private Sub txtThreshold_KeyPress(KeyAscii As Integer)
    chkAutoThreshold.Value = vbUnchecked
End Sub
