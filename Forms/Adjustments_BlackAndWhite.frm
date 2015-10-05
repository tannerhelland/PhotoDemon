VERSION 5.00
Begin VB.Form FormMonochrome 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Monochrome Conversion"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12150
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
   ScaleWidth      =   810
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.sliderTextCombo sltThreshold 
      Height          =   720
      Left            =   6000
      TabIndex        =   6
      Top             =   960
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   1270
      Caption         =   "threshold"
      Min             =   1
      Max             =   254
      Value           =   127
      NotchPosition   =   2
      NotchValueCustom=   127
   End
   Begin PhotoDemon.smartCheckBox chkAutoThreshold 
      Height          =   330
      Left            =   6120
      TabIndex        =   5
      Top             =   1860
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   582
      Caption         =   "automatically calculate threshold"
      Value           =   0
   End
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
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2880
      Width           =   4935
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.colorSelector colorPicker 
      Height          =   615
      Index           =   0
      Left            =   6120
      TabIndex        =   7
      Top             =   3840
      Width           =   2775
      _ExtentX        =   9763
      _ExtentY        =   1085
      curColor        =   0
   End
   Begin PhotoDemon.colorSelector colorPicker 
      Height          =   615
      Index           =   1
      Left            =   9000
      TabIndex        =   8
      Top             =   3840
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1085
   End
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   12150
      _ExtentX        =   21431
      _ExtentY        =   1323
      BackColor       =   14802140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "final colors"
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
      TabIndex        =   3
      Top             =   3480
      Width           =   1155
   End
   Begin VB.Label lblDither 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "dithering technique"
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
      Top             =   2520
      Width           =   2040
   End
End
Attribute VB_Name = "FormMonochrome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Monochrome Conversion Form
'Copyright 2002-2015 by Tanner Helland
'Created: some time 2002
'Last updated: 17/August/13
'Last update: greatly simplify code by using new command bar custom control
'
'The meat of this form is in the module with the same name...look there for
' real algorithm info.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private Sub cboDither_Click()
    If CBool(chkAutoThreshold.Value) Then sltThreshold = calculateOptimalThreshold()
    updatePreview
End Sub

'When the auto threshold button is clicked, disable the scroll bar and text box and calculate the optimal value immediately
Private Sub chkAutoThreshold_Click()
    
    If CBool(chkAutoThreshold.Value) Then
        cmdBar.markPreviewStatus False
        sltThreshold = calculateOptimalThreshold()
        cmdBar.markPreviewStatus True
    End If
    
    sltThreshold.Enabled = Not CBool(chkAutoThreshold.Value)
    
    updatePreview
    
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Color to monochrome", , buildParams(sltThreshold, cboDither.ListIndex, colorPicker(0).Color, colorPicker(1).Color), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

'When resetting, set the color boxes to black and white, and the dithering combo box to 6 (Stucki)
Private Sub cmdBar_ResetClick()
    
    colorPicker(0).Color = RGB(0, 0, 0)
    colorPicker(1).Color = RGB(255, 255, 255)
    cboDither.ListIndex = 6     'Stucki dithering
    
    'Standard threshold value
    chkAutoThreshold.Value = vbUnchecked
    sltThreshold.Value = 127
    
End Sub

Private Sub colorPicker_ColorChanged(Index As Integer)
    updatePreview
End Sub

Private Sub Form_Activate()
        
    'Apply translations and visual themes
    MakeFormPretty Me
    
    'Draw the preview
    updatePreview
    
End Sub

Private Sub Form_Load()
    
    cmdBar.markPreviewStatus False
    
    'Populate the dither combobox
    cboDither.Clear
    cboDither.AddItem "None", 0
    cboDither.AddItem "Ordered (Bayer 4x4)", 1
    cboDither.AddItem "Ordered (Bayer 8x8)", 2
    cboDither.AddItem "False (Fast) Floyd-Steinberg", 3
    cboDither.AddItem "Genuine Floyd-Steinberg", 4
    cboDither.AddItem "Jarvis, Judice, and Ninke", 5
    cboDither.AddItem "Stucki", 6
    cboDither.AddItem "Burkes", 7
    cboDither.AddItem "Sierra-3", 8
    cboDither.AddItem "Two-Row Sierra", 9
    cboDither.AddItem "Sierra Lite", 10
    cboDither.AddItem "Atkinson / Classic Macintosh", 11
    cboDither.ListIndex = 6
    
    cmdBar.markPreviewStatus True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Calculate the optimal threshold for the current image
Private Function calculateOptimalThreshold() As Long

    'Create a local array and point it at the pixel data of the image
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    
    prepImageData tmpSA
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
    
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
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
    workingDIB.eraseDIB
    Set workingDIB = Nothing
            
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
    
    'Make sure our suggestion doesn't exceed the limits allowed by the tool
    If x > 254 Then x = 220
    
    calculateOptimalThreshold = x
        
End Function

'Convert an image to black and white (1-bit image)
Public Sub masterBlackWhiteConversion(ByVal cThreshold As Long, Optional ByVal DitherMethod As Long = 0, Optional ByVal lowColor As Long = &H0, Optional ByVal highColor As Long = &HFFFFFF, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)

    If Not toPreview Then Message "Converting image to two colors..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    
    prepImageData tmpSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, i As Long, j As Long
    Dim initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
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
    Dim r As Long, g As Long, b As Long
    Dim l As Long, newL As Long
    Dim xModQuick As Long
    Dim DitherTable() As Byte
    Dim xLeft As Long, xRight As Long, yDown As Long
    Dim errorVal As Double
    Dim dDivisor As Double
    
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
                If Not toPreview Then
                    If (x And progBarCheck) = 0 Then
                        If userPressedESC() Then Exit For
                        SetProgBarVal x
                    End If
                End If
            Next x
            
            
        'Ordered dither (Bayer 4x4).  Unfortunately, this routine requires a unique set of code owing to its specialized
        ' implementation. Coefficients derived from http://en.wikipedia.org/wiki/Ordered_dithering
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
                If Not toPreview Then
                    If (x And progBarCheck) = 0 Then
                        If userPressedESC() Then Exit For
                        SetProgBarVal x
                    End If
                End If
            Next x

        'Ordered dither (Bayer 8x8).  Unfortunately, this routine requires a unique set of code owing to its specialized
        ' implementation. Coefficients derived from http://en.wikipedia.org/wiki/Ordered_dithering
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
                If Not toPreview Then
                    If (x And progBarCheck) = 0 Then
                        If userPressedESC() Then Exit For
                        SetProgBarVal x
                    End If
                End If
            Next x
        
        'False Floyd-Steinberg.  Coefficients derived from http://www.efg2.com/Lab/Library/ImageProcessing/DHALF.TXT
        Case 3
        
            'First, prepare a dither table
            ReDim DitherTable(0 To 1, 0 To 1) As Byte
            
            DitherTable(1, 0) = 3
            DitherTable(0, 1) = 3
            DitherTable(1, 1) = 2
            
            dDivisor = 8
        
            'Second, mark the size of the array in the left, right, and down directions
            xLeft = 0
            xRight = 1
            yDown = 1
            
        'Genuine Floyd-Steinberg.  Coefficients derived from http://www.efg2.com/Lab/Library/ImageProcessing/DHALF.TXT
        Case 4
        
            'First, prepare a Floyd-Steinberg dither table
            ReDim DitherTable(-1 To 1, 0 To 1) As Byte
            
            DitherTable(1, 0) = 7
            DitherTable(-1, 1) = 3
            DitherTable(0, 1) = 5
            DitherTable(1, 1) = 1
            
            dDivisor = 16
        
            'Second, mark the size of the array in the left, right, and down directions
            xLeft = -1
            xRight = 1
            yDown = 1
            
        'Jarvis, Judice, Ninke.  Coefficients derived from http://www.efg2.com/Lab/Library/ImageProcessing/DHALF.TXT
        Case 5
        
            'First, prepare a dither table
            ReDim DitherTable(-2 To 2, 0 To 2) As Byte
            
            DitherTable(1, 0) = 7
            DitherTable(2, 0) = 5
            
            DitherTable(-2, 1) = 3
            DitherTable(-1, 1) = 5
            DitherTable(0, 1) = 7
            DitherTable(1, 1) = 5
            DitherTable(2, 1) = 3
            
            DitherTable(-2, 2) = 1
            DitherTable(-1, 2) = 3
            DitherTable(0, 2) = 5
            DitherTable(1, 2) = 3
            DitherTable(2, 2) = 1
            
            dDivisor = 48
        
            'Second, mark the size of the array in the left, right, and down directions
            xLeft = -2
            xRight = 2
            yDown = 2
            
        'Stucki.  Coefficients derived from http://www.efg2.com/Lab/Library/ImageProcessing/DHALF.TXT
        Case 6
        
            'First, prepare a dither table
            ReDim DitherTable(-2 To 2, 0 To 2) As Byte
            
            DitherTable(1, 0) = 8
            DitherTable(2, 0) = 4
            
            DitherTable(-2, 1) = 2
            DitherTable(-1, 1) = 4
            DitherTable(0, 1) = 8
            DitherTable(1, 1) = 4
            DitherTable(2, 1) = 2
            
            DitherTable(-2, 2) = 1
            DitherTable(-1, 2) = 2
            DitherTable(0, 2) = 4
            DitherTable(1, 2) = 2
            DitherTable(2, 2) = 1
            
            dDivisor = 42
        
            'Second, mark the size of the array in the left, right, and down directions
            xLeft = -2
            xRight = 2
            yDown = 2
            
        'Burkes.  Coefficients derived from http://www.efg2.com/Lab/Library/ImageProcessing/DHALF.TXT
        Case 7
        
            'First, prepare a dither table
            ReDim DitherTable(-2 To 2, 0 To 1) As Byte
            
            DitherTable(1, 0) = 8
            DitherTable(2, 0) = 4
            
            DitherTable(-2, 1) = 2
            DitherTable(-1, 1) = 4
            DitherTable(0, 1) = 8
            DitherTable(1, 1) = 4
            DitherTable(2, 1) = 2
            
            dDivisor = 32
        
            'Second, mark the size of the array in the left, right, and down directions
            xLeft = -2
            xRight = 2
            yDown = 1
            
        'Sierra-3.  Coefficients derived from http://www.efg2.com/Lab/Library/ImageProcessing/DHALF.TXT
        Case 8
        
            'First, prepare a dither table
            ReDim DitherTable(-2 To 2, 0 To 2) As Byte
            
            DitherTable(1, 0) = 5
            DitherTable(2, 0) = 3
            
            DitherTable(-2, 1) = 2
            DitherTable(-1, 1) = 4
            DitherTable(0, 1) = 5
            DitherTable(1, 1) = 4
            DitherTable(2, 1) = 2
            
            DitherTable(-2, 2) = 0
            DitherTable(-1, 2) = 2
            DitherTable(0, 2) = 3
            DitherTable(1, 2) = 2
            DitherTable(2, 2) = 0
            
            dDivisor = 32
        
            'Second, mark the size of the array in the left, right, and down directions
            xLeft = -2
            xRight = 2
            yDown = 2
            
        'Sierra-2.  Coefficients derived from http://www.efg2.com/Lab/Library/ImageProcessing/DHALF.TXT
        Case 9
        
            'First, prepare a dither table
            ReDim DitherTable(-2 To 2, 0 To 1) As Byte
            
            DitherTable(1, 0) = 4
            DitherTable(2, 0) = 3
            
            DitherTable(-2, 1) = 1
            DitherTable(-1, 1) = 2
            DitherTable(0, 1) = 3
            DitherTable(1, 1) = 2
            DitherTable(2, 1) = 1
            
            dDivisor = 16
        
            'Second, mark the size of the array in the left, right, and down directions
            xLeft = -2
            xRight = 2
            yDown = 1
            
        'Sierra-2-4A.  Coefficients derived from http://www.efg2.com/Lab/Library/ImageProcessing/DHALF.TXT
        Case 10
        
            'First, prepare a dither table
            ReDim DitherTable(-1 To 1, 0 To 1) As Byte
            
            DitherTable(1, 0) = 2

            DitherTable(-1, 1) = 1
            DitherTable(0, 1) = 1
            
            dDivisor = 4
        
            'Second, mark the size of the array in the left, right, and down directions
            xLeft = -1
            xRight = 1
            yDown = 1
            
        'Bill Atkinson's original Hyperdither/HyperScan algorithm.  (Note: Bill invented MacPaint, QuickDraw, and HyperCard.)
        ' This is the dithering algorithm used on the original Apple Macintosh.
        ' Coefficients derived from http://gazs.github.com/canvas-atkinson-dither/
        Case 11
        
            'First, prepare a dither table
            ReDim DitherTable(-1 To 2, 0 To 2) As Byte
            
            DitherTable(1, 0) = 1
            DitherTable(2, 0) = 1
            
            DitherTable(-1, 1) = 1
            DitherTable(0, 1) = 1
            DitherTable(1, 1) = 1
            
            DitherTable(0, 2) = 1
            
            dDivisor = 8
        
            'Second, mark the size of the array in the left, right, and down directions
            xLeft = -1
            xRight = 2
            yDown = 2
            
    End Select
    
    'If we have been asked to use a non-ordered dithering method, apply it now
    If DitherMethod >= 3 Then
    
        'First, we need a dithering table the same size as the image.  We make it of Single type to prevent rounding errors.
        ' (This uses a lot of memory, but on modern systems it shouldn't be a problem.)
        Dim dErrors() As Double
        
        ReDim dErrors(0 To workingDIB.getDIBWidth, 0 To workingDIB.getDIBHeight) As Double
        
        Dim QuickX As Long, QuickY As Long
        
        'Now loop through the image, calculating errors as we go
        For x = initX To finalX
            QuickVal = x * qvDepth
        For y = initY To finalY
        
            'Get the source pixel color values
            r = ImageData(QuickVal + 2, y)
            g = ImageData(QuickVal + 1, y)
            b = ImageData(QuickVal, y)
            
            'Convert those to a luminance value and add the value of the error at this location
            l = getLuminance(r, g, b)
            newL = l + dErrors(x, y)
            
            'Check our modified luminance value against the threshold, and set new values accordingly
            If newL >= cThreshold Then
                errorVal = newL - 255
                ImageData(QuickVal + 2, y) = highR
                ImageData(QuickVal + 1, y) = highG
                ImageData(QuickVal, y) = highB
            Else
                errorVal = newL
                ImageData(QuickVal + 2, y) = lowR
                ImageData(QuickVal + 1, y) = lowG
                ImageData(QuickVal, y) = lowB
            End If
            
            'If there is an error, spread it
            If errorVal <> 0 Then
            
                'Now, spread that error across the relevant pixels according to the dither table formula
                For i = xLeft To xRight
                For j = 0 To yDown
                
                    'First, ignore already processed pixels
                    If (j = 0) And (i <= 0) Then GoTo NextDitheredPixel
                    
                    'Second, ignore pixels that have a zero in the dither table
                    If DitherTable(i, j) = 0 Then GoTo NextDitheredPixel
                    
                    QuickX = x + i
                    QuickY = y + j
                    
                    'Next, ignore target pixels that are off the image boundary
                    If QuickX < initX Then GoTo NextDitheredPixel
                    If QuickX > finalX Then GoTo NextDitheredPixel
                    If QuickY > finalY Then GoTo NextDitheredPixel
                    
                    'If we've made it all the way here, we are able to actually spread the error to this location
                    dErrors(QuickX, QuickY) = dErrors(QuickX, QuickY) + (errorVal * (CSng(DitherTable(i, j)) / dDivisor))
                
NextDitheredPixel:     Next j
                Next i
            
            End If
                
        Next y
            If Not toPreview Then
                If (x And progBarCheck) = 0 Then
                    If userPressedESC() Then Exit For
                    SetProgBarVal x
                End If
            End If
        Next x
    
    End If
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic

End Sub

Private Sub sltThreshold_Change()
    If CBool(chkAutoThreshold.Value) Then chkAutoThreshold.Value = vbUnchecked
    updatePreview
End Sub

Private Sub updatePreview()
    If cmdBar.previewsAllowed Then masterBlackWhiteConversion sltThreshold, cboDither.ListIndex, colorPicker(0).Color, colorPicker(1).Color, True, fxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub


