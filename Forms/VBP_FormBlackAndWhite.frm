VERSION 5.00
Begin VB.Form FormBlackAndWhite 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Black/White Color Conversion"
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
   StartUpPosition =   1  'CenterOwner
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
   Begin VB.PictureBox picBWColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   9000
      ScaleHeight     =   465
      ScaleWidth      =   2745
      TabIndex        =   11
      Top             =   3360
      Width           =   2775
   End
   Begin VB.PictureBox picBWColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   6120
      ScaleHeight     =   465
      ScaleWidth      =   2745
      TabIndex        =   10
      Top             =   3360
      Width           =   2775
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
      TabIndex        =   5
      Top             =   2400
      Width           =   4935
   End
   Begin VB.HScrollBar hsThreshold 
      Height          =   255
      Left            =   6120
      Max             =   254
      Min             =   1
      TabIndex        =   2
      Top             =   960
      Value           =   127
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
      Left            =   11160
      MaxLength       =   3
      TabIndex        =   3
      Text            =   "127"
      Top             =   930
      Width           =   660
   End
   Begin VB.CheckBox chkAutoThreshold 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   " have PhotoDemon estimate the ideal threshold for this image"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   360
      Left            =   6120
      TabIndex        =   4
      Top             =   1440
      Width           =   5655
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin VB.Label lblBackground 
      Height          =   855
      Left            =   0
      TabIndex        =   12
      Top             =   5760
      Width           =   12255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "colors to use (click boxes to change):"
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
      TabIndex        =   9
      Top             =   3000
      Width           =   3945
   End
   Begin VB.Label lblDither 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "dithering technique:"
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
      TabIndex        =   8
      Top             =   2040
      Width           =   2130
   End
   Begin VB.Label lblBWDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   $"VBP_FormBlackAndWhite.frx":0000
      ForeColor       =   &H00404040&
      Height          =   735
      Left            =   6000
      TabIndex        =   7
      Top             =   4200
      Width           =   5775
   End
   Begin VB.Label lblHeader 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "threshold:"
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
      Top             =   600
      Width           =   1080
   End
End
Attribute VB_Name = "FormBlackAndWhite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Black/White Color Reduction Form
'Copyright ©2000-2013 by Tanner Helland
'Created: some time 2002
'Last updated: 28/December/12
'Last update: allow optimal threshold calculation for all dithering types
'
'The meat of this form is in the module with the same name...look there for
' real algorithm info.
'
'***************************************************************************

Option Explicit

Private Sub cboDither_Click()
    If CBool(chkAutoThreshold.Value) Then txtThreshold.Text = calculateOptimalThreshold(cboDither.ListIndex)
    masterBlackWhiteConversion txtThreshold, cboDither.ListIndex, picBWColor(0).backColor, picBWColor(1).backColor, True, fxPreview
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
    
    masterBlackWhiteConversion txtThreshold, cboDither.ListIndex, picBWColor(0).backColor, picBWColor(1).backColor, True, fxPreview
    
End Sub

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub cmdOK_Click()

    'Checking the threshold value to make sure it's valid
    If EntryValid(txtThreshold, hsThreshold.Min, hsThreshold.Max) = False Then
        AutoSelectText txtThreshold
        Exit Sub
    End If
    
    Me.Visible = False
    
    Process BWMaster, txtThreshold, cboDither.ListIndex, picBWColor(0).backColor, picBWColor(1).backColor
    
    Unload Me
    
End Sub

Private Sub Form_Activate()
  
    'Populate the dither combobox
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
    cboDither.ListIndex = 11
    DoEvents
        
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
    setHandCursor picBWColor(0)
    setHandCursor picBWColor(1)
    
    'Draw the preview
    masterBlackWhiteConversion txtThreshold, cboDither.ListIndex, picBWColor(0).backColor, picBWColor(1).backColor, True, fxPreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub hsThreshold_Change()
    copyToTextBoxI txtThreshold, hsThreshold.Value
    masterBlackWhiteConversion txtThreshold, cboDither.ListIndex, picBWColor(0).backColor, picBWColor(1).backColor, True, fxPreview
End Sub

Private Sub hsThreshold_Scroll()
    chkAutoThreshold.Value = vbUnchecked
    copyToTextBoxI txtThreshold, hsThreshold.Value
    masterBlackWhiteConversion txtThreshold, cboDither.ListIndex, picBWColor(0).backColor, picBWColor(1).backColor, True, fxPreview
End Sub

'Allow the user to select a custom color for each
Private Sub picBWColor_Click(Index As Integer)
    
    'Use a common dialog box to select a new color.  (In the future, perhaps I'll design a better custom box.)
    Dim newColor As Long
    Dim comDlg As cCommonDialog
    Set comDlg = New cCommonDialog
    newColor = picBWColor(Index).backColor
    
    If comDlg.VBChooseColor(newColor, True, True, False, Me.hWnd) Then
        picBWColor(Index).backColor = newColor
        masterBlackWhiteConversion txtThreshold, cboDither.ListIndex, picBWColor(0).backColor, picBWColor(1).backColor, True, fxPreview
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
    
    calculateOptimalThreshold = x
        
End Function

'Convert an image to black and white (1-bit image)
Public Sub masterBlackWhiteConversion(ByVal cThreshold As Long, Optional ByVal DitherMethod As Long = 0, Optional ByVal lowColor As Long = &H0, Optional ByVal highColor As Long = &HFFFFFF, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)

    If toPreview = False Then Message "Converting image to two colors..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    
    prepImageData tmpSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, i As Long, j As Long
    Dim initX As Long, initY As Long, finalX As Long, finalY As Long
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
                If toPreview = False Then
                    If (x And progBarCheck) = 0 Then SetProgBarVal x
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
                If toPreview = False Then
                    If (x And progBarCheck) = 0 Then SetProgBarVal x
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
                If toPreview = False Then
                    If (x And progBarCheck) = 0 Then SetProgBarVal x
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
        
        ReDim dErrors(0 To workingLayer.getLayerWidth, 0 To workingLayer.getLayerHeight) As Double
        
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
            If toPreview = False Then
                If (x And progBarCheck) = 0 Then SetProgBarVal x
            End If
        Next x
    
    End If
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic

End Sub

Private Sub txtThreshold_KeyUp(KeyCode As Integer, Shift As Integer)
    
    chkAutoThreshold.Value = vbUnchecked
    textValidate txtThreshold
    If EntryValid(txtThreshold, hsThreshold.Min, hsThreshold.Max, False, False) Then hsThreshold.Value = Val(txtThreshold)
        
End Sub
