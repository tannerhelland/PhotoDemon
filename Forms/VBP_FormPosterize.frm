VERSION 5.00
Begin VB.Form FormPosterize 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Posterize"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11970
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
   ScaleWidth      =   798
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   11970
      _ExtentX        =   21114
      _ExtentY        =   1323
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
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.smartCheckBox chkDither 
      Height          =   540
      Left            =   6120
      TabIndex        =   2
      Top             =   3960
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   953
      Caption         =   "apply dithering"
      Value           =   1
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
   Begin PhotoDemon.smartCheckBox chkSmartColors 
      Height          =   540
      Left            =   6120
      TabIndex        =   3
      Top             =   4560
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   953
      Caption         =   "match existing colors"
      Value           =   1
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
   Begin PhotoDemon.sliderTextCombo sltRed 
      Height          =   495
      Left            =   6000
      TabIndex        =   4
      Top             =   1440
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Min             =   2
      Max             =   64
      Value           =   6
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
   Begin PhotoDemon.sliderTextCombo sltGreen 
      Height          =   495
      Left            =   6000
      TabIndex        =   5
      Top             =   2400
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Min             =   2
      Max             =   64
      Value           =   7
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
   Begin PhotoDemon.sliderTextCombo sltBlue 
      Height          =   495
      Left            =   6000
      TabIndex        =   6
      Top             =   3360
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Min             =   2
      Max             =   64
      Value           =   6
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
   Begin VB.Label lblRedValues 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "possible red values:"
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
      Top             =   1080
      Width           =   2100
   End
   Begin VB.Label lblGreenValues 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "possible green values:"
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
      Width           =   2355
   End
   Begin VB.Label lblBlueValues 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "possible blue values:"
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
      Top             =   2985
      Width           =   2205
   End
End
Attribute VB_Name = "FormPosterize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Posterizing Effect Handler
'Copyright ©2001-2014 by Tanner Helland
'Created: 4/15/01
'Last updated: 24/August/13
'Last update: completely removed the old posterize code in favor of the per-channel approach (which was taken from
'              the old Reduce Image Colors/Indexed Color dialog)
'
'Advanced posterizing interface; it has been optimized for speed and ease-of-implementation.  It offers many more
' options than a traditional posterize dialog, which should make it more useful for achieving a desired look.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

Private Sub chkDither_Click()
    updatePreview
End Sub

Private Sub chkSmartColors_Click()
    updatePreview
End Sub

Private Sub cmdBar_OKClick()
    
    If CBool(chkDither) Then
        Process "Posterize (dithered)", , buildParams(sltRed, sltGreen, sltBlue, CBool(chkSmartColors.Value)), UNDO_LAYER
    Else
        Process "Posterize", , buildParams(sltRed, sltGreen, sltBlue, CBool(chkSmartColors.Value)), UNDO_LAYER
    End If
    
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    chkDither.Value = vbUnchecked
    chkSmartColors.Value = vbUnchecked
End Sub

Private Sub Form_Activate()
        
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
    'Create a preview
    updatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub sltBits_Change()
    updatePreview
End Sub

Private Sub updatePreview()
    
    If cmdBar.previewsAllowed Then
        If CBool(chkDither) Then
            ReduceImageColors_BitRGB_ErrorDif sltRed, sltGreen, sltBlue, CBool(chkSmartColors.Value), True, fxPreview
        Else
            ReduceImageColors_BitRGB sltRed, sltGreen, sltBlue, CBool(chkSmartColors.Value), True, fxPreview
        End If
    End If
    
End Sub

'Bit RGB color reduction (no error diffusion)
Public Sub ReduceImageColors_BitRGB(ByVal rValue As Byte, ByVal gValue As Byte, ByVal bValue As Byte, Optional ByVal smartColors As Boolean = False, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)

    If Not toPreview Then Message "Posterizing image..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    
    prepImageData tmpSA, toPreview, dstPic
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
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Color variables
    Dim r As Long, g As Long, b As Long
    Dim mR As Double, mG As Double, mB As Double
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
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x
            End If
        End If
    Next x
    
    'Intelligent Coloring requires extra work.  Perform a second loop through the image, replacing values with their
    ' computed counterparts.
    If smartColors And (Not cancelCurrentAction) Then
    
        If Not toPreview Then
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
Public Sub ReduceImageColors_BitRGB_ErrorDif(ByVal rValue As Byte, ByVal gValue As Byte, ByVal bValue As Byte, Optional ByVal smartColors As Boolean = False, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If Not toPreview Then Message "Posterizing image (with advanced dithering)..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    
    prepImageData tmpSA, toPreview, dstPic
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
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not toPreview Then
        SetProgBarMax finalY
        progBarCheck = findBestProgBarValue()
    End If
    
    'Color variables
    Dim r As Long, g As Long, b As Long
    Dim cR As Long, cG As Long, cb As Long
    Dim iR As Long, iG As Long, iB As Long
    Dim mR As Double, mG As Double, mB As Double
    Dim Er As Double, eG As Double, eB As Double
    
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
        
        r = iR + Er
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
        Er = iR - cR
        eG = iG - cG
        eB = iB - cb
        
        'Diffuse the error further (in a grid pattern) to prevent undesirable lining effects
        If (x + y) And 3 <> 0 Then
            Er = Er \ 2
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
        If Not smartColors Then
            ImageData(QuickVal + 2, y) = cR
            ImageData(QuickVal + 1, y) = cG
            ImageData(QuickVal, y) = cb
        End If
        
    Next x
        
        'At the end of each line, reset our error-tracking values
        Er = 0
        eG = 0
        eB = 0
        
        If toPreview = False Then
            If (y And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal y
            End If
        End If
    Next y
    
    'Intelligent Coloring requires extra work.  Perform a second loop through the image, replacing values with their
    ' computed counterparts.
    If smartColors And (Not cancelCurrentAction) Then
        
        If Not toPreview Then
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
            
            r = iR + Er
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
        
            Er = iR - cR
            eG = iG - cG
            eB = iB - cb
            
            'Diffuse the error further (in a grid pattern) to prevent undesirable lining effects
            If (x + y) And 3 <> 0 Then
                Er = Er \ 2
                eG = eG \ 2
                eB = eB \ 2
            End If
            
        Next x
        
            'At the end of each line, reset our error-tracking values
            Er = 0
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

Private Sub sltBlue_Change()
    updatePreview
End Sub

Private Sub sltGreen_Change()
    updatePreview
End Sub

Private Sub sltRed_Change()
    updatePreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

