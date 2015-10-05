VERSION 5.00
Begin VB.Form FormGrayscale 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Black and white"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11895
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
   ScaleHeight     =   437
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   793
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.ComboBox cboDithering 
      Appearance      =   0  'Flat
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
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   4080
      Width           =   5535
   End
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5805
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   1323
      BackColor       =   14802140
   End
   Begin PhotoDemon.sliderTextCombo sltShades 
      Height          =   720
      Left            =   6000
      TabIndex        =   11
      Top             =   2700
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   1270
      Caption         =   "number of gray shades"
      Min             =   2
      Max             =   256
      Value           =   4
      NotchPosition   =   2
      NotchValueCustom=   256
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
   Begin VB.ComboBox cboMethod 
      Appearance      =   0  'Flat
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
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1560
      Width           =   5535
   End
   Begin VB.PictureBox picDecompose 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6120
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   377
      TabIndex        =   3
      Top             =   2040
      Width           =   5655
      Begin PhotoDemon.smartOptionButton optDecompose 
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   0
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   582
         Caption         =   "minimum"
         Value           =   -1  'True
      End
      Begin PhotoDemon.smartOptionButton optDecompose 
         Height          =   360
         Index           =   1
         Left            =   2400
         TabIndex        =   7
         Top             =   0
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   582
         Caption         =   "maximum"
      End
   End
   Begin VB.PictureBox picChannel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6120
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   369
      TabIndex        =   4
      Top             =   2040
      Width           =   5535
      Begin PhotoDemon.smartOptionButton optChannel 
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   0
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   582
         Caption         =   "red"
         Value           =   -1  'True
      End
      Begin PhotoDemon.smartOptionButton optChannel 
         Height          =   360
         Index           =   1
         Left            =   1680
         TabIndex        =   9
         Top             =   0
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   582
         Caption         =   "green"
      End
      Begin PhotoDemon.smartOptionButton optChannel 
         Height          =   360
         Index           =   2
         Left            =   3360
         TabIndex        =   10
         Top             =   0
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   582
         Caption         =   "blue"
      End
   End
   Begin VB.Label lblDithering 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "dithering options"
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
      TabIndex        =   12
      Top             =   3720
      Width           =   1800
   End
   Begin VB.Label lblAlgorithm 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "style"
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
      Top             =   1200
      Width           =   480
   End
End
Attribute VB_Name = "FormGrayscale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Grayscale Conversion Handler
'Copyright 2002-2015 by Tanner Helland
'Created: 1/12/02
'Last updated: 16/February/14
'Last update: rewrite code so that all conversion methods provide an option for specific # of gray shades and/or dithering,
'              and implement a full-power dithering engine with all known diffusion dithering algorithms.
'
'Updated version of the grayscale handler; utilizes five different methods (average, ISU, desaturate, max/min decomposition,
' single color channel) with the option for variable # of gray shades with/without dithering for all available methods. A
' comprehensive dithering list is also available for all methods, should the user desire it.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Preview the current grayscale conversion technique
Private Sub updatePreview()
    If cmdBar.previewsAllowed Then masterGrayscaleFunction cboMethod.ListIndex, getExtraGrayscaleParams(cboMethod.ListIndex), sltShades.Value, cboDithering.ListIndex, True, fxPreview
End Sub

Private Sub cboDithering_Click()
    updatePreview
End Sub

Private Sub cboMethod_Click()
    UpdateVisibleControls
    updatePreview
End Sub

'Certain algorithms require additional user input.  This routine enables/disables the controls associated with a given algorithm.
Private Sub UpdateVisibleControls()
    
    Select Case cboMethod.ListIndex
        Case 3
            picDecompose.Visible = True
            picChannel.Visible = False
        Case 4
            picDecompose.Visible = False
            picChannel.Visible = True
        Case Else
            picDecompose.Visible = False
            picChannel.Visible = False
    End Select
    
End Sub

Private Sub chkDither_Click()
    updatePreview
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Black and white", False, buildParams(cboMethod.ListIndex, getExtraGrayscaleParams(cboMethod.ListIndex), sltShades.Value, cboDithering.ListIndex), UNDO_LAYER
End Sub

'Some grayscale functions require extra parameters.  Some do not.  Call this function to retrieve any extra parameters
' for a given grayscale conversion method.
Private Function getExtraGrayscaleParams(ByVal grayscaleMethod As Long) As String

    Select Case grayscaleMethod
        
        Case 0
            getExtraGrayscaleParams = " "
            
        Case 1
            getExtraGrayscaleParams = " "
            
        Case 2
            getExtraGrayscaleParams = " "
            
        Case 3
            If optDecompose(0).Value Then
                getExtraGrayscaleParams = "0"
            Else
                getExtraGrayscaleParams = "1"
            End If
            
        Case 4
            If optChannel(0).Value Then
                getExtraGrayscaleParams = "0"
            ElseIf optChannel(1).Value Then
                getExtraGrayscaleParams = "1"
            Else
                getExtraGrayscaleParams = "2"
            End If
            
        Case 5
            getExtraGrayscaleParams = Str(sltShades.Value)
            
        Case 6
            getExtraGrayscaleParams = Str(sltShades.Value)
            
    End Select

End Function

Private Sub cmdBar_RequestPreviewUpdate()
    UpdateVisibleControls
    updatePreview
End Sub

'Recommend ITU grayscale correction by default, and max shades without dithering
Private Sub cmdBar_ResetClick()
    cboMethod.ListIndex = 1
    sltShades.Value = 256
End Sub

Private Sub Form_Activate()
    
    'Apply translations and visual themes
    MakeFormPretty Me
    setArrowCursor picChannel
    setArrowCursor picDecompose
    
    'Draw the initial preview
    cmdBar.markPreviewStatus True
    updatePreview
    
End Sub


'All different grayscale (black and white) routines are handled by this single function.  As of 16 Feb '14, grayscale operations
' are divided into four params: type of transform, optional params for transform (if any), number of shades to use, and
' dithering options (if any).  This should allow the user to mix and match the various options at their leisure.
Public Sub masterGrayscaleFunction(Optional ByVal grayscaleMethod As Long, Optional ByVal additionalParams As String, Optional ByVal numOfShades As Long = 256, Optional ByVal ditheringOptions As Long = 0, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)

    If Not toPreview Then Message "Converting image to black and white..."

    'For backward compatibility reasons, check additional params and make sure its length is at least 1.
    If Len(additionalParams) = 0 Then additionalParams = " "

    'Use a parameter parse string to extract any additional parameters.
    Dim cParams As pdParamString
    Set cParams = New pdParamString
    cParams.setParamString additionalParams
    
    'Create a working copy of the relevant pixel data (with all selection transforms applied)
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    
    'Based on the options the user has provided, figure out a maximum progress bar value.  This changes depending on:
    ' - If the user wants shade reduction (as this requires another pass over the image)
    ' - If the user wants dithering (as the second pass will be done horizontally instead of vertically)
    Dim progBarMax As Long
    If numOfShades < 255 Then
    
        If ditheringOptions > 0 Then
            progBarMax = workingDIB.getDIBWidth + workingDIB.getDIBHeight
        Else
            progBarMax = workingDIB.getDIBWidth * 2
        End If
    
    Else
        progBarMax = workingDIB.getDIBWidth
    End If
    
    Dim userCanceled As Long
    
    'Different grayscale conversion methods call different individual subs
    Select Case grayscaleMethod
        
        Case 0
            userCanceled = MenuGrayscaleAverage(workingDIB, toPreview, progBarMax)
            
        Case 1
            userCanceled = MenuGrayscale(workingDIB, toPreview, progBarMax)
            
        Case 2
            userCanceled = MenuDesaturate(workingDIB, toPreview, progBarMax)
            
        Case 3
            userCanceled = MenuDecompose(cParams.GetLong(1), workingDIB, toPreview, progBarMax)
            
        Case 4
            userCanceled = MenuGrayscaleSingleChannel(cParams.GetLong(1), workingDIB, toPreview, progBarMax)
        
        'Options 5 and 6 correspond to "specific # of shades" and "specific # of shades with dithering" in old builds.
        ' To retain backwards compatibility for these options, we use a standard grayscale conversion, but with shade
        ' reduction and/or dithering enabled
        Case 5
            userCanceled = MenuGrayscale(workingDIB, toPreview, progBarMax)
            numOfShades = cParams.GetLong(1)
            ditheringOptions = 0
            
        Case 6
            userCanceled = MenuGrayscale(workingDIB, toPreview, progBarMax)
            numOfShades = cParams.GetLong(1)
            ditheringOptions = 1
            
    End Select
    
    'We now apply the user's choice of shade reduction and/or dithering.
    If (numOfShades < 255) And (userCanceled <> 0) Then
        
        Select Case ditheringOptions
        
            Case 0
                fGrayscaleCustom numOfShades, workingDIB, toPreview, progBarMax, workingDIB.getDIBWidth
                
            Case Else
                fGrayscaleCustomDither numOfShades, ditheringOptions, workingDIB, toPreview, progBarMax, workingDIB.getDIBWidth
            
        End Select
        
    End If
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingDIB
    finalizeImageData toPreview, dstPic
    
End Sub

'Reduce to X # gray shades
Public Function fGrayscaleCustom(ByVal numOfShades As Long, ByRef srcDIB As pdDIB, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long

    'Point an array at the source DIB's image data
    Dim ImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    prepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(srcSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.getDIBWidth - 1
    finalY = srcDIB.getDIBHeight - 1
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = srcDIB.getDIBColorDepth \ 8
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not suppressMessages Then
        If modifyProgBarMax = -1 Then
            SetProgBarMax finalX
        Else
            SetProgBarMax modifyProgBarMax
        End If
        progBarCheck = findBestProgBarValue()
    End If
    
    'Color variables
    Dim r As Long, g As Long, b As Long, grayVal As Long
    
    'This conversion factor is the value we need to turn grayscale values in the [0,255] range into a specific subset of values
    Dim conversionFactor As Double
    conversionFactor = (255 / (numOfShades - 1))
    
    'Build a look-up table for our custom grayscale conversion results
    Dim LookUp(0 To 255) As Byte
    
    For x = 0 To 255
        grayVal = Int((CDbl(x) / conversionFactor) + 0.5) * conversionFactor
        If grayVal > 255 Then grayVal = 255
        LookUp(x) = CByte(grayVal)
    Next x
    
    'Build another look-up table for our initial grayscale index calculation
    Dim grayLookUp(0 To 765) As Byte
    For x = 0 To 765
        grayLookUp(x) = x \ 3
    Next x
        
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        grayVal = grayLookUp(r + g + b)
        
        'Assign all color channels the new gray value
        ImageData(QuickVal + 2, y) = LookUp(grayVal)
        ImageData(QuickVal + 1, y) = LookUp(grayVal)
        ImageData(QuickVal, y) = LookUp(grayVal)
        
    Next y
        If Not suppressMessages Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x + modifyProgBarOffset
            End If
        End If
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    If cancelCurrentAction Then fGrayscaleCustom = 0 Else fGrayscaleCustom = 1
    
End Function

'Reduce to X # gray shades (dithered)
Public Function fGrayscaleCustomDither(ByVal numOfShades As Long, ByVal DitherMethod As Long, ByRef srcDIB As pdDIB, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long

    'Point an array at the source DIB's image data
    Dim ImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    prepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(srcSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.getDIBWidth - 1
    finalY = srcDIB.getDIBHeight - 1
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = srcDIB.getDIBColorDepth \ 8
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not suppressMessages Then
        If modifyProgBarMax = -1 Then
            SetProgBarMax finalY
        Else
            SetProgBarMax modifyProgBarMax
        End If
        progBarCheck = findBestProgBarValue()
    End If
    
    'Color variables
    Dim g As Long, grayVal As Long
    
    'This conversion factor is the value we need to turn grayscale values in the [0,255] range into a specific subset of values
    Dim conversionFactor As Double
    conversionFactor = (255 / (numOfShades - 1))
    
    'Build a look-up table for our custom grayscale conversion results
    Dim LookUp(0 To 255) As Long
    
    For x = 0 To 255
        grayVal = Int((CDbl(x) / conversionFactor) + 0.5) * conversionFactor
        If grayVal > 255 Then grayVal = 255
        LookUp(x) = grayVal
    Next x
    
    Dim DitherTable() As Byte
    Dim xLeft As Long, xRight As Long, yDown As Long
    Dim errorVal As Double
    Dim dDivisor As Double
    Dim l As Long, newL As Long
    
    'Start by preparing a dithering table, which is obviously dependent on the requested dithering method
    Select Case DitherMethod
        
        'No dithering
        Case 0
        
        'False Floyd-Steinberg.  Coefficients derived from http://www.efg2.com/Lab/Library/ImageProcessing/DHALF.TXT
        Case 1
        
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
        Case 2
        
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
        Case 3
        
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
        Case 4
        
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
        Case 5
        
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
        Case 6
        
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
        Case 7
        
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
        Case 8
        
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
        Case 9
        
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
    
    
    'With out dithering table complete, we can now proceed to process the image
    If DitherMethod > 0 Then
    
        'First, we need a dithering table the same size as the image.  We make it of Single type to prevent rounding errors.
        ' (This uses a lot of memory, but on modern systems it shouldn't be a problem.)
        Dim dErrors() As Single
        ReDim dErrors(0 To workingDIB.getDIBWidth, 0 To workingDIB.getDIBHeight) As Single
        
        Dim i As Long, j As Long
        
        Dim QuickX As Long, QuickY As Long
        
        'Now loop through the image, calculating errors as we go
        For y = initY To finalY
        For x = initX To finalX
        
            QuickVal = x * qvDepth
            
            'Get the source pixel color values.  Because we know the image we're handed is already going to be grayscale,
            ' we can shortcut this calculation by only grabbing the red channel.
            g = ImageData(QuickVal + 2, y)
            
            'Convert those to a luminance value and add the value of the error at this location
            l = g + dErrors(x, y)
            
            'Convert that to a lookup-table-safe luminance (e.g. 0-255)
            If l < 0 Then
                newL = 0
            ElseIf l > 255 Then
                newL = 255
            Else
                newL = l
            End If
            
            'Write the new luminance value out to the image array
            ImageData(QuickVal + 2, y) = LookUp(newL)
            ImageData(QuickVal + 1, y) = LookUp(newL)
            ImageData(QuickVal, y) = LookUp(newL)
            
            'Calculate an error for this calculation
            errorVal = l - LookUp(newL)
            
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
                
        Next x

            If Not suppressMessages Then
                If (y And progBarCheck) = 0 Then
                    If userPressedESC() Then Exit For
                    SetProgBarVal y + modifyProgBarOffset
                End If
            End If

        Next y
    
    
    
    End If
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    If cancelCurrentAction Then fGrayscaleCustomDither = 0 Else fGrayscaleCustomDither = 1
    
End Function

'Reduce to gray via (r+g+b)/3
Public Function MenuGrayscaleAverage(ByRef srcDIB As pdDIB, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long
    
    'Point an array at the source DIB's image data
    Dim ImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    prepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(srcSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.getDIBWidth - 1
    finalY = srcDIB.getDIBHeight - 1
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = srcDIB.getDIBColorDepth \ 8
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not suppressMessages Then
        If modifyProgBarMax = -1 Then
            SetProgBarMax finalX
        Else
            SetProgBarMax modifyProgBarMax
        End If
        progBarCheck = findBestProgBarValue()
    End If
    
    'Color and grayscale variables
    Dim r As Long, g As Long, b As Long
    Dim grayVal As Byte
    
    'Build a look-up table of grayscale values (faster than calculating it manually for each pixel)
    Dim grayLookUp(0 To 765) As Byte
    For x = 0 To 765
        grayLookUp(x) = x \ 3
    Next x
    
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Calculate the gray value using the look-up table
        grayVal = grayLookUp(r + g + b)
        
        'Assign that gray value to each color channel
        ImageData(QuickVal, y) = grayVal
        ImageData(QuickVal + 1, y) = grayVal
        ImageData(QuickVal + 2, y) = grayVal
        
    Next y
        If Not suppressMessages Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x + modifyProgBarOffset
            End If
        End If
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    If cancelCurrentAction Then MenuGrayscaleAverage = 0 Else MenuGrayscaleAverage = 1
    
End Function

'Reduce to gray in a more human-eye friendly manner
Public Function MenuGrayscale(ByRef srcDIB As pdDIB, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long
    
    'Point an array at the source DIB's image data
    Dim ImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    prepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(srcSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.getDIBWidth - 1
    finalY = srcDIB.getDIBHeight - 1
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = srcDIB.getDIBColorDepth \ 8
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not suppressMessages Then
        If modifyProgBarMax = -1 Then
            SetProgBarMax finalX
        Else
            SetProgBarMax modifyProgBarMax
        End If
        progBarCheck = findBestProgBarValue()
    End If
    
    'Color and grayscale variables
    Dim r As Long, g As Long, b As Long
    Dim grayVal As Long
        
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Calculate a grayscale value using the original ITU-R recommended formula (BT.709, specifically)
        grayVal = (213 * r + 715 * g + 72 * b) \ 1000
        If grayVal > 255 Then grayVal = 255
        
        'Assign that gray value to each color channel
        ImageData(QuickVal, y) = grayVal
        ImageData(QuickVal + 1, y) = grayVal
        ImageData(QuickVal + 2, y) = grayVal
        
    Next y
        If Not suppressMessages Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x + modifyProgBarOffset
            End If
        End If
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    If cancelCurrentAction Then MenuGrayscale = 0 Else MenuGrayscale = 1
    
End Function

'Reduce to gray via HSL -> convert S to 0
Public Function MenuDesaturate(ByRef srcDIB As pdDIB, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long
        
    'Point an array at the source DIB's image data
    Dim ImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    prepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(srcSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.getDIBWidth - 1
    finalY = srcDIB.getDIBHeight - 1
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = srcDIB.getDIBColorDepth \ 8
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not suppressMessages Then
        If modifyProgBarMax = -1 Then
            SetProgBarMax finalX
        Else
            SetProgBarMax modifyProgBarMax
        End If
        progBarCheck = findBestProgBarValue()
    End If
    
    'Color and grayscale variables
    Dim r As Long, g As Long, b As Long
    Dim grayVal As Byte
       
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Calculate a grayscale value by using a short-hand RGB <-> HSL conversion
        grayVal = CByte(getLuminance(r, g, b))
        
        'Assign that gray value to each color channel
        ImageData(QuickVal, y) = grayVal
        ImageData(QuickVal + 1, y) = grayVal
        ImageData(QuickVal + 2, y) = grayVal
        
    Next y
        If Not suppressMessages Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x + modifyProgBarOffset
            End If
        End If
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    If cancelCurrentAction Then MenuDesaturate = 0 Else MenuDesaturate = 1
    
End Function

'Reduce to gray by selecting the minimum (maxOrMin = 0) or maximum (maxOrMin = 1) color in each pixel
Public Function MenuDecompose(ByVal maxOrMin As Long, ByRef srcDIB As pdDIB, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long

    'Point an array at the source DIB's image data
    Dim ImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    prepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(srcSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.getDIBWidth - 1
    finalY = srcDIB.getDIBHeight - 1
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = srcDIB.getDIBColorDepth \ 8
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not suppressMessages Then
        If modifyProgBarMax = -1 Then
            SetProgBarMax finalX
        Else
            SetProgBarMax modifyProgBarMax
        End If
        progBarCheck = findBestProgBarValue()
    End If
    
    'Color and grayscale variables
    Dim r As Long, g As Long, b As Long
    Dim grayVal As Byte
        
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Find the highest or lowest of the RGB values
        If maxOrMin = 0 Then grayVal = CByte(Min3Int(r, g, b)) Else grayVal = CByte(Max3Int(r, g, b))
        
        'Assign that gray value to each color channel
        ImageData(QuickVal, y) = grayVal
        ImageData(QuickVal + 1, y) = grayVal
        ImageData(QuickVal + 2, y) = grayVal
        
    Next y
        If Not suppressMessages Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x + modifyProgBarOffset
            End If
        End If
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    If cancelCurrentAction Then MenuDecompose = 0 Else MenuDecompose = 1
    
End Function

'Reduce to gray by selecting a single color channel (represeted by cChannel: 0 = Red, 1 = Green, 2 = Blue)
Public Function MenuGrayscaleSingleChannel(ByVal cChannel As Long, ByRef srcDIB As pdDIB, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long

    'Point an array at the source DIB's image data
    Dim ImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    prepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(srcSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.getDIBWidth - 1
    finalY = srcDIB.getDIBHeight - 1
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = srcDIB.getDIBColorDepth \ 8
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not suppressMessages Then
        If modifyProgBarMax = -1 Then
            SetProgBarMax finalX
        Else
            SetProgBarMax modifyProgBarMax
        End If
        progBarCheck = findBestProgBarValue()
    End If
    
    'Color and grayscale variables
    Dim r As Long, g As Long, b As Long
    Dim grayVal As Byte
        
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Assign the gray value to a single color channel based on the value of cChannel
        Select Case cChannel
            Case 0
                grayVal = r
            Case 1
                grayVal = g
            Case 2
                grayVal = b
        End Select
        
        'Assign that gray value to each color channel
        ImageData(QuickVal, y) = grayVal
        ImageData(QuickVal + 1, y) = grayVal
        ImageData(QuickVal + 2, y) = grayVal
        
    Next y
        If Not suppressMessages Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x + modifyProgBarOffset
            End If
        End If
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    If cancelCurrentAction Then MenuGrayscaleSingleChannel = 0 Else MenuGrayscaleSingleChannel = 1
        
End Function

Private Sub Form_Load()
    
    'Suspend previews while we get the form set up
    cmdBar.markPreviewStatus False
    
    'Set up the grayscale options combo box
    cboMethod.Clear
    cboMethod.AddItem " Fastest Calculation (average value)", 0
    cboMethod.AddItem " Highest Quality (ITU Standard)", 1
    cboMethod.AddItem " Desaturate", 2
    cboMethod.AddItem " Decompose", 3
    cboMethod.AddItem " Single color channel", 4
    cboMethod.ListIndex = 1
    
    'Populate the dither combobox
    cboDithering.Clear
    cboDithering.AddItem " None", 0
    cboDithering.AddItem " False (Fast) Floyd-Steinberg", 1
    cboDithering.AddItem " Genuine Floyd-Steinberg", 2
    cboDithering.AddItem " Jarvis, Judice, and Ninke", 3
    cboDithering.AddItem " Stucki", 4
    cboDithering.AddItem " Burkes", 5
    cboDithering.AddItem " Sierra-3", 6
    cboDithering.AddItem " Two-Row Sierra", 7
    cboDithering.AddItem " Sierra Lite", 8
    cboDithering.AddItem " Atkinson / Classic Macintosh", 9
    cboDithering.ListIndex = 0
    
    'Draw an initial preview
    UpdateVisibleControls
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'When option buttons are used, update the preview accordingly
Private Sub optChannel_Click(Index As Integer)
    updatePreview
End Sub

Private Sub optDecompose_Click(Index As Integer)
    updatePreview
End Sub

Private Sub sltShades_Change()
    updatePreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub


