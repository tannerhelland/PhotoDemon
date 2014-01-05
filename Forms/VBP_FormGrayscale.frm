VERSION 5.00
Begin VB.Form FormGrayscale 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Black and White"
   ClientHeight    =   6540
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
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   793
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   11895
      _ExtentX        =   20981
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
   Begin PhotoDemon.sliderTextCombo sltShades 
      Height          =   495
      Left            =   6000
      TabIndex        =   12
      Top             =   3120
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   873
      Min             =   2
      Max             =   254
      Value           =   4
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
      TabIndex        =   6
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
      Top             =   2040
      Width           =   5535
   End
   Begin VB.PictureBox picChannel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   6120
      ScaleHeight     =   855
      ScaleWidth      =   4935
      TabIndex        =   5
      Top             =   3240
      Width           =   4935
      Begin PhotoDemon.smartOptionButton optChannel 
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   0
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   635
         Caption         =   "red"
         Value           =   -1  'True
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
      Begin PhotoDemon.smartOptionButton optChannel 
         Height          =   360
         Index           =   1
         Left            =   1560
         TabIndex        =   10
         Top             =   0
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   635
         Caption         =   "green"
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
      Begin PhotoDemon.smartOptionButton optChannel 
         Height          =   360
         Index           =   2
         Left            =   3240
         TabIndex        =   11
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   635
         Caption         =   "blue"
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
   End
   Begin VB.PictureBox picDecompose 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   6120
      ScaleHeight     =   735
      ScaleWidth      =   4815
      TabIndex        =   4
      Top             =   3240
      Width           =   4815
      Begin PhotoDemon.smartOptionButton optDecompose 
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   0
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   635
         Caption         =   "minimum"
         Value           =   -1  'True
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
      Begin PhotoDemon.smartOptionButton optDecompose 
         Height          =   360
         Index           =   1
         Left            =   2160
         TabIndex        =   8
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         Caption         =   "maximum"
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
   End
   Begin VB.Label lblAdditional 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "additional options:"
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
      Top             =   2760
      Width           =   1980
   End
   Begin VB.Label lblAlgorithm 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "style:"
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
      Top             =   1605
      Width           =   570
   End
End
Attribute VB_Name = "FormGrayscale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Grayscale Conversion Handler
'Copyright ©2002-2014 by Tanner Helland
'Created: 1/12/02
'Last updated: 22/August/13
'Last update: add new command bar user control
'
'Updated version of the grayscale handler; utilizes five different methods
'(average, ISU, desaturate, X # of shades, X # of shades dithered).
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'Preview the current grayscale conversion technique
Private Sub drawGrayscalePreview()
    If cmdBar.previewsAllowed Then masterGrayscaleFunction cboMethod.ListIndex, getExtraGrayscaleParams(cboMethod.ListIndex), True, fxPreview
End Sub

Private Sub cboMethod_Click()
    UpdateVisibleControls
    drawGrayscalePreview
End Sub

'Certain algorithms require additional user input.  This routine enables/disables the controls associated with a given algorithm.
Private Sub UpdateVisibleControls()
    
    Select Case cboMethod.ListIndex
        Case 3
            sltShades.Visible = False
            lblAdditional.Caption = g_Language.TranslateMessage("decompose using these values:")
            lblAdditional.Visible = True
            picDecompose.Visible = True
            picChannel.Visible = False
        Case 4
            sltShades.Visible = False
            lblAdditional.Caption = g_Language.TranslateMessage("use this channel:")
            lblAdditional.Visible = True
            picDecompose.Visible = False
            picChannel.Visible = True
        Case 5
            sltShades.Visible = True
            lblAdditional.Caption = g_Language.TranslateMessage("use this many shades of gray:")
            lblAdditional.Visible = True
            picDecompose.Visible = False
            picChannel.Visible = False
        Case 6
            sltShades.Visible = True
            lblAdditional.Caption = g_Language.TranslateMessage("use this many shades of gray:")
            lblAdditional.Visible = True
            picDecompose.Visible = False
            picChannel.Visible = False
        Case Else
            sltShades.Visible = False
            lblAdditional.Visible = False
            picDecompose.Visible = False
            picChannel.Visible = False
    End Select
    
End Sub

'When validating the slider/text combos, if an invalid entry is found, automatically switch to that panel
Private Sub cmdBar_ExtraValidations()
    If Not sltShades.IsValid(False) And ((cboMethod.ListIndex <> 5) And (cboMethod.ListIndex <> 6)) Then cboMethod.ListIndex = 5
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Black and white", False, buildParams(cboMethod.ListIndex, getExtraGrayscaleParams(cboMethod.ListIndex))
End Sub

'All different grayscale (black and white) routines are handled by this single function
Public Sub masterGrayscaleFunction(Optional ByVal grayscaleMethod As Long, Optional ByVal additionalParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)

    'Use a parameter parse string to extract any additional parameters.
    Dim cParams As pdParamString
    Set cParams = New pdParamString
    cParams.setParamString additionalParams

    'Different grayscale conversion methods call different individual subs
    Select Case grayscaleMethod
        
        Case 0
            MenuGrayscaleAverage toPreview, dstPic
            
        Case 1
            MenuGrayscale toPreview, dstPic
            
        Case 2
            MenuDesaturate toPreview, dstPic
            
        Case 3
            MenuDecompose cParams.GetLong(1), toPreview, dstPic
            
        Case 4
            MenuGrayscaleSingleChannel cParams.GetLong(1), toPreview, dstPic
            
        Case 5
            fGrayscaleCustom cParams.GetLong(1), toPreview, dstPic
            
        Case 6
            fGrayscaleCustomDither cParams.GetLong(1), toPreview, dstPic
            
    End Select

End Sub

'Some grayscale functions require extra parameters.  Some do not.  Call this function to retrieve any extra parameters
' for a given grayscale conversion method.
Private Function getExtraGrayscaleParams(ByVal grayscaleMethod As Long) As String

    Select Case grayscaleMethod
        
        Case 0
            getExtraGrayscaleParams = ""
            
        Case 1
            getExtraGrayscaleParams = ""
            
        Case 2
            getExtraGrayscaleParams = ""
            
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
            getExtraGrayscaleParams = CStr(sltShades.Value)
            
        Case 6
            getExtraGrayscaleParams = CStr(sltShades.Value)
            
    End Select

End Function

Private Sub cmdBar_RequestPreviewUpdate()
    UpdateVisibleControls
    drawGrayscalePreview
End Sub

'Recommend ITU grayscale correction by default
Private Sub cmdBar_ResetClick()
    cboMethod.ListIndex = 1
End Sub

Private Sub Form_Activate()
    
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    setArrowCursorToObject picChannel
    setArrowCursorToObject picDecompose
    
    'Draw the initial preview
    cmdBar.markPreviewStatus True
    drawGrayscalePreview
    
End Sub

'Reduce to X # gray shades
Public Sub fGrayscaleCustom(ByVal numOfShades As Long, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)

    If Not toPreview Then Message "Converting image to %1 shades of gray...", numOfShades
    
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
        If toPreview = False Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x
            End If
        End If
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
    
End Sub

'Reduce to X # gray shades (dithered)
Public Sub fGrayscaleCustomDither(ByVal numOfShades As Long, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)

    If Not toPreview Then Message "Converting to %1 shades of gray, with dithering...", numOfShades
    
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
    If Not toPreview Then
        SetProgBarMax finalY
        progBarCheck = findBestProgBarValue()
    End If
    
    'Color variables
    Dim r As Long, g As Long, b As Long, grayVal As Long
    
    'This conversion factor is the value we need to turn grayscale values in the [0,255] range into a specific subset of values
    Dim conversionFactor As Double
    conversionFactor = (255 / (numOfShades - 1))
    
    'Build another look-up table for our initial grayscale index calculation
    Dim grayLookUp(0 To 765) As Byte
    For x = 0 To 765
        grayLookUp(x) = x \ 3
    Next x
    
    'Unfortunately, this algorithm (unlike its non-dithering counterpart) is not well-suited to using a look-up table,
    ' so all calculations have been moved into the loop
    Dim grayTempCalc As Double
    
    'This value tracks the drifting error of our conversions, which allows us to dither
    Dim errorValue As Double
    errorValue = 0
        
    'Loop through each pixel in the image, converting values as we go
    For y = initY To finalY
    For x = initX To finalX
    
        QuickVal = x * qvDepth
    
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Look up our initial grayscale value in the table
        grayVal = grayLookUp(r + g + b)
        
        'Add the error value (a cumulative value of the difference between actual gray values and gray values we've selected) to the current gray value
        grayTempCalc = grayVal + errorValue
        
        'Rebuild our temporary calculation variable using the shade reduction formula
        grayTempCalc = Int((CSng(grayTempCalc) / conversionFactor) + 0.5) * conversionFactor
        
        'Adjust our error value to include this latest calculation
        errorValue = CLng(grayVal) + errorValue - grayTempCalc
        
        If grayTempCalc < 0 Then grayTempCalc = 0
        If grayTempCalc > 255 Then grayTempCalc = 255
        
        grayVal = CByte(grayTempCalc)
        
        'Assign all color channels the new gray value
        ImageData(QuickVal + 2, y) = grayVal
        ImageData(QuickVal + 1, y) = grayVal
        ImageData(QuickVal, y) = grayVal
        
    Next x
        
        'Reset the error value at the end of each line
        errorValue = 0
        
        If toPreview = False Then
            If (y And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal y
            End If
        End If
        
    Next y
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
    
End Sub

'Reduce to gray via (r+g+b)/3
Public Sub MenuGrayscaleAverage(Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If Not toPreview Then Message "Converting image to black and white..."
    
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
        If toPreview = False Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x
            End If
        End If
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
    
End Sub

'Reduce to gray in a more human-eye friendly manner
Public Sub MenuGrayscale(Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If Not toPreview Then Message "Converting image to black and white..."
    
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
        If toPreview = False Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x
            End If
        End If
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
    
End Sub

'Reduce to gray via HSL -> convert S to 0
Public Sub MenuDesaturate(Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
        
    If Not toPreview Then Message "Converting image to black and white..."
    
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
        If toPreview = False Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x
            End If
        End If
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
    
End Sub

'Reduce to gray by selecting the minimum (maxOrMin = 0) or maximum (maxOrMin = 1) color in each pixel
Public Sub MenuDecompose(ByVal maxOrMin As Long, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)

    If Not toPreview Then Message "Converting image to black and white..."
    
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
        If toPreview = False Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x
            End If
        End If
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
    
End Sub

'Reduce to gray by selecting a single color channel (represeted by cChannel: 0 = Red, 1 = Green, 2 = Blue)
Public Sub MenuGrayscaleSingleChannel(ByVal cChannel As Long, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)

    Dim cString As String
     
    Select Case cChannel
        Case 0
            cString = g_Language.TranslateMessage("red")
        Case 1
            cString = g_Language.TranslateMessage("green")
        Case 2
            cString = g_Language.TranslateMessage("blue")
    End Select

    If Not toPreview Then Message "Converting image to black and white using %1 values...", cString
    
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
        If toPreview = False Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x
            End If
        End If
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
        
End Sub

Private Sub Form_Load()
    
    'Suspend previews while we get the form set up
    cmdBar.markPreviewStatus False
    
    'Set up the grayscale options combo box
    cboMethod.AddItem "Fastest Calculation (average value)", 0
    cboMethod.AddItem "Highest Quality (ITU Standard)", 1
    cboMethod.AddItem "Desaturate", 2
    cboMethod.AddItem "Decompose", 3
    cboMethod.AddItem "Single color channel", 4
    cboMethod.AddItem "Specific # of shades", 5
    cboMethod.AddItem "Specific # of shades (dithered)", 6
    
    UpdateVisibleControls
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'When option buttons are used, update the preview accordingly
Private Sub optChannel_Click(Index As Integer)
    drawGrayscalePreview
End Sub

Private Sub optDecompose_Click(Index As Integer)
    drawGrayscalePreview
End Sub

Private Sub sltShades_Change()
    drawGrayscalePreview
End Sub
