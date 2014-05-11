VERSION 5.00
Begin VB.Form FormFragment
   AutoRedraw = -1 'True
   BackColor = &H80000005&
   BorderStyle = 4 'Fixed ToolWindow
   Caption = " Fragment"
   ClientHeight = 6525
   ClientLeft = -15
   ClientTop = 225
   ClientWidth = 11895
   BeginProperty Font
      Name = "Tahoma"
      Size = 8.25
      Charset = 0
      Weight = 400
      Underline = 0 'False
      Italic = 0 'False
      Strikethrough = 0 'False
   EndProperty
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MinButton = 0 'False
   ScaleHeight = 435
   ScaleMode = 3 'Pixel
   ScaleWidth = 793
   ShowInTaskbar = 0 'False
   Begin PhotoDemon.commandBar cmdBar
      Align = 2 'Align Bottom
      Height = 750
      Left = 0
      TabIndex = 0
      Top = 5775
      Width = 11895
      _ExtentX = 20981
      _ExtentY = 1323
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
         Name = "Tahoma"
         Size = 9.75
         Charset = 0
         Weight = 400
         Underline = 0 'False
         Italic = 0 'False
         Strikethrough = 0 'False
      EndProperty
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview
      Height = 5625
      Left = 120
      TabIndex = 1
      Top = 120
      Width = 5625
      _ExtentX = 9922
      _ExtentY = 9922
      DisableZoomPan = -1 'True
   End
   Begin PhotoDemon.sliderTextCombo sltDistance
      Height = 495
      Left = 6000
      TabIndex = 3
      Top = 2520
      Width = 5775
      _ExtentX = 10186
      _ExtentY = 873
      Max = 100
      Value = 8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
         Name = "Tahoma"
         Size = 9.75
         Charset = 0
         Weight = 400
         Underline = 0 'False
         Italic = 0 'False
         Strikethrough = 0 'False
      EndProperty
      ForeColor = 0
   End
   Begin PhotoDemon.sliderTextCombo sltFragments
      Height = 495
      Left = 6000
      TabIndex = 4
      Top = 1320
      Width = 5775
      _ExtentX = 10186
      _ExtentY = 873
      Min = 2
      Max = 50
      Value = 4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
         Name = "Tahoma"
         Size = 9.75
         Charset = 0
         Weight = 400
         Underline = 0 'False
         Italic = 0 'False
         Strikethrough = 0 'False
      EndProperty
      ForeColor = 0
   End
   Begin PhotoDemon.sliderTextCombo sltAngle
      Height = 495
      Left = 6000
      TabIndex = 6
      Top = 3720
      Width = 5775
      _ExtentX = 10186
      _ExtentY = 873
      Max = 360
      SigDigits = 2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
         Name = "Tahoma"
         Size = 9.75
         Charset = 0
         Weight = 400
         Underline = 0 'False
         Italic = 0 'False
         Strikethrough = 0 'False
      EndProperty
      ForeColor = 0
   End
   Begin VB.Label lblAngle
      AutoSize = -1 'True
      BackStyle = 0 'Transparent
      Caption = "angle:"
      BeginProperty Font
         Name = "Tahoma"
         Size = 12
         Charset = 0
         Weight = 400
         Underline = 0 'False
         Italic = 0 'False
         Strikethrough = 0 'False
      EndProperty
      ForeColor = &H00404040&
      Height = 285
      Left = 6000
      TabIndex = 7
      Top = 3360
      Width = 660
   End
   Begin VB.Label lblFragments
      AutoSize = -1 'True
      BackStyle = 0 'Transparent
      Caption = "# of fragments:"
      BeginProperty Font
         Name = "Tahoma"
         Size = 12
         Charset = 0
         Weight = 400
         Underline = 0 'False
         Italic = 0 'False
         Strikethrough = 0 'False
      EndProperty
      ForeColor = &H00404040&
      Height = 285
      Left = 6000
      TabIndex = 5
      Top = 960
      Width = 1695
   End
   Begin VB.Label lblDistance
      AutoSize = -1 'True
      BackStyle = 0 'Transparent
      Caption = "distance:"
      BeginProperty Font
         Name = "Tahoma"
         Size = 12
         Charset = 0
         Weight = 400
         Underline = 0 'False
         Italic = 0 'False
         Strikethrough = 0 'False
      EndProperty
      ForeColor = &H00404040&
      Height = 285
      Left = 6000
      TabIndex = 2
      Top = 2160
      Width = 945
   End
End
Attribute VB_Name = "FormFragment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Fragment Filter Dialog
'Copyright ©2014 by Audioglider
'Created: 10/May/14
'Last updated: 10/May/14
'Last update: converted fragments to point-based system and added adjustable
' # of fragments and angle.
'
'Similar to the Fragment effect from Photoshop except much more adjustable.
' We allow the user to set the number of layers and offset them the same
' distance from the origin at different positions as well as
' altering the angle. Then we merge them all all together.
'
'***************************************************************************

Option Explicit

Private Type LRGBQUAD
   rgbBlue As Long
   rgbGreen As Long
   rgbRed As Long
   rgbAlpha As Long
End Type

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'Apply a fragment filter to the active layer
Public Sub Fragment(ByVal fragments As Long, ByVal distance As Long, ByVal rotationAngle As Double, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
   
    If Not toPreview = False Then Message "Applying beer goggles..."
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Create a second local array. This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent already-processed pixels from affecting the results of later pixels.)
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.createFromExistingDIB workingDIB
    
    prepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
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
    
    'To keep processing quick, only update the progress bar when absolutely necessary. This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Setup our offset points using the provided parameters
    Dim pointOffsets() As POINTAPI
    Dim n As Long
    Dim num As Double, num2 As Double, num3 As Double
            
    num = PI_DOUBLE / CDbl(fragments)
    num2 = ((rotationAngle - 90) * PI) / 180
    
    ReDim pointOffsets(0 To fragments - 1) As POINTAPI
    For n = 0 To fragments - 1
        num3 = num2 + (num * n)
        pointOffsets(n).x = Round(CDbl(distance * -Sin(num3)))
        pointOffsets(n).y = Round(CDbl(distance * -Cos(num3)))
    Next n
    
    Dim numPoints As Long
    numPoints = UBound(pointOffsets)
    
    'Stores colors for each point
    Dim colors() As LRGBQUAD
    ReDim colors(0 To UBound(pointOffsets)) As LRGBQUAD
            
    'This look-up table will be used for alpha-blending. It contains the equivalent of any two color values [0,255] added
    ' together and divided by 2.
    Dim hLookup(0 To 510) As Byte
    For x = 0 To 510
        hLookup(x) = x \ 2
    Next x
    
    'Color variables
    Dim r As Long, g As Long, b As Long, a As Long
    Dim newR As Long, newG As Long, newB As Long, newA As Long
    
    Dim yOffset As Long, xOffset As Long
    
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Grab the current pixel values
        r = srcImageData(QuickVal + 2, y)
        g = srcImageData(QuickVal + 1, y)
        b = srcImageData(QuickVal, y)
        If qvDepth = 4 Then a = srcImageData(QuickVal + 3, y)
        
        For n = 0 To numPoints
            xOffset = x - pointOffsets(n).x
            yOffset = y - pointOffsets(n).y
            
            'Perform a bounds check
            If xOffset < 0 Then
                xOffset = x
                yOffset = y
            End If
            
            If yOffset < 0 Then
                xOffset = x
                yOffset = y
            End If
            
            If xOffset > finalX Then
                xOffset = x
                yOffset = y
            End If
            
            If yOffset > finalY Then
                xOffset = x
                yOffset = y
            End If
            
            newR = srcImageData(xOffset * qvDepth + 2, yOffset)
            newG = srcImageData(xOffset * qvDepth + 1, yOffset)
            newB = srcImageData(xOffset * qvDepth, yOffset)
            If qvDepth = 4 Then
                newA = srcImageData(xOffset * qvDepth + 3, yOffset)
                colors(n).rgbAlpha = newA
            End If
            colors(n).rgbRed = newR
            colors(n).rgbGreen = newG
            colors(n).rgbBlue = newB
        Next n
        
        'First, blend the the original color with the first layer
        ' before looping through the rest of the color array
        newR = hLookup(r + colors(0).rgbRed)
        newG = hLookup(g + colors(0).rgbGreen)
        newR = hLookup(b + colors(0).rgbBlue)
        If qvDepth = 4 Then newA = hLookup(a + colors(0).rgbAlpha)
        For n = 1 To numPoints
            newR = newR + hLookup(colors(n - 1).rgbRed + colors(n).rgbRed)
            newG = newG + hLookup(colors(n - 1).rgbGreen + colors(n).rgbGreen)
            newB = newB + hLookup(colors(n - 1).rgbBlue + colors(n).rgbBlue)
            If qvDepth = 4 Then newA = newA + hLookup(colors(n - 1).rgbAlpha + colors(n).rgbAlpha)
        Next n
        
        newR = newR \ numPoints + 1
        newG = newG \ numPoints + 1
        newB = newB \ numPoints + 1
        
        If newR > 255 Then newR = 255
        If newG > 255 Then newG = 255
        If newB > 255 Then newB = 255
        
        dstImageData(QuickVal + 2, y) = newR
        dstImageData(QuickVal + 1, y) = newG
        dstImageData(QuickVal, y) = newB
        If qvDepth = 4 Then
            newA = newA \ numPoints + 1
            If newA > 255 Then newA = 255
            dstImageData(QuickVal + 3, y) = newA
        End If
        
    Next y
        If Not toPreview = False Then
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
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
        
End Sub

Private Sub cmdBar_OKClick()
    Process "Fragment", False, buildParams(sltFragments.Value, sltDistance.Value, sltAngle.Value)
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltFragments.Value = 4
    sltDistance.Value = 8
    sltAngle.Value = 0
End Sub

Private Sub Form_Activate()
        
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
    'Render an image preview
    updatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub updatePreview()
    If cmdBar.previewsAllowed Then Fragment sltFragments, sltDistance, sltAngle, True, fxPreview
End Sub

Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

Private Sub sltAngle_Change()
    updatePreview
End Sub

Private Sub sltDistance_Change()
    updatePreview
End Sub

Private Sub sltFragments_Change()
    updatePreview
End Sub