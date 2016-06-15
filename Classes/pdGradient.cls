VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pd2DGradient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Gradient Manager
'Copyright 2015-2016 by Tanner Helland
'Created: 23/July/15
'Last updated: 04/June/16
'Last update: lots of clean-up to match the new pd2D drawing system.
'
'This class manages a single PD gradient instance.
'
'At present, it is primarily based on the capabilities of GDI+ gradient brushes.  This will likely expand in the future,
' but GDI+ provides a nice baseline feature set.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private m_GradientShape As PD_2D_GradientShape
Private m_GradientAngle As Single
Private m_GradientWrapMode As PD_2D_WrapMode

'Number of points in the current gradient
Private m_NumOfPoints As Long

'Collection of GRADIENTPOINTs
Private m_GradientPoints() As GRADIENTPOINT

'Are the points currently sorted from lowest-to-highest?  If they are, this will be set to TRUE; this allows us to skip sorting
' prior to constructing the actual brush.
Private m_IsSorted As Boolean

'All non-linear brushes are constructed with the help of a pdGraphicsPath object
Private m_Path As pdGraphicsPath

'This class is capable of serializing itself to/from XML strings
Private cSerialize As pdParamXML

'Quick and dirty memory swap APIs
Private Type tmpLong
    lngResult As Long
End Type
Private Declare Sub CopyMemoryStrict Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpDst As Long, ByVal lpSrc As Long, ByVal byteLength As Long)
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (ByVal dstPointer As Long, ByVal Length As Long, ByVal filledValue As Byte)

'GDI+ brush creation
Private Declare Function GdipCreateLineBrushFromRectWithAngle Lib "gdiplus" (ByRef srcRect As RECTF, ByVal firstRGBA As Long, ByVal secondRGBA As Long, ByVal gradAngle As Single, ByVal isAngleScalable As Long, ByVal gradientWrapMode As PD_2D_WrapMode, ByRef dstLineGradientBrush As Long) As Long
Private Declare Function GdipCreatePathGradientFromPath Lib "gdiplus" (ByVal ptrToSrcPath As Long, ByRef dstPathGradientBrush As Long) As Long
Private Declare Function GdipSetLinePresetBlend Lib "gdiplus" (ByVal hBrush As Long, ByVal ptrToFirstColor As Long, ByVal ptrToFirstPosition As Long, ByVal ptCount As Long) As Long
Private Declare Function GdipSetPathGradientPresetBlend Lib "gdiplus" (ByVal hBrush As Long, ByVal ptrToFirstColor As Long, ByVal ptrToFirstPosition As Long, ByVal ptCount As Long) As Long
Private Declare Function GdipSetPathGradientWrapMode Lib "gdiplus" (ByVal hBrush As Long, ByVal newWrapMode As PD_2D_WrapMode) As Long
Private Declare Function GdipSetPathGradientCenterPoint Lib "gdiplus" (ByVal hBrush As Long, ByRef newCenterPoints As POINTFLOAT) As Long

'Get/set individual settings.  Note that these just wrap the generic Get/SetGradientProperty functions, below.
Friend Function GetGradientShape() As PD_2D_GradientShape
    GetGradientShape = GetGradientProperty(P2_GradientShape)
End Function

Friend Sub SetGradientShape(ByVal newShape As PD_2D_GradientShape)
    SetGradientProperty P2_GradientShape, newShape
End Sub

Public Function GetGradientAngle() As Single
    GetGradientAngle = GetGradientProperty(P2_GradientAngle)
End Function

Public Sub SetGradientAngle(ByVal newAngle As Single)
    SetGradientProperty P2_GradientAngle, newAngle
End Sub

Friend Function GetGradientWrapMode() As PD_2D_WrapMode
    GetGradientWrapMode = GetGradientProperty(P2_GradientWrapMode)
End Function

Friend Sub SetGradientWrapMode(ByVal newWrapMode As PD_2D_WrapMode)
    SetGradientProperty P2_GradientWrapMode, newWrapMode
End Sub

'Get/set individual settings by enum type
Public Function GetGradientProperty(ByVal propID As PD_2D_GRADIENT_SETTINGS) As Variant
    
    Select Case propID
        
        Case P2_GradientShape
            GetGradientProperty = m_GradientShape
            
        Case P2_GradientAngle
            GetGradientProperty = m_GradientAngle
            
        Case P2_GradientWrapMode
            GetGradientProperty = m_GradientWrapMode
            
    End Select
    
End Function

Public Sub SetGradientProperty(ByVal propID As PD_2D_GRADIENT_SETTINGS, ByVal newSetting As Variant)
    
    Select Case propID
        
        Case P2_GradientShape
            m_GradientShape = CLng(newSetting)
            
        Case P2_GradientAngle
            m_GradientAngle = CSng(newSetting)
            
        Case P2_GradientWrapMode
            m_GradientWrapMode = CLng(newSetting)
            
    End Select
    
End Sub

'For interop purposes, gradients are passed around PD as XML strings.
Public Function GetGradientAsString() As String
    
    With cSerialize
        .Reset 1#
        
        'Add parameters whose size and count do not vary
        .AddParam "GradientPointCount", m_NumOfPoints
        .AddParam "GradientListAlreadySorted", m_IsSorted
        .AddParam "GradientShape", m_GradientShape
        .AddParam "GradientAngle", m_GradientAngle
        .AddParam "GradientWrapMode", m_GradientWrapMode
        
        'Add the gradient point list
        Dim i As Long, iString As String
        For i = 0 To m_NumOfPoints - 1
            iString = Trim$(Str(i))
            .AddParam "GradientPoint_" & iString & "_RGB", m_GradientPoints(i).PointRGB
            .AddParam "GradientPoint_" & iString & "_Opacity", m_GradientPoints(i).PointOpacity
            .AddParam "GradientPoint_" & iString & "_Position", m_GradientPoints(i).PointPosition
        Next i
        
    End With
    
    GetGradientAsString = cSerialize.GetParamString
    
End Function

Public Sub CreateGradientFromString(ByVal srcString As String)
    
    'If the string is empty, prep a default object
    If (Len(srcString) = 0) Then
        
        m_NumOfPoints = 2
        ReDim m_GradientPoints(0 To 1) As GRADIENTPOINT
        
        With m_GradientPoints(0)
            .PointRGB = vbBlack
            .PointOpacity = 1
            .PointPosition = 0
        End With
        
        With m_GradientPoints(1)
            .PointRGB = vbWhite
            .PointOpacity = 1
            .PointPosition = 1
        End With
        
        m_GradientShape = P2_GS_Linear
        m_GradientAngle = 0
        m_GradientWrapMode = P2_WM_Tile
        
    Else

        With cSerialize
            .SetParamString srcString
            
            'Retrieve parameters whose size and count do not vary
            m_NumOfPoints = .GetLong("GradientPointCount", 0)
            m_IsSorted = .GetBool("GradientListAlreadySorted", False)
            m_GradientShape = .GetLong("GradientShape", P2_GS_Linear)
            m_GradientAngle = .GetDouble("GradientAngle", 0)
            m_GradientWrapMode = .GetLong("GradientWrapMode", P2_WM_Tile)
            
            'Retrieve the gradient point list
            ReDim m_GradientPoints(0 To m_NumOfPoints) As GRADIENTPOINT
            
            Dim i As Long, iString As String
            For i = 0 To m_NumOfPoints - 1
                iString = Trim$(Str(i))
                m_GradientPoints(i).PointRGB = .GetLong("GradientPoint_" & iString & "_RGB", vbBlack)
                m_GradientPoints(i).PointOpacity = .GetDouble("GradientPoint_" & iString & "_Opacity", 1)
                m_GradientPoints(i).PointPosition = .GetDouble("GradientPoint_" & iString & "_Position", i / m_NumOfPoints)
            Next i
            
        End With
        
    End If
    
End Sub

'The gradient editor assembles its own list of nodes.  To simplify interaction with this class, it can simply pass its
' instances to this function.
Friend Sub CreateGradientFromPointCollection(ByVal numOfPoints As Long, ByRef srcPoints() As GRADIENTPOINT)
    
    'Start by prepping our internal collections
    m_NumOfPoints = numOfPoints
    ReDim m_GradientPoints(0 To m_NumOfPoints) As GRADIENTPOINT
    
    'Copy the source array
    CopyMemoryStrict VarPtr(m_GradientPoints(0)), VarPtr(srcPoints(0)), LenB(m_GradientPoints(0)) * m_NumOfPoints
    
    'Assume the incoming array is not sorted, then sort it
    m_IsSorted = False
    SortGradientArray
    
End Sub

Friend Sub GetCopyOfPointCollection(ByRef numOfPoints As Long, ByRef srcPoints() As GRADIENTPOINT)
    numOfPoints = m_NumOfPoints
    ReDim srcPoints(0 To m_NumOfPoints) As GRADIENTPOINT
    CopyMemoryStrict VarPtr(srcPoints(0)), VarPtr(m_GradientPoints(0)), LenB(m_GradientPoints(0)) * m_NumOfPoints
End Sub

'Sort the gradient array in ascending order.  This greatly simplifies the process of creating a matching GDI+ brush.
Private Sub SortGradientArray()
    
    'If the array is already sorted, ignore this request
    If m_IsSorted Then Exit Sub
    
    'Because these arrays tend to be short, we can sort them using a simple bubble sort.
    If (m_NumOfPoints > 1) Then
    
        Dim i As Long, j As Long, loopBound As Long
        loopBound = m_NumOfPoints - 1
        
        'Loop through all entries in the stack, sorting them as we go
        For i = 0 To loopBound
            For j = 0 To loopBound
                If m_GradientPoints(i).PointPosition < m_GradientPoints(j).PointPosition Then SwapGradientPoints i, j
            Next j
        Next i
        
    End If
    
    'Mark the array as sorted
    m_IsSorted = True

End Sub

'Helper for node sorting, above.  This is not optimized at all, but it's sufficient for PD's current usage...
Private Sub SwapGradientPoints(ByVal gIndex1 As Long, ByVal gIndex2 As Long)
    Dim tmpPoint As GRADIENTPOINT
    tmpPoint = m_GradientPoints(gIndex1)
    m_GradientPoints(gIndex1) = m_GradientPoints(gIndex2)
    m_GradientPoints(gIndex2) = tmpPoint
End Sub

'Once this class is populated correctly, you can use this function to retrieve a matching GDI+ brush handle.
' Note that a rect is *required* in order to size the gradient region correctly.
' Finally, if you only want a default linear gradient (as used in the UI, among other things), set overrideNormalSettings to TRUE.
' This will return a linear gradient brush at angle zero, suitable for previews.
'
'(Also, the caller is obviously responsible for freeing the handle when done.)
Friend Function GetBrushHandle(ByRef dstRect As RECTF, Optional ByVal overrideNormalSettings As Boolean = False, Optional ByVal customOpacity As Single = 1#) As Long
    
    'Start by sorting the array.  (GDI+ may fail if an unsorted array is passed.)
    If (Not m_IsSorted) Then SortGradientArray
    
    'Next, we need to convert the gradient points into two separate arrays: one with merged RGBA values, and one with positions
    ' (as floating-point values on the range [0, 1].
    
    'As part of this process, we also need to insert discrete points at position 0 and position 1 (if they don't already exist).
    Dim dstRGBA() As Long, dstPosition() As Single
    ReDim dstRGBA(0 To m_NumOfPoints + 1) As Long
    ReDim dstPosition(0 To m_NumOfPoints + 1) As Single
    
    Dim gpOffset As Long
    gpOffset = 0
    
    'Start by creating position 0 if it doesn't already exist
    If (m_GradientPoints(0).PointPosition <> 0) Then
        gpOffset = 1
        dstRGBA(0) = GetMergedRGBA(0, customOpacity)
        dstPosition(0) = 0
    End If
    
    'Next, copy all sorted values into their destination array positions
    Dim i As Long
    For i = 0 To m_NumOfPoints - 1
        dstRGBA(i + gpOffset) = GetMergedRGBA(i, customOpacity)
        dstPosition(i + gpOffset) = m_GradientPoints(i).PointPosition
    Next i
    
    'Finally, see if we need to add a closing point at position 1
    If m_GradientPoints(m_NumOfPoints - 1).PointPosition <> 1 Then
        gpOffset = gpOffset + 1
        dstRGBA(m_NumOfPoints - 1 + gpOffset) = GetMergedRGBA(m_NumOfPoints - 1, customOpacity)
        dstPosition(m_NumOfPoints - 1 + gpOffset) = 1
    End If
    
    'The point arrays are now successfully constructed.  The "reflection" gradient type requires one additional step.
    ' We must manually enlarge the point array, and fill the top half with mirror copies of the existing gradient.
    Dim finalPointCount As Long
    
    If (m_GradientShape = P2_GS_Reflection) And (Not overrideNormalSettings) Then
        
        Dim originalPointCount As Long
        originalPointCount = m_NumOfPoints + gpOffset
        finalPointCount = originalPointCount + (originalPointCount - 1)
        
        'Start by cutting all positions in half.
        For i = 0 To originalPointCount - 1
            dstPosition(i) = dstPosition(i) / 2
        Next i
        
        'Next, enlarge the color and position arrays
        ReDim Preserve dstRGBA(0 To finalPointCount) As Long
        ReDim Preserve dstPosition(0 To finalPointCount) As Single
        
        'Finally, fill in the new array spots with corresponding color and position values
        For i = originalPointCount To finalPointCount - 1
            dstRGBA(i) = dstRGBA(originalPointCount + (originalPointCount - i) - 2)
            dstPosition(i) = 1 - dstPosition(originalPointCount + (originalPointCount - i) - 2)
        Next i
        
    Else
        finalPointCount = m_NumOfPoints + gpOffset
    End If
    
    'We can now create the brush!  We must do this in two steps.
    Dim gdipBrush As Long
    
    'First: the user can request a default, angle = 0 linear version of the gradient for UI purposes.  Handle that case separately.
    If overrideNormalSettings Then
    
        'Create a default linear gradient brush
        If GdipCreateLineBrushFromRectWithAngle(dstRect, dstRGBA(0), dstRGBA(1), 0#, 1, P2_WM_TileFlipXY, gdipBrush) <> 0 Then
            Debug.Print "WARNING!  pd2DGradient.getBrushHandle failed during line brush creation."
        End If
        
        'Next, override the brush colors with our custom array
        If GdipSetLinePresetBlend(gdipBrush, VarPtr(dstRGBA(0)), VarPtr(dstPosition(0)), finalPointCount) <> 0 Then
            Debug.Print "WARNING!  pd2DGradient.getBrushHandle failed during line preset blending."
        End If
        
    Else
        
        'Brush creation varies by shape
        Select Case m_GradientShape
        
            Case P2_GS_Linear, P2_GS_Reflection
                
                'Create a default linear gradient brush
                If GdipCreateLineBrushFromRectWithAngle(dstRect, dstRGBA(0), dstRGBA(1), m_GradientAngle, 1, m_GradientWrapMode, gdipBrush) <> 0 Then
                    Debug.Print "WARNING!  pd2DGradient.getBrushHandle failed during line brush creation."
                End If
                
                'Next, override the brush colors with our custom array
                If GdipSetLinePresetBlend(gdipBrush, VarPtr(dstRGBA(0)), VarPtr(dstPosition(0)), finalPointCount) <> 0 Then
                    Debug.Print "WARNING!  pd2DGradient.getBrushHandle failed during line preset blending."
                End If
        
            'At present, all other gradient types are constructed via reference path
            Case Else
            
                'Reset our graphics path reference
                m_Path.ResetPath
                
                'Some paths require a modified rect (e.g. square vs rectangle, circle vs ellipse), so construct a few additional
                ' measurements in advance.
                Dim halfWidth As Double, halfHeight As Double
                halfWidth = dstRect.Width / 2
                halfHeight = dstRect.Height / 2
                
                Dim centerX As Double, centerY As Double
                centerX = dstRect.Left + halfWidth
                centerY = dstRect.Top + halfHeight
                
                Dim squareRect As RECTF, maxLength As Double, radiusRect As Double, radiusSquare As Double
                If dstRect.Width > dstRect.Height Then maxLength = halfWidth Else maxLength = halfHeight
                radiusRect = Sqr(halfWidth * halfWidth + halfHeight * halfHeight) + 0.5
                radiusSquare = Sqr(maxLength * maxLength + maxLength * maxLength) + 0.5
                                
                With squareRect
                    .Left = centerX - maxLength
                    .Top = centerY - maxLength
                    .Width = maxLength * 2
                    .Height = maxLength * 2
                End With
                
                'Construct a reference path, using the 1) target rect, and 2) gradient shape to determine path behavior
                If m_GradientShape = P2_GS_Radial Then
                    
                    'Enclose the gradient within a circle with the same radius as the bounding rect's diagonal
                    m_Path.AddEllipse_Absolute centerX - radiusRect, centerY - radiusRect, centerX + radiusRect, centerY + radiusRect
                
                ElseIf m_GradientShape = P2_GS_Rectangle Then
                    
                    'Rectangles are easy - just use the bounding rect itself!
                    m_Path.AddRectangle_RectF dstRect
                    
                ElseIf m_GradientShape = P2_GS_Diamond Then
                    
                    'I had to derive this bounding formula by hand, so sorry for the lack of an easy diagram!
                    
                    'Basically, the crucial measurement for a diamond gradient is the hypotenuse of a triangle that meets
                    ' the following criteria:
                    ' 1) right triangle
                    ' 2) 90 degree node at the center of the bounding rect
                    ' 3) two other nodes at the intersection of a line that...
                    '    a) passes through the top-left corner of the bounding rect
                    '    b) intersects the horizontal and vertical center lines of the bounding rect at 45 degrees
                    
                    'Find the sine of a 45 degree angle (which must first be converted to radians, obviously)
                    Dim sin45 As Double
                    sin45 = Sin(45 * PI_DIV_180)
                    
                    'Use that value to add together the hypotenuses of the two triangles that, together with the top-left
                    ' quadrant of the bounding rect, form the triangle described above.
                    Dim dHypotenuse As Double
                    dHypotenuse = halfWidth / sin45 + halfHeight / sin45 + 0.5
                    
                    'Construct a square, using that hypotenuse value as the square's width/height
                    With squareRect
                        .Left = centerX - dHypotenuse / 2
                        .Top = centerY - dHypotenuse / 2
                        .Width = dHypotenuse
                        .Height = dHypotenuse
                    End With
                    
                    'Rotate the square 45 degrees and bam, we have a perfectly constricted diamond region!
                    m_Path.AddRectangle_RectF squareRect
                    m_Path.RotatePathAroundItsCenter 45#
                    
                End If
                
                'Create a path gradient brush, using our path object as the reference
                If GdipCreatePathGradientFromPath(m_Path.GetCopyOfPathHandle, gdipBrush) <> 0 Then
                    Debug.Print "WARNING!  pd2DGradient.getBrushHandle failed during path brush creation."
                End If
                
                'Set a default center point (in the future we may make this editable)
                Dim centerPoint As POINTFLOAT
                centerPoint.x = centerX
                centerPoint.y = centerY
                GdipSetPathGradientCenterPoint gdipBrush, centerPoint
                
                'Next, override the brush colors with our custom array
                If GdipSetPathGradientPresetBlend(gdipBrush, VarPtr(dstRGBA(0)), VarPtr(dstPosition(0)), finalPointCount) <> 0 Then
                    Debug.Print "WARNING!  pd2DGradient.getBrushHandle failed during pathpreset blending."
                End If
                
                'Finally, set any/all custom parameters (e.g. wrap mode)
                GdipSetPathGradientWrapMode gdipBrush, m_GradientWrapMode
        
        End Select
        
    End If
    
    'Return the brush!  (Note that the caller is responsible for freeing the brush when done.)
    GetBrushHandle = gdipBrush
    
End Function

'Given an index into the gradient point array, return a merged RGBA value using the opacity.
' (This is inefficient but it doesn't matter as it's only called a handful of times.)
Private Function GetMergedRGBA(ByVal gpIndex As Long, Optional ByVal customOpacityModifier As Single = 1#) As Long
    
    Dim dstQuad As RGBQUAD
    dstQuad.Red = ExtractR(m_GradientPoints(gpIndex).PointRGB)
    dstQuad.Green = ExtractG(m_GradientPoints(gpIndex).PointRGB)
    dstQuad.Blue = ExtractB(m_GradientPoints(gpIndex).PointRGB)
    dstQuad.alpha = (m_GradientPoints(gpIndex).PointOpacity * customOpacityModifier) * 255#
    
    Dim placeHolder As tmpLong
    LSet placeHolder = dstQuad
    
    GetMergedRGBA = placeHolder.lngResult
    
End Function

'Given a position on the range [0, 1], return the gradient color at that position.  Note that a linear, non-reflected
' gradient shape is always assumed, by design.
Friend Function GetColorAtPosition_RGBA(ByVal cPosition As Single, ByRef dstRGBA As RGBQUAD) As Boolean

    'Make sure we have a filled, sorted input array
    If (m_NumOfPoints > 0) Then
        
        If (Not m_IsSorted) Then SortGradientArray
        
        Dim tmpLowLong As tmpLong, tmpHighLong As tmpLong
        
        'The requested position can lie one of three places:
        ' 1) Before node 0
        ' 2) After the final node
        ' 3) In-between two existing nodes
        
        'Tackle these possibilities in turn.
        If (cPosition < m_GradientPoints(0).PointPosition) Then
        
            'Return the base color
            tmpLowLong.lngResult = GetMergedRGBA(0)
            LSet dstRGBA = tmpLowLong
            
        ElseIf (cPosition > m_GradientPoints(m_NumOfPoints - 1).PointPosition) Then
            
            'Return the top color
            tmpHighLong.lngResult = GetMergedRGBA(m_NumOfPoints - 1)
            LSet dstRGBA = tmpHighLong
            
        'We must interpolate the color manually
        Else
        
            'Find the indices of the colors surrounding the requested position.
            Dim i As Long
            i = 1
            Do While (m_GradientPoints(i).PointPosition < cPosition)
                i = i + 1
            Loop
            
            'i now points at the node position *just greater* than the requested point.
            
            'Calculate a linear blend factor between the surrounding positions
            Dim blendFactor As Double
            blendFactor = (cPosition - m_GradientPoints(i - 1).PointPosition) / (m_GradientPoints(i).PointPosition - m_GradientPoints(i - 1).PointPosition)
            
            'Retrieve the RGBA components of both colors
            Dim tmpLow As RGBQUAD, tmpHigh As RGBQUAD
            tmpLowLong.lngResult = GetMergedRGBA(i - 1)
            tmpHighLong.lngResult = GetMergedRGBA(i)
            
            LSet tmpLow = tmpLowLong
            LSet tmpHigh = tmpHighLong
            
            'Start blending
            dstRGBA.Red = GetBlendedColors(tmpLow.Red, tmpHigh.Red, blendFactor)
            dstRGBA.Green = GetBlendedColors(tmpLow.Green, tmpHigh.Green, blendFactor)
            dstRGBA.Blue = GetBlendedColors(tmpLow.Blue, tmpHigh.Blue, blendFactor)
            dstRGBA.alpha = GetBlendedColors(tmpLow.alpha, tmpHigh.alpha, blendFactor)
        
        End If
        
        GetColorAtPosition_RGBA = True
        
    Else
        Debug.Print "WARNING!  getColorAtPosition_RGBA was called before the gradient was initialized!"
        GetColorAtPosition_RGBA = False
    End If
    

End Function

'Blend byte1 w/ byte2 based on mixRatio. mixRatio is expected to be a value between 0 and 1.
Private Function GetBlendedColors(ByVal firstColor As Byte, ByVal secondColor As Byte, ByRef mixRatio As Double) As Byte
    GetBlendedColors = ((1 - mixRatio) * firstColor) + (mixRatio * secondColor)
End Function

Private Sub Class_Initialize()
    
    'Prep a string serializer (for storing/loading gradients externally)
    Set cSerialize = New pdParamXML
    
    'Prep a path (for constructing non-linear gradients)
    Set m_Path = New pdGraphicsPath
    
    'Assume sorting has not taken place
    m_IsSorted = False
    
    'Set all other default parameters
    m_GradientShape = P2_GS_Linear
    m_GradientAngle = 0
    m_GradientWrapMode = P2_WM_Tile
    
End Sub

