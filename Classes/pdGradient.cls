VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdGradient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Gradient Manager
'Copyright 2014-2015 by Tanner Helland
'Created: 23/July/15
'Last updated: 23/July/15
'Last update: initial build
'
'This class manages a single PD gradient instance.
'
'At present, it is primarily based on the capabilities of GDI+ gradient brushes.  This will likely expand in the future,
' but GDI+ provides a nice baseline feature set so I've started with it.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Generic gradient properties
Public Enum PD_GRADIENT_SETTINGS
    pdgs_GradientShape = 0
    pdgs_GradientAngle = 1
    pdgs_GradientWrapMode = 2
End Enum

#If False Then
    Private Const pdgs_GradientShape = 0, pdgs_GradientAngle = 1, pdgs_GradientWrapMode = 2
#End If

Public Enum PD_GRADIENT_SHAPE
    pdgs_ShapeLinear = 0
    pdgs_ShapeRadial = 1
    pdgs_ShapeRectangle = 2
    pdgs_ShapeDiamond = 3
End Enum

#If False Then
    Private Const pdgs_ShapeLinear = 0, pdgs_ShapeRadial = 1
#End If

Private m_GradientShape As PD_GRADIENT_SHAPE
Private m_GradientAngle As Single
Private m_GradientWrapMode As WrapMode

'Number of points in the current gradient
Private m_NumOfPoints As Long

'Collection of pdGradientPoints
Private m_GradientPoints() As pdGradientPoint

'Are the points currently sorted from lowest-to-highest?  If they are, this will be set to TRUE; this allows us to skip sorting
' prior to constructing the actual brush.
Private m_IsSorted As Boolean

'All non-linear brushes are constructed with the help of a pdGraphicsPath object
Private m_Path As pdGraphicsPath

'This class is capable of serializing itself to/from XML strings
Private cSerialize As pdParamXML

'Quick memory management
Private Type tmpLong
    lngResult As Long
End Type
Private Declare Sub CopyMemoryStrict Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpDst As Long, ByVal lpSrc As Long, ByVal byteLength As Long)
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (ByVal dstPointer As Long, ByVal Length As Long, ByVal Fill As Byte)

'GDI+ brush creation
Private Declare Function GdipCreateLineBrushFromRectWithAngle Lib "gdiplus" (ByRef srcRect As RECTF, ByVal firstRGBA As Long, ByVal secondRGBA As Long, ByVal gradAngle As Single, ByVal isAngleScalable As Long, ByVal gradientWrapMode As WrapMode, ByRef dstLineGradientBrush As Long) As Long
Private Declare Function GdipCreatePathGradientFromPath Lib "gdiplus" (ByVal ptrToSrcPath As Long, ByRef dstPathGradientBrush As Long) As Long
Private Declare Function GdipSetLinePresetBlend Lib "gdiplus" (ByVal hBrush As Long, ByVal ptrToFirstColor As Long, ByVal ptrToFirstPosition As Long, ByVal ptCount As Long) As Long
Private Declare Function GdipSetPathGradientPresetBlend Lib "gdiplus" (ByVal hBrush As Long, ByVal ptrToFirstColor As Long, ByVal ptrToFirstPosition As Long, ByVal ptCount As Long) As Long
Private Declare Function GdipSetPathGradientWrapMode Lib "gdiplus" (ByVal hBrush As Long, ByVal newWrapMode As WrapMode) As Long

'Get/set individual settings by enum type
Public Function getGradientProperty(ByVal propID As PD_GRADIENT_SETTINGS) As Variant
    
    Select Case propID
        
        Case pdgs_GradientShape
            getGradientProperty = m_GradientShape
            
        Case pdgs_GradientAngle
            getGradientProperty = m_GradientAngle
            
        Case pdgs_GradientWrapMode
            getGradientProperty = m_GradientWrapMode
            
    End Select
    
End Function

Public Sub setGradientProperty(ByVal propID As PD_GRADIENT_SETTINGS, ByVal newSetting As Variant)
    
    Select Case propID
        
        Case pdgs_GradientShape
            m_GradientShape = CLng(newSetting)
            
        Case pdgs_GradientAngle
            m_GradientAngle = CSng(newSetting)
            
        Case pdgs_GradientWrapMode
            m_GradientWrapMode = CLng(newSetting)
            
    End Select
    
End Sub


'For interop purposes, gradients are passed around PD as XML strings.
Public Function getGradientAsString() As String
    
    With cSerialize
        .Reset 1#
        
        'Add any parameters whose size and count do not vary
        .addParam "GradientPointCount", m_NumOfPoints
        .addParam "GradientListAlreadySorted", m_IsSorted
        .addParam "GradientShape", m_GradientShape
        .addParam "GradientAngle", m_GradientAngle
        .addParam "GradientWrapMode", m_GradientWrapMode
        
        'Add the gradient point list
        Dim i As Long, iString As String
        For i = 0 To m_NumOfPoints - 1
            iString = Trim$(Str(i))
            .addParam "GradientPoint_" & iString & "_RGB", m_GradientPoints(i).pdgp_RGB
            .addParam "GradientPoint_" & iString & "_Opacity", m_GradientPoints(i).pdgp_Opacity
            .addParam "GradientPoint_" & iString & "_Position", m_GradientPoints(i).pdgp_Position
        Next i
        
    End With
    
    getGradientAsString = cSerialize.getParamString
    
End Function

Public Sub createGradientFromString(ByVal srcString As String)
    
    'If the string is empty, prep a default object
    If Len(srcString) = 0 Then
        
        m_NumOfPoints = 2
        ReDim m_GradientPoints(0 To 1) As pdGradientPoint
        
        With m_GradientPoints(0)
            .pdgp_RGB = vbBlack
            .pdgp_Opacity = 1
            .pdgp_Position = 0
        End With
        
        With m_GradientPoints(1)
            .pdgp_RGB = vbWhite
            .pdgp_Opacity = 1
            .pdgp_Position = 1
        End With
        
        m_GradientShape = pdgs_ShapeLinear
        m_GradientAngle = 0
        m_GradientWrapMode = WrapModeTile
        
    Else

        With cSerialize
            .setParamString srcString
            
            'Retrieve parameters whose size and count do not vary
            m_NumOfPoints = .GetLong("GradientPointCount", 0)
            m_IsSorted = .GetBool("GradientListAlreadySorted", False)
            m_GradientShape = .GetLong("GradientShape", pdgs_ShapeLinear)
            m_GradientAngle = .GetDouble("GradientAngle", 0)
            m_GradientWrapMode = .GetLong("GradientWrapMode", 0)
            
            'Retrieve the gradient point list
            ReDim m_GradientPoints(0 To m_NumOfPoints) As pdGradientPoint
            
            Dim i As Long, iString As String
            For i = 0 To m_NumOfPoints - 1
                iString = Trim$(Str(i))
                m_GradientPoints(i).pdgp_RGB = .GetLong("GradientPoint_" & iString & "_RGB", vbBlack)
                m_GradientPoints(i).pdgp_Opacity = .GetDouble("GradientPoint_" & iString & "_Opacity", 1)
                m_GradientPoints(i).pdgp_Position = .GetDouble("GradientPoint_" & iString & "_Position", i / m_NumOfPoints)
            Next i
            
        End With
        
    End If
    
End Sub

'The gradient editor assembles its own list of nodes.  To simplify interaction with this class, it can simply pass its
' instances to this function.
Friend Sub createGradientFromPointCollection(ByVal numOfPoints As Long, ByRef srcPoints() As pdGradientPoint)
    
    'Start by prepping our internal collections
    m_NumOfPoints = numOfPoints
    ReDim m_GradientPoints(0 To m_NumOfPoints) As pdGradientPoint
    
    'Copy the source array
    CopyMemoryStrict VarPtr(m_GradientPoints(0)), VarPtr(srcPoints(0)), LenB(m_GradientPoints(0)) * m_NumOfPoints
    
    'Assume the incoming array is not sorted, then sort it
    m_IsSorted = False
    sortGradientArray
    
End Sub

Friend Sub getCopyOfPointCollection(ByRef numOfPoints As Long, ByRef srcPoints() As pdGradientPoint)
    
    numOfPoints = m_NumOfPoints
    ReDim srcPoints(0 To m_NumOfPoints) As pdGradientPoint
    
    CopyMemoryStrict VarPtr(srcPoints(0)), VarPtr(m_GradientPoints(0)), LenB(m_GradientPoints(0)) * m_NumOfPoints
    
End Sub

'Sort the gradient array in ascending order.  This greatly simplifies the process of creating a matching GDI+ brush.
Private Sub sortGradientArray()
    
    'If the array is already sorted, ignore this request
    If m_IsSorted Then Exit Sub
    
    'Because these arrays tend to be short, we can sort them using a simple bubble sort.
    If m_NumOfPoints > 1 Then
    
        Dim i As Long, j As Long, loopBound As Long
        loopBound = m_NumOfPoints - 1
        
        'Loop through all entries in the stack, sorting them as we go
        For i = 0 To loopBound
            For j = 0 To loopBound
                If m_GradientPoints(i).pdgp_Position < m_GradientPoints(j).pdgp_Position Then swapGradientPoints i, j
            Next j
        Next i
        
    End If
    
    'Mark the array as sorted
    m_IsSorted = True

End Sub

'Helper for string sorting, above.  This is not optimized at all, but it's sufficient for PD's current usage...
Private Sub swapGradientPoints(ByVal gIndex1 As Long, ByVal gIndex2 As Long)
    Dim tmpPoint As pdGradientPoint
    tmpPoint = m_GradientPoints(gIndex1)
    m_GradientPoints(gIndex1) = m_GradientPoints(gIndex2)
    m_GradientPoints(gIndex2) = tmpPoint
End Sub

'Once this class is populated correctly, you can use this function to retrieve a matching GDI+ brush handle.
' Note that a rect is *required* in order to size the gradient region correctly.
' Finally, if you only want a default linear gradient (as used in the UI, among other things), set overrideNormalSettings to TRUE.
'
'(Also, the caller is obviously responsible for freeing the handle when done.)
Friend Function getBrushHandle(ByRef dstRect As RECTF, Optional ByVal overrideNormalSettings As Boolean = False, Optional ByVal customOpacity As Single = 1#) As Long
    
    'Start by sorting the array
    If Not m_IsSorted Then sortGradientArray
    
    'Next, we need to convert the gradient points into two separate arrays: one with merged RGBA values, and one with positions.
    ' As part of this process, we also need to insert point 0 and point 1 values if they don't exist.
    Dim dstRGBA() As Long, dstPosition() As Single
    ReDim dstRGBA(0 To m_NumOfPoints + 1) As Long
    ReDim dstPosition(0 To m_NumOfPoints + 1) As Single
    
    Dim gpOffset As Long
    gpOffset = 0
    
    'Start by creating position 0 if it doesn't already exist
    If m_GradientPoints(0).pdgp_Position <> 0 Then
        gpOffset = 1
        dstRGBA(0) = getMergedRGBA(0, customOpacity)
        dstPosition(0) = 0
    End If
    
    'Next, copy all sorted values into their destination array positions
    Dim i As Long
    For i = 0 To m_NumOfPoints - 1
        dstRGBA(i + gpOffset) = getMergedRGBA(i, customOpacity)
        dstPosition(i + gpOffset) = m_GradientPoints(i).pdgp_Position
    Next i
    
    'Finally, see if we need to add a closing point at position 1
    If m_GradientPoints(m_NumOfPoints - 1).pdgp_Position <> 1 Then
        gpOffset = gpOffset + 1
        dstRGBA(m_NumOfPoints - 1 + gpOffset) = getMergedRGBA(m_NumOfPoints - 1, customOpacity)
        dstPosition(m_NumOfPoints - 1 + gpOffset) = 1
    End If
        
    'We can now create the brush!  We do this in two steps, as gradient brushes do not have a constructor that accepts
    ' a variable color array like this.
    
    Dim gdipBrush As Long
    
    'First: the user can request a default, angle = 0 linear version of the gradient for UI purposes.  Handle that case separately.
    If overrideNormalSettings Then
    
        'Create a default linear gradient brush
        If GdipCreateLineBrushFromRectWithAngle(dstRect, dstRGBA(0), dstRGBA(m_NumOfPoints - 1 + gpOffset), 0#, 1, WrapModeTileFlipXY, gdipBrush) <> 0 Then
            Debug.Print "WARNING!  pdGradient.getBrushHandle failed during line brush creation."
        End If
        
        'Next, override the brush colors with our custom array
        If GdipSetLinePresetBlend(gdipBrush, VarPtr(dstRGBA(0)), VarPtr(dstPosition(0)), m_NumOfPoints + gpOffset) <> 0 Then
            Debug.Print "WARNING!  pdGradient.getBrushHandle failed during line preset blending."
        End If
        
    Else
        
        'Brush creation varies by shape
        Select Case m_GradientShape
        
            Case pdgs_ShapeLinear
                
                'Create a default linear gradient brush
                If GdipCreateLineBrushFromRectWithAngle(dstRect, dstRGBA(0), dstRGBA(m_NumOfPoints - 1 + gpOffset), m_GradientAngle, 1, m_GradientWrapMode, gdipBrush) <> 0 Then
                    Debug.Print "WARNING!  pdGradient.getBrushHandle failed during line brush creation."
                End If
                
                'Next, override the brush colors with our custom array
                If GdipSetLinePresetBlend(gdipBrush, VarPtr(dstRGBA(0)), VarPtr(dstPosition(0)), m_NumOfPoints + gpOffset) <> 0 Then
                    Debug.Print "WARNING!  pdGradient.getBrushHandle failed during line preset blending."
                End If
        
            'At present, all other gradient types are constructed via reference path
            Case Else
            
                'Reset our graphics path reference
                m_Path.resetPath
                
                'Some paths require a modified rect (e.g. square vs rectangle, circle vs ellipse), so construct a square version
                ' of the destination rect in advance.
                Dim squareRect As RECTF, whDiff As Single, diagDistance As Single, whDiagDiff As Single
                whDiff = dstRect.Width - dstRect.Height
                
                If whDiff >= 0 Then
                    squareRect.Width = dstRect.Width
                    squareRect.Height = dstRect.Width
                    squareRect.Left = dstRect.Left
                    squareRect.Top = dstRect.Top - (whDiff / 2)
                Else
                    squareRect.Width = dstRect.Height
                    squareRect.Height = dstRect.Height
                    squareRect.Left = dstRect.Left + (whDiff / 2)
                    squareRect.Top = dstRect.Top
                End If
                
                'We now have a square rect centered around the original path's center point.
                
                'Construct a reference path, using the 1) target rect, and 2) gradient shape to determine path behavior
                If m_GradientShape = pdgs_ShapeRadial Then
                    
                    'Expand the path again, to ensure that corner regions are covered by the brush
                    diagDistance = Sqr((dstRect.Width / 2) ^ 2 + (dstRect.Height / 2) ^ 2)
                    whDiagDiff = diagDistance - (squareRect.Width / 2) + 1
                    
                    With squareRect
                        .Left = .Left - whDiagDiff '/ 2
                        .Top = .Top - whDiagDiff '/ 2
                        .Width = .Width + whDiagDiff * 2
                        .Height = .Height + whDiagDiff * 2
                    End With
                    
                    m_Path.addEllipse_RectF squareRect
                
                ElseIf m_GradientShape = pdgs_ShapeRectangle Then
                    m_Path.addRectangle_RectF dstRect
                
                ElseIf m_GradientShape = pdgs_ShapeDiamond Then
                    
                    'First, expand the square path by its diagonal distance
                    diagDistance = Sqr((squareRect.Width / 2) ^ 2 + (squareRect.Height / 2) ^ 2)
                    whDiagDiff = diagDistance - (squareRect.Width / 2)
                    
                    With squareRect
                        .Left = .Left - whDiagDiff / 2
                        .Top = .Top - whDiagDiff / 2
                        .Width = .Width + whDiagDiff
                        .Height = .Height + whDiagDiff
                    End With
                    
                    'Use the expanded rect to construct a new diamond path
                    m_Path.addRectangle_RectF squareRect
                    m_Path.rotatePathAroundItsCenter 45#
                
                End If
                
                'Create a path gradient brush, using our path object as the reference
                If GdipCreatePathGradientFromPath(m_Path.getCopyOfPathHandle, gdipBrush) <> 0 Then
                    Debug.Print "WARNING!  pdGradient.getBrushHandle failed during path brush creation."
                End If
                
                'Next, override the brush colors with our custom array
                If GdipSetPathGradientPresetBlend(gdipBrush, VarPtr(dstRGBA(0)), VarPtr(dstPosition(0)), m_NumOfPoints + gpOffset) <> 0 Then
                    Debug.Print "WARNING!  pdGradient.getBrushHandle failed during pathpreset blending."
                End If
                
                'Finally, set any/all custom parameters (e.g. wrap mode)
                GdipSetPathGradientWrapMode gdipBrush, m_GradientWrapMode
        
        End Select
        
    End If
    
    'Return the brush!  (Note that the caller is responsible for freeing the brush when done.)
    getBrushHandle = gdipBrush
    
End Function

'Given an index into the gradient point array, return a merged RGBA value using the opacity.
' (This is inefficient but it doesn't matter as it's only called a handful of times.)
Private Function getMergedRGBA(ByVal gpIndex As Long, Optional ByVal customOpacityModifier As Single = 1#) As Long
    
    Dim dstQuad As RGBQUAD
    dstQuad.Red = ExtractR(m_GradientPoints(gpIndex).pdgp_RGB)
    dstQuad.Green = ExtractG(m_GradientPoints(gpIndex).pdgp_RGB)
    dstQuad.Blue = ExtractB(m_GradientPoints(gpIndex).pdgp_RGB)
    dstQuad.Alpha = (m_GradientPoints(gpIndex).pdgp_Opacity * customOpacityModifier) * 255#
    
    Dim placeHolder As tmpLong
    LSet placeHolder = dstQuad
    
    getMergedRGBA = placeHolder.lngResult
    
End Function

Private Sub Class_Initialize()
    
    'Prep a string serializer (for storing/loading gradients externally)
    Set cSerialize = New pdParamXML
    
    'Prep a path (for constructing non-linear gradients)
    Set m_Path = New pdGraphicsPath
    
    'Assume sorting has not taken place
    m_IsSorted = False
    
    'Set all other default parameters
    m_GradientShape = pdgs_ShapeLinear
    m_GradientAngle = 0
    m_GradientWrapMode = WrapModeTile
    
End Sub

