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

'Number of points in the current gradient
Private m_NumOfPoints As Long

'Collection of pdGradientPoints
Private m_GradientPoints() As pdGradientPoint

'Gradient settings that can be changed non-destructively
Private m_GradientAngle As Single
Private m_GradientWrapMode As WrapMode

'Are the points currently sorted from lowest-to-highest?  If they are, this will be set to TRUE; this allows us to skip sorting
' prior to constructing the actual brush.
Private m_IsSorted As Boolean

'This class is capable of serializing itself to/from XML strings
Private cSerialize As pdParamXML

'Quick memory management
Private Type tmpLong
    lngResult As Long
End Type
Private Declare Sub CopyMemoryStrict Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpDst As Long, ByVal lpSrc As Long, ByVal byteLength As Long)
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (ByVal dstPointer As Long, ByVal Length As Long, ByVal Fill As Byte)

'GDI+ brush creation
Private Declare Function GdipSetLinePresetBlend Lib "gdiplus" (ByVal hBrush As Long, ByVal ptrToFirstColor As Long, ByVal ptrToFirstPosition As Long, ByVal ptCount As Long) As Long
Private Declare Function GdipCreateLineBrushFromRectWithAngle Lib "gdiplus" (ByRef srcRect As RECTF, ByVal firstRGBA As Long, ByVal secondRGBA As Long, ByVal gradAngle As Single, ByVal isAngleScalable As Long, ByVal gradientWrapMode As WrapMode, ByRef dstLineGradientBrush As Long) As Long

'For interop purposes, gradients are passed around PD as XML strings.
Public Function getGradientAsString() As String
    
    With cSerialize
        .Reset 1#
        
        'Add any parameters whose size and count do not vary
        .addParam "GradientPointCount", m_NumOfPoints
        .addParam "GradientListAlreadySorted", m_IsSorted
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
        
        m_GradientAngle = 0
        m_GradientWrapMode = WrapModeTile
        
    Else

        With cSerialize
            .setParamString srcString
            
            'Retrieve parameters whose size and count do not vary
            m_NumOfPoints = .GetLong("GradientPointCount", 0)
            m_IsSorted = .GetBool("GradientListAlreadySorted", False)
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
' Note that a rect is required in order to set the gradient region correctly.
Friend Function getBrushHandle(ByRef dstRect As RECTF, Optional ByVal useAngleOverride As Boolean = False, Optional ByVal newAngle As Single = 0#) As Long
    
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
        dstRGBA(0) = getMergedRGBA(0)
        dstPosition(0) = 0
    End If
    
    'Next, copy all sorted values into their destination array positions
    Dim i As Long
    For i = 0 To m_NumOfPoints - 1
        dstRGBA(i + gpOffset) = getMergedRGBA(i)
        dstPosition(i + gpOffset) = m_GradientPoints(i).pdgp_Position
    Next i
    
    'Finally, see if we need to add a closing point at position 1
    If m_GradientPoints(m_NumOfPoints - 1).pdgp_Position <> 1 Then
        gpOffset = gpOffset + 1
        dstRGBA(m_NumOfPoints - 1 + gpOffset) = getMergedRGBA(m_NumOfPoints - 1)
        dstPosition(m_NumOfPoints - 1 + gpOffset) = 1
    End If
        
    'We can now create the brush!  We do this in two steps, as gradient brushes do not have a constructor that accepts
    ' a variable color array like this.
    
    'Because the user is allowed to override the gradient's internal angle (helpful for previews, among other things),
    ' calculate a final angle now.
    Dim finalAngle As Single
    If useAngleOverride Then
        finalAngle = newAngle
    Else
        finalAngle = m_GradientAngle
    End If
    
    'Start by creating a default linear gradient brush
    Dim gdipBrush As Long
    If GdipCreateLineBrushFromRectWithAngle(dstRect, dstRGBA(0), dstRGBA(m_NumOfPoints - 1 + gpOffset), finalAngle, 1, m_GradientWrapMode, gdipBrush) <> 0 Then
        Debug.Print "WARNING!  pdGradient.getBrushHandle failed during line brush creation."
    End If
    
    'Next, override the brush colors with our custom array
    If GdipSetLinePresetBlend(gdipBrush, VarPtr(dstRGBA(0)), VarPtr(dstPosition(0)), m_NumOfPoints + gpOffset) <> 0 Then
        Debug.Print "WARNING!  pdGradient.getBrushHandle failed during line preset blending."
    End If
    
    'Return the brush!  Note that the caller is responsible for freeing the brush when done.
    getBrushHandle = gdipBrush
    
End Function

'Given an index into the gradient point array, return a merged RGBA value using the opacity.
' (This is inefficient but it doesn't matter as it's only called a handful of times.)
Private Function getMergedRGBA(ByVal gpIndex As Long) As Long
    
    Dim dstQuad As RGBQUAD
    dstQuad.Red = ExtractR(m_GradientPoints(gpIndex).pdgp_RGB)
    dstQuad.Green = ExtractG(m_GradientPoints(gpIndex).pdgp_RGB)
    dstQuad.Blue = ExtractB(m_GradientPoints(gpIndex).pdgp_RGB)
    dstQuad.Alpha = m_GradientPoints(gpIndex).pdgp_Opacity * 255#
    
    Dim placeHolder As tmpLong
    LSet placeHolder = dstQuad
    
    getMergedRGBA = placeHolder.lngResult
    
End Function

Private Sub Class_Initialize()
    
    'Prep a string serializer (for storing/loading gradients externally)
    Set cSerialize = New pdParamXML
    
    'Assume sorting has not taken place
    m_IsSorted = False
    
    'Set all other default parameters
    m_GradientAngle = 0
    m_GradientWrapMode = WrapModeTile
    
End Sub

