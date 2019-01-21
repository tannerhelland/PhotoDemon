VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pd2DGradientCairo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Gradient Manager (Cairo-based)
'Copyright 2019-2019 by Tanner Helland
'Created: 03/January/19
'Last updated: 21/January/19
'Last update: correctly apply extend mode when generating a pattern handle
'
'This class manages a single PD gradient instance, with Cairo as the backend renderer.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private m_GradientShape As PD_2D_GradientShape
Private m_GradientExtend As Cairo_Extend

'Linear gradients and radial gradients both require two points.  (For a radial gradient, these points are
' often identical, unless the gradient is being "offset".)
Private m_GradientPt1 As PointFloat, m_GradientPt2 As PointFloat

'Radial gradients also require two radii values; the first one may be 0.
Private m_GradientRadius1 As Single, m_GradientRadius2 As Single

'Number of points in the current gradient
Private m_NumOfPoints As Long

'Collection of individual gradient nodes; these define a color, opacity, and position for each node
Private m_GradientPoints() As GradientPoint

'All non-linear brushes are constructed with the help of a pd2DPath object
Private m_Path As pd2DPath

'This class is capable of serializing itself to/from XML strings
Private m_Serializer As pdParamXML

'Quick and dirty memory swap APIs
Private Type tmpLong
    lngResult As Long
End Type

'When this class is instantiated, we create a default gradient object.  If this object has never been modified,
' however, we flag it to avoid things like serializing the default gradient values upon request.
Private m_GradientIsDefault As Boolean

Friend Function GetGradientExtend() As Cairo_Extend
    GetGradientExtend = m_GradientExtend
End Function

Friend Sub SetGradientExtend(ByVal newExtend As Cairo_Extend)
    m_GradientExtend = newExtend
End Sub

'Linear gradients require a start and end point
Friend Function GetGradientPoint1() As PointFloat
    GetGradientPoint1 = m_GradientPt1
End Function

Friend Sub SetGradientPoint1(ByRef srcPoint As PointFloat)
    m_GradientPt1 = srcPoint
End Sub

Friend Function GetGradientPoint2() As PointFloat
    GetGradientPoint2 = m_GradientPt2
End Function

Friend Sub SetGradientPoint2(ByRef srcPoint As PointFloat)
    m_GradientPt2 = srcPoint
End Sub

Friend Sub SetGradientRadii(ByVal radius1 As Single, ByVal radius2 As Single)
    m_GradientRadius1 = radius1
    m_GradientRadius2 = radius2
End Sub

Friend Function GetGradientShape() As PD_2D_GradientShape
    GetGradientShape = m_GradientShape
End Function

Friend Sub SetGradientShape(ByVal newShape As PD_2D_GradientShape)
    m_GradientShape = newShape
End Sub

'It's a little confusing, but a gradient string is actually comprised of two portions:
' 1) Overall gradient settings (shape, angle, etc)
' 2) A list of gradient nodes (including count, node colors, and node positions)
'
'For convenience to callers, it's important that (2) be allowed to behave as a standalone property,
' e.g. something that you can set or read in a single pass.  That's the purpose of these get/set functions.
'
'(Note that in order to maintain a consistent naming convention, this public-facing GetGradientNodes()
' function wraps a similarly named internal function; the internal function does the heavy lifting.)
Friend Function GetGradientNodes() As String
    GetGradientNodes = GetGradientNodesAsString()
End Function

Friend Sub SetGradientNodes(ByRef srcString As String)
    SetGradientNodesFromString srcString
End Sub

Private Function GetGradientNodesAsString(Optional ByVal useBase64ForNodes As Boolean = True) As String
    
    If (Not m_GradientIsDefault) Then
        
        Dim startTime As Currency
        VBHacks.GetHighResTime startTime
        
        Dim tmpXML As pdParamXML
        Set tmpXML = New pdParamXML
        tmpXML.SetParamVersion 1.1
        
        With tmpXML
            
            'Add two management parameters: the point count (which simplifies parsing), and whether we have already sorted the list.
            ' (If we have, this saves future functions from needing to perform their own sort.)
            .AddParam "GradientPointCount", m_NumOfPoints, True, True
            .AddParam "GradientNodesUseBase64", useBase64ForNodes, True, True
            
            'We now have two options for encoding actual node data.  If the caller allows us, we'll just use
            ' Base64 to dump the full node array to a string.  For very large gradients, this provides a nice
            ' speed boost over manually encoding individual nodes.
            If useBase64ForNodes Then
            
                Dim strBase64 As String
                Strings.BytesToBase64Ex VarPtr(m_GradientPoints(0)), m_NumOfPoints * LenB(m_GradientPoints(0)), strBase64
                .AddParam "GradientNodesBase64", strBase64, True, False
            
            Else
            
                'Add the gradient point list; for convenience, RGB and opacity are manually separated.
                Dim i As Long, iString As String
                For i = 0 To m_NumOfPoints - 1
                    iString = Trim$(Str$(i))
                    .AddParam "GP_RGB_" & iString, m_GradientPoints(i).PointRGB, True, True
                    .AddParam "GP_Opacity_" & iString, m_GradientPoints(i).PointOpacity, True, True
                    .AddParam "GP_Position_" & iString, m_GradientPoints(i).PointPosition, True, True
                Next i
                
            End If
            
        End With
        
        GetGradientNodesAsString = tmpXML.GetParamString()
        
        'Performance reports:
        'Debug.Print "Wrote gradient to string in " & VBHacks.GetTimeDiffNowAsString(startTime)
        
    End If
        
End Function

Private Sub SetGradientNodesFromString(ByRef srcString As String)
    
    If (LenB(srcString) = 0) Then
        Me.ResetAllProperties
        
    Else
    
        Dim tmpXML As pdParamXML
        Set tmpXML = New pdParamXML
        With tmpXML
        
            .SetParamString srcString
            
            m_NumOfPoints = .GetLong("GradientPointCount", 2)
            ReDim m_GradientPoints(0 To m_NumOfPoints) As GradientPoint
            
            Dim i As Long, iString As String
            
            'Past versions of the gradient class use slightly different class names
            If (.GetParamVersion = 1.1) Then
                
                'As of v7.0, callers can choose to save gradient data as Base64 (for a nice performance boost).
                ' Check for Base64 encoding before proceeding; if it exists, we'll obviously decode the next
                ' segment differently.
                If .GetBool("GradientNodesUseBase64", False) Then
                    
                    'Base 64 was used.  Retrieve the full tag.
                    Dim strBase64 As String
                    strBase64 = .GetString("GradientNodesBase64")
                    
                    'Translate the result directly into the target gradient points array
                    Dim bufferSize As Long
                    bufferSize = m_NumOfPoints * LenB(m_GradientPoints(0))
                    Strings.BytesFromBase64Ex VarPtr(m_GradientPoints(0)), bufferSize, strBase64
                    
                'Base64 was not used, meaning gradient nodes are encoded as a (potentially large) set of
                ' individual tags.  Parse each tag in turn.
                Else
                
                    For i = 0 To m_NumOfPoints - 1
                        iString = Trim$(Str$(i))
                        m_GradientPoints(i).PointRGB = .GetLong("GP_RGB_" & iString, vbBlack)
                        m_GradientPoints(i).PointOpacity = .GetDouble("GP_Opacity_" & iString, 100#)
                        m_GradientPoints(i).PointPosition = .GetDouble("GP_Position_" & iString, 0#)
                    Next i
                    
                End If
                
            Else
                
                For i = 0 To m_NumOfPoints - 1
                    iString = Trim$(Str$(i))
                    m_GradientPoints(i).PointRGB = .GetLong("GradientPoint_" & iString & "_RGB", vbBlack)
                    m_GradientPoints(i).PointOpacity = .GetDouble("GradientPoint_" & iString & "_Opacity", 100#)
                    m_GradientPoints(i).PointPosition = .GetDouble("GradientPoint_" & iString & "_Position", i / m_NumOfPoints)
                Next i
            
            End If
            
        End With
        
    End If
    
End Sub

'Note that the string format used by pd2DGradientCairo instances are *not* identical to regular pd2DGradient strings.
' I haven't decided exactly how to handle this yet, as PD doesn't interop between the two formats anywhere.
Friend Function GetGradientAsString(Optional ByVal useBase64ForNodes As Boolean = True) As String
    
    If (Not m_GradientIsDefault) Then
        
        If (m_Serializer Is Nothing) Then Set m_Serializer = New pdParamXML
        With m_Serializer
            .Reset 1#
            
            'Add whole-gradient parameters
            .AddParam "GradientShape", m_GradientShape
            .AddParam "GradientX1", m_GradientPt1.x
            .AddParam "GradientY1", m_GradientPt1.y
            .AddParam "GradientX2", m_GradientPt2.x
            .AddParam "GradientY2", m_GradientPt2.y
            .AddParam "GradientR1", m_GradientRadius1
            .AddParam "GradientR2", m_GradientRadius2
            .AddParam "GradientExtend", m_GradientExtend
            
            'Add the gradient point list (and associated params, like number of points) as one contiguous string
            .AddParam "GradientNodes", GetGradientNodesAsString(useBase64ForNodes), True, True
            
        End With
        
        GetGradientAsString = m_Serializer.GetParamString
        
    Else
        GetGradientAsString = vbNullString
    End If
    
End Function

Friend Sub CreateGradientFromString(ByRef srcString As String)
    
    'If the string is empty, prep a default object
    If (LenB(srcString) = 0) Then
        Me.ResetAllProperties
    Else
        
        Dim startTime As Currency
        VBHacks.GetHighResTime startTime
        
        m_GradientIsDefault = False
        
        If (m_Serializer Is Nothing) Then Set m_Serializer = New pdParamXML
        With m_Serializer
            .SetParamString srcString
            
            'Retrieve parameters whose size and count do not vary
            m_GradientShape = .GetLong("GradientShape", P2_GS_Linear)
            m_GradientPt1.x = .GetSingle("GradientX1", 0!)
            m_GradientPt1.y = .GetSingle("GradientY1", 0!)
            m_GradientPt2.x = .GetSingle("GradientX2", 0!)
            m_GradientPt2.y = .GetSingle("GradientY2", 0!)
            m_GradientRadius1 = .GetSingle("GradientR1", 0!)
            m_GradientRadius2 = .GetSingle("GradientR2", 0!)
            m_GradientExtend = .GetLong("GradientExtend", ce_ExtendPad)
            
            'There are several possible options for gradient storage:
            ' 1) New versions of PD stick gradient nodes into their own XML entry.  Retrieve these and pass them off to a
            '    dedicated node parsing function.
            If m_Serializer.DoesParamExist("GradientNodes") Then
                SetGradientNodesFromString .GetString("GradientNodes", vbNullString)
                
            ' 2) Old versions of PD stored bare node data right there in the main XML string.  Parse them manually.
            Else
            
                m_NumOfPoints = .GetLong("GradientPointCount", 0)
                
                ReDim m_GradientPoints(0 To m_NumOfPoints) As GradientPoint
                
                Dim i As Long, iString As String
                For i = 0 To m_NumOfPoints - 1
                    iString = Trim$(Str$(i))
                    m_GradientPoints(i).PointRGB = .GetLong("GradientPoint_" & iString & "_RGB", vbBlack)
                    m_GradientPoints(i).PointOpacity = .GetDouble("GradientPoint_" & iString & "_Opacity", 100#)
                    m_GradientPoints(i).PointPosition = .GetDouble("GradientPoint_" & iString & "_Position", i / m_NumOfPoints)
                Next i
                
            End If
                
        End With
        
        'Want performance reports?  Here you go:
        'Debug.Print "Created " & m_NumOfPoints & "-pt gradient from string in " & VBHacks.GetTimeDiffNowAsString(startTime)
    
    End If
    
End Sub

'This (hackish) function allows you to pass a pd2DGradient string to it, and it will attempt to create
' a matching Cairo string.  Because some gradient values are handled differently (e.g. lines are
' defined by endpoints instead of an angle and bounding rect), you will still need to manually populate
' Cairo-specific values after initializing shared properties.
Friend Sub CreateGradientFromGdipGradientString(ByRef srcString As String)
    
    'If the string is empty, prep a default object
    If (LenB(srcString) = 0) Then
        Me.ResetAllProperties
    Else
        
        Dim startTime As Currency
        VBHacks.GetHighResTime startTime
        
        m_GradientIsDefault = False
        
        If (m_Serializer Is Nothing) Then Set m_Serializer = New pdParamXML
        With m_Serializer
            .SetParamString srcString
            
            'Shape is the only non-node data currently shared between gradient instances
            m_GradientShape = .GetLong("GradientShape", P2_GS_Linear)
            
            'There are several possible options for gradient storage:
            ' 1) New versions of PD stick gradient nodes into their own XML entry.  Retrieve these and pass them off to a
            '    dedicated node parsing function.
            If m_Serializer.DoesParamExist("GradientNodes") Then
                SetGradientNodesFromString .GetString("GradientNodes", vbNullString)
                
            ' 2) Old versions of PD stored bare node data right there in the main XML string.  Parse them manually.
            Else
            
                m_NumOfPoints = .GetLong("GradientPointCount", 0)
                
                ReDim m_GradientPoints(0 To m_NumOfPoints) As GradientPoint
                
                Dim i As Long, iString As String
                For i = 0 To m_NumOfPoints - 1
                    iString = Trim$(Str$(i))
                    m_GradientPoints(i).PointRGB = .GetLong("GradientPoint_" & iString & "_RGB", vbBlack)
                    m_GradientPoints(i).PointOpacity = .GetDouble("GradientPoint_" & iString & "_Opacity", 100#)
                    m_GradientPoints(i).PointPosition = .GetDouble("GradientPoint_" & iString & "_Position", i / m_NumOfPoints)
                Next i
                
            End If
                
        End With
        
    End If
    
End Sub

'The gradient editor assembles its own list of nodes.  To simplify interaction with this class, it can simply pass its
' instances to this function.
Friend Sub CreateGradientFromPointCollection(ByVal numOfPoints As Long, ByRef srcPoints() As GradientPoint)
    
    m_GradientIsDefault = False
    
    'Start by prepping our internal collections
    If (m_NumOfPoints <> numOfPoints) Then
        m_NumOfPoints = numOfPoints
        ReDim m_GradientPoints(0 To m_NumOfPoints) As GradientPoint
    End If
    
    'Copy the source array
    CopyMemoryStrict VarPtr(m_GradientPoints(0)), VarPtr(srcPoints(0)), LenB(m_GradientPoints(0)) * m_NumOfPoints
    
End Sub

'Helper functions for quickly constructing two- and three-point gradients, without the obnoxious overhead of creating your own point
' and color arrays.
Friend Sub CreateTwoPointGradient(ByVal firstColor As Long, ByVal secondColor As Long, Optional ByVal firstOpacity As Single = 100#, Optional ByVal secondOpacity As Single = 100#)
    
    m_GradientIsDefault = False
    
    m_NumOfPoints = 2
    ReDim m_GradientPoints(0 To m_NumOfPoints) As GradientPoint
    
    m_GradientPoints(0).PointPosition = 0#
    m_GradientPoints(0).PointRGB = firstColor
    m_GradientPoints(0).PointOpacity = firstOpacity
    
    m_GradientPoints(1).PointPosition = 1#
    m_GradientPoints(1).PointRGB = secondColor
    m_GradientPoints(1).PointOpacity = secondOpacity
    
End Sub

Friend Sub CreateThreePointGradient(ByVal firstColor As Long, ByVal secondColor As Long, ByVal thirdColor As Long, Optional ByVal firstOpacity As Single = 100#, Optional ByVal secondOpacity As Single = 100#, Optional ByVal thirdOpacity As Single = 100#, Optional ByVal secondColorPosition As Single = 0.5)
    
    m_GradientIsDefault = False
    
    m_NumOfPoints = 3
    ReDim m_GradientPoints(0 To m_NumOfPoints) As GradientPoint
    
    m_GradientPoints(0).PointPosition = 0#
    m_GradientPoints(0).PointRGB = firstColor
    m_GradientPoints(0).PointOpacity = firstOpacity
    
    m_GradientPoints(1).PointPosition = secondColorPosition
    m_GradientPoints(1).PointRGB = secondColor
    m_GradientPoints(1).PointOpacity = secondOpacity
    
    m_GradientPoints(2).PointPosition = 1#
    m_GradientPoints(2).PointRGB = thirdColor
    m_GradientPoints(2).PointOpacity = thirdOpacity
    
End Sub

Friend Sub GetCopyOfPointCollection(ByRef numOfPoints As Long, ByRef srcPoints() As GradientPoint)
    If (m_NumOfPoints > 0) Then
        numOfPoints = m_NumOfPoints
        ReDim srcPoints(0 To m_NumOfPoints) As GradientPoint
        CopyMemoryStrict VarPtr(srcPoints(0)), VarPtr(m_GradientPoints(0)), LenB(m_GradientPoints(0)) * m_NumOfPoints
    End If
End Sub

Friend Function GetNumOfNodes() As Long
    GetNumOfNodes = m_NumOfPoints
End Function

Friend Sub ResetAllProperties()
    
    m_GradientShape = P2_GS_Linear
    m_GradientPt1.x = 0:    m_GradientPt1.y = 0
    m_GradientPt2.x = 0:    m_GradientPt2.y = 0
    m_GradientExtend = ce_ExtendPad
    
    m_NumOfPoints = 2
    ReDim m_GradientPoints(0 To 1) As GradientPoint
    
    With m_GradientPoints(0)
        .PointRGB = vbBlack
        .PointOpacity = 100!
        .PointPosition = 0!
    End With
    
    With m_GradientPoints(1)
        .PointRGB = vbWhite
        .PointOpacity = 100!
        .PointPosition = 1!
    End With
    
    m_GradientIsDefault = True
    
End Sub

'Once this class is populated correctly, you can use this function to retrieve a matching cairo pattern handle.
'
'As a convenience, if you only want a default linear gradient (as used in the UI, among other things),
' set forceToLinearMode to TRUE.  This will return a linear gradient brush at angle zero, suitable for
' previews or any kind of "gradient editor" UI, without modifying any of the current gradient settings.
'
'Finally, the caller is obviously responsible for freeing the handle when done.  If you interface with
' this class via pd2DPatternCairo, it will take care of that automatically - but if you access this class
' manually, make sure to free the pattern using the correct backend-specific function!
Friend Function GetPatternHandle(Optional ByVal forceToLinearMode As Boolean = False, Optional ByVal customOpacityModifier As Single = 1!) As Long
    
    Const ONE_DIV_255 As Single = 1! / 255!
    
    'Make sure a gradient has been created
    If (m_NumOfPoints <= 0) Then
        GetPatternHandle = 0
        Exit Function
    End If
    
    'Request a default pattern object, reflecting the current shape type
    If (m_GradientShape = P2_GS_Linear) Or (m_GradientShape = P2_GS_Reflection) Or forceToLinearMode Then
        GetPatternHandle = Plugin_Cairo.Pattern_CreateLinearGradient(m_GradientPt1.x, m_GradientPt1.y, m_GradientPt2.x, m_GradientPt2.y)
    ElseIf (m_GradientShape = P2_GS_Radial) Then
        GetPatternHandle = Plugin_Cairo.Pattern_CreateRadialGradient(m_GradientPt1.x, m_GradientPt1.y, m_GradientRadius1, m_GradientPt1.x, m_GradientPt1.y, m_GradientRadius2)
    ElseIf (m_GradientShape = P2_GS_Rectangle) Then
        'todo
    ElseIf (m_GradientShape = P2_GS_Diamond) Then
        'todo
    End If
    
    'Set repeat behavior
    Plugin_Cairo.Pattern_SetExtend GetPatternHandle, m_GradientExtend
    
    'Cairo needs us to submit each point in the gradient as a separate function call, and it doesn't care about order.
    
    'Note, however, that PD's "reflection" gradient type is an in-house creation, and if it's active, we need to
    ' manually assemble a "reflected" list of gradient points (while also dividing all existing points by 2).
    Dim finalPointCount As Long
    If (m_GradientShape = P2_GS_Reflection) And (Not forceToLinearMode) Then
        
        'Start by cutting all positions in half.
        Dim i As Long
        For i = 0 To m_NumOfPoints - 1
            
            With m_GradientPoints(i)
                
                .PointPosition = .PointPosition * 0.5
            
                'Add all nodes to the gradient twice: once in the halved position, and another "mirrored" across
                ' the 0.5 mark
                Plugin_Cairo.Pattern_SetStopRGBA GetPatternHandle, .PointPosition, Colors.ExtractRed(.PointRGB) * ONE_DIV_255, Colors.ExtractGreen(.PointRGB) * ONE_DIV_255, Colors.ExtractBlue(.PointRGB) * ONE_DIV_255, .PointOpacity * customOpacityModifier
                Plugin_Cairo.Pattern_SetStopRGBA GetPatternHandle, 1! - .PointPosition, Colors.ExtractRed(.PointRGB) * ONE_DIV_255, Colors.ExtractGreen(.PointRGB) * ONE_DIV_255, Colors.ExtractBlue(.PointRGB) * ONE_DIV_255, .PointOpacity * customOpacityModifier
            
            End With
            
        Next i
        
    Else
        For i = 0 To m_NumOfPoints - 1
            With m_GradientPoints(i)
                Plugin_Cairo.Pattern_SetStopRGBA GetPatternHandle, .PointPosition, Colors.ExtractRed(.PointRGB), Colors.ExtractGreen(.PointRGB), Colors.ExtractBlue(.PointRGB), .PointOpacity * 2.55 * customOpacityModifier
            End With
        Next i
    End If
    
End Function

Private Sub Class_Initialize()
    
    'Set all other default parameters
    m_GradientShape = P2_GS_Linear
    m_GradientExtend = ce_ExtendNone
    
    m_GradientIsDefault = True
    
End Sub

'All pd2D classes report errors using an internal function similar to this one.  Feel free to modify this function to
' better fit your project (e.g. perhaps it could raise an actual error event).
'
'Note that a default pd2D build simply dumps the passed error information to the Immediate window.
Private Sub InternalError(Optional ByRef errName As String = vbNullString, Optional ByRef errDescription As String = vbNullString, Optional ByVal ErrNum As Long = 0)
    Drawing2D.DEBUG_NotifyExternalError errName, errDescription, ErrNum, "pd2DGradient"
End Sub

