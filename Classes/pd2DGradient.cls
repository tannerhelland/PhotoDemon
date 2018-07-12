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
'Copyright 2015-2018 by Tanner Helland
'Created: 23/July/15
'Last updated: 08/October/17
'Last update: rework the way gradient nodes are serialized to XML, to improve performance
'
'This class manages a single PD gradient instance.
'
'At present, it is primarily based on the capabilities of GDI+ gradient brushes.  This will likely expand in the future,
' but GDI+ provides a nice baseline feature set.
'
'Also, note that unlike other pd2D classes, this class does not expose a "debug mode" parameter.  While you can use this class
' to create and edit gradients directly, I *strongly* recommend handling actual gradient brush creation through the pd2DBrush class.
' It is equipped to properly create and destroy brush handles automatically, so you don't have to worry about leaks.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private m_GradientShape As PD_2D_GradientShape
Private m_GradientAngle As Single
Private m_GradientWrapMode As PD_2D_WrapMode

'Number of points in the current gradient
Private m_NumOfPoints As Long

'Collection of individual gradient nodes; these define a color, opacity, and position for each node
Private m_GradientPoints() As GradientPoint

'Are the points currently sorted from lowest-to-highest?  If they are, this will be set to TRUE; this allows us to skip sorting
' prior to constructing the actual brush.
Private m_IsSorted As Boolean

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

'Get/set individual settings.  Note that these just wrap the generic Get/SetGradientProperty functions, below.
Friend Function GetGradientShape() As PD_2D_GradientShape
    GetGradientShape = Me.GetGradientProperty(P2_GradientShape)
End Function

Friend Sub SetGradientShape(ByVal newShape As PD_2D_GradientShape)
    SetGradientProperty P2_GradientShape, newShape
End Sub

Friend Function GetGradientAngle() As Single
    GetGradientAngle = Me.GetGradientProperty(P2_GradientAngle)
End Function

Friend Sub SetGradientAngle(ByVal newAngle As Single)
    SetGradientProperty P2_GradientAngle, newAngle
End Sub

Friend Function GetGradientWrapMode() As PD_2D_WrapMode
    GetGradientWrapMode = Me.GetGradientProperty(P2_GradientWrapMode)
End Function

Friend Sub SetGradientWrapMode(ByVal newWrapMode As PD_2D_WrapMode)
    SetGradientProperty P2_GradientWrapMode, newWrapMode
End Sub

'It's a little confusing, but a gradient string is actually comprised of two portions:
' 1) Overall gradient settings (shape, angle, etc)
' 2) A list of gradient nodes (including count, node colors, and node positions)
'
'For convenience to callers, it's important that (2) be allowed to behave as a standalone property, e.g. something that
' you can set or read in a single pass.  That's the purpose of these get/set functions.
'
'(Note that in order to maintain a consistent naming convention, this public-facing GetGradientNodes() function wraps
' a similarly named internal function; the internal function does the heavy lifting.)
Friend Function GetGradientNodes() As String
    GetGradientNodes = Me.GetGradientProperty(P2_GradientNodes)
End Function

Friend Sub SetGradientNodes(ByRef srcString As String)
    SetGradientProperty P2_GradientNodes, srcString
End Sub

'Get/set individual settings by enum type
Friend Function GetGradientProperty(ByVal propID As PD_2D_GRADIENT_SETTINGS) As Variant
    
    Select Case propID
        
        Case P2_GradientShape
            GetGradientProperty = m_GradientShape
            
        Case P2_GradientAngle
            GetGradientProperty = m_GradientAngle
            
        Case P2_GradientWrapMode
            GetGradientProperty = m_GradientWrapMode
            
        Case P2_GradientNodes
            GetGradientProperty = GetGradientNodesAsString
            
    End Select
    
End Function

Friend Sub SetGradientProperty(ByVal propID As PD_2D_GRADIENT_SETTINGS, ByVal newSetting As Variant)
    
    Select Case propID
        
        Case P2_GradientShape
            m_GradientShape = CLng(newSetting)
            
        Case P2_GradientAngle
            m_GradientAngle = CSng(newSetting)
            
        Case P2_GradientWrapMode
            m_GradientWrapMode = CLng(newSetting)
            
        Case P2_GradientNodes
            SetGradientNodesFromString CStr(newSetting)
            m_IsSorted = False
            
    End Select
    
    m_GradientIsDefault = False
    
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
            .AddParam "GradientListAlreadySorted", m_IsSorted, True, True
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
                    iString = Trim$(Str(i))
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
            m_IsSorted = .GetBool("GradientListAlreadySorted", False)
                
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
                        iString = Trim$(Str(i))
                        m_GradientPoints(i).PointRGB = .GetLong("GP_RGB_" & iString, vbBlack)
                        m_GradientPoints(i).PointOpacity = .GetDouble("GP_Opacity_" & iString, 100#)
                        m_GradientPoints(i).PointPosition = .GetDouble("GP_Position_" & iString, 0#)
                    Next i
                    
                End If
                
            Else
                
                For i = 0 To m_NumOfPoints - 1
                    iString = Trim$(Str(i))
                    m_GradientPoints(i).PointRGB = .GetLong("GradientPoint_" & iString & "_RGB", vbBlack)
                    m_GradientPoints(i).PointOpacity = .GetDouble("GradientPoint_" & iString & "_Opacity", 100#)
                    m_GradientPoints(i).PointPosition = .GetDouble("GradientPoint_" & iString & "_Position", i / m_NumOfPoints)
                Next i
            
            End If
            
        End With
        
    End If
    
End Sub

'For interop purposes, gradients are passed around PD as XML strings.  To improve performance, you can specify
' an optional parameter to use Base64 when encoding node data (which can be quite large).  For small gradients,
' Base64 won't gain you anything, but for huge gradients with many data points, the gains can be large.
Friend Function GetGradientAsString(Optional ByVal useBase64ForNodes As Boolean = True) As String
    
    If (Not m_GradientIsDefault) Then
        
        If (m_Serializer Is Nothing) Then Set m_Serializer = New pdParamXML
        With m_Serializer
            .Reset 1#
            
            'Add whole-gradient parameters
            .AddParam "GradientShape", m_GradientShape
            .AddParam "GradientAngle", m_GradientAngle
            .AddParam "GradientWrapMode", m_GradientWrapMode
            
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
            m_GradientAngle = .GetDouble("GradientAngle", 0)
            m_GradientWrapMode = .GetLong("GradientWrapMode", P2_WM_Tile)
            
            'There are several possible options for gradient storage:
            ' 1) New versions of PD stick gradient nodes into their own XML entry.  Retrieve these and pass them off to a
            '    dedicated node parsing function.
            If m_Serializer.DoesParamExist("GradientNodes") Then
                SetGradientNodesFromString .GetString("GradientNodes", vbNullString)
                
            ' 2) Old versions of PD stored bare node data right there in the main XML string.  Parse them manually.
            Else
            
                m_NumOfPoints = .GetLong("GradientPointCount", 0)
                m_IsSorted = .GetBool("GradientListAlreadySorted", False)
            
                ReDim m_GradientPoints(0 To m_NumOfPoints) As GradientPoint
                
                Dim i As Long, iString As String
                For i = 0 To m_NumOfPoints - 1
                    iString = Trim$(Str(i))
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
    
    'Assume the incoming array is not sorted, then sort it
    m_IsSorted = False
    SortGradientArray
    
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
    
    m_IsSorted = True
    
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
    
    If (secondColorPosition = 0.5) Then
        m_IsSorted = True
    Else
        m_IsSorted = False
        SortGradientArray
    End If
    
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
    m_GradientAngle = 0
    m_GradientWrapMode = P2_WM_Tile
    
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

'Sort the gradient array in ascending order.  This greatly simplifies the process of creating a matching brush.
' (TODO: optimize the algorithm to pre-check for a sorted array; most of the time, that will be the case, so we can
'        skip the sorting step entirely.)
Private Sub SortGradientArray()
    
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    'If the array is already sorted, ignore this request
    If m_IsSorted Then Exit Sub
    
    'If the array only contains one node, ignore this request.
    If (m_NumOfPoints > 1) Then
    
        'Small node counts can use a simple insertion sort.
        If (m_NumOfPoints < 20) Then
        
            Dim i As Long, j As Long
            Dim tmpSort As GradientPoint, searchCont As Boolean
            i = 1
            
            Do While (i < m_NumOfPoints)
                tmpSort = m_GradientPoints(i)
                j = i - 1
                
                'Because VB6 doesn't short-circuit And statements, we have to split this check into separate parts.
                searchCont = False
                If (j >= 0) Then searchCont = (m_GradientPoints(j).PointPosition > tmpSort.PointPosition)
                
                Do While searchCont
                    m_GradientPoints(j + 1) = m_GradientPoints(j)
                    j = j - 1
                    searchCont = False
                    If (j >= 0) Then searchCont = (m_GradientPoints(j).PointPosition > tmpSort.PointPosition)
                Loop
                
                m_GradientPoints(j + 1) = tmpSort
                i = i + 1
                
            Loop
        
        'For larger node counts, switch to QuickSort.
        Else
            QuickSortNodes
        End If
        
    End If
    
    'Mark the array as sorted
    m_IsSorted = True
    
    'Want timing reports?  Here you go:
    'Debug.Print "Gradient array sorted in " & VBHacks.GetTimeDiffNowAsString(startTime)

End Sub

'Use QuickSort to sort gradient nodes.  This becomes important as the number of nodes increases.
Private Sub QuickSortNodes()
    QSInner m_GradientPoints, 0, m_NumOfPoints - 1
End Sub

Private Sub QSInner(ByRef srcGradientNodes() As GradientPoint, ByVal lowVal As Long, ByVal highVal As Long)
    
    'Ignore the search request if the bounds are mismatched
    If (lowVal < highVal) Then
        
        'Sort some sub-portion of the list, and use the returned pivot to repeat the sort process
        Dim j As Long
        j = QSPartition(srcGradientNodes, lowVal, highVal)
        QSInner srcGradientNodes, lowVal, j - 1
        QSInner srcGradientNodes, j + 1, highVal
        
    End If
    
End Sub

Private Function QSPartition(ByRef srcGradientNodes() As GradientPoint, ByVal lowVal As Long, ByVal highVal As Long) As Long
    
    Dim i As Long, j As Long
    i = lowVal
    j = highVal + 1
    
    Dim v As Single
    v = srcGradientNodes(lowVal).PointPosition
    
    Dim tmpSort As GradientPoint
    
    Do
        
        'Compare the pivot against points beneath it
        Do
            i = i + 1
            If (i = highVal) Then Exit Do
        Loop While (srcGradientNodes(i).PointPosition < v)
        
        'Compare the pivot against points above it
        Do
            j = j - 1
        Loop While (v < srcGradientNodes(j).PointPosition)
        
        'When the pivot arrives at its final location, exit
        If (i >= j) Then Exit Do
        
        'Swap the values at indexes i and j
        tmpSort = srcGradientNodes(j)
        srcGradientNodes(j) = srcGradientNodes(i)
        srcGradientNodes(i) = tmpSort
        
    Loop
    
    'Move the pivot value into its final location
    tmpSort = srcGradientNodes(j)
    srcGradientNodes(j) = srcGradientNodes(lowVal)
    srcGradientNodes(lowVal) = tmpSort
    
    'Return the pivot's final position
    QSPartition = j
    
End Function

'Once this class is populated correctly, you can use this function to retrieve a matching brush handle.  (At present, only GDI+
' is supported as a backend.)
'
'Note that a rect is *required* in order to size the gradient region correctly.  (The rect defines where the gradient starts and ends.)
'
'As a convenience, if you only want a default linear gradient (as used in the UI, among other things), set forceToLinearMode to TRUE.
' This will return a linear gradient brush at angle zero, suitable for previews or any kind of "gradient editor" UI, without modifying
' any of the current gradient settings.
'
'Finally, the caller is obviously responsible for freeing the handle when done.  If you interface with this class via pd2DBrush,
' it will take care of that automatically - but if you access this class manually, make sure to free the brush using the
' correct backend-specific function!
Friend Function GetBrushHandle(ByRef dstRect As RectF, Optional ByVal forceToLinearMode As Boolean = False, Optional ByVal customOpacityModifier As Single = 1!) As Long
    
    'Make sure a gradient has been created
    If (m_NumOfPoints <= 0) Then
        GetBrushHandle = 0
        Exit Function
    End If
    
    'Start by sorting the point array from lowest to highest.
    If (Not m_IsSorted) Then SortGradientArray
    
    'Next, we need to convert the gradient points into two separate arrays: one with merged RGBA values, and one with positions
    ' (as floating-point values on the range [0, 1]).  GDI+ accepts these as pointers to separate arrays, rather than a unified type.
    ' (As part of this process, we will also insert discrete points at position 0 and position 1 if they don't already exist.)
    Dim dstRGBA() As Long, dstPosition() As Single
    ReDim dstRGBA(0 To m_NumOfPoints + 1) As Long
    ReDim dstPosition(0 To m_NumOfPoints + 1) As Single
    
    Dim gpOffset As Long
    gpOffset = 0
    
    'Creating a discrete point at position 0.0 if it doesn't already exist
    If (m_GradientPoints(0).PointPosition <> 0) Then
        gpOffset = 1
        dstRGBA(0) = GetMergedRGBA(0, customOpacityModifier)
        dstPosition(0) = 0
    End If
    
    'Next, copy all sorted values into their destination array positions
    Dim i As Long
    For i = 0 To m_NumOfPoints - 1
        dstRGBA(i + gpOffset) = GetMergedRGBA(i, customOpacityModifier)
        dstPosition(i + gpOffset) = m_GradientPoints(i).PointPosition
    Next i
    
    'Finally, see if we need to add a discrete point at position 1.0
    If m_GradientPoints(m_NumOfPoints - 1).PointPosition <> 1 Then
        gpOffset = gpOffset + 1
        dstRGBA(m_NumOfPoints - 1 + gpOffset) = GetMergedRGBA(m_NumOfPoints - 1, customOpacityModifier)
        dstPosition(m_NumOfPoints - 1 + gpOffset) = 1
    End If
    
    'The point arrays are now successfully constructed.
    
    'Before proceeding, note that the "reflection" gradient type requires one additional step (it's a custom-built feature).
    ' If active, we must manually enlarge the point array, and fill the top half with mirror copies of the existing gradient.
    Dim finalPointCount As Long
    If (m_GradientShape = P2_GS_Reflection) And (Not forceToLinearMode) Then
        
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
    ' (IMPORTANT NOTE: at present, only GDI+ brushes are created.  Future backends will need to split out their behavior here.)
    Dim gdipBrush As Long
    
    'First: the user can request a default, angle = 0 linear version of the gradient for UI purposes.  Handle that case separately.
    If forceToLinearMode Then
    
        'Create a default linear gradient brush.  Note that GDI+ only allows you to supply two colors, so we'll just supply
        ' two dummy colors.
        gdipBrush = GDI_Plus.GetGDIPlusLinearBrushHandle(dstRect, dstRGBA(0), dstRGBA(1), 0#, 1, P2_WM_TileFlipXY)
        If (gdipBrush = 0) Then
            Debug.Print "WARNING!  pd2DGradient.GetBrushHandle failed during line brush creation."
        Else
        
            'Now that we have a brush handle, override the dummy colors with our full color array
            If Not GDI_Plus.OverrideGDIPlusLinearGradient(gdipBrush, VarPtr(dstRGBA(0)), VarPtr(dstPosition(0)), finalPointCount) Then Debug.Print "WARNING!  pd2DGradient.GetBrushHandle failed to override the default gradient brush colors."
        End If
        
    Else
        
        'Brush creation varies by shape
        Select Case m_GradientShape
        
            Case P2_GS_Linear, P2_GS_Reflection
                gdipBrush = GDI_Plus.GetGDIPlusLinearBrushHandle(dstRect, dstRGBA(0), dstRGBA(1), m_GradientAngle, 1, m_GradientWrapMode)
                If (gdipBrush = 0) Then
                    Debug.Print "WARNING!  pd2DGradient.GetBrushHandle failed during line brush creation."
                    Debug.Print "Debug info: " & Me.GetGradientAsString(False)
                Else
                    If Not GDI_Plus.OverrideGDIPlusLinearGradient(gdipBrush, VarPtr(dstRGBA(0)), VarPtr(dstPosition(0)), finalPointCount) Then Debug.Print "WARNING!  pd2DGradient.GetBrushHandle failed to override the default gradient brush colors."
                End If
                
            'At present, all other gradient types are constructed as "path gradients".  These gradients require a GDI+ GraphicsPath
            ' object that defines the external border of the gradient.
            Case Else
            
                'Reset our graphics path reference
                If (m_Path Is Nothing) Then Set m_Path = New pd2DPath Else m_Path.ResetPath
                
                'Some paths require a modified rect (e.g. square vs rectangle, circle vs ellipse), so construct a few additional
                ' measurements in advance.
                Dim halfWidth As Double, halfHeight As Double
                halfWidth = dstRect.Width / 2
                halfHeight = dstRect.Height / 2
                
                Dim centerX As Double, centerY As Double
                centerX = dstRect.Left + halfWidth
                centerY = dstRect.Top + halfHeight
                
                Dim squareRect As RectF, maxLength As Double, radiusRect As Double, radiusSquare As Double
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
                If (m_GradientShape = P2_GS_Radial) Then
                    
                    'Enclose the gradient within a circle with the same radius as the bounding rect's diagonal
                    m_Path.AddEllipse_Absolute centerX - radiusRect, centerY - radiusRect, centerX + radiusRect, centerY + radiusRect
                
                ElseIf (m_GradientShape = P2_GS_Rectangle) Then
                    
                    'Rectangles are easy - just use the bounding rect itself!
                    m_Path.AddRectangle_RectF dstRect
                    
                ElseIf (m_GradientShape = P2_GS_Diamond) Then
                    
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
                    
                    'Rotate the square 45 degrees and bam, we have a perfectly constructed diamond region!
                    m_Path.AddRectangle_RectF squareRect
                    m_Path.RotatePathAroundItsCenter 45#
                    
                End If
                
                'Create a path gradient brush, using our path object as the reference
                gdipBrush = GDI_Plus.GetGDIPlusPathBrushHandle(m_Path.GetHandle)
                If (gdipBrush <> 0) Then
                    
                    'Set a default center point (in the future, we should expose this to the user)
                    GDI_Plus.SetGDIPlusPathBrushCenter gdipBrush, centerX, centerY
                    
                    'Next, override the brush colors with our custom array
                    If Not GDI_Plus.OverrideGDIPlusPathGradient(gdipBrush, VarPtr(dstRGBA(0)), VarPtr(dstPosition(0)), finalPointCount) Then Debug.Print "WARNING!  pd2DGradient.GetBrushHandle failed to override the default gradient brush colors."
                    
                    'Finally, set any/all custom parameters (e.g. wrap mode)
                    GDI_Plus.SetGDIPlusPathBrushWrap gdipBrush, m_GradientWrapMode
                
                Else
                    Debug.Print "WARNING!  pd2DGradient.GetBrushHandle failed to create a valid path brush handle."
                End If
                
        End Select
        
    End If
    
    'Return the brush!  (Note that the caller is responsible for freeing the brush when done.)
    GetBrushHandle = gdipBrush
    
End Function

'Given an index into the gradient point array, return a merged RGBA value using the opacity.
' (This is inefficient but it doesn't matter as it's only called a handful of times.)
Private Function GetMergedRGBA(ByVal gpIndex As Long, Optional ByVal customOpacityModifier As Single = 1!) As Long
    
    Dim dstQuad As RGBQuad
    dstQuad.Red = Colors.ExtractRed(m_GradientPoints(gpIndex).PointRGB)
    dstQuad.Green = Colors.ExtractGreen(m_GradientPoints(gpIndex).PointRGB)
    dstQuad.Blue = Colors.ExtractBlue(m_GradientPoints(gpIndex).PointRGB)
    dstQuad.Alpha = (m_GradientPoints(gpIndex).PointOpacity * customOpacityModifier) * 2.55
    
    Dim placeHolder As tmpLong
    LSet placeHolder = dstQuad
    
    GetMergedRGBA = placeHolder.lngResult
    
End Function

'Given a position on the range [0, 1], return the gradient color at that position.  Note that a linear, non-reflected
' gradient shape is always assumed, by design.
Friend Function GetColorAtPosition_RGBA(ByVal cPosition As Single, ByRef dstRGBA As RGBQuad) As Boolean

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
            Dim tmpLow As RGBQuad, tmpHigh As RGBQuad
            tmpLowLong.lngResult = GetMergedRGBA(i - 1)
            tmpHighLong.lngResult = GetMergedRGBA(i)
            
            LSet tmpLow = tmpLowLong
            LSet tmpHigh = tmpHighLong
            
            'Start blending
            dstRGBA.Red = GetBlendedColors(tmpLow.Red, tmpHigh.Red, blendFactor)
            dstRGBA.Green = GetBlendedColors(tmpLow.Green, tmpHigh.Green, blendFactor)
            dstRGBA.Blue = GetBlendedColors(tmpLow.Blue, tmpHigh.Blue, blendFactor)
            dstRGBA.Alpha = GetBlendedColors(tmpLow.Alpha, tmpHigh.Alpha, blendFactor)
        
        End If
        
        GetColorAtPosition_RGBA = True
        
    Else
        Debug.Print "WARNING!  getColorAtPosition_RGBA was called before the gradient was initialized!"
        GetColorAtPosition_RGBA = False
    End If
    

End Function

'Blend byte1 w/ byte2 based on mixRatio. mixRatio is expected to be a value between 0 and 1.
Private Function GetBlendedColors(ByVal firstColor As Byte, ByVal secondColor As Byte, ByRef mixRatio As Double) As Byte
    GetBlendedColors = ((1# - mixRatio) * firstColor) + (mixRatio * secondColor)
End Function

Private Sub Class_Initialize()
    
    'Assume sorting has not taken place
    m_IsSorted = False
    
    'Set all other default parameters
    m_GradientShape = P2_GS_Linear
    m_GradientAngle = 0
    m_GradientWrapMode = P2_WM_Tile
    
    m_GradientIsDefault = True
    
End Sub

'All pd2D classes report errors using an internal function similar to this one.  Feel free to modify this function to
' better fit your project (e.g. perhaps it could raise an actual error event).
'
'Note that a default pd2D build simply dumps the passed error information to the Immediate window.
Private Sub InternalError(Optional ByRef errName As String = vbNullString, Optional ByRef errDescription As String = vbNullString, Optional ByVal ErrNum As Long = 0)
    Drawing2D.DEBUG_NotifyExternalError errName, errDescription, ErrNum, "pd2DGradient"
End Sub
