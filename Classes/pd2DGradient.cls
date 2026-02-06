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
'Copyright 2015-2026 by Tanner Helland
'Created: 23/July/15
'Last updated: 14/January/21
'Last update: allow the user to set a custom center point for radial gradients (necessary for rendering
'              vector shape gradients in the new PSP importer)
'
'This class manages a single PD gradient instance.
'
'Unlike other pd2D classes, this class does not expose a "debug mode" parameter.  While you can use
' this class to create and edit gradients directly, I *strongly* recommend handling actual gradient
' brush creation through the pd2DBrush class. It is equipped to properly create and destroy brush
' handles automatically, so you don't have to worry about leaks.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Declare Function GdipSetPathGradientCenterPoint Lib "gdiplus" (ByVal hBrush As Long, ByRef newCenterPoint As PointFloat) As GP_Result

Private m_GradientShape As PD_2D_GradientShape
Private m_GradientAngle As Single
Private m_GradientWrapMode As PD_2D_WrapMode
Private m_GradientGammaMode As Boolean
Private m_UseCustomCenterPoint As Boolean, m_GradientCenterPoint As PointFloat

'Number of points in the current gradient
Private m_NumOfPoints As Long

'Collection of individual gradient nodes; these define a color, opacity, and position for each node
Private m_GradientPoints() As GradientPoint

'Are the points currently sorted from lowest-to-highest?  If they are, this will be set to TRUE;
' this allows us to skip sorting prior to constructing the actual brush.
Private m_IsSorted As Boolean

'Gradient name is only relevant if the gradient is loaded from file; internal gradients are *not* named
Private m_GradientName As String

'All non-linear brushes are constructed with the help of a pd2DPath object
Private m_Path As pd2DPath

'This class is capable of serializing itself to/from XML strings
Private m_Serializer As pdSerialize

'Quick and dirty memory swap APIs
Private Type tmpLong
    lngResult As Long
End Type

'When this class is instantiated, we create a default gradient object.  If this object has never been modified,
' however, we flag it to avoid things like serializing the default gradient values upon request.
Private m_GradientIsDefault As Boolean

'Get/set individual settings.  Note that these just wrap the generic Get/SetGradientProperty functions, below.
Friend Function GetGradientShape() As PD_2D_GradientShape
    GetGradientShape = m_GradientShape
End Function

Friend Sub SetGradientShape(ByVal newShape As PD_2D_GradientShape)
    m_GradientShape = newShape
    m_GradientIsDefault = False
End Sub

Friend Function GetGradientAngle() As Single
    GetGradientAngle = m_GradientAngle
End Function

Friend Sub SetGradientAngle(ByVal newAngle As Single)
    m_GradientAngle = newAngle
    m_GradientIsDefault = False
End Sub

'Only applies to radial gradients.  For convenience, center position is defined as offsets on the
' (typical) range [0, 1], where [0.5, 0.5] is the default.  Center position is *always* calculated
' against the supplied boundary rect.  (You can supply values outside [0, 1], but the class will
' always warn you about this as there are limited use-cases for this.)
Friend Function GetCenterOffset() As PointFloat
    GetCenterOffset = m_GradientCenterPoint
End Function

Friend Function SetCenterOffset(ByRef newCenterX As Single, ByVal newCenterY As Single) As Boolean
    
    Const funcName As String = "SetCenterOffset"
    
    If (newCenterX < 0!) Or (newCenterX > 1!) Then InternalError funcName, "center ratios should be on the range [0, 1]: " & newCenterX
    If (newCenterY < 0!) Or (newCenterY > 1!) Then InternalError funcName, "center ratios should be on the range [0, 1]: " & newCenterY
    If (newCenterX = 0.5!) And (newCenterY = 0.5!) Then m_UseCustomCenterPoint = False Else m_UseCustomCenterPoint = True
        
    m_GradientCenterPoint.x = newCenterX
    m_GradientCenterPoint.y = newCenterY
    
End Function

'By default, radial gradients will always be centered over the centroid of the filled shape.
Friend Sub SetCenterOffset_Default()
    m_UseCustomCenterPoint = False
    m_GradientCenterPoint.x = 0.5!
    m_GradientCenterPoint.y = 0.5!
End Sub

Friend Function GetGradientGammaMode() As Boolean
    GetGradientGammaMode = m_GradientGammaMode
End Function

Friend Sub SetGradientGammaMode(ByVal newMode As Boolean)
    m_GradientGammaMode = newMode
End Sub

Friend Function GetGradientWrapMode() As PD_2D_WrapMode
    GetGradientWrapMode = m_GradientWrapMode
End Function

Friend Sub SetGradientWrapMode(ByVal newWrapMode As PD_2D_WrapMode)
    m_GradientWrapMode = newWrapMode
    m_GradientIsDefault = False
End Sub

Friend Function GetGradientName() As String
    GetGradientName = m_GradientName
End Function

Friend Sub SetGradientName(ByRef srcName As String)
    m_GradientName = srcName
    m_GradientIsDefault = False
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
    GetGradientNodes = GetGradientNodesAsString()
End Function

Friend Sub SetGradientNodes(ByRef srcString As String)
    SetGradientNodesFromString srcString
    m_IsSorted = False
    m_GradientIsDefault = False
End Sub

Private Function GetGradientNodesAsString(Optional ByVal useBase64ForNodes As Boolean = True) As String
    
    If (Not m_GradientIsDefault) Then
        
        Dim tmpXML As pdSerialize
        Set tmpXML = New pdSerialize
        tmpXML.SetParamVersion 2#
        
        With tmpXML
            
            'Add two management parameters: the point count (which simplifies parsing), and whether we have already sorted the list.
            ' (If we have, this saves future functions from needing to perform their own sort.)
            .AddParam "num-points", m_NumOfPoints, True, True
            .AddParam "sorted", m_IsSorted, True, True
            .AddParam "use-base64", useBase64ForNodes, True, True
            
            If (m_NumOfPoints > 0) Then
                
                'We now have two options for encoding actual node data.  If the caller allows us, we'll just use
                ' Base64 to dump the full node array to a string.  For very large gradients, this provides a nice
                ' speed boost over manually encoding individual nodes.
                If useBase64ForNodes Then
                    Dim strBase64 As String
                    Strings.BytesToBase64Ex VarPtr(m_GradientPoints(0)), m_NumOfPoints * LenB(m_GradientPoints(0)), strBase64
                    .AddParam "nodes-base64", strBase64, True, False
                Else
                
                    'Add the gradient point list; for convenience, RGB and opacity are manually separated.
                    Dim i As Long, sTmp As String
                    For i = 0 To m_NumOfPoints - 1
                        sTmp = Trim$(Str$(i))
                        .AddParam "node-rgb-" & sTmp, m_GradientPoints(i).PointRGB, True, True
                        .AddParam "node-opacity-" & sTmp, m_GradientPoints(i).PointOpacity, True, True
                        .AddParam "node-pos-" & sTmp, m_GradientPoints(i).PointPosition, True, True
                    Next i
                    
                End If
                
            Else
                InternalError "GetGradientNodesAsString", "0 points in gradient " & m_GradientName
            End If
            
        End With
        
        GetGradientNodesAsString = tmpXML.GetParamString()
        
    End If
        
End Function

Private Sub SetGradientNodesFromString(ByRef srcXML As String)
    
    If (LenB(srcXML) = 0) Then
        Me.ResetAllProperties
    Else
    
        Dim tmpXML As pdSerialize
        Set tmpXML = New pdSerialize
        tmpXML.SetParamString srcXML
        
        'Check for legacy parameter structures
        If (tmpXML.GetParamVersion() < 2#) Then
            SetGradientNodesFromString_Legacy tmpXML
        
        'This is a modern XML packet
        Else
            
            With tmpXML
                
                m_NumOfPoints = .GetLong("num-points", 2, True)
                If (m_NumOfPoints <= 0) Then Exit Sub
                ReDim m_GradientPoints(0 To m_NumOfPoints) As GradientPoint
                
                m_IsSorted = .GetBool("sorted", False, True)
                
                'Internally, we don't serialize point lists to XML - instead, we just dump them to
                ' Base64 for a large perf boost.
                If .GetBool("use-base64", False, True) Then
                    
                    'Translate the result directly into the target gradient points array
                    Dim dstBufferSize As Long
                    dstBufferSize = m_NumOfPoints * LenB(m_GradientPoints(0))
                    
                    Dim strBase64 As String
                    strBase64 = .GetString("nodes-base64", vbNullString, True)
                    Strings.BytesFromBase64Ex VarPtr(m_GradientPoints(0)), dstBufferSize, strBase64
                    
                'Base64 was not used, meaning gradient nodes are encoded as a (potentially large) set of
                ' individual tags.  Parse each tag in turn.
                Else
                    
                    Dim i As Long, sTmp As String
                    For i = 0 To m_NumOfPoints - 1
                        sTmp = Trim$(Str$(i))
                        m_GradientPoints(i).PointRGB = .GetLong("node-rgb-" & sTmp, vbBlack, True)
                        m_GradientPoints(i).PointOpacity = .GetDouble("node-opacity-" & sTmp, 100#, True)
                        m_GradientPoints(i).PointPosition = .GetDouble("node-pos-" & sTmp, 0#, True)
                    Next i
                    
                End If
                
            End With
            
        '/Modern vs legacy param string
        End If
    
    '/Non-zero-length input
    End If
    
End Sub

'Legacy XML parameter handler.  Do *not* pass a null pdSerialize object to this class (it will crash).
Private Sub SetGradientNodesFromString_Legacy(ByRef tmpXML As pdSerialize)

    With tmpXML
        
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
                Dim dstBufferSize As Long
                dstBufferSize = m_NumOfPoints * LenB(m_GradientPoints(0))
                Strings.BytesFromBase64Ex VarPtr(m_GradientPoints(0)), dstBufferSize, strBase64
                
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
            
End Sub

'For interop purposes, gradients are passed around PD as XML strings.  To improve performance, you can specify
' an optional parameter to use Base64 when encoding node data (which can be quite large).  For small gradients,
' Base64 won't gain you anything, but for huge gradients with many data points, the gains can be large.
Friend Function GetGradientAsString(Optional ByVal useBase64ForNodes As Boolean = True) As String
    
    If (Not m_GradientIsDefault) Then
        
        If (m_Serializer Is Nothing) Then Set m_Serializer = New pdSerialize
        With m_Serializer
            
            .Reset 2#
            
            'Add whole-gradient parameters
            .AddParam "shape", Drawing2D.XML_GetNameOfGradientShape(m_GradientShape), True, True
            .AddParam "angle", m_GradientAngle, True, True
            .AddParam "wrap-mode", Drawing2D.XML_GetNameOfWrapMode(m_GradientWrapMode), True, True
            .AddParam "gamma-mode", m_GradientGammaMode, True, True
            If (LenB(m_GradientName) <> 0) Then .AddParam "name", m_GradientName, True
            .AddParam "custom-center", m_UseCustomCenterPoint, True, True
            If m_UseCustomCenterPoint Then
                .AddParam "custom-center-x", m_GradientCenterPoint.x
                .AddParam "custom-center-y", m_GradientCenterPoint.y
            End If
            
            'Add the gradient point list (and associated params, like number of points) as one contiguous string
            .AddParam "nodes", GetGradientNodesAsString(useBase64ForNodes), True, useBase64ForNodes
            
        End With
        
        GetGradientAsString = m_Serializer.GetParamString()
        
    Else
        GetGradientAsString = vbNullString
    End If
    
End Function

Friend Sub CreateGradientFromString(ByRef srcString As String)

    'If the string is empty, prep a default gradient object
    If (LenB(Trim$(srcString)) = 0) Then
        Me.ResetAllProperties
    Else
        
        m_GradientIsDefault = False
        
        If (m_Serializer Is Nothing) Then Set m_Serializer = New pdSerialize
        m_Serializer.SetParamString srcString
        
        'Check for legacy gradient implementations
        If (m_Serializer.GetParamVersion < 2#) Then
            CreateGradientFromString_Legacy
        Else
            
            With m_Serializer
                
                'Retrieve parameters whose size and count do not vary
                m_GradientShape = Drawing2D.XML_GetGradientShapeFromName(.GetString("shape", , True))
                m_GradientAngle = .GetDouble("angle", 0, True)
                m_GradientWrapMode = Drawing2D.XML_GetWrapModeFromName(.GetString("wrap-mode", , True))
                m_GradientName = .GetString("name", vbNullString, True)
                m_GradientGammaMode = .GetBool("gamma-mode", False, True)
                m_UseCustomCenterPoint = .GetBool("custom-center", False, True)
                If m_UseCustomCenterPoint Then
                    m_GradientCenterPoint.x = .GetSingle("custom-center-x", 0.5!, True)
                    m_GradientCenterPoint.y = .GetSingle("custom-center-y", 0.5!, True)
                Else
                    m_GradientCenterPoint.x = 0.5!
                    m_GradientCenterPoint.y = 0.5!
                End If
                
                'Retrieve the gradient node collection using a separate function
                SetGradientNodesFromString .GetString("nodes", vbNullString)
                    
            End With
        
        End If
        
        'Regardless of using the legacy or modern deserializer, perform a failsafe check for
        ' node count and reset as necessary.
        If (m_NumOfPoints < 2) Then Me.ResetAllProperties
        
    End If
    
End Sub

'Legacy (versions < 2.0) gradient deserializer.  You must ensure m_Serializer is initialized with a
' valid gradient string before calling this function!
Private Sub CreateGradientFromString_Legacy()

    With m_Serializer
        
        'Retrieve parameters whose size and count do not vary
        m_GradientShape = .GetLong("GradientShape", P2_GS_Linear)
        m_GradientAngle = .GetDouble("GradientAngle", 0)
        m_GradientWrapMode = .GetLong("GradientWrapMode", P2_WM_Tile)
        m_GradientName = .GetString("GradientName", vbNullString)
        m_GradientGammaMode = .GetBool("GradientGammaMode", False)
        Me.SetCenterOffset_Default
        
        'There are several possible options for gradient storage:
        ' 1) New versions of PD stick gradient nodes into their own XML entry.  Retrieve these and pass them off to a
        '    dedicated node parsing function.
        If .DoesParamExist("GradientNodes") Then
            SetGradientNodesFromString .GetString("GradientNodes", vbNullString)
            
        ' 2) Old versions of PD stored bare node data right there in the main XML string.  Parse them manually.
        Else
        
            m_NumOfPoints = .GetLong("GradientPointCount", 0)
            m_IsSorted = .GetBool("GradientListAlreadySorted", False)
        
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
    
End Sub

'The gradient editor assembles its own list of nodes.  To simplify interaction with this class,
' you can simply pass its instances to this function.
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
Friend Sub CreateTwoPointGradient(ByVal firstColor As Long, ByVal secondColor As Long, Optional ByVal firstOpacity As Single = 100!, Optional ByVal secondOpacity As Single = 100!)
    
    m_GradientIsDefault = False
    
    m_NumOfPoints = 2
    ReDim m_GradientPoints(0 To m_NumOfPoints) As GradientPoint
    
    m_GradientPoints(0).PointPosition = 0!
    m_GradientPoints(0).PointRGB = firstColor
    m_GradientPoints(0).PointOpacity = firstOpacity
    
    m_GradientPoints(1).PointPosition = 1!
    m_GradientPoints(1).PointRGB = secondColor
    m_GradientPoints(1).PointOpacity = secondOpacity
    
    m_IsSorted = True
    
End Sub

Friend Sub CreateThreePointGradient(ByVal firstColor As Long, ByVal secondColor As Long, ByVal thirdColor As Long, Optional ByVal firstOpacity As Single = 100!, Optional ByVal secondOpacity As Single = 100!, Optional ByVal thirdOpacity As Single = 100!, Optional ByVal secondColorPosition As Single = 0.5!)
    
    m_GradientIsDefault = False
    
    m_NumOfPoints = 3
    ReDim m_GradientPoints(0 To m_NumOfPoints) As GradientPoint
    
    m_GradientPoints(0).PointPosition = 0!
    m_GradientPoints(0).PointRGB = firstColor
    m_GradientPoints(0).PointOpacity = firstOpacity
    
    m_GradientPoints(1).PointPosition = secondColorPosition
    m_GradientPoints(1).PointRGB = secondColor
    m_GradientPoints(1).PointOpacity = secondOpacity
    
    m_GradientPoints(2).PointPosition = 1!
    m_GradientPoints(2).PointRGB = thirdColor
    m_GradientPoints(2).PointOpacity = thirdOpacity
    
    If (secondColorPosition = 0.5!) Then
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

'The Gradient Editor UI allows users to sort gradients by various criteria.  This is helpful for
' categorizing e.g. gradients with averaged "red" shades vs "blue" shades.  Such sorting requires us
' to provide average HSL values for the gradient; those are calculated here.
Friend Sub GetAverageHSL(ByRef dstH As Single, ByRef dstS As Single, ByRef dstL As Single)

    If (m_NumOfPoints > 0) Then
        
        'For each point, extract RGB values, calculate corresponding HSL values, and store running averages
        Dim srcR As Long, srcG As Long, srcB As Long
        Dim h As Double, s As Double, l As Double
        Dim hTotal As Double, sTotal As Double, lTotal As Double
        Dim i As Long
        For i = 0 To m_NumOfPoints - 1
            srcR = Colors.ExtractRed(m_GradientPoints(i).PointRGB)
            srcG = Colors.ExtractGreen(m_GradientPoints(i).PointRGB)
            srcB = Colors.ExtractBlue(m_GradientPoints(i).PointRGB)
            Colors.ImpreciseRGBtoHSL srcR, srcG, srcB, h, s, l
            hTotal = hTotal + h
            sTotal = sTotal + s
            lTotal = lTotal + l
        Next i
        
        dstH = hTotal / m_NumOfPoints
        dstS = sTotal / m_NumOfPoints
        dstL = lTotal / m_NumOfPoints
    
    End If

End Sub

Friend Function GetFirstColor() As Long
    If (m_NumOfPoints > 0) Then
        If (Not m_IsSorted) Then SortGradientArray
        VBHacks.GetMem4_Ptr VarPtr(m_GradientPoints(0)), VarPtr(GetFirstColor)
    End If
End Function

Friend Function GetLastColor() As Long
    If (m_NumOfPoints > 0) Then
        If (Not m_IsSorted) Then SortGradientArray
        VBHacks.GetMem4_Ptr VarPtr(m_GradientPoints(m_NumOfPoints - 1)), VarPtr(GetLastColor)
    End If
End Function

Friend Function GetLastColorQuad() As RGBQuad
    If (m_NumOfPoints > 0) Then
        If (Not m_IsSorted) Then SortGradientArray
        With GetLastColorQuad
            .Red = Colors.ExtractRed(m_GradientPoints(m_NumOfPoints - 1).PointRGB)
            .Green = Colors.ExtractGreen(m_GradientPoints(m_NumOfPoints - 1).PointRGB)
            .Blue = Colors.ExtractBlue(m_GradientPoints(m_NumOfPoints - 1).PointRGB)
            .Alpha = Int(m_GradientPoints(m_NumOfPoints - 1).PointOpacity * 2.55 + 0.5)
        End With
    End If
End Function

Friend Function GetNumOfNodes() As Long
    GetNumOfNodes = m_NumOfPoints
End Function

Friend Function LoadGradientFromFile(ByRef srcFilename As String) As Boolean
    
    'First, we need to wrap a stream object around the source file.  If the source file is svgz,
    ' we first want to transparently decompress the file into a buffer, *then* wrap that buffer.
    ' (At present, svgz files remain TODO!)
    Dim cStream As New pdStream
    Set cStream = New pdStream
    If cStream.StartStream(PD_SM_FileBacked, PD_SA_ReadOnly, srcFilename) Then
    
        Me.ResetAllProperties
        
        'Attempt SVG first
        If Strings.StringsEqual(Files.FileGetExtension(srcFilename), "svg", True) Then LoadGradientFromFile = LoadGradient_SVG(cStream)
        
        'Next, attempt GIMP format
        If (Not LoadGradientFromFile) Then LoadGradientFromFile = LoadGradient_GIMP(cStream)
        
        'Gradients loaded from file should *not* use gamma-corrected blending
        If LoadGradientFromFile Then
            m_GradientGammaMode = False
            m_GradientIsDefault = False
        End If
        
    Else
        InternalError "LoadGradientFromFile", "failed to start stream"
        LoadGradientFromFile = False
    End If

End Function

'Given a valid path to a GIMP-format .ggr file, populate this gradient object with the saved color and position list
' (Make sure to point the stream at a file or memory location before calling this function.)
Private Function LoadGradient_GIMP(ByRef srcStream As pdStream) As Boolean
    
    On Error GoTo InvalidGradient
    If (srcStream Is Nothing) Then Exit Function
    
    'Because GIMP gradients are just text files, it's easier to parse them as strings
    srcStream.SetPosition 0, FILE_BEGIN
    
    Dim rawFileString As String
    rawFileString = srcStream.ReadString_UnknownEncoding(srcStream.GetStreamSize())
    
    If (LenB(rawFileString) <> 0) Then
        
        'GIMP gradient files always start with the text "GIMP Gradient"
        If Strings.StringsEqual(Left$(rawFileString, 13), "GIMP Gradient", True) Then
            
            'This appears to be a valid GIMP gradient file.  Hypothetically, line order should be fixed,
            ' with nodes listed in ascending order, but this function parses the file as if line order
            ' is *not* fixed.  Let me know if you encounter a file in the wild where this approach fails.
            
            'To simplify processing, split the string by lines.
            Dim fileLines As pdStringStack
            Set fileLines = New pdStringStack
            fileLines.CreateFromMultilineString rawFileString
            
            Dim curLine As String
            Const SPACE_CHAR As String = " "
            
            'Do some basic sanity tests.  We already checked line 0 for the "GIMP Gradient" ID,
            ' so the next line should contain a gradient name.
            curLine = fileLines.GetString(1)
            If Strings.StringsEqual(Left$(curLine, 5), "Name:", True) Then Me.SetGradientName Trim$(Right$(curLine, Len(curLine) - 5))
            
            'The next line is the number of gradient nodes in the file.
            Dim numOfGimpNodes As Long
            curLine = Trim$(fileLines.GetString(2))
            numOfGimpNodes = CLng(curLine)
            
            'Make sure at least one node is listed; less than this is (obviously) invalid, but just one node is
            ' okay because GIMP defines gradient nodes using *both* a left and right endpoint.
            If (numOfGimpNodes >= 1) Then
                
                'The selection of this particular "magic" node number is explained below, in the description
                ' of GIMP gradient node behavior (which is quite different from ours).
                ReDim m_GradientPoints(0 To numOfGimpNodes * 2 + 1) As GradientPoint
                m_NumOfPoints = 0
                
                Dim srcColorText() As String, srcColorFloat() As Double
                ReDim srcColorFloat(0) As Double
                
                'Parse each node line in turn.  (Normally we would pop lines like a stack, but as gradient
                ' nodes are already likely to be in order, it speeds up the subsequent sort process if we
                ' preserve existing line order - so instead of popping lines, we iterate them in order.)
                Dim i As Long, j As Long
                For i = 0 To numOfGimpNodes - 1
                    
                    'Translate i to a usable index into the string table (the first 3 lines were already parsed, remember?)
                    curLine = Trim$(fileLines.GetString(3 + i))
                    
                    'AFAIK, there is no formal GIMP spec for gradient files.  One thing we want to standardize
                    ' (to simplify parsing) is replacing tab chars with space chars; VB's lack of a generic
                    ' "whitespace" identifier makes this approach the least of several evils.
                    If (InStr(1, curLine, vbTab, vbBinaryCompare) <> 0) Then curLine = Replace$(curLine, vbTab, SPACE_CHAR, , , vbBinaryCompare)
                    
                    'Empty lines, if encountered, can be ignored
                    If (LenB(curLine) = 0) Then
                        'Do nothing
                    
                    'Comment lines in other GIMP files start with a #; to be safe, check for these and ignore
                    ElseIf Strings.StringsEqual(Left$(curLine, 1), "#", False) Then
                        'Do nothing
                    
                    'Other lines should only be color descriptors
                    Else
                        
                        'Color descriptor lines contain a minimum of 11 floating-point values, and possible 13 (or more)
                        ' values, with the extra values being GIMP-specific IDs indicating type (the shape of the blend
                        ' between two nodes) and color mode (RGB, HSV counter-clockwise, and HSV clockwise).  We do not
                        ' currently attempt to process these extra values, although it could be added if enough users
                        ' request it.
                        
                        'Here is an example of a typical GIMP gradient color line:
                        '0.386835 0.430897 0.470785 0.977143 0.634945 0.718332 1.000000 0.391757 0.449753 0.484848 1.000000 0 0
                        
                        'The first three values are the gradient's left-point, mid-point, and right-point.
                        'The next four values are RGBA values for the left end-point
                        'The next four values are RGBA values for the right end-point
                        
                        'This format is problematic for PhotoDemon, which behaves like most other software and defines
                        ' gradients as a series of standalone nodes, each with their own RGBA values.
                        
                        'As such, we need to translate the GIMP values into usable PhotoDemon ones.
                        
                        'The way we do this is simple:
                        '1) Create nodes for all left endpoints using left colors
                        '2) Auto-calculate midpoint color values by interpolating left/right colors,
                        '    and manually create matching nodes for them.
                        '3) Ignore right end-point colors and positions, except for the final node.
                        
                        'Start by attempting to split the line by space chars
                        srcColorText = Split(curLine, SPACE_CHAR)
                        
                        'Make sure we retrieved enough values to avoid OOB errors during parsing
                        If (UBound(srcColorText) < 10) Then
                            InternalError "LoadGradient_GIMP", "bad color found on line " & CStr(i + 3)
                            GoTo BadLineColor
                        End If
                        
                        'Next, attempt to convert all text entries to valid floating-point values on
                        ' the range [0.0, 1.0].
                        If (UBound(srcColorFloat) <> UBound(srcColorText)) Then ReDim srcColorFloat(0 To UBound(srcColorText))
                        For j = 0 To UBound(srcColorText)
                            srcColorFloat(j) = TextSupport.CDblCustom(srcColorText(j))
                            
                            'For color, opacity, and position entries, ensure the color value is valid
                            If (j <= 10) Then
                                If (srcColorFloat(j) < 0#) Or (srcColorFloat(j) > 1#) Then
                                    InternalError "LoadGradient_GIMP", "bad color found on line " & CStr(i + 3)
                                    GoTo BadLineColor
                                End If
                            End If
                        Next j
                        
                        'If we're still here, we hypothetically have usable data.  Start producing gradient nodes
                        ' from the result!  (Note that this section uses some "magic numbers" for indices into the
                        ' split line string.  If you have questions, review the sample GIMP gradient color line,
                        ' above, to see why we use the indices we do for specific values.)
                        
                        'First is the left endpoint.  This value is always generated.
                        With m_GradientPoints(m_NumOfPoints)
                            .PointPosition = srcColorFloat(0)
                            .PointRGB = RGB(Int(srcColorFloat(3) * 255# + 0.5), Int(srcColorFloat(4) * 255# + 0.5), Int(srcColorFloat(5) * 255# + 0.5))
                            .PointOpacity = srcColorFloat(6) * 100#
                        End With
                        
                        m_NumOfPoints = m_NumOfPoints + 1
                        
                        'Next is the midpoint.  If the midpoint's coordinate is anything other than the halfway
                        ' value between the left and right endpoints, we will generate it as a unique node with
                        ' its own unique position.
                        If (Abs(srcColorFloat(1) - (srcColorFloat(0) + srcColorFloat(2)) * 0.5) >= 0.000002) Then
                            
                            'Because of the way GIMP gradient segments work, gradient files may contain
                            ' segments with identical start and end positions (to forcibly set a whole
                            ' chunk of the gradient to a solid color).  Check for this state, and if found,
                            ' do not generate a manual gradient midpoint.
                            If (srcColorFloat(3) <> srcColorFloat(7)) Or (srcColorFloat(4) <> srcColorFloat(8)) Or (srcColorFloat(5) <> srcColorFloat(9)) Or (srcColorFloat(6) <> srcColorFloat(10)) Then
                            
                                Dim newR As Double, newG As Double, newB As Double, newA As Double
                                With m_GradientPoints(m_NumOfPoints)
                                    .PointPosition = srcColorFloat(1)
                                    newR = (srcColorFloat(3) + srcColorFloat(7)) * 0.5
                                    newG = (srcColorFloat(4) + srcColorFloat(8)) * 0.5
                                    newB = (srcColorFloat(5) + srcColorFloat(9)) * 0.5
                                    newA = (srcColorFloat(6) + srcColorFloat(10)) * 0.5
                                    .PointRGB = RGB(Int(newR * 255# + 0.5), Int(newG * 255# + 0.5), Int(newB * 255# + 0.5))
                                    .PointOpacity = newA * 100#
                                End With
    
                                m_NumOfPoints = m_NumOfPoints + 1
                                
                            End If

                        End If
                        
                        'If this is the last point in the file, add the final *right* endpoint to the list
                        If (i = numOfGimpNodes - 1) Then
                            With m_GradientPoints(m_NumOfPoints)
                                .PointPosition = srcColorFloat(2)
                                .PointRGB = RGB(Int(srcColorFloat(7) * 255# + 0.5), Int(srcColorFloat(8) * 255# + 0.5), Int(srcColorFloat(9) * 255# + 0.5))
                                .PointOpacity = srcColorFloat(10) * 100#
                            End With
                            m_NumOfPoints = m_NumOfPoints + 1
                        End If
                    
                    '/end check line type
                    End If
BadLineColor:
                Next i
                
                'If we haven't errored out, consider this a valid parse
                If (m_NumOfPoints >= 2) Then
                    ReDim Preserve m_GradientPoints(0 To m_NumOfPoints - 1) As GradientPoint
                    LoadGradient_GIMP = True
                Else
                    LoadGradient_GIMP = False
                End If
                
            '/end check for at least two valid nodes
            End If
            
            'Note that the current gradient is *not* guaranteed to be sorted (although it likely is,
            ' as GIMP tends to list points in ascending order).
            m_IsSorted = False
            
        '/end basic file validation; file is not a GIMP gradient
        Else
            LoadGradient_GIMP = False
        End If
    
    End If
    
    Exit Function

InvalidGradient:
    LoadGradient_GIMP = False

End Function

'Given a valid path to an SVG file with at least one defined <linearGradient> object,
' populate this gradient object with the saved color and position list.
'
' (Make sure to point the stream at a file or memory location before calling this function.)
Private Function LoadGradient_SVG(ByRef srcStream As pdStream) As Boolean
    
    On Error GoTo InvalidGradient
    If (srcStream Is Nothing) Then Exit Function
    
    'Because SVG gradients are just text files, it's easier to parse them as strings
    srcStream.SetPosition 0, FILE_BEGIN
    
    Dim rawFileString As String
    rawFileString = srcStream.ReadString_UnknownEncoding(srcStream.GetStreamSize())
    
    If (LenB(rawFileString) <> 0) Then
        
        'Validate and load the XML
        Dim xmlEngine As Object 'MSXML2.DOMDocument
        Set xmlEngine = CreateObject("MSXML2.DOMDocument")
        xmlEngine.async = False
        xmlEngine.validateOnParse = True
        
        If xmlEngine.loadXML(rawFileString) Then
        
            'The SVG file appears to be usable.  Look for the first linear gradient object
            ' in the file.  (If multiple gradients are embedded, we don't currently have a
            ' good way to deal with them - so for now, users will have to settle for loading
            ' the first stored gradient, only.  We could possibly raise a dialog in the
            ' future to allow them to select an embedded gradient, but that's a huge UI
            ' investment for a minor feature... so it'll have to wait.)
            Dim embeddedGradients As Object 'MSXML2.IXMLDOMNodeList
            Set embeddedGradients = xmlEngine.getElementsByTagName("linearGradient")
            
            'If a linear gradient wasn't found, look for a radial gradient
            If (embeddedGradients.Length <= 0) Then Set embeddedGradients = xmlEngine.getElementsByTagName("radialGradient")
            
            'Hopefully we found at least *one* gradient in the file...
            If (embeddedGradients.Length > 0) Then
            
                'We now want to iterate all "stop" attributes inside the linear gradient object
                Dim listOfStops As Object 'MSXML2.IXMLDOMNodeList
                Set listOfStops = embeddedGradients(0).childNodes
                
                'Look for a name, which is typically stored as the "id" of the gradient.
                ' (This is true for e.g. Krita, but not necessarily for Inkscape SVGs,
                '  which may just supply a generic ID like "linearGradient1234")
                On Error GoTo 0: On Error Resume Next
                m_GradientName = vbNullString
                m_GradientName = embeddedGradients(0).Attributes.getNamedItem("id").NodeValue
                On Error GoTo 0: On Error GoTo InvalidGradient
                
                'The number of child nodes are also the "upper limit" of possible node objects.
                ReDim m_GradientPoints(0 To listOfStops.Length - 1) As GradientPoint
                m_NumOfPoints = 0
                
                Dim i As Long, j As Long, k As Long
                For i = 0 To listOfStops.Length - 1
                    
                    'If this is a stop, parse its attributes and add it to the current node list
                    If Strings.StringsEqual(listOfStops(i).baseName, "stop", True) Then
                        
                        'Populate a default opacity value, in case an attribute is missing.
                        ' (Default colors and/or position values are 0.)
                        m_GradientPoints(m_NumOfPoints).PointOpacity = 100#
                        
                        On Error GoTo 0: On Error GoTo BadStopNode
                        
                        Dim stopProps As Object 'MSXML2.IXMLDOMNode
                        Dim stopAttribName As String
                        
                        For j = 0 To listOfStops(i).Attributes.Length - 1
                            
                            Set stopProps = listOfStops(i).Attributes(j)
                            stopAttribName = stopProps.nodeName
                            
                            'Offsets must always be declared as standalone attributes, but note that they
                            ' *can* be described as percents instead of fractions.
                            If Strings.StringsEqual(stopAttribName, "offset", True) Then m_GradientPoints(m_NumOfPoints).PointPosition = SVG_GetFloatFromString(stopProps.NodeValue)
                            
                            'It's easier for us if stop color and offset are stored individualy...
                            If Strings.StringsEqual(stopAttribName, "stop-color", True) Then Colors.GetColorFromString stopProps.NodeValue, m_GradientPoints(m_NumOfPoints).PointRGB
                            If Strings.StringsEqual(stopAttribName, "stop-opacity", True) Then m_GradientPoints(m_NumOfPoints).PointOpacity = SVG_GetFloatFromString(stopProps.NodeValue) * 100#
                            
                            '...but they may also be crammed into a single "style" attribute
                            If Strings.StringsEqual(stopAttribName, "style", True) Then
                                
                                'Unfortunately, I don't know if MSXML can be coerced into parsing a style
                                ' attribute for us.  It's easy enough to parse ourselves, at least.
                                
                                'Break the style element into discrete descriptors.
                                Dim styleEntries() As String, styleSubEntries() As String
                                If InStr(1, stopProps.NodeValue, ";") Then
                                    styleEntries = Split(stopProps.NodeValue, ";")
                                Else
                                    ReDim styleEntries(0) As String
                                    styleEntries(0) = stopProps.NodeValue
                                End If
                                
                                For k = 0 To UBound(styleEntries)
                                    If (LenB(styleEntries(k)) > 0) Then
                                        styleSubEntries = Split(styleEntries(k), ":")
                                        If Strings.StringsEqual(styleSubEntries(0), "stop-color", True) Then Colors.GetColorFromString styleSubEntries(1), m_GradientPoints(m_NumOfPoints).PointRGB
                                        If Strings.StringsEqual(styleSubEntries(0), "stop-opacity", True) Then m_GradientPoints(m_NumOfPoints).PointOpacity = SVG_GetFloatFromString(styleSubEntries(1)) * 100#
                                    End If
                                Next k
                                
                            End If
BadStopNode:
                            On Error GoTo 0: On Error GoTo InvalidGradient
                        Next j
                        
                        m_NumOfPoints = m_NumOfPoints + 1
                        
                    End If
                
                Next i
                
                'If we haven't errored out, and at least two valid stops were found, consider this a valid parse
                If (m_NumOfPoints >= 2) Then
                    ReDim Preserve m_GradientPoints(0 To m_NumOfPoints - 1) As GradientPoint
                    LoadGradient_SVG = True
                Else
                    InternalError "LoadGradient_SVG", "not enough valid stops"
                    LoadGradient_SVG = False
                End If
                
                'Note that the current gradient is *not* guaranteed to be sorted (although it likely is,
                ' as e.g. Inkscape tends to write nodes in ascending order).
                m_IsSorted = False
            
            '/no linear or radial gradients found
            Else
                InternalError "LoadGradient_SVG", "no gradients in file"
                LoadGradient_SVG = False
            End If
            
        '/end basic file validation; file is not SVG
        Else
            InternalError "LoadGradient_SVG", "couldn't load XML"
            LoadGradient_SVG = False
        End If
    
    '/end sanity check on file contents being non-null
    Else
        LoadGradient_SVG = False
    End If
    
    Exit Function

InvalidGradient:
    LoadGradient_SVG = False

End Function

'Gradient attributes in an SVG file can be floats or percents; this function helpfully covers
' both cases, and always returns a proper float value on the range [0.0, 1.0].
Private Function SVG_GetFloatFromString(ByVal srcValue As String) As Double
    If (InStr(1, srcValue, "%") <> 0) Then
        SVG_GetFloatFromString = TextSupport.CDblCustom(Left$(srcValue, InStr(1, srcValue, "%") - 1)) * 0.01
    Else
        SVG_GetFloatFromString = TextSupport.CDblCustom(srcValue)
    End If
    If (SVG_GetFloatFromString < 0#) Then SVG_GetFloatFromString = 0#
    If (SVG_GetFloatFromString > 1#) Then SVG_GetFloatFromString = 1#
End Function

'Save the current gradient to a GIMP-format .ggr file.  The target file, if it exists, will be forcibly overwritten - plan accordingly!
Friend Function SaveGradient_GIMP(ByRef dstFilename As String) As Boolean
    
    'Sanity checks; all possible fail states are not currently covered, so "use at your own risk", etc
    If (m_NumOfPoints <= 1) Then Exit Function
    
    'Because GIMP gradients are just text files, it's easier to assemble them as strings
    Dim dstText As pdString
    Set dstText = New pdString
    
    'GIMP gradient files always start with the text "GIMP Gradient"
    dstText.AppendLine "GIMP Gradient"
    
    'Next is the gradient name, if any
    If (LenB(m_GradientName) <> 0) Then
        dstText.AppendLine "Name: " & m_GradientName
    Else
        dstText.AppendLine "Name: " & g_Language.TranslateMessage("Unnamed gradient")
    End If
    
    'Next is the number of nodes in the gradient.  Note that GIMP differs from most photo editors because
    ' it defines gradients by a series of *segments*, not *nodes* - and each segment has its own left,
    ' midpoint, and right coordinates, plus left and right colors.  Because of this, we actually write
    ' one *less* node than we actually have, as the final node is instead stored as the right endpoint
    ' of the next-to-last segment in the file.
    dstText.AppendLine CStr(m_NumOfPoints - 1)
    
    'We now need to write each node in turn, but note that we will be writing these in *pairs*,
    ' because we must write them as *segments*, not *nodes*.  (Note also that all values are written
    ' as text representations of floating-point values, forced to 7 significant digits.)
    Dim i As Long
    For i = 0 To m_NumOfPoints - 2
        
        'First, place the left, midpoint, and right values
        dstText.Append GetGIMPStringValue(m_GradientPoints(i).PointPosition)
        dstText.Append GetGIMPStringValue((m_GradientPoints(i).PointPosition + m_GradientPoints(i + 1).PointPosition) * 0.5)
        dstText.Append GetGIMPStringValue(m_GradientPoints(i + 1).PointPosition)
        
        'Next, the RGBA values of the current point
        dstText.Append GetGIMPStringValue(CDbl(Colors.ExtractRed(m_GradientPoints(i).PointRGB)) / 255#)
        dstText.Append GetGIMPStringValue(CDbl(Colors.ExtractGreen(m_GradientPoints(i).PointRGB)) / 255#)
        dstText.Append GetGIMPStringValue(CDbl(Colors.ExtractBlue(m_GradientPoints(i).PointRGB)) / 255#)
        dstText.Append GetGIMPStringValue(CDbl(m_GradientPoints(i).PointOpacity) * 0.01)
        
        '...followed by the RGBA values of the next point
        dstText.Append GetGIMPStringValue(CDbl(Colors.ExtractRed(m_GradientPoints(i + 1).PointRGB)) / 255#)
        dstText.Append GetGIMPStringValue(CDbl(Colors.ExtractGreen(m_GradientPoints(i + 1).PointRGB)) / 255#)
        dstText.Append GetGIMPStringValue(CDbl(Colors.ExtractBlue(m_GradientPoints(i + 1).PointRGB)) / 255#)
        dstText.Append GetGIMPStringValue(CDbl(m_GradientPoints(i + 1).PointOpacity) * 0.01)
        
        'Immediately following these values are GIMP-specific type and color mode identifiers.  We can simply
        ' hard-code these.
        dstText.AppendLine "0 0"
        
    Next i
    
    'And that's that!  Dump the completed string out to file.
    SaveGradient_GIMP = Files.FileSaveAsText(dstText.ToString(), dstFilename, True, False)
    
End Function

'When writing GIMP gradient files, we have to supply specifically formatted number strings; this function
' takes care of that for you.  Note that it forcibly appends a trailing space, by design.
Private Function GetGIMPStringValue(ByVal srcValue As Double) As String
    GetGIMPStringValue = Format$(srcValue, "0.000000")
    If (InStr(1, GetGIMPStringValue, ",") <> 0) Then GetGIMPStringValue = Replace$(GetGIMPStringValue, ",", ".")
    GetGIMPStringValue = GetGIMPStringValue & " "
End Function

'Save the current gradient to a mini SVG file.  The target file, if it exists, will be forcibly overwritten - plan accordingly!
Friend Function SaveGradient_SVG(ByRef dstFilename As String) As Boolean
    
    'Sanity checks; all possible fail states are not currently covered, so "use at your own risk", etc
    If (m_NumOfPoints <= 1) Then Exit Function
    
    'Because SVG gradients are ultimately just text files, it's easier to assemble them as strings
    Dim dstText As pdString
    Set dstText = New pdString
    
    'Start by appending standard XML headers
    dstText.AppendLine "<?xml version=""1.0"" encoding=""UTF-8"" ?>"
    dstText.AppendLine "<!-- Created with PhotoDemon (https://photodemon.org) -->"
    dstText.AppendLine "<svg>"
    
    'Next, a header for the gradient object itself.  Note that some attributes are simply placeholders
    ' to ensure the gradient is a valid SVG item; photo editors do not actually use these attributes
    ' when rendering the gradient inside an app.
    dstText.Append vbTab & "<linearGradient id="""
    If (LenB(m_GradientName) <> 0) Then dstText.Append m_GradientName Else dstText.Append g_Language.TranslateMessage("Unnamed gradient")
    dstText.AppendLine """ gradientUnits=""objectBoundingBox"" spreadMethod=""pad"" >"
    
    'Write each gradient node in turn
    Dim i As Long
    For i = 0 To m_NumOfPoints - 1
        dstText.Append vbTab & vbTab & "<stop "
        dstText.Append "stop-color=""#" & Colors.GetHexStringFromRGB(m_GradientPoints(i).PointRGB) & """ "
        dstText.Append "stop-opacity=""" & TextSupport.FormatInvariant(m_GradientPoints(i).PointOpacity * 0.01, "0.######") & """ "
        dstText.Append "offset=""" & TextSupport.FormatInvariant(m_GradientPoints(i).PointPosition, "0.######") & """ "
        dstText.AppendLine "/>"
    Next i
    
    'Finish off the XML
    dstText.AppendLine vbTab & "</linearGradient>"
    dstText.AppendLine "</svg>"
    
    'And that's that!  Dump the completed string out to file.
    SaveGradient_SVG = Files.FileSaveAsText(dstText.ToString(), dstFilename, True, False)
    
    Exit Function

InvalidGradient:
    SaveGradient_SVG = False

End Function

Friend Sub ResetAllProperties()
    
    m_GradientShape = P2_GS_Linear
    m_GradientAngle = 0!
    m_GradientWrapMode = P2_WM_Tile
    m_GradientGammaMode = False
    
    m_NumOfPoints = 2
    ReDim m_GradientPoints(0 To 1) As GradientPoint
    m_IsSorted = True
    
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
    Me.SetCenterOffset_Default
    
End Sub

Friend Sub ReverseGradient()
    
    If (m_NumOfPoints < 2) Then Exit Sub
    m_GradientIsDefault = False
    
    'Make sure points are sorted
    If (Not m_IsSorted) Then SortGradientArray
    
    Dim i As Long
    For i = 0 To m_NumOfPoints - 1
        m_GradientPoints(i).PointPosition = 1! - m_GradientPoints(i).PointPosition
    Next i
    
    'Sort again
    m_IsSorted = False
    SortGradientArray
    
End Sub

'Sort the gradient array in ascending order.  This greatly simplifies the process of creating a matching brush.
' (TODO: optimize the algorithm to pre-check for a sorted array; most of the time, that will be the case, so we can
'        skip the sorting step entirely.)
Private Sub SortGradientArray()
    
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

'Once this class is populated with all desired settings, you can use this function to retrieve a
' matching brush handle.
'
'Note that a rect is *required* in order to size the gradient region correctly.  (The rect defines
' where the gradient starts and ends.)
'
'As a convenience, if you only want a default linear gradient (as used in the UI, among other things),
' set forceToLinearMode to TRUE.  This will return a linear gradient brush at angle zero, suitable for
' previews or any kind of "interactive" UI, without modifying any of the current gradient settings.
'
'Finally, the caller is obviously responsible for freeing the handle when done.  If you interface with
' this class via pd2DBrush, it will take care of that automatically - but if you access this class manually,
' make sure to free the brush using the correct backend-specific function!
Friend Function GetBrushHandle(ByRef dstRect As RectF, Optional ByVal forceToLinearMode As Boolean = False, Optional ByVal customOpacityModifier As Single = 1!) As Long
    
    Const funcName As String = "GetBrushHandle"
    
    'Make sure a gradient has been created
    If (m_NumOfPoints <= 0) Then
        GetBrushHandle = 0
        InternalError funcName, "no gradient points added!"
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
    
    'TODO: add a "repeat" parameter, for convenience?
    
    'We can now create the brush!  We must do this in two steps.
    ' (IMPORTANT NOTE: at present, only GDI+ brushes are created.  Future backends will need to split out their behavior here.)
    Dim gdipBrush As Long
    
    'First: the user can request a default, angle = 0 linear version of the gradient for UI purposes.  Handle that case separately.
    If forceToLinearMode Then
    
        'Create a default linear gradient brush.  Note that GDI+ only allows you to supply two colors, so we'll just supply
        ' two dummy colors.
        gdipBrush = GDI_Plus.GetGDIPlusLinearBrushHandle(dstRect, dstRGBA(0), dstRGBA(1), 0!, 0, P2_WM_TileFlipXY)
        If (gdipBrush = 0) Then
            InternalError funcName, "forceToLinearMode: couldn't create linear gradient brush"
        Else
        
            'Now that we have a brush handle, override the dummy colors with our full color array
            If Not GDI_Plus.OverrideGDIPlusLinearGradient(gdipBrush, VarPtr(dstRGBA(0)), VarPtr(dstPosition(0)), finalPointCount) Then
                InternalError funcName, "failed to override default gradient brush colors"
            End If
            
            'Set gamma mode, as relevant
            If m_GradientGammaMode Then GDI_Plus.GDIPlus_LineGradientSetGamma gdipBrush, True
            
        End If
        
    Else
        
        'Brush creation varies by shape
        Select Case m_GradientShape
        
            Case P2_GS_Linear, P2_GS_Reflection
                gdipBrush = GDI_Plus.GetGDIPlusLinearBrushHandle(dstRect, dstRGBA(0), dstRGBA(1), m_GradientAngle, 0, m_GradientWrapMode)
                If (gdipBrush = 0) Then
                    InternalError funcName, "failed to create linear brush"
                Else
                    If Not GDI_Plus.OverrideGDIPlusLinearGradient(gdipBrush, VarPtr(dstRGBA(0)), VarPtr(dstPosition(0)), finalPointCount) Then
                        InternalError funcName, "failed to override default gradient brush colors"
                    End If
                    If m_GradientGammaMode Then GDI_Plus.GDIPlus_LineGradientSetGamma gdipBrush, True
                End If
                
            'All other gradient types are constructed as "path gradients".
            ' These gradients require a GDI+ GraphicsPath object that defines the external border of the gradient.
            Case Else
            
                'Reset our graphics path reference
                If (m_Path Is Nothing) Then Set m_Path = New pd2DPath Else m_Path.ResetPath
                
                'Some paths require a modified rect (e.g. square vs rectangle, circle vs ellipse),
                ' so construct a few additional measurements in advance.
                Dim halfWidth As Double, halfHeight As Double
                halfWidth = dstRect.Width / 2
                halfHeight = dstRect.Height / 2
                
                Dim centerX As Double, centerY As Double
                centerX = dstRect.Left + halfWidth
                centerY = dstRect.Top + halfHeight
                
                Dim squareRect As RectF, MaxSize As Double, radiusRect As Double, radiusSquare As Double
                If (dstRect.Width > dstRect.Height) Then MaxSize = halfWidth Else MaxSize = halfHeight
                radiusRect = Sqr(halfWidth * halfWidth + halfHeight * halfHeight) + 0.5
                radiusSquare = Sqr(MaxSize * MaxSize + MaxSize * MaxSize) + 0.5
                                
                With squareRect
                    .Left = centerX - MaxSize
                    .Top = centerY - MaxSize
                    .Width = MaxSize * 2
                    .Height = MaxSize * 2
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
                    
                    'Basically, the crucial measurement for a diamond gradient is the hypotenuse of a triangle
                    ' that meets the following criteria:
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
                    m_Path.RotatePathAroundItsCenter 45!
                    
                End If
                
                'Create a path gradient brush, using our path object as the reference
                gdipBrush = GDI_Plus.GetGDIPlusPathBrushHandle(m_Path.GetHandle)
                If (gdipBrush <> 0) Then
                    
                    'Set the center point
                    Dim tmpCenterPoint As PointFloat
                    If m_UseCustomCenterPoint Then
                        tmpCenterPoint.x = dstRect.Left + (dstRect.Width * m_GradientCenterPoint.x)
                        tmpCenterPoint.y = dstRect.Top + (dstRect.Height * m_GradientCenterPoint.y)
                    Else
                        tmpCenterPoint.x = centerX
                        tmpCenterPoint.y = centerY
                        GDI_Plus.SetGDIPlusPathBrushCenter gdipBrush, centerX, centerY
                    End If
                    
                    If (GdipSetPathGradientCenterPoint(gdipBrush, tmpCenterPoint) <> GP_OK) Then
                        InternalError funcName, "failed to set brush center point"
                    End If
                    
                    'Next, override the brush colors with our custom array
                    If Not GDI_Plus.OverrideGDIPlusPathGradient(gdipBrush, VarPtr(dstRGBA(0)), VarPtr(dstPosition(0)), finalPointCount) Then
                        InternalError funcName, "failed to override default gradient brush colors"
                    End If
                    
                    'Finally, set any/all custom parameters (e.g. wrap mode)
                    GDI_Plus.SetGDIPlusPathBrushWrap gdipBrush, m_GradientWrapMode
                    If m_GradientGammaMode Then GDI_Plus.GDIPlus_PathGradientSetGamma gdipBrush, True
                
                Else
                    InternalError funcName, "failed to create path brush handle"
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

'Given a position on the range [0, 1], return the gradient color at that position.  Note that a linear,
' non-reflected gradient shape is always assumed, by design.
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
        GetColorAtPosition_RGBA = False
    End If
    

End Function

'Return a Long-type lookup table that describes the contents of the gradient.  The table size can be defined
' by the caller; larger tables obviously include more details, but they may be pointless depending on the
' number of colors in the gradient and/or their positions.  (For example, a 2-color gradient from 0.0 to 1.0
' won't benefit from more than 256 entries in the lookup table, because the underlying data uses 8-bit channels.)
Friend Function GetLookupTable(ByRef dstArray() As Long, Optional ByVal sizeOfTable As Long = 1024) As Boolean

    'Use GDI+ to render a 1-byte-tall sample of the current gradient.
    ' (Note that we deliberately set the DIB to be 1-pixel larger than the actual size of the
    ' lookup table; this lets us skip specialized edge-handling workarounds for GDI+ rendering bugs).
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    tmpDIB.CreateBlank sizeOfTable, 1, 32, 0, 0
    
    Dim bRect As RectF
    bRect.Left = 0!
    bRect.Top = 0!
    bRect.Width = sizeOfTable
    bRect.Height = 1!
    
    Dim gradHandle As Long
    gradHandle = Me.GetBrushHandle(bRect, True)
    
    Dim imgHandle As Long
    imgHandle = GDI_Plus.GetGDIPlusGraphicsFromDC_Fast(tmpDIB.GetDIBDC)
    
    GDI_Plus.GDIPlus_FillRectF imgHandle, gradHandle, bRect.Left, bRect.Top, bRect.Width, bRect.Height
    GDI_Plus.ReleaseGDIPlusGraphics imgHandle
    GDI_Plus.ReleaseGDIPlusBrush gradHandle
    
    'Simply copy the contents of the temporary image into the lookup table
    ReDim dstArray(0 To sizeOfTable - 1) As Long
    CopyMemoryStrict VarPtr(dstArray(0)), tmpDIB.GetDIBPointer, sizeOfTable * 4

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
    Me.SetCenterOffset_Default
    
End Sub

'All pd2D classes report errors using an internal function similar to this one.
' Feel free to modify this function to better fit your project
' (for example, maybe you prefer to raise an actual error event).
'
'Note that by default, pd2D build simply dumps all error information to the Immediate window.
Private Sub InternalError(ByRef errFunction As String, ByRef errDescription As String, Optional ByVal errNum As Long = 0)
    Drawing2D.DEBUG_NotifyError "pd2DGradient", errFunction, errDescription, errNum
End Sub
