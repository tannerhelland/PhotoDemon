VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdKDTreeNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'This node's color
Private m_Color As RGBQuad

'The depth of this node determines how we split it when assigning children (as we must rotate
' between x/y/z axes).  This value can also be one of several special constants.
Private m_Depth As Long
Private Const NODE_EMPTY As Long = -1

'Child nodes; may not be initialized!
Private m_Left As pdKDTreeNode, m_Right As pdKDTreeNode

Friend Function GetRed() As Long
    GetRed = m_Color.Red
End Function

Friend Function GetGreen() As Long
    GetGreen = m_Color.Green
End Function

Friend Function GetBlue() As Long
    GetBlue = m_Color.Blue
End Function

'Insert may be called recursively
Friend Sub InsertNode(ByRef srcColor As RGBQuad, Optional ByVal depth As Long = 0)
    
    'MOD is slower than a simple branch; cycle between values 0/1/2 for red/green/blue axes
    If (depth > 2) Then depth = 0
    
    'If this node has not yet been assigned a color, store the passed color locally.
    If (m_Depth = NODE_EMPTY) Then
        m_Color = srcColor
        m_Depth = depth
        
    'This node already contains a color.  See if the passed color is "higher" or "lower" than our own color -
    ' and note that we compare differently depending on the current depth.
    Else
        
        'Do a quick search for equality; if found, return our color and exit.
        ' (Note the manual short-circuiting to work around VB deficiencies.)
        If (srcColor.Red = m_Color.Red) Then
            If (srcColor.Green = m_Color.Green) Then
                If (srcColor.Blue = m_Color.Blue) Then
                    Exit Sub
                End If
            End If
        End If
        
        'Rotate between RGB axes
        Select Case depth
            
            'Red
            Case 0
            
                'Left branch
                If (srcColor.Red < m_Color.Red) Then
                    If (m_Left Is Nothing) Then Set m_Left = New pdKDTreeNode
                    m_Left.InsertNode srcColor, depth + 1
                    
                'Right branch
                Else
                    If (m_Right Is Nothing) Then Set m_Right = New pdKDTreeNode
                    m_Right.InsertNode srcColor, depth + 1
                
                End If
            
            'Green
            Case 1
                If (srcColor.Green < m_Color.Green) Then
                    If (m_Left Is Nothing) Then Set m_Left = New pdKDTreeNode
                    m_Left.InsertNode srcColor, depth + 1
                Else
                    If (m_Right Is Nothing) Then Set m_Right = New pdKDTreeNode
                    m_Right.InsertNode srcColor, depth + 1
                End If
            
            'Blue
            Case 2
                If (srcColor.Blue < m_Color.Blue) Then
                    If (m_Left Is Nothing) Then Set m_Left = New pdKDTreeNode
                    m_Left.InsertNode srcColor, depth + 1
                Else
                    If (m_Right Is Nothing) Then Set m_Right = New pdKDTreeNode
                    m_Right.InsertNode srcColor, depth + 1
                End If
            
        End Select
        
    End If

End Sub

'Return the distance between nearest color to some source color,
Friend Function NearestColor(ByRef srcColor As RGBQuad, ByRef curBestColor As RGBQuad, ByRef curBestDistance As Double) As Double
    
    'Search until we find a leaf node containing the "best match", while accounting for the possibility of
    ' "perfect matches".
    
    'Look for a "perfect match" between the target color color and our color
    If (srcColor.Red = m_Color.Red) Then
        If (srcColor.Green = m_Color.Green) Then
            If (srcColor.Blue = m_Color.Blue) Then
                
                'Perfect match!  Return our color and exit
                NearestColor = 0#
                curBestColor = m_Color
                curBestDistance = 0#
                Exit Function
                
            End If
        End If
    End If
    
    'If we're still here, this color does not perfectly match our color.  Continue searching until we reach
    ' a leaf node; once that happens, we'll "unwind" the recursion to check for better matches in alternate
    ' directions.
    Dim nearDirection As pdKDTreeNode, farDirection As pdKDTreeNode
    
    'Rotate between RGB axes
    Select Case m_Depth
        
        'Red
        Case 0
            If (srcColor.Red < m_Color.Red) Then
                Set nearDirection = m_Left
                Set farDirection = m_Right
            Else
                Set nearDirection = m_Right
                Set farDirection = m_Left
            End If
            
        'Green
        Case 1
            If (srcColor.Green < m_Color.Green) Then
                Set nearDirection = m_Left
                Set farDirection = m_Right
            Else
                Set nearDirection = m_Right
                Set farDirection = m_Left
            End If
        
        'Blue
        Case 2
            If (srcColor.Blue < m_Color.Blue) Then
                Set nearDirection = m_Left
                Set farDirection = m_Right
            Else
                Set nearDirection = m_Right
                Set farDirection = m_Left
            End If
        
    End Select
    
    'We now have references to the "nearer node" and "further node" (in terms of which direction our
    ' point is most likely to lie inside).  Query the near point for the closest point, and note that
    ' it will update the shared curBestColor and curBestDistance values.
    If (nearDirection Is Nothing) Then
    
        'We don't have any closer points in this direction, meaning we are the leaf node.
        ' Store our value and distance as the current "best match".
        Dim tmpDistance As Double
        tmpDistance = DistanceToMe(srcColor)
        
        If (tmpDistance < curBestDistance) Then
            curBestDistance = tmpDistance
            curBestColor = m_Color
        End If
    
    'We are not a leaf node; query our child for its "best match"
    Else
        nearDirection.NearestColor srcColor, curBestColor, curBestDistance
        
        'Perform a failsafe check against our internal color
        tmpDistance = DistanceToMe(srcColor)
        
        If (tmpDistance < curBestDistance) Then
            curBestDistance = tmpDistance
            curBestColor = m_Color
        End If
        
    End If
    
    'Now that recursion has unwound to our current position, see if there's a possibility that our
    ' "far direction" region could possibly contain a point that is closer to the target point than
    ' our current "best point".  This is trivial to calculate as the closest "theoretical point" for
    ' e.g. a split on the red plane, is a pixel where the maximum (on left) or minimum (on right) red
    ' value, plus identical green and blue values, on *that* side of the plane is closer than the
    ' current best match.
    If (Not farDirection Is Nothing) Then
        
        Dim checkLeft As Boolean
        
        Select Case m_Depth
            
            'Red
            Case 0
                If (srcColor.Red < m_Color.Red) Then
                    Set nearDirection = m_Left
                    Set farDirection = m_Right
                    checkLeft = True
                Else
                    Set nearDirection = m_Right
                    Set farDirection = m_Left
                End If
                
            'Green
            Case 1
                If (srcColor.Green < m_Color.Green) Then
                    Set nearDirection = m_Left
                    Set farDirection = m_Right
                    checkLeft = True
                Else
                    Set nearDirection = m_Right
                    Set farDirection = m_Left
                End If
            
            'Blue
            Case 2
                If (srcColor.Blue < m_Color.Blue) Then
                    Set nearDirection = m_Left
                    Set farDirection = m_Right
                    checkLeft = True
                Else
                    Set nearDirection = m_Right
                    Set farDirection = m_Left
                End If
            
        End Select
        
        'If the closest "theoretical" point in the far direction is closer than the current best match,
        ' query that child for a best match.
        If (TheoreticalDistanceToMe(srcColor, True) < curBestDistance) Then
            Debug.Print "yes"
            farDirection.NearestColor srcColor, curBestColor, curBestDistance
        Else
            Debug.Print "no"
        End If
        
    End If
    
End Function

Private Function DistanceToMe(ByRef compPoint As RGBQuad) As Double
    Dim r As Long, g As Long, b As Long
    r = CLng(compPoint.Red) - m_Color.Red
    g = CLng(compPoint.Green) - m_Color.Green
    b = CLng(compPoint.Blue) - m_Color.Blue
    DistanceToMe = r * r + g * g + b * b
End Function

Private Function TheoreticalDistanceToMe(ByRef compPoint As RGBQuad, ByVal checkLeft As Boolean) As Double
    
    Dim testVal As Long
    
    Select Case m_Depth
        
        'Red
        Case 0
            If checkLeft Then
                testVal = CLng(compPoint.Red) - 1
                If (testVal < 0) Then testVal = 0
                testVal = testVal - m_Color.Red
            Else
                testVal = CLng(compPoint.Red) - m_Color.Red
            End If
        
        'Green
        Case 1
            If checkLeft Then
                testVal = CLng(compPoint.Green) - 1
                If (testVal < 0) Then testVal = 0
                testVal = testVal - m_Color.Green
            Else
                testVal = CLng(compPoint.Green) - m_Color.Green
            End If
        
        'Blue
        Case 2
            If checkLeft Then
                testVal = CLng(compPoint.Blue) - 1
                If (testVal < 0) Then testVal = 0
                testVal = testVal - m_Color.Blue
            Else
                testVal = CLng(compPoint.Blue) - m_Color.Blue
            End If
        
    End Select
    
    TheoreticalDistanceToMe = testVal * testVal
    
End Function

Private Sub Class_Initialize()
    m_Depth = NODE_EMPTY
End Sub
