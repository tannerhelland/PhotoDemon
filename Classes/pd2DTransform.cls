VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pd2DTransform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon 2D Transformation manager
'Copyright 2014-2016 by Tanner Helland
'Created: 09/May/15 (though assembled from various parts written much earlier)
'Last updated: 22/June/15
'Last update: add x/y shear support
'
'This class is a VB6-friendly wrapper for the GDI+ Matrix object.  It is not intended as a comprehensive interface;
' instead, I add functions to it as I need them for various PD elements.
'
'The GDI+ matrix handle is created on-demand, and automatically released at class termination.  The matrix handle is
' persistent by design, so the caller never needs to manually create or delete it.  If you need to start over, use
' the ResetMatrix() function.  (Note that resetting the matrix may cause a new handle value to be created.  For this
' reason, you must never cache the transformation handle returned by the GetHandle() function.)
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************
Option Explicit

'The backend used to manage this object.  At present, this is immutable; once set, it cannot be migrated.
Private m_TransformBackend As PD_2D_RENDERING_BACKEND

'When debug mode is active, object creation and destruction is reported to the central Drawing2D module
Private m_DebugMode As Boolean

'Persistent handle to the transformation associated with this instance
Private m_TransformHandle As Long

'Return a copy of the handle.  The caller should not free or otherwise modify the handle; this function exists solely
' for interop with other pd2D objects.
Public Function GetHandle(Optional ByVal createAsNecessary As Boolean = True) As Long
    If (createAsNecessary And (m_TransformHandle = 0)) Then
        If CreateTransform Then GetHandle = m_TransformHandle Else GetHandle = 0
    Else
        GetHandle = m_TransformHandle
    End If
End Function

Public Function GetTransformBackend() As PD_2D_RENDERING_BACKEND
    GetTransformBackend = m_TransformBackend
End Function

Public Function HasTransform() As Boolean
    HasTransform = CBool(m_TransformHandle <> 0)
End Function

Public Function CloneExistingTransform(ByVal srcTransform As pd2DTransform) As Boolean
    
    If (Not srcTransform Is Nothing) Then
    
        Me.ReleaseTransform
        
        Select Case m_TransformBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                m_TransformHandle = GDI_Plus.GDIPlus_MatrixClone(srcTransform.GetHandle)
        End Select
        
        CloneExistingTransform = CBool(m_TransformHandle <> 0)
        
    Else
        CloneExistingTransform = False
    End If
    
End Function

'Apply a translation amount to the matrix; any x/y amounts are valid
Friend Function ApplyTranslation(ByVal xDistance As Single, ByVal yDistance As Single, Optional ByVal transformOrder As PD_2D_TransformOrder = P2_TO_Append) As Boolean
    If (m_TransformHandle = 0) Then CreateTransform
    Select Case m_TransformBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
            ApplyTranslation = GDI_Plus.GDIPlus_MatrixTranslate(m_TransformHandle, xDistance, yDistance, transformOrder)
    End Select
End Function

'Apply a scaling amount to the matrix; any x/y amounts are valid
Friend Function ApplyScaling(ByVal xScaleFactor As Single, ByVal yScaleFactor As Single, Optional ByVal transformOrder As PD_2D_TransformOrder = P2_TO_Append) As Boolean
    If (m_TransformHandle = 0) Then CreateTransform
    Select Case m_TransformBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
            ApplyScaling = GDI_Plus.GDIPlus_MatrixScale(m_TransformHandle, xScaleFactor, yScaleFactor, transformOrder)
    End Select
End Function

'Mirror the matrix in the x and/or y direction
Friend Function ApplyMirror(ByVal mirrorHorizontal As Boolean, ByVal mirrorVertical As Boolean, Optional ByVal transformOrder As PD_2D_TransformOrder = P2_TO_Append) As Boolean
    If (m_TransformHandle = 0) Then CreateTransform
    Select Case m_TransformBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
            If mirrorHorizontal Then ApplyMirror = Me.ApplyScaling(-1, 1, transformOrder)
            If mirrorVertical Then ApplyMirror = Me.ApplyScaling(1, -1, transformOrder)
    End Select
End Function

'Apply a rotation to the matrix; units are degrees, and positive degrees indicate clockwise transformations.
' An optional center x/y can be provided as well.
' IMPORTANT NOTE: if a custom center point is specified, the optional transformOrder parameter is ignored.  (This is necessary
'                 because we must prepend and append custom translation operations.)
Friend Function ApplyRotation(ByVal rotationAngle As Single, Optional ByVal centerX As Single = 0, Optional ByVal centerY As Single = 0, Optional ByVal transformOrder As PD_2D_TransformOrder = P2_TO_Append) As Boolean
    
    If (m_TransformHandle = 0) Then CreateTransform
    
    Select Case m_TransformBackend
    
        Case P2_DefaultBackend, P2_GDIPlusBackend
        
            'If a different center point is desired, we must translate the matrix before and after rotation.
            ' (The optional transformOrder parameter is ignored out of necessity.)
            If (centerX <> 0) Or (centerY <> 0) Then
                
                'Center the matrix over the rotation center point
                Me.ApplyTranslation -1 * centerX, -1 * centerY, P2_TO_Append
                
                'Apply the rotation
                ApplyRotation = GDI_Plus.GDIPlus_MatrixRotate(m_TransformHandle, rotationAngle, P2_TO_Append)
                
                'Translate back into its original position
                Me.ApplyTranslation centerX, centerY, P2_TO_Append
            
            Else
                ApplyRotation = GDI_Plus.GDIPlus_MatrixRotate(m_TransformHandle, rotationAngle, transformOrder)
            End If
    
    End Select
    
End Function

'Apply x/y shear factors to the matrix.  An optional center x/y can also be specified.
' IMPORTANT NOTE: if a custom center point is specified, the optional transformOrder parameter is ignored.  (This is necessary
'                 because we must prepend and append custom translation operations.)
Friend Function ApplyShear(ByVal shearX As Single, ByVal shearY As Single, Optional ByVal centerX As Single = 0, Optional ByVal centerY As Single = 0, Optional ByVal transformOrder As PD_2D_TransformOrder = P2_TO_Append) As Boolean
    
    If (m_TransformHandle = 0) Then CreateTransform
    
    Select Case m_TransformBackend
    
        Case P2_DefaultBackend, P2_GDIPlusBackend
        
            'If a different center point is desired, we must translate the matrix before and after shearing.
            ' (The optional transformOrder parameter is ignored out of necessity.)
            If (centerX <> 0) Or (centerY <> 0) Then
                
                'Center the matrix over the shear center point
                Me.ApplyTranslation -1 * centerX, -1 * centerY, P2_TO_Append
                
                'Apply the shear
                ApplyShear = GDI_Plus.GDIPlus_MatrixShear(m_TransformHandle, shearX, shearY, P2_TO_Append)
                
                'Translate back into its original position
                Me.ApplyTranslation centerX, centerY, P2_TO_Append
            
            Else
                ApplyShear = GDI_Plus.GDIPlus_MatrixShear(m_TransformHandle, shearX, shearY, transformOrder)
            End If
    
    End Select
    
End Function

'Translate an arbitrary (x, y) coordinate pair by the current matrix
Friend Function ApplyTransformToPointF(ByRef srcPoint As POINTFLOAT) As Boolean
    If (m_TransformHandle = 0) Then CreateTransform
    Select Case m_TransformBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
            ApplyTransformToPointF = GDI_Plus.GDIPlus_MatrixTransformListOfPoints(m_TransformHandle, VarPtr(srcPoint), 1&)
    End Select
End Function

Friend Function ApplyTransformToXY(ByRef srcX As Single, ByRef srcY As Single) As Boolean
    
    If (m_TransformHandle = 0) Then CreateTransform
    
    Select Case m_TransformBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
    
            'Create a temporary POINTFLOAT struct
            Dim tmpPoint As POINTFLOAT
            tmpPoint.x = srcX
            tmpPoint.y = srcY
            
            'Apply the matrix transformation
            ApplyTransformToXY = GDI_Plus.GDIPlus_MatrixTransformListOfPoints(m_TransformHandle, VarPtr(tmpPoint), 1&)
            
            'Return the translated values
            srcX = tmpPoint.x
            srcY = tmpPoint.y
    
    End Select
    
End Function

'Invert the matrix.  Note that this is a function, and it returns FALSE if the matrix is not invertible.  (This can happen if
' the matrix has been set to all zeroes, for example, by bogus scaling code.)
Public Function InvertTransform() As Boolean
    
    If (m_TransformHandle = 0) Then CreateTransform
    
    Select Case m_TransformBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
    
            'Make sure the matrix is invertible
            If GDI_Plus.GDIPlus_MatrixIsInvertible(m_TransformHandle) Then
                InvertTransform = GDI_Plus.GDIPlus_MatrixInvert(m_TransformHandle)
            Else
                InvertTransform = False
            End If
    
    End Select
    
End Function

'Reset the matrix.  Note that this a new matrix will be auto-created by other functions, as necessary.
Public Sub Reset()
    If (m_TransformHandle <> 0) Then ReleaseTransform
End Sub

'Create an actual transform handle using the current backend and the current transform settings.
' NOTE: by design, this function is not exposed externally, because the caller *never* needs to call this directly.
'       If GetTransformHandle is called and transform doesn't yet exist, it will be auto-created (using an
'       identity matrix).  Similarly, if a matrix operation is applied to this class but the base matrix doesn't
'       exist, it will also be auto-created.
Private Function CreateTransform() As Boolean

    If (m_TransformHandle <> 0) Then Me.ReleaseTransform
    
    'Call the backend-specific creation function
    Select Case m_TransformBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
            m_TransformHandle = GDI_Plus.GDIPlus_MatrixCreate()
    End Select
    
    'When debug mode is active, all object creations are reported back to the master Drawing2D module
    CreateTransform = CBool(m_TransformHandle <> 0)
    If (CreateTransform And m_DebugMode) Then Drawing2D.DEBUG_NotifyTransformCountChange m_TransformBackend, True
    
End Function

'Free the current matrix handle.
Public Function ReleaseTransform() As Boolean
    
    If (m_TransformHandle <> 0) Then
        
        'Call the backend-specific release function
        Select Case m_TransformBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                ReleaseTransform = GDI_Plus.GDIPlus_MatrixDelete(m_TransformHandle)
        End Select
        
        'After a successful release, we must always reset the class-level handle to match, and during debug mode,
        ' the central Drawing2D module also needs to be notified.
        If ReleaseTransform Then
            m_TransformHandle = 0
            If (m_DebugMode) Then Drawing2D.DEBUG_NotifyTransformCountChange m_TransformBackend, False
        End If
        
    Else
        ReleaseTransform = False
    End If
    
End Function

'You only need to manually initialize a backend if you plan on using a NON-DEFAULT backend.
' (Objects with the default backend are initialized automatically.)
Friend Function InitializeTransformBackend(Optional ByVal targetBackend As PD_2D_RENDERING_BACKEND = P2_DefaultBackend) As Boolean
    m_TransformBackend = targetBackend
    InitializeTransformBackend = True
End Function

'When debug mode is active, this class will report object creation and destruction back to the master Drawing2D module.
Public Sub SetDebugMode(ByVal newMode As Boolean)
    m_DebugMode = newMode
End Sub

Private Sub Class_Initialize()
    Me.Reset
End Sub

Private Sub Class_Terminate()
    Me.ReleaseTransform
End Sub

'All pd2D classes report errors using an internal function similar to this one.  Feel free to modify this function to
' better fit your project (e.g. perhaps it could raise an actual error event).
'
'Note that a default pd2D build simply dumps the passed error information to the Immediate window.
Private Sub InternalError(Optional ByRef errName As String = vbNullString, Optional ByRef errDescription As String = vbNullString, Optional ByVal ErrNum As Long = 0)
    Drawing2D.DEBUG_NotifyExternalError errName, errDescription, ErrNum, "pd2DTransform"
End Sub
