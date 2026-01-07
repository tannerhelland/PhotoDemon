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
'Copyright 2014-2026 by Tanner Helland
'Created: 09/May/15 (though assembled from various parts written much earlier)
'Last updated: 02/March/22
'Last update: add helper function to retrieve matrix as an array of bare floats
'
'This class is a VB6-friendly wrapper for the GDI+ Matrix object.  It is not intended as a comprehensive interface;
' instead, I add functions to it as I need them for various PD elements.
'
'The GDI+ matrix handle is created on-demand, and automatically released at class termination.  The matrix handle is
' persistent by design, so the caller never needs to manually create or delete it.  If you need to start over, use
' the ResetMatrix() function.  (Note that resetting the matrix may cause a new handle value to be created.  For this
' reason, you must never cache the transformation handle returned by the GetHandle() function.)
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************
Option Explicit

'GDI+ declares
Private Declare Function GdipCloneMatrix Lib "gdiplus" (ByVal srcMatrix As Long, ByRef dstMatrix As Long) As GP_Result
Private Declare Function GdipCreateMatrix Lib "gdiplus" (ByRef dstMatrix As Long) As GP_Result
'Private Declare Function GdipCreateMatrix2 Lib "gdiplus" (ByVal mM11 As Single, ByVal mM12 As Single, ByVal mM21 As Single, ByVal mM22 As Single, ByVal mDx As Single, ByVal mDy As Single, ByRef dstMatrix As Long) As GP_Result
Private Declare Function GdipDeleteMatrix Lib "gdiplus" (ByVal hMatrix As Long) As GP_Result
Private Declare Function GdipGetMatrixElements Lib "gdiplus" (ByVal hMatrix As Long, ByVal ptrTo6ElementFloatArray As Long) As GP_Result
Private Declare Function GdipInvertMatrix Lib "gdiplus" (ByVal hMatrix As Long) As GP_Result
Private Declare Function GdipIsMatrixInvertible Lib "gdiplus" (ByVal hMatrix As Long, ByRef dstResult As Long) As Long
Private Declare Function GdipRotateMatrix Lib "gdiplus" (ByVal hMatrix As Long, ByVal rotateAngle As Single, ByVal mOrder As GP_MatrixOrder) As GP_Result
Private Declare Function GdipScaleMatrix Lib "gdiplus" (ByVal hMatrix As Long, ByVal scaleX As Single, ByVal scaleY As Single, ByVal mOrder As GP_MatrixOrder) As GP_Result
Private Declare Function GdipShearMatrix Lib "gdiplus" (ByVal hMatrix As Long, ByVal shearX As Single, ByVal shearY As Single, ByVal mOrder As GP_MatrixOrder) As GP_Result
Private Declare Function GdipTransformMatrixPoints Lib "gdiplus" (ByVal hMatrix As Long, ByVal ptrToFirstPointF As Long, ByVal numOfPoints As Long) As GP_Result
Private Declare Function GdipTranslateMatrix Lib "gdiplus" (ByVal hMatrix As Long, ByVal offsetX As Single, ByVal offsetY As Single, ByVal mOrder As GP_MatrixOrder) As GP_Result

'Persistent handle to the transformation associated with this instance
Private m_TransformHandle As Long

'Return a copy of the handle.  The caller should not free or otherwise modify the handle; this function exists solely
' for interop with other pd2D objects.
Friend Function GetHandle(Optional ByVal createAsNecessary As Boolean = True) As Long
    If (m_TransformHandle <> 0) Then
        GetHandle = m_TransformHandle
    ElseIf createAsNecessary Then
        If CreateTransform() Then GetHandle = m_TransformHandle
    End If
End Function

Friend Function HasTransform() As Boolean
    HasTransform = (m_TransformHandle <> 0)
End Function

Friend Function CloneExistingTransform(ByVal srcTransform As pd2DTransform) As Boolean
    
    If (Not srcTransform Is Nothing) Then
    
        Me.ReleaseTransform
        
        CloneExistingTransform = (GdipCloneMatrix(srcTransform.GetHandle, m_TransformHandle) = GP_OK)
        
        If CloneExistingTransform Then
            CloneExistingTransform = (m_TransformHandle <> 0)
        Else
            InternalError "CloneExistingTransform", "GDI+ failure"
        End If
        
        If (CloneExistingTransform And PD2D_DEBUG_MODE) Then Drawing2D.DEBUG_NotifyTransformCountChange True
        
    Else
        CloneExistingTransform = False
    End If
    
End Function

'Apply a translation amount to the matrix; any x/y amounts are valid
Friend Function ApplyTranslation(ByVal xDistance As Single, ByVal yDistance As Single, Optional ByVal transformOrder As PD_2D_TransformOrder = P2_TO_Append) As Boolean
    
    'For all transform operations, a transform object must first exist
    If (m_TransformHandle = 0) Then CreateTransform
    
    Dim gpResult As GP_Result
    gpResult = GdipTranslateMatrix(m_TransformHandle, xDistance, yDistance, transformOrder)
    ApplyTranslation = (gpResult = GP_OK)
    If (Not ApplyTranslation) Then InternalError "ApplyTranslation", "GDI+ failure", gpResult
    
End Function

'Translate the entire path by some amount in polar coordinates (angle + radius).
Friend Function ApplyTranslation_Polar(ByVal translateAngle As Single, ByVal translateRadius As Single, Optional ByVal angleIsInDegrees As Boolean = True, Optional ByVal transformOrder As PD_2D_TransformOrder = P2_TO_Append) As Boolean
    If angleIsInDegrees Then translateAngle = PDMath.DegreesToRadians(translateAngle)
    ApplyTranslation_Polar = Me.ApplyTranslation(translateRadius * Cos(translateAngle), translateRadius * Sin(translateAngle), transformOrder)
End Function

'Apply a scaling amount to the matrix
Friend Function ApplyScaling(ByVal xScaleFactor As Single, ByVal yScaleFactor As Single, Optional ByVal centerX As Single = 0!, Optional ByVal centerY As Single = 0!, Optional ByVal transformOrder As PD_2D_TransformOrder = P2_TO_Append) As Boolean
    
    If (m_TransformHandle = 0) Then CreateTransform
    
    Dim gpResult As GP_Result
    
    'If a different center point is desired, we must translate the matrix before and after rotation.
    ' (The optional transformOrder parameter is ignored out of necessity.)
    If (centerX <> 0!) Or (centerY <> 0!) Then
        
        'Center the matrix over the rotation center point
        ApplyScaling = Me.ApplyTranslation(-1! * centerX, -1! * centerY, P2_TO_Append)
        
        'Apply scaling
        gpResult = GdipScaleMatrix(m_TransformHandle, xScaleFactor, yScaleFactor, transformOrder)
        ApplyScaling = ApplyScaling And (gpResult = GP_OK)
        
        'Translate back into its original position
        ApplyScaling = ApplyScaling And Me.ApplyTranslation(centerX, centerY, P2_TO_Append)
    
    Else
        gpResult = GdipScaleMatrix(m_TransformHandle, xScaleFactor, yScaleFactor, transformOrder)
        ApplyScaling = (gpResult = GP_OK)
    End If
    
    If (Not ApplyScaling) Then InternalError "ApplyScaling", "GDI+ failure", gpResult
    
End Function

'Mirror the matrix in the x and/or y direction
Friend Function ApplyMirror(ByVal mirrorHorizontal As Boolean, ByVal mirrorVertical As Boolean, Optional ByVal transformOrder As PD_2D_TransformOrder = P2_TO_Append) As Boolean
    If mirrorHorizontal Then ApplyMirror = Me.ApplyScaling(-1!, 1!, , , transformOrder)
    If mirrorVertical Then ApplyMirror = Me.ApplyScaling(1!, -1!, , , transformOrder)
End Function

'Apply a rotation to the matrix; units are degrees, and positive degrees indicate clockwise transformations.
' An optional center x/y can be provided as well.
' IMPORTANT NOTE: if a custom center point is specified, the optional transformOrder parameter is ignored.
'                 (This is necessary because we must prepend and append custom translation operations.)
Friend Function ApplyRotation(ByVal rotationAngle As Single, Optional ByVal centerX As Single = 0!, Optional ByVal centerY As Single = 0!, Optional ByVal transformOrder As PD_2D_TransformOrder = P2_TO_Append) As Boolean
    
    If (m_TransformHandle = 0) Then CreateTransform
    
    Dim gpResult As GP_Result
    
    'If a different center point is desired, we must translate the matrix before and after rotation.
    ' (The optional transformOrder parameter is ignored out of necessity.)
    If (centerX <> 0!) Or (centerY <> 0!) Then
        
        'Center the matrix over the rotation center point
        ApplyRotation = Me.ApplyTranslation(-1! * centerX, -1! * centerY, P2_TO_Append)
        
        'Apply the rotation
        gpResult = GdipRotateMatrix(m_TransformHandle, rotationAngle, P2_TO_Append)
        ApplyRotation = ApplyRotation And (gpResult = GP_OK)
        
        'Translate back into its original position
        ApplyRotation = ApplyRotation And Me.ApplyTranslation(centerX, centerY, P2_TO_Append)
    
    Else
        gpResult = GdipRotateMatrix(m_TransformHandle, rotationAngle, transformOrder)
        ApplyRotation = (gpResult = GP_OK)
    End If
    
    If (Not ApplyRotation) Then InternalError "ApplyRotation", "GDI+ failure", gpResult
    
End Function

'Apply x/y shear factors to the matrix.  An optional center x/y can also be specified.
' IMPORTANT NOTE: if a custom center point is specified, the optional transformOrder parameter is ignored.  (This is necessary
'                 because we must prepend and append custom translation operations.)
Friend Function ApplyShear(ByVal shearX As Single, ByVal shearY As Single, Optional ByVal centerX As Single = 0!, Optional ByVal centerY As Single = 0!, Optional ByVal transformOrder As PD_2D_TransformOrder = P2_TO_Append) As Boolean
    
    If (m_TransformHandle = 0) Then CreateTransform
    
    Dim gpResult As GP_Result
    
    'If a different center point is desired, we must translate the matrix before and after shearing.
    ' (The optional transformOrder parameter is ignored out of necessity.)
    If (centerX <> 0!) Or (centerY <> 0!) Then
        
        'Center the matrix over the shear center point
        ApplyShear = Me.ApplyTranslation(-1! * centerX, -1! * centerY, P2_TO_Append)
        
        'Apply the shear
        gpResult = GdipShearMatrix(m_TransformHandle, shearX, shearY, P2_TO_Append)
        ApplyShear = ApplyShear And (gpResult = GP_OK)
        
        'Translate back into its original position
        ApplyShear = ApplyShear And Me.ApplyTranslation(centerX, centerY, P2_TO_Append)
    
    Else
        gpResult = GdipShearMatrix(m_TransformHandle, shearX, shearY, transformOrder)
        ApplyShear = (gpResult = GP_OK)
    End If
    
    If (Not ApplyShear) Then InternalError "ApplyShear", "GDI+ failure", gpResult
    
End Function

'Apply the current transform to any number of arbitrary (x, y) coordinate pair(s)
Friend Function ApplyTransformToPointFs(ByVal ptrToFirstPointF As Long, ByVal numOfPoints As Long) As Boolean
    
    If (m_TransformHandle = 0) Then CreateTransform
    
    Dim gpResult As GP_Result
    gpResult = GdipTransformMatrixPoints(m_TransformHandle, ptrToFirstPointF, numOfPoints)
    ApplyTransformToPointFs = (gpResult = GP_OK)
    If (Not ApplyTransformToPointFs) Then InternalError "ApplyTransformToPointFs", "GDI+ failure", gpResult
    
End Function

Friend Function ApplyTransformToPointF(ByRef srcPoint As PointFloat) As Boolean
    
    If (m_TransformHandle = 0) Then CreateTransform
    
    Dim gpResult As GP_Result
    gpResult = GdipTransformMatrixPoints(m_TransformHandle, VarPtr(srcPoint), 1)
    ApplyTransformToPointF = (gpResult = GP_OK)
    If (Not ApplyTransformToPointF) Then InternalError "ApplyTransformToPointF", "GDI+ failure", gpResult
    
End Function

Friend Function ApplyTransformToXY(ByRef srcX As Single, ByRef srcY As Single) As Boolean
    
    If (m_TransformHandle = 0) Then CreateTransform

    'Create a temporary POINTFLOAT struct
    Dim tmpPoint As PointFloat
    tmpPoint.x = srcX
    tmpPoint.y = srcY
    
    'Apply the matrix transformation
    ApplyTransformToXY = Me.ApplyTransformToPointF(tmpPoint)
    
    'Return the translated values
    srcX = tmpPoint.x
    srcY = tmpPoint.y
    
End Function

'Return matrix points as a float array.  The array will be forcibly resized to (0, 5),
' and points will be returned in the order m11, m12, m21, m22, m31, m32,
' where mij denotes the element in row i, column j.
Friend Function GetMatrixPoints(ByRef dstPoints() As Single) As Boolean
    
    'The matrix always has six points.
    ReDim dstPoints(0 To 5) As Single
    
    If (m_TransformHandle = 0) Then CreateTransform
    GetMatrixPoints = (GdipGetMatrixElements(m_TransformHandle, VarPtr(dstPoints(0))) = GP_OK)
    
End Function

'Invert the matrix.  Note that this is a function, and it returns FALSE if the matrix is not invertible.  (This can happen if
' the matrix has been set to all zeroes, for example, by bogus scaling code.)
Friend Function InvertTransform() As Boolean
    
    If (m_TransformHandle = 0) Then CreateTransform
    InvertTransform = False
    
    'Make sure the matrix is invertible
    Dim tmpLong As Long
    If (GdipIsMatrixInvertible(m_TransformHandle, tmpLong) = GP_OK) Then
        If (tmpLong <> 0) Then InvertTransform = (GdipInvertMatrix(m_TransformHandle) = GP_OK)
    Else
        InternalError "InvertTransform", "GDI+ failure"
    End If
    
End Function

'Reset the matrix.  Note that this a new matrix will be auto-created by other functions, as necessary.
Friend Sub Reset()
    If (m_TransformHandle <> 0) Then Me.ReleaseTransform
End Sub

'Create an actual transform handle using the current backend and the current transform settings.
' NOTE: by design, this function is not exposed externally, because the caller *never* needs to call this directly.
'       If GetTransformHandle is called and transform doesn't yet exist, it will be auto-created (using an
'       identity matrix).  Similarly, if a matrix operation is applied to this class but the base matrix doesn't
'       exist, it will also be auto-created.
Private Function CreateTransform() As Boolean

    If (m_TransformHandle <> 0) Then Me.ReleaseTransform
    
    Dim gpResult As GP_Result
    gpResult = GdipCreateMatrix(m_TransformHandle)
    
    'When debug mode is active, all object creations are reported back to the central Drawing2D module
    CreateTransform = (gpResult = GP_OK) And (m_TransformHandle <> 0)
    If (CreateTransform And PD2D_DEBUG_MODE) Then Drawing2D.DEBUG_NotifyTransformCountChange True
    
End Function

'Free the current matrix handle.
Friend Function ReleaseTransform() As Boolean
    
    If (m_TransformHandle <> 0) Then
        
        ReleaseTransform = (GdipDeleteMatrix(m_TransformHandle) = GP_OK)
        
        'After a successful release, we must always reset the class-level handle to match, and during debug mode,
        ' the central Drawing2D module also needs to be notified.
        If ReleaseTransform Then
            m_TransformHandle = 0
            If PD2D_DEBUG_MODE Then Drawing2D.DEBUG_NotifyTransformCountChange False
        End If
        
    Else
        ReleaseTransform = True
    End If
    
End Function

Private Sub Class_Initialize()
    Me.Reset
End Sub

Private Sub Class_Terminate()
    Me.ReleaseTransform
End Sub

'All pd2D classes report errors using an internal function similar to this one.
' Feel free to modify this function to better fit your project
' (for example, maybe you prefer to raise an actual error event).
'
'Note that by default, pd2D build simply dumps all error information to the Immediate window.
Private Sub InternalError(ByRef errFunction As String, ByRef errDescription As String, Optional ByVal errNum As Long = 0)
    Drawing2D.DEBUG_NotifyError "pd2DTransform", errFunction, errDescription, errNum
End Sub
