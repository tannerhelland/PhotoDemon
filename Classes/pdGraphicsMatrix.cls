VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdGraphicsMatrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Graphics Matrix Wrapper
'Copyright 2014-2015 by Tanner Helland
'Created: 09/May/15 (though assembled from various parts written much earlier)
'Last updated: 22/June/15
'Last update: add x/y shear support
'
'This class is a VB6-friendly wrapper for the GDI+ Matrix object.  It is not intended as a comprehensive interface;
' instead, I add functions to it as I need them for various PD elements.
'
'The GDI+ matrix handle is created at class initialization, and released at class termination.  The matrix handle is
' persistent by design, so no functions are provided for recreating or deleting it.  If you need to start over, use
' the resetMatrix() function.  (Note that resetting the matrix may cause a new handle value to be created.  For this
' reason, you should not cache the matrix handle returned by the getMatrixHandle() function.)
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************
Option Explicit

Public Enum MatrixOrder
    MatrixOrderPrepend = 0
    MatrixOrderAppend = 1
End Enum

#If False Then
    Private Const MatrixOrderPrepend = 0, MatrixOrderAppend = 1
#End If

Private Declare Function GdipCreateMatrix Lib "gdiplus" (ByRef dstMatrix As Long) As Long
Private Declare Function GdipCloneMatrix Lib "gdiplus" (ByVal srcMatrix As Long, ByRef dstMatrix As Long) As Long
Private Declare Function GdipDeleteMatrix Lib "gdiplus" (ByVal hMatrix As Long) As Long
Private Declare Function GdipRotateMatrix Lib "gdiplus" (ByVal hMatrix As Long, ByVal rotateAngle As Single, ByVal mOrder As MatrixOrder) As Long
Private Declare Function GdipScaleMatrix Lib "gdiplus" (ByVal hMatrix As Long, ByVal scaleX As Single, ByVal scaleY As Single, ByVal mOrder As MatrixOrder) As Long
Private Declare Function GdipShearMatrix Lib "gdiplus" (ByVal hMatrix As Long, ByVal shearX As Single, ByVal shearY As Single, ByVal mOrder As MatrixOrder) As Long
Private Declare Function GdipTranslateMatrix Lib "gdiplus" (ByVal hMatrix As Long, ByVal offsetX As Single, ByVal offsetY As Single, ByVal mOrder As MatrixOrder) As Long
Private Declare Function GdipTransformMatrixPoints Lib "gdiplus" (ByVal hMatrix As Long, ByVal ptrToFirstPointF As Long, ByVal numOfPoints As Long) As Long
Private Declare Function GdipIsMatrixInvertible Lib "gdiplus" (ByVal hMatrix As Long, ByRef dstResult As Long) As Long
Private Declare Function GdipInvertMatrix Lib "gdiplus" (ByVal hMatrix As Long) As Long

'Persistent handle to the matrix associated with this instance
Private m_MatrixHandle As Long

'Return a copy of the handle.  The caller should not free or otherwise modify the handle; this function exists solely for plugging
' the matrix into other GDI+ functions.
Public Function getMatrixHandle() As Long
    getMatrixHandle = m_MatrixHandle
End Function

Public Sub createFromExistingHandle(ByVal srcHandle As Long)
    deleteMatrix
    GdipCloneMatrix srcHandle, m_MatrixHandle
End Sub

'Return a full copy of this object
Public Sub getCopyOfMatrix(ByRef dstMatrix As pdGraphicsMatrix)
    
    If dstMatrix Is Nothing Then Set dstMatrix = New pdGraphicsMatrix
    dstMatrix.createFromExistingHandle m_MatrixHandle
    
End Sub

'Apply a translation amount to the matrix; any x/y amounts are valid
Friend Sub TranslateMatrix(ByVal xDistance As Single, ByVal yDistance As Single, Optional ByVal transformOrder As MatrixOrder = MatrixOrderAppend)
    GdipTranslateMatrix m_MatrixHandle, xDistance, yDistance, transformOrder
End Sub

'Apply a scaling amount to the matrix; any x/y amounts are valid
Friend Sub ScaleMatrix(ByVal xScaleFactor As Single, ByVal yScaleFactor As Single, Optional ByVal transformOrder As MatrixOrder = MatrixOrderAppend)
    GdipScaleMatrix m_MatrixHandle, xScaleFactor, yScaleFactor, transformOrder
End Sub

'Mirror the matrix in the x and/or y direction
Friend Sub MirrorMatrix(ByVal mirrorHorizontal As Boolean, ByVal mirrorVertical As Boolean, Optional ByVal transformOrder As MatrixOrder = MatrixOrderAppend)
    If mirrorHorizontal Then ScaleMatrix -1, 1, transformOrder
    If mirrorVertical Then ScaleMatrix 1, -1, transformOrder
End Sub

'Apply a rotation to the matrix; units are degrees, and positive degrees indicate clockwise transformations.
' An optional center x/y can be provided as well.
' IMPORTANT NOTE: if a custom center point is specified, the optional transformOrder parameter is ignored.  (This is necessary
'                 because we must prepend and append custom translation operations.)
Friend Sub RotateMatrix(ByVal rotateAmount As Single, Optional ByVal centerX As Single = 0, Optional ByVal centerY As Single = 0, Optional ByVal transformOrder As MatrixOrder = MatrixOrderAppend)
    
    'If a different center point is desired, we must translate the matrix before and after rotation.
    ' (The optional transformOrder parameter is ignored out of necessity.)
    If (centerX <> 0) Or (centerY <> 0) Then
        
        'Center the matrix over the rotation center point
        TranslateMatrix -1 * centerX, -1 * centerY, MatrixOrderAppend
        
        'Apply the rotation
        GdipRotateMatrix m_MatrixHandle, rotateAmount, MatrixOrderAppend
        
        'Translate back into its original position
        TranslateMatrix centerX, centerY, MatrixOrderAppend
    
    Else
        GdipRotateMatrix m_MatrixHandle, rotateAmount, transformOrder
    End If
    
End Sub

'Apply x/y shear factors to the matrix.  An optional center x/y can also be specified.
' IMPORTANT NOTE: if a custom center point is specified, the optional transformOrder parameter is ignored.  (This is necessary
'                 because we must prepend and append custom translation operations.)
Friend Sub ShearMatrix(ByVal shearX As Single, ByVal shearY As Single, Optional ByVal centerX As Single = 0, Optional ByVal centerY As Single = 0, Optional ByVal transformOrder As MatrixOrder = MatrixOrderAppend)
    
    'If a different center point is desired, we must translate the matrix before and after rotation.
    ' (The optional transformOrder parameter is ignored out of necessity.)
    If (centerX <> 0) Or (centerY <> 0) Then
        
        'Center the matrix over the shear center point
        TranslateMatrix -1 * centerX, -1 * centerY, MatrixOrderAppend
        
        'Apply the shear
        GdipShearMatrix m_MatrixHandle, shearX, shearY, MatrixOrderAppend
        
        'Translate back into its original position
        TranslateMatrix centerX, centerY, MatrixOrderAppend
    
    Else
        GdipShearMatrix m_MatrixHandle, shearX, shearY, transformOrder
    End If
    
End Sub

'Translate an arbitrary (x, y) coordinate pair by the current matrix
Friend Sub applyMatrixToPointF(ByRef srcPoint As POINTFLOAT)
    GdipTransformMatrixPoints m_MatrixHandle, VarPtr(srcPoint), 1
End Sub

Friend Sub applyMatrixToXYPair(ByRef srcX As Single, ByRef srcY As Single)
    
    'Create a temporary POINTFLOAT struct
    Dim tmpPoint As POINTFLOAT
    
    With tmpPoint
        .x = srcX
        .y = srcY
    End With
    
    'Apply the matrix transformation
    GdipTransformMatrixPoints m_MatrixHandle, VarPtr(tmpPoint), 1
    
    'Return the translated values
    srcX = tmpPoint.x
    srcY = tmpPoint.y
    
End Sub

'Invert the matrix.  Note that this is a function, and it returns FALSE if the matrix is not invertible.  (This can happen if
' the matrix has been set to all zeroes, for example, by bogus scaling code.)
Public Function InvertMatrix() As Boolean

    'Make sure the matrix is invertible
    Dim retMatrixInvertible As Long
    If GdipIsMatrixInvertible(m_MatrixHandle, retMatrixInvertible) = 0 Then
        InvertMatrix = CBool(retMatrixInvertible = 1)
    Else
        InvertMatrix = False
    End If
    
    If InvertMatrix Then GdipInvertMatrix m_MatrixHandle
    
End Function

'Reset the matrix.  Note that this will create a new matrix as necessary.
Public Sub ResetToIdentityMatrix()
    
    'Delete the current matrix, if any
    deleteMatrix
    
    'Create a new identity matrix
    GdipCreateMatrix m_MatrixHandle
    
End Sub

'Free the current matrix handle.  Note that modification functions will fail if a new matrix is not created prior to calling them!
Private Sub deleteMatrix()
    
    If m_MatrixHandle <> 0 Then
        GdipDeleteMatrix m_MatrixHandle
        m_MatrixHandle = 0
    End If
    
End Sub

Private Sub Class_Initialize()
    
    'An associated matrix handle is created and destroyed alongside the class
    ResetToIdentityMatrix
    
End Sub

Private Sub Class_Terminate()
    
    'Free the matrix
    deleteMatrix
    
End Sub
