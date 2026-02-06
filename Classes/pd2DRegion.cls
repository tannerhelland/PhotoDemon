VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pd2DRegion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Region Class
'Copyright 2016-2026 by Tanner Helland
'Created: 18/June/16 (but assembled from many bits written earlier)
'Last updated: 19/June/16
'Last update: continue expanding functionality
'
'This class manages a single region instance.  Unlike other classes, this class does not delay construction of
' the underlying object until it's actually requested; as soon as the region is modified in any way, a default
' (infinitely large) region will be created, and subsequent requests are applied against that region.
' This approach greatly simplifies the class's design, but note that it may cause region counts in the central
' debugger to appear higher than other types of pd2D objects.
'
'At present, this class is primarily based on the capabilities of GDI+.  This may change going forward,
' but because GDI+ provides a nice baseline feature set, that's where I started.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'GDI+ declares
Private Declare Function GdipCloneRegion Lib "gdiplus" (ByVal srcRegion As Long, ByRef dstRegion As Long) As GP_Result
Private Declare Function GdipCombineRegionRect Lib "gdiplus" (ByVal hRegion As Long, ByRef srcRectF As RectF, ByVal dstCombineMode As GP_CombineMode) As GP_Result
Private Declare Function GdipCombineRegionRectI Lib "gdiplus" (ByVal hRegion As Long, ByRef srcRectL As RectL, ByVal dstCombineMode As GP_CombineMode) As GP_Result
Private Declare Function GdipCombineRegionRegion Lib "gdiplus" (ByVal dstRegion As Long, ByVal srcRegion As Long, ByVal dstCombineMode As GP_CombineMode) As GP_Result
Private Declare Function GdipCombineRegionPath Lib "gdiplus" (ByVal dstRegion As Long, ByVal srcPath As Long, ByVal dstCombineMode As GP_CombineMode) As GP_Result
Private Declare Function GdipCreateRegion Lib "gdiplus" (ByRef dstRegion As Long) As GP_Result
'Private Declare Function GdipCreateRegionPath Lib "gdiplus" (ByVal hPath As Long, ByRef hRegion As Long) As GP_Result
Private Declare Function GdipDeleteRegion Lib "gdiplus" (ByVal hRegion As Long) As GP_Result
Private Declare Function GdipGetRegionBounds Lib "gdiplus" (ByVal hRegion As Long, ByVal hGraphics As Long, ByRef dstRectF As RectF) As GP_Result
Private Declare Function GdipGetRegionBoundsI Lib "gdiplus" (ByVal hRegion As Long, ByVal hGraphics As Long, ByRef dstRectL As RectL) As GP_Result
Private Declare Function GdipGetRegionHRgn Lib "gdiplus" (ByVal hRegion As Long, ByVal hGraphics As Long, ByRef dstHRgn As Long) As GP_Result
Private Declare Function GdipGetRegionScansCount Lib "gdiplus" (ByVal hRegion As Long, ByRef dstRectCount As Long, ByVal hMatrix As Long) As GP_Result
Private Declare Function GdipGetRegionScans Lib "gdiplus" (ByVal hRegion As Long, ByVal ptrDstRectFs As Long, ByRef numRects As Long, ByVal hMatrix As Long) As GP_Result
Private Declare Function GdipGetRegionScansI Lib "gdiplus" (ByVal hRegion As Long, ByVal ptrDstRectLs As Long, ByRef numRects As Long, ByVal hMatrix As Long) As GP_Result
Private Declare Function GdipIsEmptyRegion Lib "gdiplus" (ByVal srcRegion As Long, ByVal srcGraphics As Long, ByRef dstResult As Long) As GP_Result
Private Declare Function GdipIsEqualRegion Lib "gdiplus" (ByVal srcRegion1 As Long, ByVal srcRegion2 As Long, ByVal srcGraphics As Long, ByRef dstResult As Long) As GP_Result
Private Declare Function GdipIsInfiniteRegion Lib "gdiplus" (ByVal srcRegion As Long, ByVal srcGraphics As Long, ByRef dstResult As Long) As GP_Result
Private Declare Function GdipIsVisibleRegionPoint Lib "gdiplus" (ByVal hRegion As Long, ByVal x As Single, ByVal y As Single, ByVal hGraphicsOptional As Long, ByRef dstResult As Long) As GP_Result
Private Declare Function GdipIsVisibleRegionPointI Lib "gdiplus" (ByVal hRegion As Long, ByVal x As Long, ByVal y As Long, ByVal hGraphicsOptional As Long, ByRef dstResult As Long) As GP_Result
Private Declare Function GdipSetEmpty Lib "gdiplus" (ByVal hRegion As Long) As GP_Result
Private Declare Function GdipSetInfinite Lib "gdiplus" (ByVal hRegion As Long) As GP_Result

'This class is not yet capable of serializing itself to/from XML strings, but it may be possible in the future...
'Private cSerialize As pdSerialize

'Once a region has been created, this handle value will be non-zero
Private m_RegionHandle As Long

Friend Function AddRectangleF(ByVal rLeft As Single, ByVal rTop As Single, ByVal rWidth As Single, ByVal rHeight As Single, Optional ByVal useCombineMode As PD_2D_CombineMode = P2_CM_Replace) As Boolean
    
    If (m_RegionHandle = 0) Then Me.CreateRegion
    
    Dim tmpRectF As RectF
    tmpRectF.Left = rLeft
    tmpRectF.Top = rTop
    tmpRectF.Width = rWidth
    tmpRectF.Height = rHeight
    
    AddRectangleF = (GdipCombineRegionRect(m_RegionHandle, tmpRectF, useCombineMode) = GP_OK)
    If (Not AddRectangleF) Then InternalError "AddRectangleF", "GDI+ failure"
    
End Function

Friend Function AddRectangle_FromRectF(ByRef srcRectF As RectF, Optional ByVal useCombineMode As PD_2D_CombineMode = P2_CM_Replace) As Boolean
    If (m_RegionHandle = 0) Then Me.CreateRegion
    AddRectangle_FromRectF = (GdipCombineRegionRect(m_RegionHandle, srcRectF, useCombineMode) = GP_OK)
    If (Not AddRectangle_FromRectF) Then InternalError "AddRectangle_FromRectF", "GDI+ failure"
End Function

Friend Function AddRectangle_FromRectL(ByRef srcRectL As RectL, Optional ByVal useCombineMode As PD_2D_CombineMode = P2_CM_Replace) As Boolean
    If (m_RegionHandle = 0) Then Me.CreateRegion
    AddRectangle_FromRectL = (GdipCombineRegionRectI(m_RegionHandle, srcRectL, useCombineMode) = GP_OK)
    If (Not AddRectangle_FromRectL) Then InternalError "AddRectangle_FromRectL", "GDI+ failure"
End Function

Friend Function AddRegion(ByRef srcRegion As pd2DRegion, Optional ByVal useCombineMode As PD_2D_CombineMode = P2_CM_Replace) As Boolean
    If (m_RegionHandle = 0) Then Me.CreateRegion
    AddRegion = (GdipCombineRegionRegion(m_RegionHandle, srcRegion.GetHandle, useCombineMode) = GP_OK)
    If (Not AddRegion) Then InternalError "AddRegion", "GDI+ failure"
End Function

Friend Function AddPath(ByRef srcPath As pd2DPath, Optional ByVal useCombineMode As PD_2D_CombineMode = P2_CM_Replace) As Boolean
    If (m_RegionHandle = 0) Then Me.CreateRegion
    AddPath = (GdipCombineRegionPath(m_RegionHandle, srcPath.GetHandle, useCombineMode) = GP_OK)
    If (Not AddPath) Then InternalError "AddPath", "GDI+ failure"
End Function

Friend Function CloneRegion(ByRef srcRegion As pd2DRegion) As Boolean
    If (m_RegionHandle <> 0) Then Me.ReleaseRegion
    CloneRegion = (GdipCloneRegion(srcRegion.GetHandle, m_RegionHandle) = GP_OK)
    If (Not CloneRegion) Then InternalError "CloneRegion", "GDI+ failure"
End Function

Friend Function GetRegionBoundsF() As RectF
    If (m_RegionHandle <> 0) Then
        Dim tmpReturn As GP_Result
        tmpReturn = GdipGetRegionBounds(m_RegionHandle, 0&, GetRegionBoundsF)
        If (tmpReturn <> GP_OK) Then InternalError "GetRegionBoundsF", "GDI+ failure", tmpReturn
    Else
        InternalError "GetRegionBoundsF", "null handle"
    End If
End Function

Friend Function GetRegionBoundsL() As RectL
    If (m_RegionHandle <> 0) Then
        Dim tmpReturn As GP_Result
        tmpReturn = GdipGetRegionBoundsI(m_RegionHandle, 0&, GetRegionBoundsL)
        If (tmpReturn <> GP_OK) Then InternalError "GetRegionBoundsL", "GDI+ failure", tmpReturn
    Else
        InternalError "GetRegionBoundsL", "null handle"
    End If
End Function

'TODO: is graphics object required by GDI+ call?  Docs are unclear...
Friend Function IsPointInRegion(ByVal srcX As Single, ByVal srcY As Single) As Boolean
    If (m_RegionHandle <> 0) Then
        Dim tmpResult As Long, tmpReturn As GP_Result
        tmpReturn = GdipIsVisibleRegionPoint(m_RegionHandle, srcX, srcY, 0&, tmpResult)
        If (tmpReturn = GP_OK) Then
            IsPointInRegion = (tmpResult <> 0)
        Else
            InternalError "IsPointInRegion", "GDI+ failure", tmpReturn
        End If
    Else
        InternalError "IsPointInRegion", "null handle"
    End If
End Function

'TODO: is graphics object required by GDI+ call?  Docs are unclear...
Friend Function IsPointInRegionL(ByVal srcX As Long, ByVal srcY As Long) As Boolean
    If (m_RegionHandle <> 0) Then
        Dim tmpResult As Long, tmpReturn As GP_Result
        tmpReturn = GdipIsVisibleRegionPointI(m_RegionHandle, srcX, srcY, 0&, tmpResult)
        If (tmpReturn = GP_OK) Then
            IsPointInRegionL = (tmpResult <> 0)
        Else
            InternalError "IsPointInRegionL", "GDI+ failure", tmpReturn
        End If
    Else
        InternalError "IsPointInRegionL", "null handle"
    End If
End Function

'Don't use this function unless you know what you're doing.  This will copy (*NOT* clone)
' an external GDI+ region handle.  That handle *must not be in use elsewhere*, because if
' it is deleted, this class has no way of detecting that.  (This function is only used when
' a convoluted API does not allow us to directly grab a returned handle ourselves.)
Friend Sub AssignExternalHandleDirectly(ByVal srcHandle As Long)
    If (m_RegionHandle <> 0) Then Me.ReleaseRegion
    m_RegionHandle = srcHandle
End Sub

Friend Function MakeRegionInfinite() As Boolean
    If (m_RegionHandle = 0) Then Me.CreateRegion
    MakeRegionInfinite = (GdipSetInfinite(m_RegionHandle) = GP_OK)
End Function

Friend Function MakeRegionEmpty() As Boolean
    If (m_RegionHandle = 0) Then Me.CreateRegion
    MakeRegionEmpty = (GdipSetEmpty(m_RegionHandle) = GP_OK)
End Function

'TODO: is graphics object required by GDI+ call?  Docs are unclear...
Friend Function IsRegionInfinite() As Boolean
    If (m_RegionHandle <> 0) Then
        Dim tmpResult As Long, tmpReturn As GP_Result
        tmpReturn = GdipIsInfiniteRegion(m_RegionHandle, 0&, tmpResult)
        If (tmpReturn = GP_OK) Then
            IsRegionInfinite = (tmpResult <> 0)
        Else
            InternalError "IsRegionInfinite", "GDI+ failure", tmpReturn
        End If
    Else
        InternalError "IsRegionInfinite", "null handle"
    End If
End Function

Friend Function IsRegionEmpty() As Boolean
    
    If (m_RegionHandle <> 0) Then
    
        'Unlike some GDI+ flat APIs, a graphics object is *required* by this call.
        ' (Passing null will result in error #2, bad parameters)
        Dim tmpResult As Long, tmpGraphics As Long, tmpReturn As GP_Result
        tmpGraphics = GDI_Plus.GetGDIPlusGraphicsFromHWnd(OS.ThunderMainHWnd)
        tmpReturn = GdipIsEmptyRegion(m_RegionHandle, tmpGraphics, tmpResult)
        GDI_Plus.ReleaseGDIPlusGraphics tmpGraphics
        
        If (tmpReturn = GP_OK) Then
            IsRegionEmpty = (tmpResult <> 0)
        Else
            InternalError "IsRegionEmpty", "GDI+ failure", tmpReturn
        End If
    
    'No region handle...
    Else
        InternalError "IsRegionEmpty", "null handle"
    End If
    
End Function

'TODO: is graphics object required by GDI+ call?  Docs are unclear...
Friend Function IsEqual(ByRef srcRegion As pd2DRegion) As Boolean
    If (m_RegionHandle = 0) Then Me.CreateRegion
    Dim tmpResult As Long, tmpReturn As GP_Result
    tmpReturn = GdipIsEqualRegion(m_RegionHandle, srcRegion.GetHandle, 0&, tmpResult)
    If (tmpReturn = GP_OK) Then
        IsEqual = (tmpResult <> 0)
    Else
        InternalError "IsEqual", "GDI+ failure", tmpReturn
    End If
End Function

Friend Function GetHandle(Optional ByVal createAsNecessary As Boolean = True) As Long
    If (createAsNecessary And (m_RegionHandle = 0)) Then
        If Me.CreateRegion() Then GetHandle = m_RegionHandle Else GetHandle = 0
    Else
        GetHandle = m_RegionHandle
    End If
End Function

'Retrieve a copy of this region in GDI format.  Note that the caller is responsible for freeing the hRgn
' via DeleteObject when finished.
Friend Function GetRegionAsHRgn() As Long
    Dim tmpReturn As GP_Result
    tmpReturn = GdipGetRegionHRgn(m_RegionHandle, 0&, GetRegionAsHRgn)
    If (tmpReturn <> GP_OK) Then InternalError "GetRegionAsHRgn", "GDI+ failure", tmpReturn
End Function

'Return the current region as an array of RectFs.  The dstRectFs() array is guaranteed to be dimmed as
' [0, dstNumRects - 1] on a successful call.
Friend Function GetRegionAsRectFs(ByRef dstNumRects As Long, ByRef dstRectFs() As RectF, Optional ByVal hTransformMatrix As Long = 0&) As Boolean
    
    If (m_RegionHandle = 0) Then Exit Function
    
    'This function requires a transform matrix; generate one if one is not provided
    Dim tmpMatrix As pd2DTransform
    If (hTransformMatrix = 0) Then
        Set tmpMatrix = New pd2DTransform
        hTransformMatrix = tmpMatrix.GetHandle(True)
    End If
    
    GetRegionAsRectFs = (GdipGetRegionScansCount(m_RegionHandle, dstNumRects, hTransformMatrix) = GP_OK)
    If GetRegionAsRectFs And (dstNumRects > 0) Then
        ReDim dstRectFs(0 To dstNumRects - 1) As RectF
        GetRegionAsRectFs = (GdipGetRegionScans(m_RegionHandle, VarPtr(dstRectFs(0)), dstNumRects, hTransformMatrix) = GP_OK)
    End If
    
End Function

'Return the current region as an array of RectFs.  The dstRectFs() array is guaranteed to be dimmed as
' [0, dstNumRects - 1] on a successful call.
Friend Function GetRegionAsRectLs(ByRef dstNumRects As Long, ByRef dstRectLs() As RectL_WH, Optional ByVal hTransformMatrix As Long = 0&) As Boolean
    
    If (m_RegionHandle = 0) Then Exit Function
    
    'This function requires a transform matrix; generate one if one is not provided
    Dim tmpMatrix As pd2DTransform
    If (hTransformMatrix = 0) Then
        Set tmpMatrix = New pd2DTransform
        hTransformMatrix = tmpMatrix.GetHandle(True)
    End If
    
    GetRegionAsRectLs = (GdipGetRegionScansCount(m_RegionHandle, dstNumRects, hTransformMatrix) = GP_OK)
    If GetRegionAsRectLs And (dstNumRects > 0) Then
        ReDim dstRectLs(0 To dstNumRects - 1) As RectL_WH
        GetRegionAsRectLs = (GdipGetRegionScansI(m_RegionHandle, VarPtr(dstRectLs(0)), dstNumRects, hTransformMatrix) = GP_OK)
    End If
    
End Function

Friend Function HasRegion() As Boolean
    HasRegion = (m_RegionHandle <> 0)
End Function

'Create an actual region handle using the current backend and the current region settings.
' NOTE: the caller doesn't *need* to call this directly.  If GetRegionHandle is called and the region doesn't
'       yet exist, it will be auto-created.
Friend Function CreateRegion() As Boolean

    If (m_RegionHandle <> 0) Then Me.ReleaseRegion
    If (GdipCreateRegion(m_RegionHandle) = GP_OK) Then
    
        CreateRegion = (m_RegionHandle <> 0)
        
        'When debug mode is active, all object creations are reported back to the central Drawing2D module
        If (CreateRegion And PD2D_DEBUG_MODE) Then Drawing2D.DEBUG_NotifyRegionCountChange True
        
    Else
        InternalError "CreateRegion", "GDI+ failure"
    End If
    
End Function

Friend Function ReleaseRegion() As Boolean
    
    ReleaseRegion = True
    
    If (m_RegionHandle <> 0) Then
        
        ReleaseRegion = (GdipDeleteRegion(m_RegionHandle) = GP_OK)
        
        'After a successful release, we must always reset the class-level handle to match,
        ' and during debug mode the central Drawing2D module also needs to be notified.
        If ReleaseRegion Then
            m_RegionHandle = 0
            If PD2D_DEBUG_MODE Then Drawing2D.DEBUG_NotifyRegionCountChange False
        Else
            InternalError "ReleaseRegion", "GDI+ failure"
        End If
    
    End If
    
End Function

Friend Sub ResetAllProperties()
    Me.ReleaseRegion
End Sub

Private Sub Class_Initialize()
    Me.ResetAllProperties
End Sub

Private Sub Class_Terminate()
    Me.ReleaseRegion
End Sub

'All pd2D classes report errors using an internal function similar to this one.
' Feel free to modify this function to better fit your project
' (for example, maybe you prefer to raise an actual error event).
'
'Note that by default, pd2D build simply dumps all error information to the Immediate window.
Private Sub InternalError(ByRef errFunction As String, ByRef errDescription As String, Optional ByVal errNum As Long = 0)
    Drawing2D.DEBUG_NotifyError "pd2DRegion", errFunction, errDescription, errNum
End Sub

