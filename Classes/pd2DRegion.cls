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
'Copyright 2016-2020 by Tanner Helland
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

'This class is not yet capable of serializing itself to/from XML strings, but it may be possible in the future...
'Private cSerialize As pdSerialize

'When debug mode is active, object creation and destruction is reported to the central Drawing2D module
Private m_DebugMode As Boolean

'Once a region has been created, this handle value will be non-zero
Private m_RegionHandle As Long

Friend Function AddRectangleF(ByVal rLeft As Single, ByVal rTop As Single, ByVal rWidth As Single, ByVal rHeight As Single, Optional ByVal useCombineMode As PD_2D_CombineMode = P2_CM_Replace) As Boolean
    
    If (m_RegionHandle = 0) Then Me.CreateRegion
    
    Dim tmpRectF As RectF
    tmpRectF.Left = rLeft
    tmpRectF.Top = rTop
    tmpRectF.Width = rWidth
    tmpRectF.Height = rHeight
    
    AddRectangleF = GDI_Plus.GDIPlus_RegionAddRectF(m_RegionHandle, tmpRectF, useCombineMode)
    
End Function

Friend Function AddRectangle_FromRectF(ByRef srcRectF As RectF, Optional ByVal useCombineMode As PD_2D_CombineMode = P2_CM_Replace) As Boolean
    If (m_RegionHandle = 0) Then Me.CreateRegion
    AddRectangle_FromRectF = GDI_Plus.GDIPlus_RegionAddRectF(m_RegionHandle, srcRectF, useCombineMode)
End Function

Friend Function AddRectangle_FromRectL(ByRef srcRectL As RectL, Optional ByVal useCombineMode As PD_2D_CombineMode = P2_CM_Replace) As Boolean
    If (m_RegionHandle = 0) Then Me.CreateRegion
    AddRectangle_FromRectL = GDI_Plus.GDIPlus_RegionAddRectL(m_RegionHandle, srcRectL, useCombineMode)
End Function

Friend Function AddRegion(ByRef srcRegion As pd2DRegion, Optional ByVal useCombineMode As PD_2D_CombineMode = P2_CM_Replace) As Boolean
    If (m_RegionHandle = 0) Then Me.CreateRegion
    AddRegion = GDI_Plus.GDIPlus_RegionAddRegion(m_RegionHandle, srcRegion.GetHandle, useCombineMode)
End Function

Friend Function AddPath(ByRef srcPath As pd2DPath, Optional ByVal useCombineMode As PD_2D_CombineMode = P2_CM_Replace) As Boolean
    If (m_RegionHandle = 0) Then Me.CreateRegion
    AddPath = GDI_Plus.GDIPlus_RegionAddPath(m_RegionHandle, srcPath.GetHandle, useCombineMode)
End Function

Friend Function CloneRegion(ByRef srcRegion As pd2DRegion) As Boolean
    If (m_RegionHandle <> 0) Then Me.ReleaseRegion
    CloneRegion = GDI_Plus.GDIPlus_RegionClone(srcRegion.GetHandle, m_RegionHandle)
End Function

Friend Function GetRegionBoundsF() As RectF
    If (m_RegionHandle <> 0) Then GetRegionBoundsF = GDI_Plus.GDIPlus_RegionGetClipRectF(m_RegionHandle)
End Function

Friend Function GetRegionBoundsL() As RectL
    If (m_RegionHandle <> 0) Then GetRegionBoundsL = GDI_Plus.GDIPlus_RegionGetClipRectI(m_RegionHandle)
End Function

Friend Function IsPointInRegion(ByVal srcX As Single, ByVal srcY As Single) As Boolean
    If (m_RegionHandle <> 0) Then IsPointInRegion = GDI_Plus.GDIPlus_RegionContainsPoint(m_RegionHandle, srcX, srcY)
End Function

Friend Function IsPointInRegionInt(ByVal srcX As Long, ByVal srcY As Long) As Boolean
    If (m_RegionHandle <> 0) Then IsPointInRegionInt = GDI_Plus.GDIPlus_RegionContainsPointI(m_RegionHandle, srcX, srcY)
End Function

'Don't use this function unless you know what you're doing.  This will copy (*not* clone) an external handle from the
' current backend.  That handle *must not be in use elsewhere*, because if it is deleted, this class has no way of
' detecting that.  (This function is only used when a convoluted API does not allow us to directly grab a returned
' handle ourselves.)
Friend Sub AssignExternalHandleDirectly(ByVal srcHandle As Long)
    If (m_RegionHandle <> 0) Then Me.ReleaseRegion
    m_RegionHandle = srcHandle
End Sub

Friend Function MakeRegionInfinite() As Boolean
    If (m_RegionHandle = 0) Then Me.CreateRegion
    MakeRegionInfinite = GDI_Plus.GDIPlus_RegionSetInfinite(m_RegionHandle)
End Function

Friend Function MakeRegionEmpty() As Boolean
    If (m_RegionHandle = 0) Then Me.CreateRegion
    MakeRegionEmpty = GDI_Plus.GDIPlus_RegionSetEmpty(m_RegionHandle)
End Function

Friend Function IsRegionInfinite() As Boolean
    If (m_RegionHandle = 0) Then Me.CreateRegion
    IsRegionInfinite = GDI_Plus.GDIPlus_RegionIsInfinite(m_RegionHandle)
End Function

Friend Function IsRegionEmpty() As Boolean
    If (m_RegionHandle = 0) Then Me.CreateRegion
    IsRegionEmpty = GDI_Plus.GDIPlus_RegionIsEmpty(m_RegionHandle)
End Function

Friend Function IsEqual(ByRef srcRegion As pd2DRegion) As Boolean
    If (m_RegionHandle = 0) Then Me.CreateRegion
    IsEqual = GDI_Plus.GDIPlus_RegionsAreEqual(srcRegion.GetHandle, m_RegionHandle)
End Function

Friend Function GetHandle(Optional ByVal createAsNecessary As Boolean = True) As Long
    If (createAsNecessary And (m_RegionHandle = 0)) Then
        If CreateRegion() Then GetHandle = m_RegionHandle Else GetHandle = 0
    Else
        GetHandle = m_RegionHandle
    End If
End Function

'Retrieve a copy of this region in GDI format.  Note that the caller is responsible for freeing the hRgn
' via DeleteObject when finished.
Friend Function GetRegionAsHRgn() As Long
    GetRegionAsHRgn = GDI_Plus.GDIPlus_RegionGetHRgn(m_RegionHandle)
End Function

Friend Function HasRegion() As Boolean
    HasRegion = (m_RegionHandle <> 0)
End Function

'Create an actual region handle using the current backend and the current region settings.
' NOTE: the caller doesn't *need* to call this directly.  If GetRegionHandle is called and the region doesn't
'       yet exist, it will be auto-created.
Friend Function CreateRegion() As Boolean

    If (m_RegionHandle <> 0) Then Me.ReleaseRegion
    m_RegionHandle = GDI_Plus.GetGDIPlusRegionHandle()
    
    'When debug mode is active, all object creations are reported back to the master Drawing2D module
    CreateRegion = (m_RegionHandle <> 0)
    If (CreateRegion And m_DebugMode) Then Drawing2D.DEBUG_NotifyRegionCountChange True
    
End Function

Friend Function ReleaseRegion() As Boolean
    
    If (m_RegionHandle <> 0) Then
        
        ReleaseRegion = GDI_Plus.ReleaseGDIPlusRegion(m_RegionHandle)
        
        'After a successful release, we must always reset the class-level handle to match, and during debug mode,
        ' the central Drawing2D module also needs to be notified.
        If ReleaseRegion Then
            m_RegionHandle = 0
            If (m_DebugMode) Then Drawing2D.DEBUG_NotifyRegionCountChange False
        End If
    
    Else
        ReleaseRegion = True
    End If
    
End Function

Friend Sub ResetAllProperties()
    Me.ReleaseRegion
End Sub

'When debug mode is active, this class will report object creation and destruction back to the master Drawing2D module.
Friend Sub SetDebugMode(ByVal newMode As Boolean)
    m_DebugMode = newMode
End Sub

Private Sub Class_Initialize()
    Me.ResetAllProperties
    m_DebugMode = Drawing2D.GetLibraryDebugMode()
End Sub

Private Sub Class_Terminate()
    Me.ReleaseRegion
End Sub

'All pd2D classes report errors using an internal function similar to this one.  Feel free to modify this function to
' better fit your project (e.g. perhaps it could raise an actual error event).
'
'Note that a default pd2D build simply dumps the passed error information to the Immediate window.
Private Sub InternalError(Optional ByRef errName As String = vbNullString, Optional ByRef errDescription As String = vbNullString, Optional ByVal ErrNum As Long = 0)
    Drawing2D.DEBUG_NotifyExternalError errName, errDescription, ErrNum, "pd2DRegion"
End Sub
