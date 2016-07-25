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
'Copyright 2016-2016 by Tanner Helland
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
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'This class is not yet capable of serializing itself to/from XML strings, but it may be possible in the future...
Private cSerialize As pdParamXML

'The backend used to manage this region.  At present, this is immutable; once set, it cannot be migrated.
Private m_RegionBackend As PD_2D_RENDERING_BACKEND

'When debug mode is active, object creation and destruction is reported to the central Drawing2D module
Private m_DebugMode As Boolean

'Once a region has been created, this handle value will be non-zero
Private m_RegionHandle As Long

Friend Function AddRectangleF(ByVal rLeft As Single, ByVal rTop As Single, ByVal rWidth As Single, ByVal rHeight As Single, Optional ByVal useCombineMode As PD_2D_CombineMode = P2_CM_Replace) As Boolean
    
    If (m_RegionHandle = 0) Then Me.CreateRegion
    
    Dim tmpRectF As RECTF
    tmpRectF.Left = rLeft
    tmpRectF.Top = rTop
    tmpRectF.Width = rWidth
    tmpRectF.Height = rHeight
    
    Select Case m_RegionBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
            AddRectangleF = GDI_Plus.GDIPlus_RegionAddRectF(m_RegionHandle, tmpRectF, useCombineMode)
    End Select
    
End Function

Friend Function AddRectangle_FromRectF(ByRef srcRectF As RECTF, Optional ByVal useCombineMode As PD_2D_CombineMode = P2_CM_Replace) As Boolean
    
    If (m_RegionHandle = 0) Then Me.CreateRegion
    
    Select Case m_RegionBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
            AddRectangle_FromRectF = GDI_Plus.GDIPlus_RegionAddRectF(m_RegionHandle, srcRectF, useCombineMode)
    End Select
    
End Function

Friend Function AddRectangle_FromRectL(ByRef srcRectL As RECTL, Optional ByVal useCombineMode As PD_2D_CombineMode = P2_CM_Replace) As Boolean
    
    If (m_RegionHandle = 0) Then Me.CreateRegion
    
    Select Case m_RegionBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
            AddRectangle_FromRectL = GDI_Plus.GDIPlus_RegionAddRectL(m_RegionHandle, srcRectL, useCombineMode)
    End Select
    
End Function

Friend Function AddRegion(ByRef srcRegion As pd2DRegion, Optional ByVal useCombineMode As PD_2D_CombineMode = P2_CM_Replace) As Boolean
    
    If (m_RegionHandle = 0) Then Me.CreateRegion
    
    Select Case m_RegionBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
            AddRegion = GDI_Plus.GDIPlus_RegionAddRegion(m_RegionHandle, srcRegion.GetHandle, useCombineMode)
    End Select
    
End Function

Friend Function AddPath(ByRef srcPath As pd2DPath, Optional ByVal useCombineMode As PD_2D_CombineMode = P2_CM_Replace) As Boolean
    
    If (m_RegionHandle = 0) Then Me.CreateRegion
    
    Select Case m_RegionBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
            AddPath = GDI_Plus.GDIPlus_RegionAddPath(m_RegionHandle, srcPath.GetHandle, useCombineMode)
    End Select
    
End Function

Friend Function CloneRegion(ByRef srcRegion As pd2DRegion) As Boolean

    If (m_RegionHandle <> 0) Then Me.ReleaseRegion
    
    Select Case m_RegionBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
            CloneRegion = GDI_Plus.GDIPlus_RegionClone(srcRegion.GetHandle, m_RegionHandle)
    End Select
    
End Function

Friend Function GetRegionBoundsF() As RECTF
    If (m_RegionHandle <> 0) Then
        Select Case m_RegionBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                GetRegionBoundsF = GDI_Plus.GDIPlus_RegionGetClipRectF(m_RegionHandle)
        End Select
    End If
End Function

Friend Function GetRegionBoundsL() As RECTL
    If (m_RegionHandle <> 0) Then
        Select Case m_RegionBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                GetRegionBoundsL = GDI_Plus.GDIPlus_RegionGetClipRectI(m_RegionHandle)
        End Select
    End If
End Function

'Don't use this function unless you know what you're doing.  This will copy (*not* clone) an external handle from the
' current backend.  That handle *must not be in use elsewhere*, because if it is deleted, this class has no way of
' detecting that.  (This function is only used when a convoluted API does not allow us to directly grab a returned
' handle ourselves.)
Public Sub AssignExternalHandleDirectly(ByVal srcHandle As Long)
    If (m_RegionHandle <> 0) Then Me.ReleaseRegion
    m_RegionHandle = srcHandle
End Sub

Public Function MakeRegionInfinite() As Boolean

    If (m_RegionHandle = 0) Then Me.CreateRegion
    
    Select Case m_RegionBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
             MakeRegionInfinite = GDI_Plus.GDIPlus_RegionSetInfinite(m_RegionHandle)
    End Select
    
End Function

Public Function MakeRegionEmpty() As Boolean

    If (m_RegionHandle = 0) Then Me.CreateRegion
    
    Select Case m_RegionBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
             MakeRegionEmpty = GDI_Plus.GDIPlus_RegionSetEmpty(m_RegionHandle)
    End Select
    
End Function

Public Function IsRegionInfinite() As Boolean

    If (m_RegionHandle = 0) Then Me.CreateRegion
    
    Select Case m_RegionBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
             IsRegionInfinite = GDI_Plus.GDIPlus_RegionIsInfinite(m_RegionHandle)
    End Select
    
End Function

Public Function IsRegionEmpty() As Boolean

    If (m_RegionHandle = 0) Then Me.CreateRegion
    
    Select Case m_RegionBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
             IsRegionEmpty = GDI_Plus.GDIPlus_RegionIsEmpty(m_RegionHandle)
    End Select
    
End Function

Public Function IsEqual(ByRef srcRegion As pd2DRegion) As Boolean
    
    If (m_RegionHandle = 0) Then Me.CreateRegion
    
    Select Case m_RegionBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
             IsEqual = GDI_Plus.GDIPlus_RegionsAreEqual(srcRegion.GetHandle, m_RegionHandle)
    End Select
    
End Function

Public Function GetHandle(Optional ByVal createAsNecessary As Boolean = True) As Long
    If (createAsNecessary And (m_RegionHandle = 0)) Then
        If CreateRegion() Then GetHandle = m_RegionHandle Else GetHandle = 0
    Else
        GetHandle = m_RegionHandle
    End If
End Function

Public Function GetRegionBackend() As PD_2D_RENDERING_BACKEND
    GetRegionBackend = m_RegionBackend
End Function

Public Function HasRegion() As Boolean
    HasRegion = CBool(m_RegionHandle <> 0)
End Function

'You only need to manually initialize a backend if you plan on using a NON-DEFAULT backend.
' (Objects with the default backend are initialized automatically.)
Friend Function InitializeRegionBackend(Optional ByVal targetBackend As PD_2D_RENDERING_BACKEND = P2_DefaultBackend) As Boolean
    m_RegionBackend = targetBackend
    InitializeRegionBackend = True
End Function

'Create an actual region handle using the current backend and the current region settings.
' NOTE: the caller doesn't *need* to call this directly.  If GetRegionHandle is called and the region doesn't
'       yet exist, it will be auto-created.
Friend Function CreateRegion() As Boolean

    If (m_RegionHandle <> 0) Then Me.ReleaseRegion
    
    'Call the backend-specific creation function
    Select Case m_RegionBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
            m_RegionHandle = GDI_Plus.GetGDIPlusRegionHandle()
    End Select
    
    'When debug mode is active, all object creations are reported back to the master Drawing2D module
    CreateRegion = CBool(m_RegionHandle <> 0)
    If (CreateRegion And m_DebugMode) Then Drawing2D.DEBUG_NotifyRegionCountChange m_RegionBackend, True
    
End Function

Public Function ReleaseRegion() As Boolean
    
    If (m_RegionHandle <> 0) Then
        
        'Call the backend-specific release function
        Select Case m_RegionBackend
            Case P2_DefaultBackend, P2_GDIPlusBackend
                ReleaseRegion = GDI_Plus.ReleaseGDIPlusRegion(m_RegionHandle)
        End Select
        
        'After a successful release, we must always reset the class-level handle to match, and during debug mode,
        ' the central Drawing2D module also needs to be notified.
        If ReleaseRegion Then
            m_RegionHandle = 0
            If (m_DebugMode) Then Drawing2D.DEBUG_NotifyRegionCountChange m_RegionBackend, False
        End If
    
    Else
        ReleaseRegion = True
    End If
    
End Function

Public Sub ResetAllProperties()
    Me.ReleaseRegion
End Sub

'When debug mode is active, this class will report object creation and destruction back to the master Drawing2D module.
Public Sub SetDebugMode(ByVal newMode As Boolean)
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

