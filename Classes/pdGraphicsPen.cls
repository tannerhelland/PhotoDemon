VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pd2DPen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Graphics Outline ("Pen" in GDI/GDI+ parlance) Class
'Copyright 2014-2016 by Tanner Helland
'Created: 04/July/15 (but assembled from many bits written earlier)
'Last updated: 13/May/16
'Last update: total overhaul
'
'This class manages a single GDI+ WAPI pen instance.  (Pens are used to trace shapes/paths.)
'
'At present, it is primarily based on the capabilities of GDI+.  This may change going forward, but because GDI+ provides
' a nice baseline feature set, we started with it.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit
    
'Current pen style (solid, dashed, dotted, etc)
Private m_PenStyle As GP_DashStyle

'Basic pen settings
Private m_PenColor As Long
Private m_PenOpacity As Single
Private m_PenWidth As Single

'Advanced pen settings
Private m_PenLineJoin As GP_LineJoin
Private m_PenLineCap As GP_LineCap
Private m_PenDashCap As GP_DashCap
Private m_PenMiterLimit As Single
Private m_PenAlignment As GP_PenAlignment
Private m_PenStartCap As GP_LineCap
Private m_PenEndCap As GP_LineCap

'Pen transform properties (TBD)

'This class is capable of serializing itself to/from XML strings
Private cSerialize As pdParamXML

'The backend used to render this pen.  At present, this is immutable; once set, it cannot be migrated
Private m_PenBackend As PD_2D_RENDERING_BACKEND

'When debug mode is active, object creation and destruction is reported to the central Drawing2D module
Private m_DebugMode As Boolean

'Once a pen has been created, this handle value will be non-zero
Private m_PenHandle As Long

'Get/set individual settings.  Note that these just wrap the generic Get/SetPenProperty functions, below.
Public Function GetPenStyle() As GP_DashStyle
    GetPenStyle = Me.GetPenProperty(PD2D_PenStyle)
End Function

Public Function GetPenColor() As Long
    GetPenColor = Me.GetPenProperty(PD2D_PenColor)
End Function

Public Function GetPenOpacity() As Single
    GetPenOpacity = Me.GetPenProperty(PD2D_PenOpacity)
End Function

Public Function GetPenWidth() As Single
    GetPenWidth = Me.GetPenProperty(PD2D_PenWidth)
End Function

Public Function GetPenLineJoin() As GP_LineJoin
    GetPenLineJoin = Me.GetPenProperty(PD2D_PenLineJoin)
End Function

Public Function GetPenLineCap() As GP_LineCap
    GetPenLineCap = Me.GetPenProperty(PD2D_PenLineCap)
End Function

Public Function GetPenDashCap() As GP_DashCap
    GetPenDashCap = Me.GetPenProperty(PD2D_PenDashCap)
End Function

Public Function GetPenMiterLimit() As Single
    GetPenMiterLimit = Me.GetPenProperty(PD2D_PenMiterLimit)
End Function

Public Function GetPenAlignment() As GP_PenAlignment
    GetPenAlignment = Me.GetPenProperty(PD2D_PenAlignment)
End Function

Public Function GetPenStartCap() As GP_LineCap
    GetPenStartCap = Me.GetPenProperty(PD2D_PenStartCap)
End Function

Public Function GetPenEndCap() As GP_LineCap
    GetPenEndCap = Me.GetPenProperty(PD2D_PenEndCap)
End Function

Public Sub SetPenStyle(ByVal newSetting As GP_DashStyle)
    Me.SetPenProperty PD2D_PenStyle, newSetting
End Sub

Public Sub SetPenColor(ByVal newSetting As Long)
    Me.SetPenProperty PD2D_PenColor, newSetting
End Sub

Public Sub SetPenOpacity(ByVal newSetting As Single)
    Me.SetPenProperty PD2D_PenOpacity, newSetting
End Sub

Public Sub SetPenWidth(ByVal newSetting As Single)
     Me.SetPenProperty PD2D_PenWidth, newSetting
End Sub

Public Sub SetPenLineJoin(ByVal newSetting As GP_LineJoin)
    Me.SetPenProperty PD2D_PenLineJoin, newSetting
End Sub

Public Sub SetPenLineCap(ByVal newSetting As GP_LineCap)
    Me.SetPenProperty PD2D_PenLineCap, newSetting
End Sub

Public Sub SetPenDashCap(ByVal newSetting As GP_DashCap)
     Me.SetPenProperty PD2D_PenDashCap, newSetting
End Sub

Public Sub SetPenMiterLimit(ByVal newSetting As Single)
     Me.SetPenProperty PD2D_PenMiterLimit, newSetting
End Sub

Public Sub SetPenAlignment(ByVal newSetting As GP_PenAlignment)
     Me.SetPenProperty PD2D_PenAlignment, newSetting
End Sub

Public Sub SetPenStartCap(ByVal newSetting As GP_LineCap)
     Me.SetPenProperty PD2D_PenStartCap, newSetting
End Sub

Public Sub SetPenEndCap(ByVal newSetting As GP_LineCap)
     Me.SetPenProperty PD2D_PenEndCap, newSetting
End Sub

'Get/set individual settings by enum type
Public Function GetPenProperty(ByVal propID As PD_2D_PEN_SETTINGS) As Variant
    
    Select Case propID
        
        Case PD2D_PenStyle
            GetPenProperty = m_PenStyle
        
        Case PD2D_PenColor
            GetPenProperty = m_PenColor
            
        Case PD2D_PenOpacity
            GetPenProperty = m_PenOpacity
            
        Case PD2D_PenWidth
            GetPenProperty = m_PenWidth
            
        Case PD2D_PenLineJoin
            GetPenProperty = m_PenLineJoin
            
        Case PD2D_PenLineCap
            GetPenProperty = m_PenLineCap
            
        Case PD2D_PenDashCap
            GetPenProperty = m_PenDashCap
            
        Case PD2D_PenMiterLimit
            GetPenProperty = m_PenMiterLimit
            
        Case PD2D_PenAlignment
            GetPenProperty = m_PenAlignment
            
        Case PD2D_PenStartCap
            GetPenProperty = m_PenStartCap
        
        Case PD2D_PenEndCap
            GetPenProperty = m_PenEndCap
            
    End Select
    
End Function

Public Sub SetPenProperty(ByVal propID As PD_2D_PEN_SETTINGS, ByVal newSetting As Variant)
    
    Select Case propID
        
        Case PD2D_PenStyle
            m_PenStyle = CLng(newSetting)
        
        Case PD2D_PenColor
            m_PenColor = CLng(newSetting)
            
        Case PD2D_PenOpacity
            m_PenOpacity = CSng(newSetting)
            
        Case PD2D_PenWidth
            m_PenWidth = CSng(newSetting)
            
        Case PD2D_PenLineJoin
            m_PenLineJoin = CLng(newSetting)
            
        Case PD2D_PenLineCap
            m_PenLineCap = CLng(newSetting)
            m_PenStartCap = CLng(newSetting)
            m_PenEndCap = CLng(newSetting)
            
        Case PD2D_PenDashCap
            m_PenDashCap = CLng(newSetting)
            
        Case PD2D_PenMiterLimit
            m_PenMiterLimit = CSng(newSetting)
            
        Case PD2D_PenAlignment
            m_PenAlignment = CLng(newSetting)
            
    End Select
    
    RelayPropertyChangeToBackend propID
    
End Sub

Private Sub RelayPropertyChangeToBackend(ByVal propID As PD_2D_PEN_SETTINGS)
    
    'Property changes only need to be relayed if we've already created a handle
    If (m_PenHandle <> 0) Then
    
        'Relays obviously vary by the active backend!
        Select Case m_PenBackend
            
            Case PD2D_DefaultBackend, PD2D_GDIPlusBackend
                
                Select Case propID
                    Case PD2D_PenStyle
                        GDI_Plus.SetGDIPlusPenProperty m_PenHandle, propID, m_PenStyle
                        
                    Case PD2D_PenColor
                        GDI_Plus.SetGDIPlusPenProperty m_PenHandle, propID, m_PenColor
                        
                    Case PD2D_PenOpacity
                        GDI_Plus.SetGDIPlusPenProperty m_PenHandle, propID, m_PenOpacity
                        
                    Case PD2D_PenWidth
                        GDI_Plus.SetGDIPlusPenProperty m_PenHandle, propID, m_PenWidth
                        
                    Case PD2D_PenLineJoin
                        GDI_Plus.SetGDIPlusPenProperty m_PenHandle, propID, m_PenLineJoin
                        
                    Case PD2D_PenLineCap
                        GDI_Plus.SetGDIPlusPenProperty m_PenHandle, propID, m_PenLineCap
                        
                    Case PD2D_PenDashCap
                        GDI_Plus.SetGDIPlusPenProperty m_PenHandle, propID, m_PenDashCap
                        
                    Case PD2D_PenMiterLimit
                        GDI_Plus.SetGDIPlusPenProperty m_PenHandle, propID, m_PenMiterLimit
                        
                    Case PD2D_PenAlignment
                        GDI_Plus.SetGDIPlusPenProperty m_PenHandle, propID, m_PenAlignment
                        
                End Select
                
            'In the future, other backends can be handled here...
            Case Else
                
        End Select
    
    End If
        
End Sub

'For interop purposes, pens are passed around PD as XML strings.
Public Function GetPenPropertiesAsXML() As String
    
    If (cSerialize Is Nothing) Then Set cSerialize = New pdParamXML
    With cSerialize
        .Reset 1#
        .AddParam "PenMode", m_PenStyle
        .AddParam "PenColor", m_PenColor
        .AddParam "PenOpacity", m_PenOpacity
        .AddParam "PenWidth", m_PenWidth
        .AddParam "PenLineJoin", m_PenLineJoin
        .AddParam "PenLineCap", m_PenLineCap
        .AddParam "PenDashCap", m_PenDashCap
        .AddParam "PenMiterLimit", m_PenMiterLimit
        .AddParam "PenAlignment", m_PenAlignment
        .AddParam "PenStartCap", m_PenStartCap
        .AddParam "PenEndCap", m_PenEndCap
    End With
    
    GetPenPropertiesAsXML = cSerialize.GetParamString
    
End Function

Public Sub SetPenPropertiesFromXML(ByRef srcString As String)
    
    'If the string is empty, prep a default object
    If Len(srcString) = 0 Then
        Me.ResetAllProperties
    Else
    
        If (cSerialize Is Nothing) Then Set cSerialize = New pdParamXML
        With cSerialize
            .SetParamString srcString
            Me.SetPenProperty PD2D_PenStyle, .GetLong("PenMode", GP_DS_Solid)
            Me.SetPenProperty PD2D_PenColor, .GetLong("PenColor", vbBlack)
            Me.SetPenProperty PD2D_PenOpacity, .GetDouble("PenOpacity", 100)
            Me.SetPenProperty PD2D_PenWidth, .GetDouble("PenWidth", 1#)
            Me.SetPenProperty PD2D_PenLineJoin, .GetLong("PenLineJoin", GP_LJ_Miter)
            Me.SetPenProperty PD2D_PenLineCap, .GetLong("PenLineCap", GP_LC_Flat)
            Me.SetPenProperty PD2D_PenDashCap, .GetLong("PenDashCap", GP_DC_Flat)
            Me.SetPenProperty PD2D_PenMiterLimit, .GetDouble("PenMiterLimit", 3#)
            Me.SetPenProperty PD2D_PenAlignment, .GetLong("PenAlignment", GP_PA_Center)
            Me.SetPenProperty PD2D_PenStartCap, .GetLong("PenStartCap", m_PenLineCap)
            Me.SetPenProperty PD2D_PenEndCap, .GetLong("PenEndCap", m_PenLineCap)
        End With
        
    End If
    
End Sub

Public Function GetHandle(Optional ByVal createAsNecessary As Boolean = True) As Long
    If (createAsNecessary And (m_PenHandle = 0)) Then
        If CreatePen() Then GetHandle = m_PenHandle Else GetHandle = 0
    Else
        GetHandle = m_PenHandle
    End If
End Function

Public Function GetPenBackend() As PD_2D_RENDERING_BACKEND
    GetPenBackend = m_PenBackend
End Function

Public Function HasPen() As Boolean
    HasPen = CBool(m_PenHandle <> 0)
End Function

'You only need to manually initialize a pen backend if you plan on using a non-default backend.  (Pens with the
' default backend are initialized automatically.)
Friend Function InitializePenBackend(Optional ByVal targetBackend As PD_2D_RENDERING_BACKEND = PD2D_DefaultBackend) As Boolean
    m_PenBackend = targetBackend
    InitializePenBackend = True
End Function

'Create an actual pen handle using the current backend and the current pen settings.
' NOTE: the caller doesn't actually *need* to call this directly.  If GetPenHandle is called and the pen doesn't yet exist,
'       it will be auto-created.
Public Function CreatePen() As Boolean
    
    If (m_PenHandle <> 0) Then ReleasePen
    
    'Call the backend-specific surface creation function
    Select Case m_PenBackend
        
        Case PD2D_DefaultBackend, PD2D_GDIPlusBackend
            m_PenHandle = GDI_Plus.GetGDIPlusPenHandle(m_PenColor, m_PenOpacity * 2.55, m_PenWidth, m_PenLineCap, m_PenLineJoin, m_PenStyle, m_PenMiterLimit, m_PenAlignment)
            If (m_PenHandle <> 0) Then
                GDI_Plus.SetGDIPlusPenProperty m_PenHandle, PD2D_PenDashCap, m_PenDashCap
                If (m_PenStartCap <> m_PenLineCap) Then GDI_Plus.SetGDIPlusPenProperty m_PenHandle, PD2D_PenStartCap, m_PenStartCap
                If (m_PenEndCap <> m_PenLineCap) Then GDI_Plus.SetGDIPlusPenProperty m_PenHandle, PD2D_PenEndCap, m_PenEndCap
            End If
            CreatePen = (m_PenHandle <> 0)
            
    End Select
    
    'When debug mode is active, all object creations are reported back to the master Drawing2D module
    If (CreatePen And m_DebugMode) Then Drawing2D.DEBUG_NotifyPenCountChange m_PenBackend, True
    
End Function

Public Function ReleasePen() As Boolean
    
    If (m_PenHandle <> 0) Then
        
        'Call the backend-specific release function
        Select Case m_PenBackend
            Case PD2D_DefaultBackend, PD2D_GDIPlusBackend
                ReleasePen = GDI_Plus.ReleaseGDIPlusPen(m_PenHandle)
        End Select
        
        'After a successful release, we must always reset the class-level handle to match, and during debug mode,
        ' the central Drawing2D module also needs to be notified.
        If ReleasePen Then
            m_PenHandle = 0
            If (m_DebugMode) Then Drawing2D.DEBUG_NotifyPenCountChange m_PenBackend, False
        End If
    
    Else
        ReleasePen = True
    End If
    
End Function

Public Sub ResetAllProperties()
    Me.SetPenProperty PD2D_PenStyle, GP_DS_Solid
    Me.SetPenProperty PD2D_PenColor, vbBlack
    Me.SetPenProperty PD2D_PenOpacity, 100
    Me.SetPenProperty PD2D_PenWidth, 1#
    Me.SetPenProperty PD2D_PenLineJoin, GP_LJ_Miter
    Me.SetPenProperty PD2D_PenLineCap, GP_LC_Flat
    Me.SetPenProperty PD2D_PenDashCap, GP_DC_Flat
    Me.SetPenProperty PD2D_PenMiterLimit, 3#
    Me.SetPenProperty PD2D_PenAlignment, GP_PA_Center
    Me.SetPenProperty PD2D_PenStartCap, GP_LC_Flat
    Me.SetPenProperty PD2D_PenEndCap, GP_LC_Flat
End Sub

'When debug mode is active, this class will report object creation and destruction back to the master Drawing2D module.
Public Sub SetDebugMode(ByVal newMode As Boolean)
    m_DebugMode = newMode
End Sub

Private Sub Class_Initialize()
    Me.ResetAllProperties
End Sub

Private Sub Class_Terminate()
    Me.ReleasePen
End Sub
