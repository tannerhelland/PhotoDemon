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
'Copyright 2014-2018 by Tanner Helland
'Created: 04/July/15 (but assembled from many bits written earlier)
'Last updated: 13/May/16
'Last update: total overhaul
'
'This class manages a single GDI+ WAPI pen instance.  Pens are used to trace shapes and paths.
'
'At present, this class is primarily based on the capabilities of GDI+.  This may change going forward,
' but because GDI+ provides a nice baseline feature set, that's where we started.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit
    
'Current pen style (solid, dashed, dotted, etc)
Private m_PenStyle As PD_2D_DashStyle

'Basic pen settings
Private m_PenColor As Long
Private m_PenOpacity As Single
Private m_PenWidth As Single

'Advanced pen settings
Private m_PenLineJoin As PD_2D_LineJoin
Private m_PenLineCap As PD_2D_LineCap
Private m_PenDashCap As PD_2D_DashCap
Private m_PenMiterLimit As Single
Private m_PenAlignment As PD_2D_PenAlignment
Private m_PenStartCap As PD_2D_LineCap
Private m_PenEndCap As PD_2D_LineCap
Private m_PenDashOffset As Single

'Pen transform properties (TBD)

'This class is capable of serializing itself to/from XML strings
Private cSerialize As pdParamXML

'When debug mode is active, object creation and destruction is reported to the central Drawing2D module
Private m_DebugMode As Boolean

'Once a pen has been created, this handle value will be non-zero
Private m_PenHandle As Long

'Get/set individual settings.  Note that these typically just wrap the generic Get/SetPenProperty functions, below.
Friend Function GetPenAlignment() As PD_2D_PenAlignment
    GetPenAlignment = Me.GetPenProperty(P2_PenAlignment)
End Function

Friend Function GetPenStyle() As PD_2D_DashStyle
    GetPenStyle = Me.GetPenProperty(P2_PenStyle)
End Function

Friend Function GetPenColor() As Long
    GetPenColor = Me.GetPenProperty(P2_PenColor)
End Function

Friend Function GetPenDashCap() As PD_2D_DashCap
    GetPenDashCap = Me.GetPenProperty(P2_PenDashCap)
End Function

Friend Function GetPenDashOffset() As Single
    GetPenDashOffset = Me.GetPenProperty(P2_PenDashOffset)
End Function

Friend Function GetPenEndCap() As PD_2D_LineCap
    GetPenEndCap = Me.GetPenProperty(P2_PenEndCap)
End Function

Friend Function GetPenLineCap() As PD_2D_LineCap
    GetPenLineCap = Me.GetPenProperty(P2_PenLineCap)
End Function

Friend Function GetPenLineJoin() As PD_2D_LineJoin
    GetPenLineJoin = Me.GetPenProperty(P2_PenLineJoin)
End Function

Friend Function GetPenMiterLimit() As Single
    GetPenMiterLimit = Me.GetPenProperty(P2_PenMiterLimit)
End Function

Friend Function GetPenOpacity() As Single
    GetPenOpacity = Me.GetPenProperty(P2_PenOpacity)
End Function

Friend Function GetPenStartCap() As PD_2D_LineCap
    GetPenStartCap = Me.GetPenProperty(P2_PenStartCap)
End Function

Friend Function GetPenWidth() As Single
    GetPenWidth = Me.GetPenProperty(P2_PenWidth)
End Function

Friend Sub SetPenAlignment(ByVal newSetting As PD_2D_PenAlignment)
     Me.SetPenProperty P2_PenAlignment, newSetting
End Sub

Friend Sub SetPenColor(ByVal newSetting As Long)
    Me.SetPenProperty P2_PenColor, newSetting
End Sub

Friend Sub SetPenDashCap(ByVal newSetting As PD_2D_DashCap)
     Me.SetPenProperty P2_PenDashCap, newSetting
End Sub

Friend Sub SetPenDashOffset(ByVal newSetting As Single)
    Me.SetPenProperty P2_PenDashOffset, newSetting
End Sub

Friend Sub SetPenEndCap(ByVal newSetting As PD_2D_LineCap)
     Me.SetPenProperty P2_PenEndCap, newSetting
End Sub

Friend Sub SetPenLineCap(ByVal newSetting As PD_2D_LineCap)
    Me.SetPenProperty P2_PenLineCap, newSetting
End Sub

Friend Sub SetPenLineJoin(ByVal newSetting As PD_2D_LineJoin)
    Me.SetPenProperty P2_PenLineJoin, newSetting
End Sub

Friend Sub SetPenMiterLimit(ByVal newSetting As Single)
     Me.SetPenProperty P2_PenMiterLimit, newSetting
End Sub

Friend Sub SetPenOpacity(ByVal newSetting As Single)
    Me.SetPenProperty P2_PenOpacity, newSetting
End Sub

Friend Sub SetPenStartCap(ByVal newSetting As PD_2D_LineCap)
     Me.SetPenProperty P2_PenStartCap, newSetting
End Sub

Friend Sub SetPenStyle(ByVal newSetting As PD_2D_DashStyle)
    Me.SetPenProperty P2_PenStyle, newSetting
End Sub

Friend Sub SetPenWidth(ByVal newSetting As Single)
     Me.SetPenProperty P2_PenWidth, newSetting
End Sub

'Unsafe subs are subs that let you *set* a pen property, but that property cannot be safely stored as part of the
' current pen's properties.  (In this case, dash arrays would need to be translated into some kind of string list,
' and it's not worth the trouble at present.)
'
'Subs like this offer better performance, but note that the GetPenPropertiesAsXML sub *will ignore this setting*!
Friend Sub SetPenDashes_UNSAFE(ByVal ptrToDashLengths As Long, ByVal numOfDashes As Long)
    If (m_PenHandle <> 0) Then
        GDI_Plus.GDIPlus_PenSetDashArray m_PenHandle, ptrToDashLengths, numOfDashes
    End If
End Sub

'Get/set individual settings by enum type
Friend Function GetPenProperty(ByVal propID As PD_2D_PEN_SETTINGS) As Variant
    
    Select Case propID
        
        Case P2_PenAlignment
            GetPenProperty = m_PenAlignment
        
        Case P2_PenColor
            GetPenProperty = m_PenColor
            
        Case P2_PenDashCap
            GetPenProperty = m_PenDashCap
        
        Case P2_PenDashOffset
            GetPenProperty = m_PenDashOffset
            
        Case P2_PenEndCap
            GetPenProperty = m_PenEndCap
        
        Case P2_PenLineCap
            GetPenProperty = m_PenLineCap
        
        Case P2_PenLineJoin
            GetPenProperty = m_PenLineJoin
            
        Case P2_PenMiterLimit
            GetPenProperty = m_PenMiterLimit
        
        Case P2_PenOpacity
            GetPenProperty = m_PenOpacity
            
        Case P2_PenStartCap
            GetPenProperty = m_PenStartCap
        
        Case P2_PenStyle
            GetPenProperty = m_PenStyle
        
        Case P2_PenWidth
            GetPenProperty = m_PenWidth
            
    End Select
    
End Function

Friend Sub SetPenProperty(ByVal propID As PD_2D_PEN_SETTINGS, ByVal newSetting As Variant)
    
    Select Case propID
        
        Case P2_PenAlignment
            m_PenAlignment = CLng(newSetting)
        
        Case P2_PenColor
            m_PenColor = CLng(newSetting)
        
        Case P2_PenDashCap
            m_PenDashCap = CLng(newSetting)
            
        Case P2_PenDashOffset
            m_PenDashOffset = CSng(newSetting)
        
        Case P2_PenEndCap
            m_PenEndCap = CLng(newSetting)
        
        Case P2_PenLineCap
            m_PenLineCap = CLng(newSetting)
            m_PenStartCap = CLng(newSetting)
            m_PenEndCap = CLng(newSetting)
        
        Case P2_PenLineJoin
            m_PenLineJoin = CLng(newSetting)
        
        Case P2_PenMiterLimit
            m_PenMiterLimit = CSng(newSetting)
        
        Case P2_PenOpacity
            m_PenOpacity = CSng(newSetting)
        
        Case P2_PenStartCap
            m_PenStartCap = CLng(newSetting)
            
        Case P2_PenStyle
            m_PenStyle = CLng(newSetting)
            
        Case P2_PenWidth
            m_PenWidth = CSng(newSetting)
            
    End Select
    
    RelayPropertyChangeToBackend propID
    
End Sub

Private Sub RelayPropertyChangeToBackend(ByVal propID As PD_2D_PEN_SETTINGS)
    
    'Property changes only need to be relayed if we've already created a handle
    If (m_PenHandle <> 0) Then
    
        Select Case propID
                
            Case P2_PenAlignment
                GDI_Plus.SetGDIPlusPenProperty m_PenHandle, propID, m_PenAlignment
                
            Case P2_PenColor
                GDI_Plus.SetGDIPlusPenProperty m_PenHandle, propID, m_PenColor
                
            Case P2_PenDashCap
                GDI_Plus.SetGDIPlusPenProperty m_PenHandle, propID, ConvertPDDashCapToGDIPDashCap(m_PenDashCap)
                    
            Case P2_PenDashOffset
                GDI_Plus.GDIPlus_PenSetDashOffset m_PenHandle, m_PenDashOffset
                
            Case P2_PenEndCap
                GDI_Plus.SetGDIPlusPenProperty m_PenHandle, P2_PenEndCap, m_PenEndCap
                
            Case P2_PenLineCap
                GDI_Plus.SetGDIPlusPenProperty m_PenHandle, propID, m_PenLineCap
                
            Case P2_PenLineJoin
                GDI_Plus.SetGDIPlusPenProperty m_PenHandle, propID, m_PenLineJoin
                
            Case P2_PenMiterLimit
                GDI_Plus.SetGDIPlusPenProperty m_PenHandle, propID, m_PenMiterLimit
                
            Case P2_PenOpacity
                GDI_Plus.SetGDIPlusPenProperty m_PenHandle, propID, m_PenOpacity
                
            Case P2_PenStartCap
                GDI_Plus.SetGDIPlusPenProperty m_PenHandle, P2_PenStartCap, m_PenStartCap
            
            Case P2_PenStyle
                GDI_Plus.SetGDIPlusPenProperty m_PenHandle, propID, m_PenStyle
                
            Case P2_PenWidth
                GDI_Plus.SetGDIPlusPenProperty m_PenHandle, propID, m_PenWidth
                
        End Select
        
    End If
        
End Sub

'For interop purposes, pens are passed around PD as XML strings.
Friend Function GetPenPropertiesAsXML() As String
    
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
        .AddParam "PenDashOffset", m_PenDashOffset
    End With
    
    GetPenPropertiesAsXML = cSerialize.GetParamString
    
End Function

Friend Sub SetPenPropertiesFromXML(ByRef srcString As String)
    
    'If the string is empty, prep a default object
    If Len(srcString) = 0 Then
        Me.ResetAllProperties
    Else
    
        If (cSerialize Is Nothing) Then Set cSerialize = New pdParamXML
        With cSerialize
            .SetParamString srcString
            Me.SetPenProperty P2_PenStyle, .GetLong("PenMode", P2_DS_Solid)
            Me.SetPenProperty P2_PenColor, .GetLong("PenColor", vbBlack)
            Me.SetPenProperty P2_PenOpacity, .GetDouble("PenOpacity", 100)
            Me.SetPenProperty P2_PenWidth, .GetDouble("PenWidth", 1#)
            Me.SetPenProperty P2_PenLineJoin, .GetLong("PenLineJoin", P2_LJ_Miter)
            Me.SetPenProperty P2_PenLineCap, .GetLong("PenLineCap", P2_LC_Flat)
            Me.SetPenProperty P2_PenDashCap, .GetLong("PenDashCap", P2_DC_Flat)
            Me.SetPenProperty P2_PenMiterLimit, .GetDouble("PenMiterLimit", 3#)
            Me.SetPenProperty P2_PenAlignment, .GetLong("PenAlignment", P2_PA_Center)
            Me.SetPenProperty P2_PenStartCap, .GetLong("PenStartCap", m_PenLineCap)
            Me.SetPenProperty P2_PenEndCap, .GetLong("PenEndCap", m_PenLineCap)
            Me.SetPenProperty P2_PenDashOffset, .GetSingle("PenDashOffset", m_PenDashOffset)
        End With
        
    End If
    
End Sub

Friend Function GetHandle(Optional ByVal createAsNecessary As Boolean = True) As Long
    If (createAsNecessary And (m_PenHandle = 0)) Then
        If CreatePen() Then GetHandle = m_PenHandle Else GetHandle = 0
    Else
        GetHandle = m_PenHandle
    End If
End Function

Friend Function HasPen() As Boolean
    HasPen = (m_PenHandle <> 0)
End Function

'Clone an existing pd2DPen instance.  This may not work for complicated pens like pens created from brushes;
' for that, we'd need to look into a different solution.  (Standard pens will clone just fine, though.)
Friend Function ClonePen(ByRef srcPen As pd2DPen) As Boolean
    If (srcPen Is Nothing) Then
        Me.ResetAllProperties
    Else
        Me.SetPenPropertiesFromXML srcPen.GetPenPropertiesAsXML()
    End If
End Function

'Create an actual pen handle using the current backend and the current pen settings.
' NOTE: the caller doesn't actually *need* to call this directly.  If GetPenHandle is called and the pen doesn't yet exist,
'       it will be auto-created.
Friend Function CreatePen() As Boolean
    
    If (m_PenHandle <> 0) Then ReleasePen
    
    m_PenHandle = GDI_Plus.GetGDIPlusPenHandle(m_PenColor, m_PenOpacity * 2.55, m_PenWidth, m_PenLineCap, m_PenLineJoin, m_PenStyle, m_PenMiterLimit, m_PenAlignment)
            
    If (m_PenHandle <> 0) Then
        GDI_Plus.SetGDIPlusPenProperty m_PenHandle, P2_PenDashCap, ConvertPDDashCapToGDIPDashCap(m_PenDashCap)
        If (m_PenStartCap <> m_PenLineCap) Then GDI_Plus.SetGDIPlusPenProperty m_PenHandle, P2_PenStartCap, m_PenStartCap
        If (m_PenEndCap <> m_PenLineCap) Then GDI_Plus.SetGDIPlusPenProperty m_PenHandle, P2_PenEndCap, m_PenEndCap
        GDI_Plus.GDIPlus_PenSetDashOffset m_PenHandle, m_PenDashOffset
    End If
    CreatePen = (m_PenHandle <> 0)
    
    'When debug mode is active, all object creations are reported back to the master Drawing2D module
    If (CreatePen And m_DebugMode) Then Drawing2D.DEBUG_NotifyPenCountChange True
    
End Function

'So, some notes about this function:
' 1) Yes, it will create a pen with the same attributes as the brush it receives.  Very cool!
' 2) Because pens-from-brushes are a uniquely weird use-case, these pens obey slightly different rules.
'    Specifically, you *must* observe some special criteria when creating them:
'    - You *must* set a pen width prior to creation.  Changing the pen width post-creation is not guaranteed to work.
'    - You *cannot* serialize the pen to/from a string.  (Instead, serialize the source pd2DBrush object.)
'    - This function only works with the GDI+ backend.
Friend Function CreatePenFromBrush(ByRef srcBrush As pd2DBrush) As Boolean
    
    If (m_PenHandle <> 0) Then ReleasePen
    
    m_PenHandle = GDI_Plus.GetGDIPlusPenFromBrush(srcBrush.GetHandle, m_PenWidth)
    CreatePenFromBrush = (m_PenHandle <> 0)
    
    'When debug mode is active, all object creations are reported back to the master Drawing2D module
    If (CreatePenFromBrush And m_DebugMode) Then Drawing2D.DEBUG_NotifyPenCountChange True
    
End Function

Friend Function ReleasePen() As Boolean
    
    If (m_PenHandle <> 0) Then
        
        ReleasePen = GDI_Plus.ReleaseGDIPlusPen(m_PenHandle)
        
        'After a successful release, we must always reset the class-level handle to match, and during debug mode,
        ' the central Drawing2D module also needs to be notified.
        If ReleasePen Then
            m_PenHandle = 0
            If (m_DebugMode) Then Drawing2D.DEBUG_NotifyPenCountChange False
        End If
    
    Else
        ReleasePen = True
    End If
    
End Function

Friend Sub ResetAllProperties()
    Me.SetPenStyle P2_DS_Solid
    Me.SetPenColor vbBlack
    Me.SetPenOpacity 100!
    Me.SetPenWidth 1!
    Me.SetPenLineJoin P2_LJ_Miter
    Me.SetPenLineCap P2_LC_Flat
    Me.SetPenDashCap P2_DC_Flat
    Me.SetPenMiterLimit 3#
    Me.SetPenAlignment P2_PA_Center
    Me.SetPenStartCap P2_LC_Flat
    Me.SetPenEndCap P2_LC_Flat
    Me.SetPenDashOffset 0!
End Sub

'When debug mode is active, this class will report object creation and destruction back to the master Drawing2D module.
Friend Sub SetDebugMode(ByVal newMode As Boolean)
    m_DebugMode = newMode
End Sub

'Some GDI+ enums differ from their PD equivalents.  Use the functions below for convenience conversions.
Private Function ConvertPDDashCapToGDIPDashCap(ByVal pdDashCap As PD_2D_DashCap) As GP_DashCap
    
    Select Case pdDashCap
    
        Case P2_DC_Square
            ConvertPDDashCapToGDIPDashCap = GP_DC_Square
            
        Case Else
            ConvertPDDashCapToGDIPDashCap = pdDashCap
            
    End Select
    
End Function

Private Sub Class_Initialize()
    Me.ResetAllProperties
    m_DebugMode = Drawing2D.GetLibraryDebugMode()
End Sub

Private Sub Class_Terminate()
    Me.ReleasePen
End Sub

'All pd2D classes report errors using an internal function similar to this one.  Feel free to modify this function to
' better fit your project (e.g. perhaps it could raise an actual error event).
'
'Note that a default pd2D build simply dumps the passed error information to the Immediate window.
Private Sub InternalError(Optional ByRef errName As String = vbNullString, Optional ByRef errDescription As String = vbNullString, Optional ByVal ErrNum As Long = 0)
    Drawing2D.DEBUG_NotifyExternalError errName, errDescription, ErrNum, "pd2DPen"
End Sub
