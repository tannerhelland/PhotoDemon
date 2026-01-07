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
'Copyright 2014-2026 by Tanner Helland
'Created: 04/July/15 (but assembled from many bits written earlier)
'Last updated: 25/November/20
'Last update: finish inlining GDI+ calls
'
'This class manages a single GDI+ WAPI pen instance.  Pens are used to trace shapes and paths.
'
'At present, this class is primarily based on the capabilities of GDI+.  This may change going forward,
' but because GDI+ provides a nice baseline feature set, that's where we started.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'GDI+ declares
Private Declare Function GdipCreatePen1 Lib "gdiplus" (ByVal srcColor As Long, ByVal srcWidth As Single, ByVal srcUnit As GP_Unit, ByRef dstPen As Long) As GP_Result
Private Declare Function GdipCreatePenFromBrush Lib "gdiplus" Alias "GdipCreatePen2" (ByVal srcBrush As Long, ByVal penWidth As Single, ByVal srcUnit As GP_Unit, ByRef dstPen As Long) As GP_Result
Private Declare Function GdipDeletePen Lib "gdiplus" (ByVal hPen As Long) As GP_Result
Private Declare Function GdipSetPenColor Lib "gdiplus" (ByVal hPen As Long, ByVal pARGBColor As Long) As GP_Result
Private Declare Function GdipSetPenDashArray Lib "gdiplus" (ByVal hPen As Long, ByVal ptrToDashArray As Long, ByVal numOfDashes As Long) As GP_Result
Private Declare Function GdipSetPenDashCap Lib "gdiplus" Alias "GdipSetPenDashCap197819" (ByVal hPen As Long, ByVal newCap As GP_DashCap) As GP_Result
Private Declare Function GdipSetPenDashOffset Lib "gdiplus" (ByVal hPen As Long, ByVal newPenOffset As Single) As GP_Result
Private Declare Function GdipSetPenDashStyle Lib "gdiplus" (ByVal hPen As Long, ByVal newDashStyle As GP_DashStyle) As GP_Result
Private Declare Function GdipSetPenEndCap Lib "gdiplus" (ByVal hPen As Long, ByVal endCap As GP_LineCap) As GP_Result
Private Declare Function GdipSetPenLineCap Lib "gdiplus" Alias "GdipSetPenLineCap197819" (ByVal hPen As Long, ByVal startCap As GP_LineCap, ByVal endCap As GP_LineCap, ByVal dashCap As GP_DashCap) As GP_Result
Private Declare Function GdipSetPenLineJoin Lib "gdiplus" (ByVal hPen As Long, ByVal newLineJoin As GP_LineJoin) As GP_Result
Private Declare Function GdipSetPenMiterLimit Lib "gdiplus" (ByVal hPen As Long, ByVal newMiterLimit As Single) As GP_Result
Private Declare Function GdipSetPenStartCap Lib "gdiplus" (ByVal hPen As Long, ByVal startCap As GP_LineCap) As GP_Result
Private Declare Function GdipSetPenWidth Lib "gdiplus" (ByVal hPen As Long, ByVal penWidth As Single) As GP_Result

'GDI+ hypothetically supports pens that are both centered along the path (standard behavior),
' or offset from the path (non-standard).  Offset pens are extremely unreliable, with many potential
' failure conditions, so this class does not attempt to support them.
'Private Declare Function GdipSetPenMode Lib "gdiplus" (ByVal hPen As Long, ByVal penMode As GP_PenAlignment) As GP_Result

'Current pen style (solid, dashed, dotted, etc)
Private m_PenStyle As PD_2D_DashStyle

'Basic pen settings
Private m_PenColorOpacity As RGBQuad    'GDI+ requires RGB quads; we translate between this format as necessary
Private m_PenColorLong As Long          'Must always be manually synched against m_PenColorOpacity
Private m_PenWidth As Single

'Advanced pen settings
Private m_PenLineJoin As PD_2D_LineJoin
Private m_PenLineCap As PD_2D_LineCap
Private m_PenDashCap As PD_2D_DashCap
Private m_PenMiterLimit As Single
Private m_PenStartCap As PD_2D_LineCap
Private m_PenEndCap As PD_2D_LineCap
Private m_PenDashOffset As Single

'Pen transform properties (TBD)

'This class is capable of serializing itself to/from XML strings
Private m_Serialize As pdSerialize

'Once a pen has been created, this handle value will be non-zero
Private m_PenHandle As Long

'Get/set individual settings.
Friend Function GetPenColor() As Long
    GetPenColor = RGB(m_PenColorOpacity.Red, m_PenColorOpacity.Green, m_PenColorOpacity.Blue)
End Function

Friend Function GetPenColorRGBA() As RGBQuad
    GetPenColorRGBA = m_PenColorOpacity
End Function

Friend Function GetPenDashCap() As PD_2D_DashCap
    GetPenDashCap = m_PenDashCap
End Function

Friend Function GetPenDashOffset() As Single
    GetPenDashOffset = m_PenDashOffset
End Function

Friend Function GetPenEndCap() As PD_2D_LineCap
    GetPenEndCap = m_PenEndCap
End Function

Friend Function GetPenLineCap() As PD_2D_LineCap
    GetPenLineCap = m_PenLineCap
End Function

Friend Function GetPenLineJoin() As PD_2D_LineJoin
    GetPenLineJoin = m_PenLineJoin
End Function

Friend Function GetPenMiterLimit() As Single
    GetPenMiterLimit = m_PenMiterLimit
End Function

Friend Function GetPenOpacity() As Single
    GetPenOpacity = m_PenColorOpacity.Alpha / 2.55!
End Function

Friend Function GetPenStartCap() As PD_2D_LineCap
    GetPenStartCap = m_PenStartCap
End Function

Friend Function GetPenStyle() As PD_2D_DashStyle
    GetPenStyle = m_PenStyle
End Function

Friend Function GetPenWidth() As Single
    GetPenWidth = m_PenWidth
End Function

'Color does *not* currently support OLE color constants; you need to manually uncomment
' the OleTranslateColor call, below, if you care about this (PD does not, and it's obviously
' faster to skip the function call)
Friend Function SetPenColor(ByVal newSetting As Long) As Boolean
    
    'See documentation for OleTranslateColor before enabling this line
    'If OleTranslateColor(newSetting, 0, newSetting) Then 'failure - do what you want!
    
    m_PenColorOpacity.Red = Colors.ExtractRed(newSetting)
    m_PenColorOpacity.Green = Colors.ExtractGreen(newSetting)
    m_PenColorOpacity.Blue = Colors.ExtractBlue(newSetting)
    
    'Synchronize against our long-type value
    GetMem4 VarPtr(m_PenColorOpacity), m_PenColorLong
    
    'Apply immediately
    If (m_PenHandle <> 0) Then SetPenColor = (GdipSetPenColor(m_PenHandle, m_PenColorLong) = GP_OK)
    
End Function

'Color does *not* currently support OLE color constants; you need to manually add a call to
' OleTranslateColor if you want to blindly use system color constants
Friend Function SetPenColorRGBA(ByRef newSetting As RGBQuad) As Boolean
    
    m_PenColorOpacity = newSetting
    
    'Synchronize against our long-type value
    GetMem4 VarPtr(m_PenColorOpacity), m_PenColorLong
    
    'Apply immediately
    If (m_PenHandle <> 0) Then SetPenColorRGBA = (GdipSetPenColor(m_PenHandle, m_PenColorLong) = GP_OK)
    
End Function

Friend Function SetPenDashCap(ByVal newSetting As PD_2D_DashCap) As Boolean
    m_PenDashCap = newSetting
    If (m_PenHandle <> 0) Then SetPenDashCap = (GdipSetPenDashCap(m_PenHandle, ConvertPDDashCapToGDIPDashCap(newSetting)) = GP_OK)
End Function

Friend Function SetPenDashOffset(ByVal newSetting As Single) As Boolean
    m_PenDashOffset = newSetting
    If (m_PenHandle <> 0) Then SetPenDashOffset = (GdipSetPenDashOffset(m_PenHandle, newSetting) = GP_OK)
End Function

Friend Function SetPenEndCap(ByVal newSetting As PD_2D_LineCap) As Boolean
    m_PenEndCap = newSetting
    If (m_PenHandle <> 0) Then SetPenEndCap = (GdipSetPenEndCap(m_PenHandle, newSetting) = GP_OK)
End Function

'Convenience function to set start, end, and generic linecap properties all at once
Friend Function SetPenLineCap(ByVal newSetting As PD_2D_LineCap) As Boolean
    m_PenLineCap = newSetting
    m_PenStartCap = m_PenLineCap
    m_PenEndCap = m_PenLineCap
    If (m_PenHandle <> 0) Then SetPenLineCap = (GdipSetPenLineCap(m_PenHandle, newSetting, newSetting, ConvertPDDashCapToGDIPDashCap(m_PenDashCap)) = GP_OK)
End Function

Friend Function SetPenLineJoin(ByVal newSetting As PD_2D_LineJoin) As Boolean
    m_PenLineJoin = newSetting
    If (m_PenHandle <> 0) Then SetPenLineJoin = (GdipSetPenLineJoin(m_PenHandle, newSetting) = GP_OK)
End Function

Friend Function SetPenMiterLimit(ByVal newSetting As Single) As Boolean
    m_PenMiterLimit = newSetting
    If (m_PenHandle <> 0) Then SetPenMiterLimit = (GdipSetPenMiterLimit(m_PenHandle, newSetting) = GP_OK)
End Function

'NOTE!  PEN OPACITY setting is treated as a single on the range [0, 100], *not* as a byte on the range [0, 255]
Friend Function SetPenOpacity(ByVal newSetting As Single) As Boolean
    m_PenColorOpacity.Alpha = Int(newSetting * 2.55! + 0.5!)
    GetMem4 VarPtr(m_PenColorOpacity), m_PenColorLong
    If (m_PenHandle <> 0) Then SetPenOpacity = (GdipSetPenColor(m_PenHandle, m_PenColorLong) = GP_OK)
End Function

Friend Function SetPenStartCap(ByVal newSetting As PD_2D_LineCap) As Boolean
    m_PenStartCap = newSetting
    If (m_PenHandle <> 0) Then SetPenStartCap = (GdipSetPenStartCap(m_PenHandle, newSetting) = GP_OK)
End Function

Friend Function SetPenStyle(ByVal newSetting As PD_2D_DashStyle) As Boolean
    m_PenStyle = newSetting
    If (m_PenHandle <> 0) Then SetPenStyle = (GdipSetPenDashStyle(m_PenHandle, newSetting) = GP_OK)
End Function

Friend Function SetPenWidth(ByVal newSetting As Single) As Boolean
    m_PenWidth = newSetting
    If (m_PenHandle <> 0) Then SetPenWidth = (GdipSetPenWidth(m_PenHandle, newSetting) = GP_OK)
End Function

'Unsafe subs have very strict rules.
' 1) They can only be called after a pen has been created (e.g. a non-zero pen handle MUST exist).
' 2) They let you *set* a pen property, but that property cannot be *retrieved again* (usually because
'     said property has a complex format)
' 3) The property you set will *not* be persisted if/when this pen is serialized or copied.  You must
'     manually persist and/or copy any unsafe properties.
'
'The above limitations could be worked around with some effort, but I don't have a use for this behavior
' so I haven't tackled it (yet).
Friend Function SetPenDashes_UNSAFE(ByVal ptrToDashLengths As Long, ByVal numOfDashes As Long) As Boolean
    If (m_PenHandle <> 0) Then
        SetPenDashes_UNSAFE = (GdipSetPenDashArray(m_PenHandle, ptrToDashLengths, numOfDashes) = GP_OK)
    Else
        InternalError "SetPenDashes_UNSAFE", "_UNSAFE functions require non-null handle"
    End If
End Function

'For interop purposes, pens are passed around PD as XML strings.
Friend Function GetPenPropertiesAsXML() As String
    
    If (m_Serialize Is Nothing) Then Set m_Serialize = New pdSerialize
    With m_Serialize
        .Reset 1#
        .AddParam "pen-color", RGB(m_PenColorOpacity.Red, m_PenColorOpacity.Green, m_PenColorOpacity.Blue), True, True
        .AddParam "pen-opacity", m_PenColorOpacity.Alpha / 2.55!, True, True
        .AddParam "pen-width", m_PenWidth, True, True
        .AddParam "pen-line-join", Drawing2D.XML_GetNameOfLineJoin(m_PenLineJoin), True, True
        .AddParam "pen-miter-limit", m_PenMiterLimit, True, True
        .AddParam "pen-style", Drawing2D.XML_GetNameOfDashStyle(m_PenStyle), True, True
        .AddParam "pen-line-cap", Drawing2D.XML_GetNameOfLineCap(m_PenLineCap), True, True
        .AddParam "pen-start-cap", Drawing2D.XML_GetNameOfLineCap(m_PenStartCap), True, True
        .AddParam "pen-end-cap", Drawing2D.XML_GetNameOfLineCap(m_PenEndCap), True, True
        .AddParam "pen-dash-cap", Drawing2D.XML_GetNameOfDashCap(m_PenDashCap), True, True
        .AddParam "pen-dash-offset", m_PenDashOffset, True, True
    End With
    
    GetPenPropertiesAsXML = m_Serialize.GetParamString
    
End Function

Friend Sub SetPenPropertiesFromXML(ByRef srcString As String)
    
    'If the string is empty, prep a default object
    If (LenB(srcString) = 0) Then
        Me.ResetAllProperties
    Else
    
        If (m_Serialize Is Nothing) Then Set m_Serialize = New pdSerialize
        With m_Serialize
            .SetParamString srcString
            
            'Check for modern property storage
            If .DoesParamExist("pen-color") Then
                
                Me.SetPenColor .GetLong("pen-color", vbBlack, True)
                Me.SetPenOpacity .GetSingle("pen-opacity", 100!, True)
                Me.SetPenWidth .GetSingle("pen-width", 1!, True)
                Me.SetPenLineJoin Drawing2D.XML_GetLineJoinFromName(.GetString("pen-line-join", , True))
                Me.SetPenMiterLimit .GetSingle("pen-miter-limit", 3!, True)
                Me.SetPenStyle Drawing2D.XML_GetDashStyleFromName(.GetString("pen-style", , True))
                Me.SetPenLineCap Drawing2D.XML_GetLineCapFromName(.GetString("pen-line-cap", , True))
                Me.SetPenStartCap Drawing2D.XML_GetLineCapFromName(.GetString("pen-start-cap", , True))
                Me.SetPenEndCap Drawing2D.XML_GetLineCapFromName(.GetString("pen-end-cap", , True))
                Me.SetPenDashCap Drawing2D.XML_GetDashCapFromName(.GetString("pen-dash-cap", , True))
                Me.SetPenDashOffset .GetSingle("pen-dash-offset", m_PenDashOffset, True)
                
            'Use legacy naming scheme
            Else
                Me.SetPenStyle .GetLong("PenMode", P2_DS_Solid)
                Me.SetPenColor .GetLong("PenColor", vbBlack)
                Me.SetPenOpacity .GetDouble("PenOpacity", 100)
                Me.SetPenWidth .GetDouble("PenWidth", 1#)
                Me.SetPenLineJoin .GetLong("PenLineJoin", P2_LJ_Miter)
                Me.SetPenLineCap .GetLong("PenLineCap", P2_LC_Flat)
                Me.SetPenDashCap .GetLong("PenDashCap", P2_DC_Flat)
                Me.SetPenMiterLimit .GetDouble("PenMiterLimit", 3#)
                Me.SetPenStartCap .GetLong("PenStartCap", m_PenLineCap)
                Me.SetPenEndCap .GetLong("PenEndCap", m_PenLineCap)
                Me.SetPenDashOffset .GetSingle("PenDashOffset", m_PenDashOffset)
            End If
            
        End With
        
    End If
    
End Sub

Friend Function GetHandle(Optional ByVal createAsNecessary As Boolean = True) As Long
    If (m_PenHandle <> 0) Then
        GetHandle = m_PenHandle
    ElseIf createAsNecessary Then
        If CreatePen() Then GetHandle = m_PenHandle
    End If
End Function

Friend Function HasPen() As Boolean
    HasPen = (m_PenHandle <> 0)
End Function

'Clone an existing pd2DPen instance.  This does not work on pens created with _UNSAFE suffixes,
' or pens created from raster sources (e.g. brushes); for those, you'll need a custom solution.
Friend Function ClonePen(ByRef srcPen As pd2DPen) As Boolean
    If (srcPen Is Nothing) Then
        Me.ResetAllProperties
    Else
        Me.SetPenPropertiesFromXML srcPen.GetPenPropertiesAsXML()
    End If
End Function

'Create an actual pen handle using the current backend and the current pen settings.
' NOTE: the caller doesn't actually *need* to call this directly.  If GetPenHandle() is called
'       and the pen doesn't yet exist, it will be auto-created.
Friend Function CreatePen() As Boolean
    
    If (m_PenHandle <> 0) Then Me.ReleasePen
    
    'Create the base pen, but note that only *some* settings are specified in this initial call.
    If (GdipCreatePen1(m_PenColorLong, m_PenWidth, GP_U_Pixel, m_PenHandle) = GP_OK) Then
        
        CreatePen = (m_PenHandle <> 0)
        If CreatePen Then
            
            'A base pen was created successfully, but only color, opacity, and width have been set.
            
            'Now we need to relay all other settings to the pen.
            
            'Because VB6 doesn't provide a good mechanism for declaring constants from an enum,
            ' I perform manual checks against known default GDI+ properties here.  (This makes pen
            ' creation faster for default solid pens.)
            If (m_PenLineCap <> P2_LC_Flat) Then CreatePen = CreatePen And Me.SetPenLineCap(m_PenLineCap)
            If (m_PenStartCap <> m_PenLineCap) Then CreatePen = CreatePen And Me.SetPenStartCap(m_PenStartCap)
            If (m_PenEndCap <> m_PenLineCap) Then CreatePen = CreatePen And Me.SetPenEndCap(m_PenEndCap)
            If (m_PenLineJoin <> P2_LJ_Miter) Then CreatePen = CreatePen And Me.SetPenLineJoin(m_PenLineJoin) Else CreatePen = CreatePen And Me.SetPenMiterLimit(m_PenMiterLimit)
            If (m_PenStyle <> P2_DS_Solid) Then
                CreatePen = CreatePen And Me.SetPenStyle(m_PenStyle)
                CreatePen = CreatePen And Me.SetPenDashCap(m_PenDashCap)
                CreatePen = CreatePen And Me.SetPenDashOffset(m_PenDashOffset)
            End If
            
        Else
            InternalError "CreatePen", "null pen handle after creation"
        End If
        
    Else
        InternalError "CreatePen", "GdipCreatePen1 failed"
    End If
    
    'When debug mode is active, all object creations are reported back to the central Drawing2D module
    If PD2D_DEBUG_MODE Then
        If CreatePen Then
            Drawing2D.DEBUG_NotifyPenCountChange True
        Else
            InternalError "CreatePen", "create failed; reason unknown"
        End If
    End If
    
End Function

'So, some notes about this function:
' 1) Yes, it will create a pen with the same attributes as the brush it receives.  Very cool!
' 2) Because pens-from-brushes are a uniquely weird use-case, these pens obey slightly different rules.
'    Specifically, you *must* observe some special criteria when creating them:
'    - You *must* set a pen width prior to creation.  Changing the pen width post-creation is not
'       guaranteed to work (because that portion of the bitmap may not have been cached at create-time).
'    - You *cannot* successfully serialize a brush-pen to/from a string.  (Instead, serialize the source
'       pd2DBrush object.)
Friend Function CreatePenFromBrush(ByRef srcBrush As pd2DBrush) As Boolean
    
    If (m_PenHandle <> 0) Then Me.ReleasePen
    
    CreatePenFromBrush = (GdipCreatePenFromBrush(srcBrush.GetHandle, m_PenWidth, GP_U_Pixel, m_PenHandle) = GP_OK)
    If CreatePenFromBrush Then CreatePenFromBrush = (m_PenHandle <> 0)
    
    'When debug mode is active, all object creations are reported back to the central Drawing2D module
    If (CreatePenFromBrush And PD2D_DEBUG_MODE) Then Drawing2D.DEBUG_NotifyPenCountChange True
    
End Function

Friend Function ReleasePen() As Boolean
    
    If (m_PenHandle <> 0) Then
        
        ReleasePen = (GdipDeletePen(m_PenHandle) = GP_OK)
        
        'After a successful release, we must always reset the class-level handle to match,
        ' and during debug mode, the central Drawing2D module also needs to be notified.
        If ReleasePen Then
            m_PenHandle = 0
            If PD2D_DEBUG_MODE Then Drawing2D.DEBUG_NotifyPenCountChange False
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
    Me.SetPenStartCap P2_LC_Flat
    Me.SetPenEndCap P2_LC_Flat
    Me.SetPenDashOffset 0!
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
End Sub

Private Sub Class_Terminate()
    Me.ReleasePen
End Sub

'All pd2D classes report errors using an internal function similar to this one.
' Feel free to modify this function to better fit your project
' (for example, maybe you prefer to raise an actual error event).
'
'Note that by default, pd2D build simply dumps all error information to the Immediate window.
Private Sub InternalError(ByRef errFunction As String, ByRef errDescription As String, Optional ByVal errNum As Long = 0)
    Drawing2D.DEBUG_NotifyError "pd2DPen", errFunction, errDescription, errNum
End Sub

