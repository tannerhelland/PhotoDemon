VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdGraphicsPen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Graphics Outline ("Pen" in GDI/GDI+ parlance) Class
'Copyright 2014-2015 by Tanner Helland
'Created: 04/July/15 (but assembled from many bits written earlier)
'Last updated: 04/July/15
'Last update: start migrating code from other places in PD to this singular class
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

'Pens support a *lot* of internal settings.
Public Enum PD_GRAPHICS_PEN_SETTINGS
    pgps_PenMode = 0
    pgps_PenColor = 1
    pgps_PenOpacity = 2
    pgps_PenWidth = 3
    pgps_PenLineJoin = 4
    pgps_PenLineCap = 5
    pgps_PenDashCap = 6
    pgps_PenMiterLimit = 7
    pgps_PenAlignment = 8
End Enum

#If False Then
    Const pgps_PenMode = 0, pgps_PenColor = 1, pgps_PenOpacity = 2, pgps_PenWidth = 3, pgps_PenLineJoin = 4
    Const pgps_PenLineCap = 5, pgps_PenDashCap = 6, pgps_PenMiterLimit = 7, pgps_PenAlignment = 8
#End If

'Current pen mode (solid vs dashes, dots, etc)
Private m_PenMode As Long

'Basic pen settings
Private m_PenColor As Long
Private m_PenOpacity As Single
Private m_PenWidth As Single

'Advanced pen settings
Private m_PenLineJoin As Long
Private m_PenLineCap As Long
Private m_PenDashCap As Long
Private m_PenMiterLimit As Single
Private m_PenAlignment As Long

'Some GDI+-specific functions are declared here, as they're of limited use elsewhere
Private Declare Function GdipSetPenDashCap197819 Lib "gdiplus" (ByVal srcPen As Long, ByVal newDashCap As Long) As Long

'This class is capable of serializing itself to/from XML strings
Private cSerialize As pdParamXML


'Pen transforms (TBD)

'Get/set individual settings by enum type
Public Function getPenProperty(ByVal propID As PD_GRAPHICS_PEN_SETTINGS) As Variant
    
    Select Case propID
        
        Case pgps_PenMode
            getPenProperty = m_PenMode
        
        Case pgps_PenColor
            getPenProperty = m_PenColor
            
        Case pgps_PenOpacity
            getPenProperty = m_PenOpacity
            
        Case pgps_PenWidth
            getPenProperty = m_PenWidth
            
        Case pgps_PenLineJoin
            getPenProperty = m_PenLineJoin
            
        Case pgps_PenLineCap
            getPenProperty = m_PenLineCap
            
        Case pgps_PenDashCap
            getPenProperty = m_PenDashCap
            
        Case pgps_PenMiterLimit
            getPenProperty = m_PenMiterLimit
            
        Case pgps_PenAlignment
            getPenProperty = m_PenAlignment
            
    End Select
    
End Function

Public Sub setPenProperty(ByVal propID As PD_GRAPHICS_PEN_SETTINGS, ByVal newSetting As Variant)
    
    Select Case propID
        
        Case pgps_PenMode
            m_PenMode = CLng(newSetting)
        
        Case pgps_PenColor
            m_PenColor = CLng(newSetting)
            
        Case pgps_PenOpacity
            m_PenOpacity = CSng(newSetting)
            
        Case pgps_PenWidth
            m_PenWidth = CSng(newSetting)
            
        Case pgps_PenLineJoin
            m_PenLineJoin = CLng(newSetting)
            
        Case pgps_PenLineCap
            m_PenLineCap = CLng(newSetting)
            
        Case pgps_PenDashCap
            m_PenDashCap = CLng(newSetting)
            
        Case pgps_PenMiterLimit
            m_PenMiterLimit = CSng(newSetting)
            
        Case pgps_PenAlignment
            m_PenAlignment = CLng(newSetting)
            
    End Select
    
End Sub

'For interop purposes, pens are passed around PD as XML strings.
Public Function getPenAsString() As String
    
    With cSerialize
        .Reset 1#
        .addParam "PenMode", m_PenMode
        .addParam "PenColor", m_PenColor
        .addParam "PenOpacity", m_PenOpacity
        .addParam "PenWidth", m_PenWidth
        .addParam "PenLineJoin", m_PenLineJoin
        .addParam "PenLineCap", m_PenLineCap
        .addParam "PenDashCap", m_PenDashCap
        .addParam "PenMiterLimit", m_PenMiterLimit
        .addParam "PenAlignment", m_PenAlignment
    End With
    
    getPenAsString = cSerialize.getParamString
    
End Function

Public Sub createPenFromString(ByRef srcString As String)
    
    'If the string is empty, prep a default object
    If Len(srcString) = 0 Then
        
        m_PenMode = 0
        m_PenColor = vbBlack
        m_PenOpacity = 100
        m_PenWidth = 1
        m_PenLineJoin = 0
        m_PenLineCap = 0
        m_PenDashCap = 0
        m_PenMiterLimit = 10#
        m_PenAlignment = PenAlignmentCenter
        
    Else
    
        'For now, we use a pdParamString class to make this possible.  Eventually, we will be moving to a new format.
        With cSerialize
            .setParamString srcString
            m_PenMode = .GetLong("PenMode", 0)
            m_PenColor = .GetLong("PenColor", vbBlack)
            m_PenOpacity = .GetDouble("PenOpacity", 100)
            m_PenWidth = .GetDouble("PenWidth", 1)
            m_PenLineJoin = .GetLong("PenLineJoin", 0)
            m_PenLineCap = .GetLong("PenLineCap", 0)
            m_PenDashCap = .GetLong("PenDashCap", 0)
            m_PenMiterLimit = .GetDouble("PenMiterLimit", 10#)
            m_PenAlignment = .GetLong("PenAlignment", PenAlignmentCenter)
        End With
        
    End If
    
End Sub

'Return an actual GDI+ pen handle matching the current settings.
' Note that the caller is responsible for freeing that handle once done.
Public Function getPenHandle() As Long
    
    getPenHandle = GDI_Plus.getGDIPlusPenHandle(m_PenColor, m_PenOpacity * 2.55, m_PenWidth, m_PenLineCap, m_PenLineJoin, m_PenMode, m_PenMiterLimit, m_PenAlignment)
    
    'Set some additional settings manually
    GdipSetPenDashCap197819 getPenHandle, m_PenDashCap
    
End Function

'Helper function to simplify pen destruction
Public Sub releasePenHandle(ByVal srcPen As Long)
    GDI_Plus.releaseGDIPlusPen srcPen
End Sub

Private Sub Class_Initialize()
    
    'Prep a string serializer
    Set cSerialize = New pdParamXML
    
End Sub
