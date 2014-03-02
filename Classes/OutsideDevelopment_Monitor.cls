VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Note: this file has been heavily modified for use within PhotoDemon.

'This class was originally written by Mike Dixon.

'You may download the original version of this code from the following link (good as of November '12):
' http://support.microsoft.com/kb/194578


Option Explicit

' --------------------------------------------------------------------------
'               Copyright (C) 1998 Microsoft Corporation                   '
' --------------------------------------------------------------------------
' You have a royalty-free right to use, modify, reproduce and distribute   '
' the Sample Application Files (and/or any modified version) in any way    '
' you find useful, provided that you agree that Microsoft has no warranty, '
' obligations or liability for any Sample Application Files.               '
' --------------------------------------------------------------------------
' Written by Mike Dixon (mikedix@microsoft.com)                            '
' --------------------------------------------------------------------------


'===================================================
'Monitor Class, Contains information about a monitor
'All values should be stored as pixels
'===================================================

Private m_Handle        As Long
Private m_Left          As Long
Private m_Right         As Long
Private m_Top           As Long
Private m_Bottom        As Long
Private m_WorkLeft      As Long
Private m_Workright     As Long
Private m_Worktop       As Long
Private m_Workbottom    As Long
Private m_Width         As Long
Private m_Height        As Long
Private m_Name          As String
Private m_FriendlyName  As String
Private m_Description   As String

'Addition by Tanner: EDID-based properties.  This is the only way I know to retrieve actual monitor size, and it's great
' for retrieving other detailed monitor info.  The way I've implemented it at present is to simply pass the EDID
' (if found) at creation time.  The setEDID function handles all parsing duties, and a copy of the EDID is stored
' "just in case".
Private m_hasEDID As Boolean
Private m_EDID() As Byte
Private m_EDIDWidthMM As Long
Private m_EDIDHeightMM As Long
Private m_EDIDWidthInches As Double
Private m_EDIDHeightInches As Double
Private m_EDIDDiagonalInches As Double

Private m_Primary As Boolean

'Retrieve the monitors size, as a string (useful for displaying to the user)
Public Function getMonitorSizeAsString() As String

    If m_hasEDID Then
        getMonitorSizeAsString = Format(m_EDIDDiagonalInches, "#0.0") & """"
    Else
        getMonitorSizeAsString = ""
    End If

End Function

Public Function hasEDID() As Boolean
    hasEDID = m_hasEDID
End Function

'At creation time, the clsMonitors class will attempt to retrieve an EDID for this monitor.  Regardless of success
' or failure, it will call this function.
Public Sub setEDID(edidArray() As Byte, Optional ByVal noEDIDFound As Boolean = False)

    'If no EDID could be found, mark the failure state and exit.
    If noEDIDFound Then
        Debug.Print "WARNING: Monitor EDID could not be retrieved; physical dimensions unknown"
        m_hasEDID = False
        Exit Sub
    End If
    
    'If we already have an EDID, exit
    If m_hasEDID Then Exit Sub
    
    m_hasEDID = True
    ReDim m_EDID(0 To UBound(edidArray)) As Byte
    
    'Note: CopyMemory would also work just fine here
    Dim i As Long
    For i = 0 To UBound(edidArray)
        m_EDID(i) = edidArray(i)
    Next i
    
    'With the EDID successfully retrieved, parse it for useful values.
    
    'Start with the monitor's physical dimensions, specifically the monitor's physical size in MM.
    m_EDIDWidthMM = ((m_EDID(68) And &HF0) * 16) + m_EDID(66)
    m_EDIDHeightMM = ((m_EDID(68) And &HF) * 256) + m_EDID(67)
    
    'For convenience, store a matching size in inches
    m_EDIDWidthInches = (m_EDIDWidthMM / 25.4)
    m_EDIDHeightInches = (m_EDIDHeightMM / 25.4)
    m_EDIDDiagonalInches = Sqr(m_EDIDWidthInches * m_EDIDWidthInches + m_EDIDHeightInches * m_EDIDHeightInches)
    
    'Debug.Print m_EDIDWidthInches, m_EDIDHeightInches, m_EDIDDiagonalInches
    
    'Next, we will try to retrieve a monitor name.  Monitor name is an optional descriptor (but provided 99.9% of the time) in one of the
    ' four VESA Descriptor blocks of the EDID.  Those descriptor blocks are found in locations:
    ' 54–71, 72–89, 90–107, 108–125
    ' Location of the name is not guaranteed in advance, so each block must be parsed individually until found.
    
'    Dim descriptorText As String
'    descriptorText = ""
'
'    For i = 54 To 71
'        descriptorText = descriptorText & Chr(m_EDID(i))
'    Next i
'
'    Debug.Print "Descriptor 1: " & descriptorText
'
'    descriptorText = ""
'    For i = 72 To 89
'        descriptorText = descriptorText & Chr(m_EDID(i))
'    Next i
'
'    Debug.Print "Descriptor 2: " & descriptorText
'
'    descriptorText = ""
'    For i = 90 To 107
'        descriptorText = descriptorText & Chr(m_EDID(i))
'    Next i
'
'    Debug.Print "Descriptor 3: " & descriptorText
'
'    descriptorText = ""
'    For i = 108 To 125
'        descriptorText = descriptorText & Chr(m_EDID(i))
'    Next i
'
'    Debug.Print "Descriptor 4: " & descriptorText
    
End Sub

Public Property Get isPrimary() As Boolean
    isPrimary = m_Primary
End Property

Public Property Let isPrimary(lPrimary As Boolean)
    m_Primary = lPrimary
End Property

Public Property Get Description() As String
    Description = m_Description
End Property

Public Property Let Description(lDescription As String)
    m_Description = lDescription
End Property

Public Property Get friendlyName() As String
    friendlyName = m_FriendlyName
End Property

Public Property Let friendlyName(lFriendlyName As String)
    m_FriendlyName = lFriendlyName
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Name(lName As String)
    m_Name = lName
End Property

Public Property Get Handle() As Long
    Handle = m_Handle
End Property

Public Property Let Handle(lHandle As Long)
    m_Handle = lHandle
End Property

Public Property Get Height() As Long
    Height = m_Height
End Property

Public Property Let Height(h As Long)
    m_Height = h
End Property

Public Property Get Width() As Long
    Width = m_Width
End Property

Public Property Let Width(w As Long)
    m_Width = w
End Property

Public Property Get Left() As Long
    Left = m_Left
End Property

Public Property Let Left(l As Long)
    m_Left = l
End Property

Public Property Get Right() As Long
    Right = m_Right
End Property

Public Property Let Right(r As Long)
    m_Right = r
End Property

Public Property Get Top() As Long
    Top = m_Top
End Property

Public Property Let Top(t As Long)
    m_Top = t
End Property

Public Property Get Bottom() As Long
    Bottom = m_Bottom
End Property

Public Property Let Bottom(b As Long)
    m_Bottom = b
End Property

Public Property Get WorkLeft() As Long
    WorkLeft = m_WorkLeft
End Property

Public Property Let WorkLeft(l As Long)
    m_WorkLeft = l
End Property

Public Property Get WorkRight() As Long
    WorkRight = m_Workright
End Property

Public Property Let WorkRight(r As Long)
    m_Workright = r
End Property

Public Property Get WorkTop() As Long
    WorkTop = m_Worktop
End Property

Public Property Let WorkTop(t As Long)
    m_Worktop = t
End Property

Public Property Get WorkBottom() As Long
    WorkBottom = m_Workbottom
End Property

Public Property Let WorkBottom(b As Long)
    m_Workbottom = b
End Property
