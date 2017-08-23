VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdObjectList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PD Object List Manager
'Copyright 2017-2017 by Tanner Helland
'Created: 18/August/17
'Last updated: 18/August/17
'Last update: initial build
'
'This class is exclusively used by the Tabs module.  Look there for implementation details.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'The hWnd of the parent (the form hosting this collection)
Private m_ParentHWnd As Long

'hWnds of child controls
Private Const INIT_NUM_OF_CONTROLS As Long = 16

'We may not need to use a custom type for this - an hWnd may ultimately be enough - but during debugging,
' it's nice to store some extra bits of information
Private Type PD_Control
    hWnd As Long
    ctlName As String
    ctlType As PD_ControlType
End Type

Private m_Controls() As PD_Control
Private m_numOfControls As Long

Friend Sub SetParentHWnd(ByVal hWnd As Long)
    m_ParentHWnd = hWnd
End Sub

Friend Function GetParentHWnd() As Long
    GetParentHWnd = m_ParentHWnd
End Function

Friend Function DoesHWndExist(ByVal srcHwnd As Long) As Boolean
    
    DoesHWndExist = False
    
    If (m_numOfControls > 0) Then
        Dim i As Long
        For i = 0 To m_numOfControls
            If (m_Controls(i).hWnd = srcHwnd) Then
                DoesHWndExist = True
                Exit For
            End If
        Next i
    End If
    
End Function

Friend Function DoesTypeOfControlExist(ByVal srcType As PD_ControlType) As Boolean
    
    DoesTypeOfControlExist = False
    
    If (m_numOfControls > 0) Then
        Dim i As Long
        For i = 0 To m_numOfControls
            If (m_Controls(i).ctlType = srcType) Then
                DoesTypeOfControlExist = True
                Exit For
            End If
        Next i
    End If
    
End Function

Friend Function GetFirstHWndForType(ByVal srcType As PD_ControlType) As Long

    GetFirstHWndForType = 0
    
    If (m_numOfControls > 0) Then
        Dim i As Long
        For i = 0 To m_numOfControls
            If (m_Controls(i).ctlType = srcType) Then
                GetFirstHWndForType = m_Controls(i).hWnd
                Exit For
            End If
        Next i
    End If

End Function

'Load individual controls using this function; we'll retrieve whatever we need from them
Friend Sub NotifyChildControl(ByRef childControl As Object)

    'Make sure we have room for this hWnd
    If (m_numOfControls = 0) Then
        ReDim m_Controls(0 To INIT_NUM_OF_CONTROLS - 1) As PD_Control
    Else
        If (m_numOfControls > UBound(m_Controls)) Then ReDim Preserve m_Controls(0 To m_numOfControls * 2 - 1) As PD_Control
    End If
    
    With m_Controls(m_numOfControls)
        .hWnd = childControl.hWnd
        .ctlName = childControl.GetControlName()
        .ctlType = childControl.GetControlType()
    End With
    
    m_numOfControls = m_numOfControls + 1
    
End Sub

Friend Sub PrintDebugList()
    
    If (m_numOfControls > 0) Then
        
        Dim i As Long
        For i = 0 To m_numOfControls - 1
            Debug.Print m_Controls(i).ctlName, m_Controls(i).hWnd, UserControls.GetNameOfControlType(m_Controls(i).ctlType)
        Next i
        
    End If
    
End Sub
