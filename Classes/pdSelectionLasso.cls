VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdSelection_Lasso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Selection Manager (Lasso)
'Copyright 2012-2021 by Tanner Helland
'Created: 25/September/12
'Last updated: 15/February/21
'Last update: work on splitting shape-specific selection tools into their own classes
'
'For many years, PhotoDemon implemented all selection tools inside a single (giant!) class.
' In 2021, work began on migrating this to a more sustainable, extensible system.  This class
' is part of that effort.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'All selection shapes must return a unique, immutable name
Private Const THIS_SHAPE_NAME As String = "lasso"

'These properties must be supported by *every* selection type.
Private m_SelPropArea As PD_SelectionArea
Private m_SelPropSmoothing As PD_EdgeSmoothing
Private m_SelPropFeatherRadius As Single
Private m_SelPropBorderWidth As Single

'Other properties unique to this shape can be referenced here
Private m_NumOfLassoPoints As Long, m_LassoPoints() As PointFloat


'Each selection shape has additional properties specific to that shape.  (For example, rectangular
' selections have a "rounded corner radius" property unique to them.)  To improve extensibility,
' properties are stored in a dictionary.  All properties are relayed to the current selection shape
' instance, but each instance only needs to use the properties they care about.
Private m_PropertyDict As pdDictionary

'When move transforms are applied to a lasso selection, this class maintains an "original" copy of
' the path's outline.  Any transformation values are applied to this "original" copy, and the copy
' is erased when the mouse is released.
Private m_LassoPointsBackup() As PointFloat

'This class tracks lasso state; "closed "means the mouse has been released and the lasso is fixed
Private m_LassoClosed As Boolean

Public Function GetShapeID() As PD_SelectionShape
    GetShapeID = ss_Lasso
End Function

Public Function GetShapeName() As String
    GetShapeName = THIS_SHAPE_NAME
End Function

'Some shapes can be transformed after their creation (e.g. a square can be resized).
' Other shapes cannot (e.g. a "magic wand" selection).  Return TRUE if this shape can
' handle run-time transform events.
Public Function GetShapeTransformable() As Boolean
    GetShapeTransformable = False
End Function

'Initialize this shape from a serialized XML string.  The XML string may contain items this
' class doesn't use; that's fine.  The parent class pulls some properties for itself, so this
' class simply needs to grab items it recognizes while ignoring the rest.
'
'Note also that all coordinates in PD selections are stored as relative values on the scale [0, 1].
' This is deliberate so that selections saved/loaded to file (e.g. during macro recording) can be
' auto-scaled to the current image's size.  The parent image's dimensions must be passed to this
' function for correct coordinate translation.
'
'Returns: TRUE if the initialized selection is different from the current one.  This clue can be
' used to rebuild the current selection mask.
Public Function InitializeFromXML(ByRef srcXML As String, ByVal parentImageWidth As Long, ByVal parentImageHeight As Long) As Boolean

    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString srcXML
    
    'These properties are universal and must be supported by *all* selection types
    m_SelPropArea = cParams.GetLong("SelectionArea", sa_Interior)
    m_SelPropSmoothing = cParams.GetLong("SelectionSmoothing", es_None)
    m_SelPropFeatherRadius = cParams.GetSingle("SelectionFeatheringRadius", 0!)
    m_SelPropBorderWidth = cParams.GetSingle("SelectionBorderWidth", 1!)
    
    'Any additional properties are specific to this shape.
    ' (For lasso selections, that means the list of lasso points.)
    '
    'Originally there was also a selection parameter here for "selection smooth stroke";
    ' I've since dropped this parameter is it was not especially helpful and it introduced some
    ' potential issues, especially on sharp corners.
    'SetSelectionProperty sp_SmoothStroke, cParams.GetDouble("SelectionSmoothStroke", 0#)
    
    m_NumOfLassoPoints = cParams.GetLong("SelectionNumOfPoints", 0)
    If (m_NumOfLassoPoints > 0) Then
    
        ReDim m_LassoPoints(0 To m_NumOfLassoPoints - 1) As PointFloat
        
        'Retrieve all remaining lasso points
        If (m_NumOfLassoPoints > 0) Then
            Dim i As Long
            For i = 0 To m_NumOfLassoPoints - 1
                With m_LassoPoints(i)
                    .x = PDMath.TranslateValue_RelToAbs(cParams.GetDouble("SelectionPointX" & CStr(i + 1), 0#), parentImageWidth)
                    .y = PDMath.TranslateValue_RelToAbs(cParams.GetDouble("SelectionPointY" & CStr(i + 1), 0#), parentImageHeight)
                End With
            Next i
        End If
        
    Else
        Const ARBITRARY_INIT_LASSO_POINTS As Long = 128
        ReDim m_LassoPoints(0 To ARBITRARY_INIT_LASSO_POINTS - 1) As PointFloat
    End If

End Function

'Return all of this selection's important settings as an XML packet.  If the current selection
' shape is vector-based, this packet can be used to exactly re-create this selection.
' (See also: "InitializeFromXML()" above)
Public Function SerializeToXML() As String

End Function

Private Sub Class_Initialize()
    Set m_PropertyDict = New pdDictionary
End Sub
