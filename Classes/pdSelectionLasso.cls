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

'Each selection shape has various properties specific to that shape.  To improve extensibility,
' properties are stored in a dictionary.  All properties are relayed to the current selection shape
' instance, but each instance only needs to use the properties they care about.
Private m_PropertyDict As pdDictionary

'Unique to lasso selections is a variable-length list of exterior path points.
Private m_NumOfLassoPoints As Long, m_LassoPoints() As PointFloat

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

'Initialize this shape from a
