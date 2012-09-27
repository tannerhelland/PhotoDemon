VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Selection class
'Copyright ©2011-2012 by Tanner Helland
'Created: 25/September/12
'Last updated: 27/September/12
'Last update: continued work on initial build
'
'This class handles all selections in PhotoDemon.  At present, selections are simply rectangular areas - but this will change
' in the future as new selection modes are added.
'
'All selections are treated as a subset of pixels within a rectangular area.  The x1, y1, x2, y2 coordinates define opposite
' corners of the selection area.  There is no guarantee that these coordinates will appear in a given corner - so they must
' be checked manually.  When a selection is set, those values are translated into Left, Top, Width and Height values, which
' are in turn used for everything from cropping to applying filters.
'
'Note also that this class is treated as a subset of pdLayer().  Right now each layer object only contains one selection,
' but there's no reason it couldn't store multiples in the future.  (Although I'm not sure what purpose that would serve.)
'
'***************************************************************************

Option Explicit

Private x1 As Long, x2 As Long
Private y1 As Long, y2 As Long

Private lockedIn As Boolean

Public Left As Long, Top As Long
Public Width As Long, Height As Long

'Takes x and y coordinates (from a _MouseDown event, typically) and stores them internally
Public Sub setInitialCoordinates(ByVal x As Long, y As Long)
    
    'If new initial coordinates are being set, this selection must be "unlocked"
    lockedIn = False
    
    x1 = x
    y1 = y
    x2 = x
    y2 = y
    
End Sub

'Takes x and y coordinates (from a _MouseMove event, typically) and stores them internally
Public Sub setAdditionalCoordinates(ByVal x As Long, y As Long)
    x2 = x
    y2 = y
End Sub

'Has this selection been locked in?
Public Function isLockedIn() As Boolean
    isLockedIn = lockedIn
End Function

'"Lock-in" a selection.  Typically this is prompted by a _MouseUp event
Public Sub lockIn()
    
    'Mark this selection as locked-in
    lockedIn = True
    
    'Calculate left, top width, height based off the current x1/x2/y1/y2 values
    If x1 < x2 Then
        Left = x1
        Width = x2 - x1
    Else
        Left = x2
        Width = x1 - x2
    End If
    
    If y1 < y2 Then
        Top = y1
        Height = y2 - y1
    Else
        Top = y2
        Height = y1 - y2
    End If
    
End Sub

'Render an intermediate version of the selection onto the canvas.  This uses the x1/x2/y1/y2 values, and is optimized for speed.
Public Sub renderIntermediate()

End Sub

'Render the final, locked in selection onto a given image.  This emphasizes quality over
Public Sub renderFinal()

End Sub

Private Sub Class_Initialize()
    lockedIn = False
End Sub
