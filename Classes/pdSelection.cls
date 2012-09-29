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
'Last updated: 28/September/12
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

'SafeArray types for pointing VB arrays at arbitrary memory locations (in our case, bitmap data)
Private Type SAFEARRAYBOUND
    cElements As Long
    lBound   As Long
End Type

Private Type SAFEARRAY2D
    cDims      As Integer
    fFeatures  As Integer
    cbElements As Long
    cLocks     As Long
    pvData     As Long
    Bounds(1)  As SAFEARRAYBOUND
End Type

'SafeArray API functions
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal byteLength As Long)
Private Declare Function VarPtrArray Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long

'DC API functions
Private Declare Function GetDC Lib "user32" (ByVal HWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal HWnd As Long, ByVal hDC As Long) As Long

'Object API functions
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

'Pen drawing
Private Const PS_SOLID = 0
Private Const PS_DASH = 1
Private Const PS_DOT = 2
Private Const PS_DASHDOT = 3
Private Const PS_DASHDOTDOT = 4

Private Const RGN_AND = 1
Private Const RGN_OR = 2
Private Const RGN_XOR = 3
Private Const RGN_DIFF = 4
Private Const RGN_COPY = 5

Private Const HS_DIAGCROSS = 5

Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long

'Rectangle drawing
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long

'What type of selection is this?
Public Enum SelectionType
    sRectangle = 0
    sSquare = 1
    sCircle = 2
End Enum

Private sType As SelectionType

'The outer boundaries of the current selection
Private x1 As Long, x2 As Long
Private y1 As Long, y2 As Long

'When a selection is "locked in", the x and y values are converted to these values
Public selLeft As Long, selTop As Long
Public selWidth As Long, selHeight As Long

'Is this selection "locked in"?
Private lockedIn As Boolean

'Set a selection type
Public Sub setSelectionType(ByVal selType As SelectionType)
    sType = selType
End Sub

'Takes x and y coordinates (from a _MouseDown event, typically) and stores them internally
Public Sub setInitialCoordinates(ByVal x As Long, y As Long)
    
    'If new initial coordinates are being set, this selection must be "unlocked"
    lockedIn = False
    
    x1 = x
    y1 = y
    
    If sType = sRectangle Then
        x2 = x
        y2 = y
    End If
    
    updateInternalCoords
    
End Sub

'Takes x and y coordinates (from a _MouseMove event, typically) and stores them internally
Public Sub setAdditionalCoordinates(ByVal x As Long, y As Long)
    x2 = x
    y2 = y
    
    updateInternalCoords
    
End Sub

'Has this selection been locked in?
Public Function isLockedIn() As Boolean
    isLockedIn = lockedIn
End Function

Private Sub updateInternalCoords()

    'Calculate left, top width, height based off the current x1/x2/y1/y2 values
    If x1 < x2 Then
        selLeft = x1
        selWidth = x2 - x1
    Else
        selLeft = x2
        selWidth = x1 - x2
    End If
    
    If y1 < y2 Then
        selTop = y1
        selHeight = y2 - y1
    Else
        selTop = y2
        selHeight = y1 - y2
    End If

End Sub

'"Lock-in" a selection.  Typically this is prompted by a _MouseUp event
Public Sub lockIn()
    
    'Mark this selection as locked-in
    lockedIn = True
    
    updateInternalCoords
    
End Sub

'"Unlock" a selection
Public Sub lockRelease()

    lockedIn = False

End Sub

'Render an intermediate version of the selection onto the canvas.  This uses the x1/x2/y1/y2 values, and is optimized for speed.
Public Sub renderIntermediate(ByRef dstForm As Form, ByVal imgLeft As Long, ByVal imgTop As Long)

    'Before drawing, we need to convert our "image" coordinates back to "viewport" coordinates
    Static ZoomVal As Single
    ZoomVal = Zoom.ZoomArray(pdImages(dstForm.Tag).CurrentZoomValue)
    
    Static dstX1 As Single, dstX2 As Single
    Static dstY1 As Single, dstY2 As Single
    
    dstX1 = (x1 * ZoomVal) + imgLeft - dstForm.HScroll.Value
    dstX2 = (x2 * ZoomVal) + imgLeft - dstForm.HScroll.Value
    dstY1 = (y1 * ZoomVal) + imgTop - dstForm.VScroll.Value
    dstY2 = (y2 * ZoomVal) + imgTop - dstForm.VScroll.Value
    
    'Since we'll be using the API to draw our selection area, we need to initialize several brushes
    Dim hPen As Long, hOldPen As Long
    Dim hBrush As Long, hOldBrush As Long
    
    hPen = CreatePen(PS_DOT, 0, RGB(255, 255, 255))
    
    'Draw a box onto this image's front buffer
    dstForm.FrontBuffer.Line (dstX2, dstY2)-(dstX2, dstY1)
    dstForm.FrontBuffer.Line (dstX2, dstY2)-(dstX1, dstY2)
    dstForm.FrontBuffer.Line (dstX1, dstY1)-(dstX2, dstY1)
    dstForm.FrontBuffer.Line (dstX1, dstY1)-(dstX1, dstY2)
    
End Sub

'Render the final, locked in selection onto a given image.  This emphasizes quality over speed.
Public Sub renderFinal(ByRef dstForm As Form, ByVal imgLeft As Long, ByVal imgTop As Long, ByVal imgWidth As Long, ByVal imgHeight As Long)

    'Make sure the selection is locked in.  If it isn't, we probably meant to draw an intermediate region.  Call that and exit.
    If lockedIn = False Then
        renderIntermediate dstForm, imgLeft, imgTop
        Exit Sub
    End If
    
    'Before drawing, we need to convert our "image" coordinates back to "viewport" coordinates
    Static ZoomVal As Single
    ZoomVal = Zoom.ZoomArray(pdImages(dstForm.Tag).CurrentZoomValue)
    
    Static dstLeft As Single, dstRight As Single
    Static dstTop As Single, dstBottom As Single
        
    dstLeft = ((selLeft - dstForm.HScroll.Value) * ZoomVal) + imgLeft
    dstRight = ((selLeft + selWidth - dstForm.HScroll.Value) * ZoomVal) + imgLeft
    dstTop = ((selTop - dstForm.VScroll.Value) * ZoomVal) + imgTop
    dstBottom = ((selTop + selHeight - dstForm.VScroll.Value) * ZoomVal) + imgTop
    
    'Get a pointer to the target layer's pixel bits
    Dim iData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepLayerSafeArray pdImages(dstForm.Tag).backBuffer, tmpSA
    CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
    
    Static x As Long, y As Long
    Static QuickVal As Long
    
    'Use a look-up table to dim the non-active parts of the image
    Dim hData(0 To 255) As Byte
    
    For x = 0 To 255
        hData(x) = x \ 2
    Next x
    
    'Dim the non-active parts of the image
    Dim rightBound As Long, bottomBound As Long
    rightBound = imgLeft + imgWidth - 1
    bottomBound = imgTop + imgHeight - 1
    
    If rightBound > pdImages(dstForm.Tag).backBuffer.getLayerWidth - 1 Then rightBound = pdImages(dstForm.Tag).backBuffer.getLayerWidth - 1
    If bottomBound > pdImages(dstForm.Tag).backBuffer.getLayerHeight - 1 Then bottomBound = pdImages(dstForm.Tag).backBuffer.getLayerHeight - 1
    
    For x = imgLeft To rightBound
        QuickVal = x * 3
    For y = imgTop To bottomBound
    
        If (x < dstLeft) Or (x > dstRight) Or (y < dstTop) Or (y > dstBottom) Then
            iData(QuickVal + 2, y) = hData(iData(QuickVal + 2, y))
            iData(QuickVal + 1, y) = hData(iData(QuickVal + 1, y))
            iData(QuickVal, y) = hData(iData(QuickVal, y))
        End If
        
    Next y
    Next x
    
    'With our work complete, point iData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(iData), 0&, 4
    Erase iData
    
End Sub

Private Sub Class_Initialize()
    
    lockedIn = False
    
End Sub

'Sometimes a selection needs to directly access the bits of a target layer.  This routine sets up the corresponding SafeArray.
Private Sub prepLayerSafeArray(ByRef srcLayer As pdLayer, ByRef tmpSA As SAFEARRAY2D)
    With tmpSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lBound = 0
        .Bounds(0).cElements = srcLayer.getLayerHeight
        .Bounds(1).lBound = 0
        .Bounds(1).cElements = srcLayer.getLayerArrayWidth
        .pvData = srcLayer.getLayerDIBits
    End With
End Sub
