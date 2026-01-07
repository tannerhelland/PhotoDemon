VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdViewport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'Image Viewport Support Class
'Copyright 2001-2026 by Tanner Helland
'Created: 15/April/01
'Last updated: 13/September/15
'Last update: overhaul the class completely as part of prepping for paint tools
'
'While many viewport actions and operations are universal regardless of image,
' some elements are image-specific - for example, the positioning of the image on
' the canvas, which is contingent on the image's size, the current zoom, and the
' size of the viewport/canvas.
'
'Rather than let these tracking variables float around inside the parent pdImage
' object, they are stored inside this dedicated viewport class.  Eventually, zoom
' and other viewport-related factors should probably be moved here as well.
'
'This class was completely overhauled prior to the 7.0 release, as part of prepping
' for paint tools.  It now stores much more viewport-specific data, which greatly
' simplifies tasks like coordinate space conversions.  Almost everything in this
' class is created by the ViewportEngine module, so look there for details on how
' various structs are calculated and filled.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'At present, PD only offers one viewport to the user.  (Hypothetically, multiple viewports
' could be active in the future.)  This value stores the zoom *index* of this current image
' in the main viewport.  Why an index, you ask, instead of an absolute zoom value?  Because
' the absolute zoom value may change due to reasons inaccessible to this object.  (For
' example, if the zoom is set to "fit image", and the main PD window is resized, the
' absolute zoom value will have changed - but the zoom *index* won't have.)
Private m_ZoomIndex As Long

'As of version 7.0, the first stage of the viewport pipeline uses RectF structs to store the calculated
' coordinate regions.  See the ViewportEngine module for more details on how these structs are used.
Private m_ImageRect_CanvasCoords As RectF
Private m_ImageRect_TranslatedCanvasCoords As RectF
Private m_CanvasRect_ImageCoords As RectF
Private m_CanvasRect_ActualPixels As RectF

'The area of the canvas covered by the image is stored here.  Note that a separate boolean is also stored,
' on the off chance that the canvas and image do not currently intersect.
Private m_CanvasAndImageIntersect As Boolean
Private m_CanvasImageIntersectRect As RectF
Private m_SrcImageIntersectRect As RectF

'Current scroll bar values.  These are saved/restored when the active image changes.
Private m_HScrollValue As Long, m_VScrollValue As Long

Friend Function GetHScrollValue() As Long
    GetHScrollValue = m_HScrollValue
End Function

Friend Function GetVScrollValue() As Long
    GetVScrollValue = m_VScrollValue
End Function

Friend Function GetIntersectState() As Boolean
    GetIntersectState = m_CanvasAndImageIntersect
End Function

Friend Sub GetIntersectRectCanvas(ByRef dstRectF As RectF)
    dstRectF = m_CanvasImageIntersectRect
End Sub

Friend Sub GetIntersectRectImage(ByRef dstRectF As RectF)
    dstRectF = m_SrcImageIntersectRect
End Sub

Friend Sub GetImageRectCanvasCoords(ByRef dstRectF As RectF)
    dstRectF = m_ImageRect_CanvasCoords
End Sub

Friend Sub GetImageRectTranslated(ByRef dstRectF As RectF)
    dstRectF = m_ImageRect_TranslatedCanvasCoords
End Sub

Friend Sub GetCanvasRectImageCoords(ByRef dstRectF As RectF)
    dstRectF = m_CanvasRect_ImageCoords
End Sub

Friend Sub GetCanvasRectActualPixels(ByRef dstRectF As RectF)
    dstRectF = m_CanvasRect_ActualPixels
End Sub

Friend Function GetZoomIndex() As Long
    GetZoomIndex = m_ZoomIndex
End Function

Friend Sub SetHScrollValue(ByVal newValue As Long)
    m_HScrollValue = newValue
End Sub

Friend Sub SetVScrollValue(ByVal newValue As Long)
    m_VScrollValue = newValue
End Sub

Friend Sub SetIntersectState(ByRef newState As Boolean)
    m_CanvasAndImageIntersect = newState
End Sub

Friend Sub SetIntersectRectCanvas(ByRef srcRectF As RectF)
    m_CanvasImageIntersectRect = srcRectF
End Sub

Friend Sub SetIntersectRectImage(ByRef srcRectF As RectF)
    m_SrcImageIntersectRect = srcRectF
End Sub

Friend Sub SetImageRectCanvasCoords(ByRef srcRectF As RectF)
    m_ImageRect_CanvasCoords = srcRectF
End Sub

Friend Sub SetImageRectTranslated(ByRef srcRectF As RectF)
    m_ImageRect_TranslatedCanvasCoords = srcRectF
End Sub

Friend Sub SetCanvasRectImageCoords(ByRef srcRectF As RectF)
    m_CanvasRect_ImageCoords = srcRectF
End Sub

Friend Sub SetCanvasRectActualPixels(ByRef srcRectF As RectF)
    m_CanvasRect_ActualPixels = srcRectF
End Sub

Friend Sub SetZoomIndex(ByVal newIndex As Long)
    m_ZoomIndex = newIndex
End Sub
