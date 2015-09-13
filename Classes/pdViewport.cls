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
'Copyright 2001-2015 by Tanner Helland
'Created: 15/April/01
'Last updated: 05/December/13
'Last update: started migrating certain viewport functions and tracking code into this standalone class.  The long-term
'              goal is to move all viewport functionality into this class, and to abstract out the rendering
'              destination code.  This would allow the program to create arbitrary viewport renderers for any
'              target object or DC, which would greatly increase flexibility over the current fixed system.
'
'While many viewport actions and operations are universal regardless of image, some elements are image-specific -
' for example, the positioning of the image on the canvas, which is contingent on the image's size, the current zoom,
' and the size of the viewport.
'
'Rather than keep these tracking variables floating around loose inside the parent pdImage object, they are now stored
' here.  Eventually, zoom and other viewport factors will be moved to this class as well.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'As of version 7.0, the first stage of the viewport pipeline uses RectF structs to store the calculated
' coordinate regions.  See the Viewport_Engine module for more details on how these structs are used.
Private ImageRect_CanvasCoords As RECTF
Private ImageRect_TranslatedCanvasCoords As RECTF
Private CanvasRect_ImageCoords As RECTF
Private CanvasRect_ActualPixels As RECTF

'The area of the canvas covered by the image is stored here.  Note that a separate boolean is also stored,
' on the off chance that the canvas and image do not currently intersect.
Private doCanvasAndImageIntersect As Boolean
Private CanvasImageIntersect As RECTF

Friend Sub setIntersectState(ByRef NewState As Boolean)
    doCanvasAndImageIntersect = NewState
End Sub

Friend Sub setIntersectRect(ByRef srcRectF As RECTF)
    CanvasImageIntersect = srcRectF
End Sub

Friend Sub setImageRectCanvasCoords(ByRef srcRectF As RECTF)
    ImageRect_CanvasCoords = srcRectF
End Sub

Friend Sub setImageRectTranslated(ByRef srcRectF As RECTF)
    ImageRect_TranslatedCanvasCoords = srcRectF
End Sub

Friend Sub setCanvasRectImageCoords(ByRef srcRectF As RECTF)
    CanvasRect_ImageCoords = srcRectF
End Sub

Friend Sub setCanvasRectActualPixels(ByRef srcRectF As RECTF)
    CanvasRect_ActualPixels = srcRectF
End Sub

Friend Function getIntersectState() As Boolean
    getIntersectState = doCanvasAndImageIntersect
End Function

Friend Sub getIntersectRect(ByRef dstRectF As RECTF)
    dstRectF = CanvasImageIntersect
End Sub

Friend Sub getImageRectCanvasCoords(ByRef dstRectF As RECTF)
    dstRectF = ImageRect_CanvasCoords
End Sub

Friend Sub getImageRectTranslated(ByRef dstRectF As RECTF)
    dstRectF = ImageRect_TranslatedCanvasCoords
End Sub

Friend Sub getCanvasRectImageCoords(ByRef dstRectF As RECTF)
    dstRectF = CanvasRect_ImageCoords
End Sub

Friend Sub getCanvasRectActualPixels(ByRef dstRectF As RECTF)
    dstRectF = CanvasRect_ActualPixels
End Sub

