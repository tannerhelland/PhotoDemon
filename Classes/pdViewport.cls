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
'Copyright ©2001-2014 by Tanner Helland
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

'These values track the size and position of the image, as seen from the viewport.  These values are...
' - SET by the PrepareViewport function
' - READ by the ScrollViewport and RenderViewport functions
Public targetWidth As Double
Public targetHeight As Double
Public targetTop As Long
Public targetLeft As Long

'These additional offsets are added to all viewport calculations.  The offsets will be > 0 if:
' - the image window status bar is visible
' - rulers are active
Private verticalOffset As Long
Private horizontalOffset As Long

'The vertical and horizontal offsets can be further broken down into specific sides/directions
Private topOffset As Long
Private bottomOffset As Long

'Get/set the horizontal and vertical offset values
Public Function getVerticalOffset() As Long
    getVerticalOffset = topOffset + bottomOffset
End Function

'Get/set specific directional offsets
Public Function getTopOffset() As Long
    getTopOffset = topOffset
End Function

Public Function setTopOffset(ByVal newOffset As Long) As Long
    topOffset = newOffset
End Function

Public Function getBottomOffset() As Long
    getBottomOffset = bottomOffset
End Function

Public Function setBottomOffset(ByVal newOffset As Long) As Long
    bottomOffset = newOffset
End Function
