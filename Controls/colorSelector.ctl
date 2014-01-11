VERSION 5.00
Begin VB.UserControl colorSelector 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   1710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   114
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ToolboxBitmap   =   "colorSelector.ctx":0000
End
Attribute VB_Name = "colorSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Color Selector custom control
'Copyright ©2013-2014 by Tanner Helland
'Created: 17/August/13
'Last updated: 04/September/13
'Last update: external displayColorSelection function, which forces the control to display the color selection
'             window (without requiring a user click)
'
'This thin user control is basically an empty control that when clicked, displays a color selection window.  If a
' color is selected (e.g. Cancel is not pressed), it updates its back color to match, and raises a "ColorChanged"
' event.
'
'Though simple, this control solves a lot of problems.  It is especially helpful for improving interaction with the
' command bar user control, as it easily supports color reset/randomize/preset events.  It is also nice to be able
' to update a single master function for color selection, then have the change propagate to all tool windows.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************



Option Explicit

'This control doesn't really do anything interesting, besides allow a color to be selected.
Public Event ColorChanged()

'A specialized mouse class is used to handle the hand cursor for this control
Private mouseHandler As bluMouseEvents

'The control's current color
Private curColor As OLE_COLOR

'At present, all this control does is store a color value
Public Property Get Color() As OLE_COLOR
    Color = curColor
End Property

Public Property Let Color(ByVal newColor As OLE_COLOR)
    curColor = newColor
    UserControl.BackColor = curColor
    drawControlBorders
    PropertyChanged "Color"
    RaiseEvent ColorChanged
End Property

'Outside functions can call this to force a display of the color window
Public Sub displayColorSelection()
    UserControl_Click
End Sub

Private Sub UserControl_Click()

    'Use the default color dialog to select a new color
    Dim newColor As Long
    If showColorDialog(newColor, CLng(curColor)) Then
        Color = newColor
    End If
    
End Sub

Private Sub UserControl_Initialize()

    drawControlBorders
    
    If g_UserModeFix Then
        Set mouseHandler = New bluMouseEvents
        mouseHandler.Attach UserControl.hWnd ', UserControl.Parent.hWnd
        mouseHandler.MousePointer = IDC_HAND
    End If
    
End Sub

Private Sub UserControl_InitProperties()
    curColor = RGB(255, 255, 255)
    Color = curColor
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    curColor = PropBag.ReadProperty("curColor", RGB(255, 255, 255))
    Color = curColor
End Sub

Private Sub UserControl_Resize()
    drawControlBorders
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "curColor", curColor, RGB(255, 255, 255)
End Sub

'For flexibility, we draw our own borders.  I may decide to change this behavior in the future...
Private Sub drawControlBorders()
        
    'For color management to work, we must pre-render the control onto a DIB, then copy the DIB to the screen.
    ' Using VB's internal draw commands leads to unpredictable results.
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    
    tmpDIB.createBlank UserControl.ScaleWidth, UserControl.ScaleHeight, 24, UserControl.BackColor
    
    'Use the API to draw borders around the control
    GDIPlusDrawLineToDC tmpDIB.getDIBDC, 0, 0, UserControl.ScaleWidth - 1, 0, vbBlack
    GDIPlusDrawLineToDC tmpDIB.getDIBDC, UserControl.ScaleWidth - 1, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, vbBlack
    GDIPlusDrawLineToDC tmpDIB.getDIBDC, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, 0, UserControl.ScaleHeight - 1, vbBlack
    GDIPlusDrawLineToDC tmpDIB.getDIBDC, 0, UserControl.ScaleHeight - 1, 0, 0, vbBlack
    
    'Render the backcolor to the control; doing it this way ensures color management works.  (Note that we use a
    ' g_UserModeFix check to prevent color management from firing at compile-time.)
    If g_UserModeFix Then turnOnDefaultColorManagement UserControl.hDC, UserControl.hWnd
    BitBlt UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, tmpDIB.getDIBDC, 0, 0, vbSrcCopy
    UserControl.Picture = UserControl.Image
    UserControl.Refresh
    
End Sub
