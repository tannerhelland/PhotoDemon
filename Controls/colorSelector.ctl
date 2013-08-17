VERSION 5.00
Begin VB.UserControl colorSelector 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   1710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
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
    UserControl.backColor = curColor
    drawControlBorders
    PropertyChanged "Color"
    RaiseEvent ColorChanged
End Property

Private Sub UserControl_Click()

    'Use the default color dialog to select a new color
    Dim newColor As Long
    If showColorDialog(newColor, UserControl.Parent, CLng(curColor)) Then
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

Private Sub drawControlBorders()
    'For flexibility, we draw our own borders.  I may decide to change this behavior in the future...
    UserControl.Cls
    UserControl.Line (0, 0)-(UserControl.ScaleWidth - 1, 0)
    UserControl.Line (UserControl.ScaleWidth - 1, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1)
    UserControl.Line (UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1)-(0, UserControl.ScaleHeight - 1)
    UserControl.Line (0, UserControl.ScaleHeight - 1)-(0, 0)
    UserControl.Refresh
End Sub
