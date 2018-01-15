VERSION 5.00
Begin VB.Form layerpanel_Navigator 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   201
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   304
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.pdNavigator nvgMain 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1296
   End
End
Attribute VB_Name = "layerpanel_Navigator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Navigation/Overview Tool Panel
'Copyright 2015-2018 by Tanner Helland
'Created: 15/October/15
'Last updated: 15/October/15
'Last update: initial build
'
'As part of the 7.0 release, PD's right-side panel gained a lot of new functionality.  To simplify the code for
' the new panel, each chunk of related settings (e.g. layer, nav, color selector) was moved to its own subpanel.
'
'This form is the subpanel for the navigator/overview panel.
'
'The most interesting object on this form is the navigator user control, which synchronizes with the viewport to
' allow the user to quickly move around the image regardless of zoom.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'The value of all controls on this form are saved and loaded to file by this class
Private WithEvents lastUsedSettings As pdLastUsedSettings
Attribute lastUsedSettings.VB_VarHelpID = -1

Private Sub Form_Load()
    
    'Load any last-used settings for this form
    Set lastUsedSettings = New pdLastUsedSettings
    lastUsedSettings.SetParentForm Me
    lastUsedSettings.LoadAllControlValues
    
    'Update everything against the current theme.  This will also set tooltips for various controls.
    UpdateAgainstCurrentTheme
    
    'Reflow the interface to match its current size
    ReflowInterface
    
End Sub

'Whenever this panel is resized, we must reflow all objects to fit the available space.
Private Sub ReflowInterface()

    'For now, make the navigator UC the same size as the underlying form
    If (Me.ScaleWidth > 10) Then
        nvgMain.Move 0, 0, Me.ScaleWidth - FixDPI(10), Me.ScaleHeight
    End If
    
End Sub

'Updating against the current theme accomplishes a number of things:
' 1) All user-drawn controls are redrawn according to the current g_Themer settings.
' 2) All tooltips and captions are translated according to the current language.
' 3) ApplyThemeAndTranslations is called, which redraws the form itself according to any theme and/or system settings.
'
'This function is called at least once, at Form_Load, but can be called again if the active language or theme changes.
Public Sub UpdateAgainstCurrentTheme()
    
    'Start by redrawing the form according to current theme and translation settings.  (This function also takes care of
    ' any common controls that may still exist in the program.)
    ApplyThemeAndTranslations Me
    
    'Reflow the interface, to account for any language changes.
    ReflowInterface
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    'Save all last-used settings to file
    If Not (lastUsedSettings Is Nothing) Then
        lastUsedSettings.SaveAllControlValues
        lastUsedSettings.SetParentForm Nothing
    End If
    
End Sub

Private Sub Form_Resize()
    ReflowInterface
End Sub

'The navigator will periodically request new thumbnails.  Supply them whenever requested.
Private Sub nvgMain_RequestUpdatedThumbnail(ByRef thumbDIB As pdDIB, ByRef thumbX As Single, ByRef thumbY As Single)
    
    If (g_OpenImageCount > 0) Then
        
        'The thumbDIB passed to this function will always be sized to the largest size the navigator can physically support.
        ' Our job is to place a composited copy of the current image inside that DIB, automatically centered as necessary.
        Dim thumbImageWidth As Long, thumbImageHeight As Long
        
        'Start by determining proper dimensions for the resized thumbnail image.
        PDMath.ConvertAspectRatio pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height, thumbDIB.GetDIBWidth, thumbDIB.GetDIBHeight, thumbImageWidth, thumbImageHeight
        
        'From there, solve for the top-left corner of the centered image
        If (thumbImageWidth < thumbDIB.GetDIBWidth) Then
            thumbX = (thumbDIB.GetDIBWidth - thumbImageWidth) * 0.5
        Else
            thumbX = 0!
        End If
        
        If (thumbImageHeight < thumbDIB.GetDIBHeight) Then
            thumbY = (thumbDIB.GetDIBHeight - thumbImageHeight) * 0.5
        Else
            thumbY = 0!
        End If
        
        Dim dstRectF As RectF
        With dstRectF
            .Left = thumbX
            .Top = thumbY
            .Width = thumbImageWidth
            .Height = thumbImageHeight
        End With
        
        'Request a copy of the current image thumbnail, at the size and offset we've calculated
        pdImages(g_CurrentImage).RequestThumbnail thumbDIB, , False, VarPtr(dstRectF)
        
    Else
        thumbX = 0!
        thumbY = 0!
    End If
    
End Sub
