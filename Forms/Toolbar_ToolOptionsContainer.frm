VERSION 5.00
Begin VB.Form toolbar_Options 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   " Tools"
   ClientHeight    =   1470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13515
   ControlBox      =   0   'False
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
   Icon            =   "Toolbar_ToolOptionsContainer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   98
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   901
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
End
Attribute VB_Name = "toolbar_Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Tools Toolbox
'Copyright 2013-2026 by Tanner Helland
'Created: 03/October/13
'Last updated: 31/October/16
'Last update: minor code cleanup
'
'This form was initially integrated into the main MDI form.  In fall 2013, PhotoDemon left behind the MDI model,
' and all toolbars were moved to their own forms.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'When a tool panel is active, this value will mirror the tool panel's hWnd
Private m_PanelHWnd As Long

'To better support high-DPI OS settings, we use a system-based move/size handler
Private WithEvents m_WindowSize As pdWindowSize
Attribute m_WindowSize.VB_VarHelpID = -1

Private Sub Form_Load()
    
    If PDMain.IsProgramRunning() Then
        Set m_WindowSize = New pdWindowSize
        m_WindowSize.AttachToHWnd Me.hWnd, True
    End If
    
    'Update everything against the current theme.  This will also set tooltips for various controls.
    UpdateAgainstCurrentTheme
    
End Sub

'Toolbars can never be unloaded, EXCEPT when the whole program is going down.  Check for the program-wide closing flag prior
' to exiting; if it is not found, cancel the unload and simply hide this form.  (Note that the ToggleToolboxVisibility sub
' will also keep this toolbar's Window menu entry in sync with the form's current visibility.)
Private Sub Form_Unload(Cancel As Integer)
    If g_ProgramShuttingDown Then
        ReleaseFormTheming Me
        Set m_WindowSize = Nothing
    Else
        PDDebug.LogAction "WARNING!  toolbar_Options was unloaded prematurely - why??"
        Cancel = True
    End If
End Sub

Public Sub NotifyChildPanelHWnd(ByVal srcHWnd As Long)
    m_PanelHWnd = srcHWnd
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
    
End Sub

Private Sub m_WindowSize_WindowResize(ByVal newWidth As Long, ByVal newHeight As Long)
    If (m_PanelHWnd <> 0) Then g_WindowManager.SetSizeByHWnd m_PanelHWnd, newWidth, newHeight, False
End Sub
