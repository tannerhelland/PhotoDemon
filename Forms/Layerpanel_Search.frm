VERSION 5.00
Begin VB.Form layerpanel_Search 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
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
   Icon            =   "Layerpanel_Search.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   33
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   304
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.pdSearchBar srchMain 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   450
   End
End
Attribute VB_Name = "layerpanel_Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Search Tool Panel
'Copyright 2019-2021 by Tanner Helland
'Created: 25/April/19
'Last updated: 25/April/19
'Last update: initial build
'
'PhotoDemon has a lot of tools and menus.  It can be hard to find things.  This search bar is meant to help.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'The value of all controls on this form are saved and loaded to file by this class
Private WithEvents lastUsedSettings As pdLastUsedSettings
Attribute lastUsedSettings.VB_VarHelpID = -1

'Some search terms are managed by the menu manager.  Others are managed by the tool manager.
' We want to condense these into a single list of search terms, and when a search result is returned,
' we'll match it up against its relevant source stack.
Private m_AllSearchTerms As pdStringStack
Private m_MenuSearchTerms As pdStringStack
Private m_ToolSearchTerms As pdStringStack
Private m_ToolActions As pdStringStack

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

    'For now, make the search box roughly the same size as the underlying form
    If (Me.ScaleWidth > 10) Then
        srchMain.SetPositionAndSize Interface.FixDPI(2), Interface.FixDPI(2), Me.ScaleWidth - Interface.FixDPI(4), srchMain.GetHeight()
    End If
    
    'Refresh the panel immediately, so the user can see the result of the resize
    If (Not g_WindowManager Is Nothing) Then g_WindowManager.ForceWindowRepaint Me.hWnd
    
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

Public Sub SetFocusToSearchBox()
    srchMain.SetFocusToSearchBox
End Sub

Private Sub lastUsedSettings_ReadCustomPresetData()
    srchMain.Text = vbNullString
End Sub

Private Sub srchMain_Click(bestSearchHit As String)
    
    srchMain.Text = bestSearchHit
    
    If (m_MenuSearchTerms.ContainsString(bestSearchHit) >= 0) Then
        Menus.ProcessDefaultAction_BySearch bestSearchHit
    
    'If the search result is *not* a menu, the menu module can't auto-map it to a corresponding
    ' action string.  Return the action string instead.
    Else
        Menus.ProcessDefaultAction_ByName m_ToolActions.GetString(m_ToolSearchTerms.ContainsString(bestSearchHit, True))
    End If
    
    'Before exiting, update the search list as available items may have changed.
    ' (For example, if the user types "undo" and hits Enter, "redo" may now be available to them.)
    UpdateSearchTerms
    
End Sub

Private Sub srchMain_GotFocusAPI()
    UpdateSearchTerms
End Sub

Private Sub srchMain_RequestSearchList()
    UpdateSearchTerms
End Sub

Private Sub UpdateSearchTerms()
    
    Set m_AllSearchTerms = New pdStringStack
    Menus.GetSearchableMenuList m_MenuSearchTerms
    m_AllSearchTerms.CloneStack m_MenuSearchTerms
    
    toolbar_Toolbox.GetListOfToolNamesAndActions m_ToolSearchTerms, m_ToolActions
    m_AllSearchTerms.AppendStack m_ToolSearchTerms
    srchMain.SetSearchList m_AllSearchTerms
    
End Sub
