VERSION 5.00
Begin VB.Form toolbar_Options 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Tools"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   13515
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   98
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   901
   ShowInTaskbar   =   0   'False
   Begin VB.Line lnSeparatorTop 
      BorderColor     =   &H80000002&
      X1              =   0
      X2              =   5000
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "toolbar_Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Tools Toolbox
'Copyright 2013-2015 by Tanner Helland
'Created: 03/October/13
'Last updated: 16/October/14
'Last update: rework all selection interface code to use the new property dictionary functions
'
'This form was initially integrated into the main MDI form.  In fall 2013, PhotoDemon left behind the MDI model,
' and all toolbars were moved to their own forms.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private Sub Form_Load()

    Dim i As Long
        
    'Update everything against the current theme.  This will also set tooltips for various controls.
    UpdateAgainstCurrentTheme
    
End Sub

'Toolbars can never be unloaded, EXCEPT when the whole program is going down.  Check for the program-wide closing flag prior
' to exiting; if it is not found, cancel the unload and simply hide this form.  (Note that the toggleToolbarVisibility sub
' will also keep this toolbar's Window menu entry in sync with the form's current visibility.)
Private Sub Form_Unload(Cancel As Integer)
    
    If g_ProgramShuttingDown Then
        ReleaseFormTheming Me
        g_WindowManager.UnregisterForm Me
    Else
        Cancel = True
        ToggleToolbarVisibility TOOLS_TOOLBOX
    End If
    
End Sub

'Updating against the current theme accomplishes a number of things:
' 1) All user-drawn controls are redrawn according to the current g_Themer settings.
' 2) All tooltips and captions are translated according to the current language.
' 3) MakeFormPretty is called, which redraws the form itself according to any theme and/or system settings.
'
'This function is called at least once, at Form_Load, but can be called again if the active language or theme changes.
Public Sub UpdateAgainstCurrentTheme()
    
    'Start by redrawing the form according to current theme and translation settings.  (This function also takes care of
    ' any common controls that may still exist in the program.)
    MakeFormPretty Me
    
    'The top separator line is colored according to the current shadow accent color
    If Not g_Themer Is Nothing Then
        lnSeparatorTop.BorderColor = g_Themer.getThemeColor(PDTC_GRAY_SHADOW)
    Else
        lnSeparatorTop.BorderColor = vbHighlight
    End If
    
End Sub
