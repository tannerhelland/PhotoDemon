VERSION 5.00
Begin VB.Form toolbar_Layers 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000002&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Layers"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3660
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   356
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   244
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "toolbar_Layers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

Private Sub Form_Load()

    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip

End Sub

'Toolbars can never be unloaded, EXCEPT when the whole program is going down.  Check for the program-wide closing flag prior
' to exiting; if it is not found, cancel the unload and simply hide this form.  (Note that the toggleToolbarVisibility sub
' will also keep this toolbar's Window menu entry in sync with the form's current visibility.)
Private Sub Form_Unload(Cancel As Integer)
    
    If g_ProgramShuttingDown Then
        ReleaseFormTheming Me
        g_WindowManager.unregisterForm Me
    Else
        Cancel = True
        toggleToolbarVisibility TOOLS_TOOLBOX
    End If
    
End Sub
