VERSION 5.00
Begin VB.Form toolbar_ImageTabs 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Images"
   ClientHeight    =   1140
   ClientLeft      =   2250
   ClientTop       =   1770
   ClientWidth     =   13725
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
   NegotiateMenus  =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   76
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   915
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdImageStrip ImageStrip 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
   End
   Begin VB.Menu mnuImageTabsContext 
      Caption         =   "&Image"
      Visible         =   0   'False
      Begin VB.Menu mnuTabstripPopup 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnuTabstripPopup 
         Caption         =   "Save copy (&lossless)"
         Index           =   1
      End
      Begin VB.Menu mnuTabstripPopup 
         Caption         =   "Save &as..."
         Index           =   2
      End
      Begin VB.Menu mnuTabstripPopup 
         Caption         =   "Revert"
         Enabled         =   0   'False
         Index           =   3
      End
      Begin VB.Menu mnuTabstripPopup 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuTabstripPopup 
         Caption         =   "Open location in E&xplorer"
         Index           =   5
      End
      Begin VB.Menu mnuTabstripPopup 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuTabstripPopup 
         Caption         =   "&Close"
         Index           =   7
      End
      Begin VB.Menu mnuTabstripPopup 
         Caption         =   "Close all except this"
         Index           =   8
      End
   End
End
Attribute VB_Name = "toolbar_ImageTabs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Image Selection ("Tab") Toolbar
'Copyright 2013-2016 by Tanner Helland
'Created: 15/October/13
'Last updated: 19/February/15
'Last updated by: Raj
'Last update: Added a close icon on hover of each thumbnail, and a context menu
'
'In fall 2013, PhotoDemon left behind the MDI model in favor of fully dockable/floatable tool and image windows.
' This required quite a new features, including a way to switch between loaded images when image windows are docked -
' which is where this form comes in.
'
'The purpose of this form is to provide a tab-like interface for switching between open images.  Please note that
' much of this form's layout and alignment is handled by PhotoDemon's window manager, so you will need to look
' there for detailed information on things like the window's positioning and alignment.
'
'To my knowledge, as of January '14 the tabstrip should work properly under all orientations and screen DPIs.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'In Feb '15, Raj added a great context menu to the tabstrip.  To help simplify menu enable/disable behavior, this enum can be used to identify
' individual menu entries.
Private Enum POPUP_MENU_ENTRIES
    POP_SAVE = 0
    POP_SAVE_COPY = 1
    POP_SAVE_AS = 2
    POP_REVERT = 3
    POP_OPEN_IN_EXPLORER = 5
    POP_CLOSE = 7
    POP_CLOSE_OTHERS = 8
End Enum

#If False Then
    Private Const POP_SAVE = 0, POP_SAVE_COPY = 1, POP_SAVE_AS = 2, POP_REVERT = 3, POP_OPEN_IN_EXPLORER = 5, POP_CLOSE = 7, POP_CLOSE_OTHERS = 8
#End If

'External functions can force a full redraw by calling this sub
Public Sub forceRedraw()
    Form_Resize
End Sub

'When the user switches images, redraw the toolbar to match the change
Public Sub NotifyNewActiveImage(ByVal pdImageIndex As Long)
    ImageStrip.NotifyNewActiveImage pdImageIndex
End Sub

'When the user somehow changes an image, they need to notify the toolbar, so that a new thumbnail can be rendered
Public Sub NotifyUpdatedImage(ByVal pdImageIndex As Long)
    ImageStrip.NotifyUpdatedImage pdImageIndex
End Sub

'Whenever a new image is loaded, it needs to be registered with the toolbar
Public Sub RegisterNewImage(ByVal pdImageIndex As Long)
    ImageStrip.AddNewThumb pdImageIndex
End Sub

'Whenever an image is unloaded, it needs to be de-registered with the toolbar
Public Sub RemoveImage(ByVal pdImageIndex As Long, Optional ByVal refreshToolbar As Boolean = True)
    ImageStrip.RemoveThumb pdImageIndex, refreshToolbar
End Sub

Private Sub Form_Load()
    Me.UpdateAgainstCurrentTheme
End Sub

'Any time this window is resized, we need to recreate the thumbnail display
Private Sub Form_Resize()
    ImageStrip.SetPositionAndSize 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
    g_WindowManager.UnregisterForm Me
End Sub

'External functions can use this to re-theme this form at run-time (important when changing languages, for example)
Public Sub requestApplyThemeAndTranslations()
    ApplyThemeAndTranslations Me
End Sub

Private Sub ImageStrip_Click(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    If (Button And pdRightButton) <> 0 Then
    
        'Enable various pop-up menu entries.  Wherever possible, we simply want to mimic the official PD menu, which saves
        ' us having to supply our own heuristics for menu enablement.
        mnuTabstripPopup(POP_SAVE).Enabled = FormMain.MnuFile(8).Enabled
        mnuTabstripPopup(POP_SAVE_COPY).Enabled = FormMain.MnuFile(9).Enabled
        mnuTabstripPopup(POP_SAVE_AS).Enabled = FormMain.MnuFile(10).Enabled
        mnuTabstripPopup(POP_REVERT).Enabled = FormMain.MnuFile(11).Enabled
        mnuTabstripPopup(POP_CLOSE).Enabled = FormMain.MnuFile(5).Enabled
        
        'Two special commands only appear in this menu: Open in Explorer, and Close Other Images
        ' Use our own enablement heuristics for these.
        
        'Open in Explorer only works if the image is currently on-disk
        mnuTabstripPopup(POP_OPEN_IN_EXPLORER).Enabled = (Len(pdImages(g_CurrentImage).locationOnDisk) > 0)
        
        'Close Other Images only works if more than one image is open.  We can determine this using the Next/Previous Image items
        ' in the Window menu
        mnuTabstripPopup(POP_CLOSE).Enabled = FormMain.MnuWindow(5).Enabled
        
        'Raise the context menu
        Me.PopupMenu mnuImageTabsContext, x:=x, y:=y
        
    End If
    
End Sub

Private Sub ImageStrip_ItemClosed(ByVal itemIndex As Long)
    Image_Canvas_Handler.FullPDImageUnload itemIndex
End Sub

Private Sub ImageStrip_ItemSelected(ByVal itemIndex As Long)
    ActivatePDImage itemIndex, "user clicked image thumbnail"
End Sub

Private Sub mnuImageTabsContext_Click()

End Sub

'All popup menu clicks are handled here
Private Sub mnuTabstripPopup_Click(Index As Integer)

    Select Case Index
        
        'Save
        Case 0
            File_Menu.MenuSave g_CurrentImage
        
        'Save copy (lossless)
        Case 1
            File_Menu.MenuSaveLosslessCopy g_CurrentImage
        
        'Save as
        Case 2
            File_Menu.MenuSaveAs g_CurrentImage
        
        'Revert
        Case 3
            
            pdImages(g_CurrentImage).undoManager.revertToLastSavedState
                        
            'Also, redraw the current child form icon
            CreateCustomFormIcons pdImages(g_CurrentImage)
            ImageStrip.NotifyUpdatedImage g_CurrentImage
        
        '(separator)
        Case 4
        
        'Open location in Explorer
        Case 5
            Dim filePath As String, shellCommand As String
            filePath = pdImages(g_CurrentImage).locationOnDisk
            shellCommand = "explorer.exe /select,""" & filePath & """"
            Shell shellCommand, vbNormalFocus
        
        '(separator)
        Case 6
        
        'Close
        Case 7
            Image_Canvas_Handler.FullPDImageUnload g_CurrentImage
        
        'Close all but this
        Case 8
            
            Dim curImageID As Long
            curImageID = pdImages(g_CurrentImage).imageID
            
            Dim i As Long
            For i = 0 To UBound(pdImages)
                If (Not pdImages(i) Is Nothing) Then
                    If pdImages(i).imageID <> curImageID Then FullPDImageUnload i
                End If
            Next i
    
    End Select

End Sub

'Updating against the current theme accomplishes a number of things:
' 1) All user-drawn controls are redrawn according to the current g_Themer settings.
' 2) All tooltips and captions are translated according to the current language.
' 3) ApplyThemeAndTranslations is called, which redraws the form itself according to any theme and/or system settings.
'
'This function is called at least once, at Form_Load, but can be called again if the active language or theme changes.
Public Sub UpdateAgainstCurrentTheme()
    ImageStrip.UpdateAgainstCurrentTheme
    ApplyThemeAndTranslations Me
End Sub
