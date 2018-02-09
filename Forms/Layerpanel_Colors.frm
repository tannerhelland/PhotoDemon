VERSION 5.00
Begin VB.Form layerpanel_Colors 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2850
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
   ScaleWidth      =   190
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.pdHistory clrHistory 
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   2520
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   344
   End
   Begin PhotoDemon.pdColorVariants clrVariants 
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1720
   End
   Begin PhotoDemon.pdColorWheel clrWheel 
      Height          =   975
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1720
      WheelWidth      =   13
   End
End
Attribute VB_Name = "layerpanel_Colors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Color Selector Tool Panel
'Copyright 2015-2018 by Tanner Helland
'Created: 15/October/15
'Last updated: 20/October/15
'Last update: actually implement color selection controls!
'
'As part of the 7.0 release, PD's right-side panel gained a lot of new functionality.  To simplify the code for
' the new panel, each chunk of related settings (e.g. layer, nav, color selector) was moved to its own subpanel.
'
'This form is the subpanel for the color selector panel.  It is currently under construction.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'The value of all controls on this form are saved and loaded to file by this class
Private WithEvents lastUsedSettings As pdLastUsedSettings
Attribute lastUsedSettings.VB_VarHelpID = -1

'To avoid nested resize calls, trackers are used
Private m_ResizeInProgress As Boolean

'We do some custom rendering in this panel (for the color history dialog), so it's helpful to cache a
' pd2DPainter object.
Private m_Painter As pd2DPainter

'When various paint tools are used on the main window, they will notify us (via window message) of what
' color was used.  We will add those colors to our history list.
Private Sub clrHistory_CustomWindowMessage(ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturn As Long)
    If (wMsg = WM_PD_PRIMARY_COLOR_APPLIED) Then clrHistory.PushNewHistoryItem CStr(wParam), , True
End Sub

Private Sub clrHistory_DrawHistoryItem(ByVal histIndex As Long, ByVal histValue As String, ByVal targetDC As Long, ByVal ptrToRectF As Long)
    
    If (Len(histValue) <> 0) And MainModule.IsProgramRunning() And (targetDC <> 0) Then
        
        If MainModule.IsProgramRunning Then
        
            Dim tmpRectF As RectF
            If (ptrToRectF <> 0) Then
            
                CopyMemory ByVal VarPtr(tmpRectF), ByVal ptrToRectF, LenB(tmpRectF)
                
                'Note that this control *is* color-managed
                Dim cmResult As Long
                ColorManagement.ApplyDisplayColorManagement_SingleColor CLng(histValue), cmResult
            
                Dim cSurface As pd2DSurface: Dim cBrush As pd2DBrush
                Drawing2D.QuickCreateSurfaceFromDC cSurface, targetDC
                Drawing2D.QuickCreateSolidBrush cBrush, cmResult
                m_Painter.FillRectangleF_FromRectF cSurface, cBrush, tmpRectF
                
                Set cSurface = Nothing: Set cBrush = Nothing
            
            End If
            
        End If
        
    End If
    
End Sub

Private Sub clrHistory_HistoryDoesntExist(ByVal histIndex As Long, histValue As String)

    Dim newColor As Long
    
    Select Case histIndex
    
        Case 0
            newColor = RGB(0, 0, 0)
        Case 1
            newColor = RGB(34, 32, 52)
        Case 2
            newColor = RGB(69, 40, 60)
        Case 3
            newColor = RGB(102, 57, 49)
        Case 4
            newColor = RGB(143, 86, 59)
        Case 5
            newColor = RGB(223, 113, 38)
        Case 6
            newColor = RGB(217, 160, 102)
        Case 7
            newColor = RGB(238, 195, 154)
        Case 8
            newColor = RGB(251, 242, 54)
        Case 9
            newColor = RGB(153, 229, 80)
        Case 10
            newColor = RGB(106, 190, 48)
        Case 11
            newColor = RGB(55, 148, 110)
        Case 12
            newColor = RGB(75, 105, 47)
        Case 13
            newColor = RGB(82, 75, 36)
        Case 14
            newColor = RGB(50, 60, 57)
        Case 15
            newColor = RGB(63, 63, 116)
        Case 16
            newColor = RGB(48, 96, 130)
        Case 17
            newColor = RGB(91, 110, 225)
        Case 18
            newColor = RGB(99, 155, 255)
        Case 19
            newColor = RGB(95, 205, 228)
        Case 20
            newColor = RGB(203, 219, 252)
        Case 21
            newColor = RGB(255, 255, 255)
        Case 22
            newColor = RGB(155, 173, 183)
        Case 23
            newColor = RGB(132, 126, 135)
        Case 24
            newColor = RGB(105, 106, 106)
        Case 25
            newColor = RGB(89, 86, 82)
        Case 26
            newColor = RGB(118, 66, 138)
        Case 27
            newColor = RGB(172, 50, 50)
        Case 28
            newColor = RGB(217, 87, 99)
        Case 29
            newColor = RGB(215, 123, 186)
        Case 30
            newColor = RGB(143, 151, 74)
        Case 31
            newColor = RGB(138, 111, 48)
        Case Else
            newColor = RGB(255, 255, 255)
            
    End Select
    
    histValue = CStr(newColor)
    
End Sub

Private Sub clrHistory_HistoryItemClicked(ByVal histIndex As Long, ByVal histValue As String)
    
    If (LenB(histValue) <> 0) Then
    
        Dim clickedColor As Long
        clickedColor = CLng(histValue)
        
        'Update the other color selectors with this color value
        clrWheel.Color = clickedColor
        clrVariants.Color = clickedColor
        
    End If
    
End Sub

Private Sub clrVariants_ColorChanged(ByVal newColor As Long, ByVal srcIsInternal As Boolean)
    
    'If the clrVariant control is where the color was actually changed (and it's not just syncing itself to some
    ' external color change), relay the new color to the neighboring color wheel.
    If srcIsInternal Then clrWheel.Color = newColor
    
    'Whenever this primary color changes, we broadcast the change throughout PD, so other color selector controls
    ' know to redraw themselves accordingly.
    UserControls.PostPDMessage WM_PD_PRIMARY_COLOR_CHANGE, newColor
    
    'We also check to see if a paint-related tool is active.  If it is, assign the new color immediately.
    Select Case g_CurrentTool
    
        Case PAINT_BASICBRUSH, PAINT_SOFTBRUSH
            Paintbrush.SetBrushSourceColor newColor
            
        Case PAINT_FILL
            FillTool.SetFillBrushColor newColor
    
    End Select
    
End Sub

Private Sub clrWheel_ColorChanged(ByVal newColor As Long, ByVal srcIsInternal As Boolean)
    If srcIsInternal Then clrVariants.Color = newColor
End Sub

Private Sub Form_Load()
    
    m_ResizeInProgress = True
    
    'Prep some items related to the color history UI
    Set m_Painter = New pd2DPainter
    clrHistory.RequestCustomSubclassing WM_PD_PRIMARY_COLOR_APPLIED, True
    
    'Load any last-used settings for this form
    Set lastUsedSettings = New pdLastUsedSettings
    lastUsedSettings.SetParentForm Me
    lastUsedSettings.LoadAllControlValues
    
    'Update everything against the current theme.  This will also set tooltips for various controls,
    ' and reflow the interface to match.
    UpdateAgainstCurrentTheme
    
    m_ResizeInProgress = False
    
End Sub

'Whenever this panel is resized, we must reflow all objects to fit the available space.
Private Sub ReflowInterface()
    
    Dim curFormWidth As Long, curFormHeight As Long
    If (g_WindowManager Is Nothing) Then
        curFormWidth = Me.ScaleWidth
        curFormHeight = Me.ScaleHeight
    Else
        curFormWidth = g_WindowManager.GetClientWidth(Me.hWnd)
        curFormHeight = g_WindowManager.GetClientHeight(Me.hWnd)
    End If
    
    'Failsafe to prevent IDE errors
    If (curFormWidth > 10) And (curFormHeight > 10) Then
        
        'Bottom-align the color history panel
        clrHistory.SetPositionAndSize 0, curFormHeight - clrHistory.GetHeight, curFormWidth, clrHistory.GetHeight
        
        'Calculate a new height available to the other controls on this panel
        curFormHeight = curFormHeight - (clrHistory.GetHeight + Interface.FixDPI(2))
        
        'Before rendering other elements, enforce a minimum size.  During startup, form size vacillates
        ' several times as this window is "fit" against its neighbors.  This can throw GDI+ rendering
        ' error messages until a final size is arrived at.
        If (curFormHeight > 50) And (curFormWidth > 50) Then
            
            'Right-align the color wheel
            clrWheel.SetPositionAndSize curFormWidth - (curFormHeight + Interface.FixDPI(10)), 0, curFormHeight, curFormHeight
            
            'Fit the variant selector into the remaining area.
            clrVariants.SetPositionAndSize 0, 0, clrWheel.GetLeft - Interface.FixDPI(10), curFormHeight
            
        End If
        
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
    If (Not m_ResizeInProgress) Then
        m_ResizeInProgress = True
        ReflowInterface
        m_ResizeInProgress = False
    End If
End Sub

Public Function GetCurrentColor()
    GetCurrentColor = clrVariants.Color
End Function

Public Sub SetCurrentColor(ByVal newR As Long, ByVal newG As Long, ByVal newB As Long)
    clrVariants.Color = RGB(newR, newG, newB)
    clrWheel.Color = RGB(newR, newG, newB)
    clrHistory.PushNewHistoryItem RGB(newR, newG, newB), , True
End Sub
