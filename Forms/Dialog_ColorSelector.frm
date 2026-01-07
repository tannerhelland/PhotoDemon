VERSION 5.00
Begin VB.Form dialog_ColorSelector 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Change color"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11535
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   403
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   769
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdHistory hstColors 
      Height          =   900
      Left            =   5070
      TabIndex        =   10
      Top             =   4290
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   1588
      FontSize        =   10
      HistoryRows     =   2
   End
   Begin PhotoDemon.pdNewOld noColor 
      Height          =   1095
      Left            =   240
      TabIndex        =   9
      Top             =   4080
      Width           =   4350
      _ExtentX        =   7673
      _ExtentY        =   1931
   End
   Begin PhotoDemon.pdSlider sldHSV 
      Height          =   375
      Index           =   0
      Left            =   6480
      TabIndex        =   6
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   661
      Max             =   359
      SliderKnobStyle =   1
      SliderTrackStyle=   5
      NotchPosition   =   2
   End
   Begin PhotoDemon.pdSlider sldRGB 
      Height          =   375
      Index           =   0
      Left            =   6480
      TabIndex        =   3
      Top             =   1920
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   661
      Max             =   255
      SliderKnobStyle =   1
      SliderTrackStyle=   2
      NotchPosition   =   2
   End
   Begin PhotoDemon.pdColorWheel clrWheel 
      Height          =   3855
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   6800
      WheelWidth      =   25
   End
   Begin PhotoDemon.pdCommandBarMini cmdBarMini 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5295
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdTextBox txtHex 
      Height          =   315
      Left            =   6600
      TabIndex        =   1
      Top             =   3735
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      Text            =   "abcdef"
   End
   Begin PhotoDemon.pdLabel lblColor 
      Height          =   360
      Index           =   8
      Left            =   5070
      Top             =   3765
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   1270
      Alignment       =   1
      Caption         =   "HTML / CSS:"
      ForeColor       =   0
      Layout          =   1
   End
   Begin PhotoDemon.pdLabel lblColor 
      Height          =   240
      Index           =   7
      Left            =   5130
      Top             =   3180
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   423
      Alignment       =   1
      Caption         =   "blue:"
      ForeColor       =   0
   End
   Begin PhotoDemon.pdLabel lblColor 
      Height          =   240
      Index           =   6
      Left            =   5115
      Top             =   2580
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   423
      Alignment       =   1
      Caption         =   "green:"
      ForeColor       =   0
   End
   Begin PhotoDemon.pdLabel lblColor 
      Height          =   240
      Index           =   5
      Left            =   5085
      Top             =   1980
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   423
      Alignment       =   1
      Caption         =   "red:"
      ForeColor       =   0
   End
   Begin PhotoDemon.pdLabel lblColor 
      Height          =   240
      Index           =   4
      Left            =   5040
      Top             =   1380
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   423
      Alignment       =   1
      Caption         =   "value:"
      ForeColor       =   0
   End
   Begin PhotoDemon.pdLabel lblColor 
      Height          =   240
      Index           =   3
      Left            =   5115
      Top             =   780
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   423
      Alignment       =   1
      Caption         =   "saturation:"
      ForeColor       =   0
   End
   Begin PhotoDemon.pdLabel lblColor 
      Height          =   240
      Index           =   2
      Left            =   5055
      Top             =   180
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   423
      Alignment       =   1
      Caption         =   "hue:"
      ForeColor       =   0
   End
   Begin PhotoDemon.pdSlider sldRGB 
      Height          =   375
      Index           =   1
      Left            =   6480
      TabIndex        =   4
      Top             =   2520
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   661
      Max             =   255
      SliderKnobStyle =   1
      SliderTrackStyle=   2
      NotchPosition   =   2
   End
   Begin PhotoDemon.pdSlider sldRGB 
      Height          =   375
      Index           =   2
      Left            =   6480
      TabIndex        =   5
      Top             =   3120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   661
      Max             =   255
      SliderKnobStyle =   1
      SliderTrackStyle=   2
      NotchPosition   =   2
   End
   Begin PhotoDemon.pdSlider sldHSV 
      Height          =   375
      Index           =   1
      Left            =   6480
      TabIndex        =   7
      Top             =   720
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   661
      Max             =   100
      SliderKnobStyle =   1
      SliderTrackStyle=   5
      NotchPosition   =   2
   End
   Begin PhotoDemon.pdSlider sldHSV 
      Height          =   375
      Index           =   2
      Left            =   6480
      TabIndex        =   8
      Top             =   1320
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   661
      Max             =   100
      SliderKnobStyle =   1
      SliderTrackStyle=   5
      NotchPosition   =   2
   End
End
Attribute VB_Name = "dialog_ColorSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Color Selection Dialog
'Copyright 2013-2026 by Tanner Helland
'Created: 11/November/13
'Last updated: 14/May/16
'Last update: improve real-time handling of hex input
'
'Basic color selection dialog.  At present, the dialog is heavily modeled after GIMP's color selection dialog.
'
'Special thank you to:
' - "DawnBringer" from the PixelJoint forums - http://pixeljoint.com/forum/forum_posts.asp?TID=16247 -
'    for the DB32 palette that provides a great default color selection on first load.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'The user input from the dialog (OK vs Cancel)
Private m_DialogResult As VbMsgBoxResult

'The original color when the dialog is first loaded; the user can restore this using the "original" box
Private m_OriginalColor As Long

'The new color selected by the user, if any.  This is cached so the caller can retrieve it at the same time
' as m_DialogResult, above.  It is only populated if the user clicks OK.
Private m_NewColor As Long

'To simplify color synchronization, the current color is parsed into RGB and HSV components, all of which
' are cached at module-level.  UI elements can grab these at any time to re-sync themselves.
Private m_CurrentColor As Long
Private m_Red As Long, m_Green As Long, m_Blue As Long
Private m_Hue As Double, m_Saturation As Double, m_Value As Double

'Changing the various text boxes resyncs the dialog, unless this parameter is set.  (We use it to prevent
' infinite resyncs.)
Private m_suspendTextResync As Boolean, m_suspendHexInput As Boolean
Private m_suspendParentNotifications As Boolean

'If a user control spawned this dialog, it will pass itself as a reference.  We can then send color updates back
' to the control, allowing for real-time updates on the screen despite a modal dialog being raised!
Private m_ParentColorSelector As pdColorSelector

'The color selector history is saved and loaded to file by this class.
' (Normally this is declared WithEvents, but this dialog doesn't require custom settings behavior.)
Private m_lastUsedSettings As pdLastUsedSettings
Attribute m_lastUsedSettings.VB_VarHelpID = -1

'The user's answer is returned via this property
Public Property Get DialogResult() As VbMsgBoxResult
    DialogResult = m_DialogResult
End Property

'The newly selected color (if any) is returned via this property
Public Property Get NewlySelectedColor() As Long
    NewlySelectedColor = m_NewColor
End Property

Private Sub clrWheel_ColorChanged(ByVal newColor As Long, ByVal srcIsInternal As Boolean)
    
    If srcIsInternal Then
    
        'Rebuild all module-level color variables to match the new color
        m_Red = Colors.ExtractRed(newColor)
        m_Green = Colors.ExtractGreen(newColor)
        m_Blue = Colors.ExtractBlue(newColor)
        
        'If this color has zero saturation (meaning it's a gray pixel), do not change the current hue
        Dim tmpHue As Double
        Colors.RGBtoHSV m_Red, m_Green, m_Blue, tmpHue, m_Saturation, m_Value
        If (m_Saturation <> 0#) Then m_Hue = tmpHue
        
        'Redraw any necessary interface elements
        SyncInterfaceToCurrentColor
        
    End If
    
End Sub

Private Sub cmdBarMini_CancelClick()
    
    'To prevent circular references, free our parent control reference immediately
    Set m_ParentColorSelector = Nothing
    
    m_DialogResult = vbCancel
    Me.Hide
    
End Sub

Private Sub cmdBarMini_OKClick()
    
    'Store the m_NewColor value (which the dialog handler will use to return the selected color)
    m_NewColor = RGB(m_Red, m_Green, m_Blue)
    
    'Push the newly selected color onto the color history stack
    hstColors.PushNewHistoryItem CStr(m_NewColor)
    
    'Save all last-used settings to file
    If (Not m_lastUsedSettings Is Nothing) Then
        m_lastUsedSettings.SaveAllControlValues
        m_lastUsedSettings.SetParentForm Nothing
    End If
    
    'To prevent circular references, free our parent control reference immediately
    Set m_ParentColorSelector = Nothing
    
    m_DialogResult = vbOK
    Me.Hide
    
End Sub

'The ShowDialog routine presents the user with this form.
Public Sub ShowDialog(ByVal initialColor As Long, Optional ByRef callingControl As pdColorSelector = Nothing, Optional ByRef callerParent As Form = Nothing)
        
    m_suspendParentNotifications = True
        
    'Store a reference to the calling control (if any)
    Set m_ParentColorSelector = callingControl
    
    'Load any last-used settings for this form
    Set m_lastUsedSettings = New pdLastUsedSettings
    m_lastUsedSettings.SetParentForm Me
    m_lastUsedSettings.LoadAllControlValues
    
    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    m_DialogResult = vbCancel
    
    'The passed color may be an OLE constant rather than an actual RGB triplet, so convert it now.
    initialColor = ConvertSystemColor(initialColor)
    
    'Cache the currentColor parameter so we can access it later
    m_OriginalColor = initialColor
    
    'Sync all current color values to the initial color
    m_suspendTextResync = True
    m_CurrentColor = initialColor
    m_Red = Colors.ExtractRed(initialColor)
    m_Green = Colors.ExtractGreen(initialColor)
    m_Blue = Colors.ExtractBlue(initialColor)
    sldRGB(0).NotchValueCustom = m_Red
    sldRGB(0).DefaultValue = m_Red
    sldRGB(1).NotchValueCustom = m_Green
    sldRGB(1).DefaultValue = m_Green
    sldRGB(2).NotchValueCustom = m_Blue
    sldRGB(2).DefaultValue = m_Blue
    
    Colors.RGBtoHSV m_Red, m_Green, m_Blue, m_Hue, m_Saturation, m_Value
    sldHSV(0).NotchValueCustom = m_Hue * 359
    sldHSV(0).DefaultValue = sldHSV(0).NotchValueCustom
    sldHSV(1).NotchValueCustom = m_Saturation * 100
    sldHSV(1).DefaultValue = sldHSV(1).NotchValueCustom
    sldHSV(2).NotchValueCustom = m_Value * 100
    sldHSV(2).DefaultValue = sldHSV(2).NotchValueCustom
    m_suspendTextResync = False
    
    'Synchronize the interface to this new color
    SyncInterfaceToCurrentColor
    
    'Make sure that the proper cursor is set
    Screen.MousePointer = 0
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me
    
    m_suspendParentNotifications = False
    
    'Display the dialog, and if the caller supplied a custom parent, center the dialog against *that* window
    If (callerParent Is Nothing) Then
        Interface.ShowPDDialog vbModal, Me, True
    Else
        Interface.ShowCustomPopup vbModal, Me, callerParent, True
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'When *all* current color values are updated and valid, use this function to synchronize the interface to match
' their appearance.
Private Sub SyncInterfaceToCurrentColor()
    
    'The integrated color wheel is easy.  Just make it match our current RGB values!
    m_suspendTextResync = True
    clrWheel.Color = RGB(m_Red, m_Green, m_Blue)
    m_suspendTextResync = False
    
    'Render the "new" and "old" color boxes on the left
    noColor.RequestRedraw True
    
    'Synchronize all text boxes to their current values
    RedrawAllTextBoxes
    
    'If we have a reference to a parent color selection user control, notify that control that the user's color
    ' has changed.
    If (Not m_ParentColorSelector Is Nothing) And (Not m_suspendParentNotifications) Then
        m_ParentColorSelector.NotifyOfLiveColorChange RGB(m_Red, m_Green, m_Blue)
    End If
    
End Sub

'Use this sub to resync all text boxes to the current RGB/HSV values
Private Sub RedrawAllTextBoxes()
    
    'We don't want the _Change events for the text boxes firing while we resync them, so we disable any resyncing in advance
    m_suspendTextResync = True
    
    'As of 7.0, new helper functions allow us to change slider values and gradient colors simultaneously.
    ' This improves performance by coalescing redraw events.
    sldRGB(0).SetGradientColorsAndValueAtOnce RGB(0, m_Green, m_Blue), RGB(255, m_Green, m_Blue), m_Red
    sldRGB(1).SetGradientColorsAndValueAtOnce RGB(m_Red, 0, m_Blue), RGB(m_Red, 255, m_Blue), m_Green
    sldRGB(2).SetGradientColorsAndValueAtOnce RGB(m_Red, m_Green, 0), RGB(m_Red, m_Green, 255), m_Blue
    
    'The HSV sliders have their own redraw code.  They do not support RGB gradients (as their gradients must
    ' be calculated in the HSV space).
    sldHSV(0).Value = m_Hue * 359
    sldHSV(1).Value = m_Saturation * 100
    sldHSV(2).Value = m_Value * 100
    
    sldHSV(0).RequestOwnerDrawChange
    sldHSV(1).RequestOwnerDrawChange
    sldHSV(2).RequestOwnerDrawChange
    
    'Update the hex representation box
    If (Not m_suspendHexInput) Then txtHex.Text = Colors.GetHexStringFromRGB(RGB(m_Red, m_Green, m_Blue))
    
    'Re-enable syncing
    m_suspendTextResync = False
    
End Sub

Private Sub hstColors_DrawHistoryItem(ByVal histIndex As Long, ByVal histValue As String, ByVal targetDC As Long, ByVal ptrToRectF As Long)
    
    If (LenB(histValue) <> 0) And PDMain.IsProgramRunning() And (targetDC <> 0) Then
    
        Dim tmpRectF As RectF
        If (ptrToRectF <> 0) Then CopyMemoryStrict VarPtr(tmpRectF), ptrToRectF, LenB(tmpRectF)
            
        'Note that this control *is* color-managed inside this dialog
        Dim cmResult As Long
        ColorManagement.ApplyDisplayColorManagement_SingleColor CLng(histValue), cmResult
    
        Dim cSurface As pd2DSurface: Dim cBrush As pd2DBrush
        Drawing2D.QuickCreateSurfaceFromDC cSurface, targetDC
        Drawing2D.QuickCreateSolidBrush cBrush, cmResult
        PD2D.FillRectangleF_FromRectF cSurface, cBrush, tmpRectF
        
        Set cSurface = Nothing: Set cBrush = Nothing
        
    End If
    
End Sub

Private Sub hstColors_HistoryDoesntExist(ByVal histIndex As Long, histValue As String)
    
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

Private Sub hstColors_HistoryItemClicked(ByVal histIndex As Long, ByVal histValue As String)
    
    If (LenB(histValue) <> 0) Then
    
        Dim clickedColor As Long
        clickedColor = CLng(histValue)
        
        'Update the current color values with the color of this box
        m_Red = Colors.ExtractRed(clickedColor)
        m_Green = Colors.ExtractGreen(clickedColor)
        m_Blue = Colors.ExtractBlue(clickedColor)
        
        'Calculate new HSV values to match
        RGBtoHSV m_Red, m_Green, m_Blue, m_Hue, m_Saturation, m_Value
        
        'Resync the interface to match the new value!
        SyncInterfaceToCurrentColor
        
    End If
    
End Sub

Private Sub noColor_DrawNewItem(ByVal targetDC As Long, ByVal ptrToRectF As Long)
    
    Dim tmpRectF As RectF
    If (ptrToRectF <> 0) Then CopyMemoryStrict VarPtr(tmpRectF), ptrToRectF, LenB(tmpRectF)
    
    If PDMain.IsProgramRunning() And (targetDC <> 0) Then
        
        'Note that this control *is* color-managed inside this dialog
        Dim cmResult As Long
        ColorManagement.ApplyDisplayColorManagement_SingleColor RGB(m_Red, m_Green, m_Blue), cmResult
        
        Dim cSurface As pd2DSurface: Dim cBrush As pd2DBrush
        Drawing2D.QuickCreateSurfaceFromDC cSurface, targetDC
        Drawing2D.QuickCreateSolidBrush cBrush, cmResult
        PD2D.FillRectangleF_FromRectF cSurface, cBrush, tmpRectF
        Set cSurface = Nothing: Set cBrush = Nothing
        
    End If
    
End Sub

Private Sub noColor_DrawOldItem(ByVal targetDC As Long, ByVal ptrToRectF As Long)

    Dim tmpRectF As RectF
    If (ptrToRectF <> 0) Then CopyMemoryStrict VarPtr(tmpRectF), ptrToRectF, LenB(tmpRectF)
    
    If PDMain.IsProgramRunning() And (targetDC <> 0) Then
        
        'Note that this control *is* color-managed inside this dialog
        Dim cmResult As Long
        ColorManagement.ApplyDisplayColorManagement_SingleColor m_OriginalColor, cmResult
        
        Dim cSurface As pd2DSurface: Dim cBrush As pd2DBrush
        Drawing2D.QuickCreateSurfaceFromDC cSurface, targetDC
        Drawing2D.QuickCreateSolidBrush cBrush, cmResult
        PD2D.FillRectangleF_FromRectF cSurface, cBrush, tmpRectF
        Set cSurface = Nothing: Set cBrush = Nothing
        
    End If
    
End Sub

Private Sub noColor_OldItemClicked()

    'Update the current color values with the color of this box
    m_Red = Colors.ExtractRed(m_OriginalColor)
    m_Green = Colors.ExtractGreen(m_OriginalColor)
    m_Blue = Colors.ExtractBlue(m_OriginalColor)
    
    'Calculate new HSV values to match
    RGBtoHSV m_Red, m_Green, m_Blue, m_Hue, m_Saturation, m_Value
    
    'Resync the interface to match the new value!
    SyncInterfaceToCurrentColor
    
End Sub

Private Sub sldHSV_Change(Index As Integer)

    If (Not m_suspendTextResync) Then
    
        'Update the current color values with the color of this box
        Select Case Index
            Case 0
                m_Hue = sldHSV(Index).Value / 359
            Case 1
                m_Saturation = sldHSV(Index).Value / 100
            Case 2
                m_Value = sldHSV(Index).Value / 100
        End Select
        
        'Calculate new rgb values to match
        Colors.HSVtoRGB m_Hue, m_Saturation, m_Value, m_Red, m_Green, m_Blue
        
        'Resync the interface to match the new value!
        SyncInterfaceToCurrentColor
        
    End If

End Sub

Private Sub sldHSV_RenderTrackImage(Index As Integer, dstDIB As pdDIB, ByVal leftBoundary As Single, ByVal rightBoundary As Single)

    'Because the HSV sliders are owner-drawn, we have to render them manually.  Note that the slider will hand us an
    ' already-prepared DIB; we just have to fill it with the gradient we want.
    
    'Before doing anything else, pre-calculate left edge and right edge colors.
    Dim leftColor As Long, rightColor As Long
    
    Dim r As Long, g As Long, b As Long
    Dim h As Double, s As Double, v As Double
    
    Select Case Index
    
        'Hue
        Case 0
            h = 0#
            s = m_Saturation
            v = m_Value
            Colors.HSVtoRGB h, s, v, r, g, b
            leftColor = RGB(r, g, b)
            h = 1#
            Colors.HSVtoRGB h, s, v, r, g, b
            rightColor = RGB(r, g, b)
        
        'Saturation
        Case 1
            h = m_Hue
            s = 0#
            v = m_Value
            Colors.HSVtoRGB h, s, v, r, g, b
            leftColor = RGB(r, g, b)
            s = 1#
            Colors.HSVtoRGB h, s, v, r, g, b
            rightColor = RGB(r, g, b)
        
        'Value
        Case 2
            h = m_Hue
            s = m_Saturation
            v = 0#
            Colors.HSVtoRGB h, s, v, r, g, b
            leftColor = RGB(r, g, b)
            v = 1#
            Colors.HSVtoRGB h, s, v, r, g, b
            rightColor = RGB(r, g, b)
        
    End Select
    
    Dim gradientValue As Double, gradientMax As Double
    gradientMax = (rightBoundary - leftBoundary)
    If (gradientMax <> 0#) Then gradientMax = 1# / gradientMax Else gradientMax = 1#
    
    Dim targetColor As Long, targetHeight As Long
    targetHeight = dstDIB.GetDIBHeight
    
    'Simple gradient-ish code implementation of drawing any individual color component
    If (dstDIB.GetDIBDC <> 0) Then
        
        Dim cSurface As pd2DSurface
        Set cSurface = New pd2DSurface
        cSurface.WrapSurfaceAroundPDDIB dstDIB
        cSurface.SetSurfaceAntialiasing P2_AA_None
        cSurface.SetSurfacePixelOffset P2_PO_Normal
        
        Dim cPen As pd2DPen
        Set cPen = New pd2DPen
        cPen.SetPenWidth 1!
    
        Dim x As Long
        For x = 0 To dstDIB.GetDIBWidth - 1
            
            If (x <= leftBoundary) Then
                targetColor = leftColor
            ElseIf (x >= rightBoundary) Then
                targetColor = rightColor
            Else
                gradientValue = (x - leftBoundary) * gradientMax
            
                If (Index = 0) Then
                    h = gradientValue
                ElseIf (Index = 1) Then
                    s = gradientValue
                ElseIf (Index = 2) Then
                    v = gradientValue
                End If
                
                Colors.HSVtoRGB h, s, v, r, g, b
                targetColor = RGB(r, g, b)
                
            End If
            
            'Draw the finished color onto the target DIB
            cPen.SetPenColor targetColor
            PD2D.DrawLineI cSurface, cPen, x, 0, x, targetHeight
            
        Next x
        
        Set cSurface = Nothing
    
    End If
    
End Sub

Private Sub sldRGB_Change(Index As Integer)
        
    If (Not m_suspendTextResync) Then
    
        'Update the current color values with the color of this box
        Select Case Index
            Case 0
                m_Red = sldRGB(Index).Value
            Case 1
                m_Green = sldRGB(Index).Value
            Case 2
                m_Blue = sldRGB(Index).Value
        End Select
        
        'Calculate new HSV values to match
        RGBtoHSV m_Red, m_Green, m_Blue, m_Hue, m_Saturation, m_Value
        
        'Resync the interface to match the new value!
        SyncInterfaceToCurrentColor
        
    End If

End Sub

'Full validation of hex input happens in its LostFocus event, but we also do a quick-and-dirty sync during change events
Private Sub txtHex_Change()
    
    If (m_suspendHexInput Or m_suspendTextResync) Then Exit Sub
    
    m_suspendHexInput = True
    
    Dim newText As String
    newText = txtHex.Text
    
    'If the hex input looks valid, update the colors to match; otherwise, ignore the text completely
    If DoesHexLookValid(newText) Then
        
        'Parse the string to calculate actual numeric values; we can use VB's Val() function for this!
        m_Red = Val("&H" & Left$(newText, 2))
        m_Green = Val("&H" & Mid$(newText, 3, 2))
        m_Blue = Val("&H" & Right$(newText, 2))
        
        'Calculate new HSV values to match
        RGBtoHSV m_Red, m_Green, m_Blue, m_Hue, m_Saturation, m_Value
        
        'Resync the interface to match the new value!
        SyncInterfaceToCurrentColor
        
    End If
    
    m_suspendHexInput = False

End Sub

Private Sub txtHex_LostFocusAPI()
    
    m_suspendHexInput = True
    
    Dim newText As String
    newText = txtHex.Text
    
    'If the hex input looks valid, update the colors to match; otherwise, ignore the text completely
    If DoesHexLookValid(newText) Then
        
        'Change the text box to match our properly formatted string
        txtHex.Text = newText
        
        'Parse the string to calculate actual numeric values; we can use VB's Val() function for this!
        m_Red = Val("&H" & Left$(newText, 2))
        m_Green = Val("&H" & Mid$(newText, 3, 2))
        m_Blue = Val("&H" & Right$(newText, 2))
        
        'Calculate new HSV values to match
        RGBtoHSV m_Red, m_Green, m_Blue, m_Hue, m_Saturation, m_Value
        
        'Resync the interface to match the new value!
        SyncInterfaceToCurrentColor
        
    Else
        txtHex.Text = Colors.GetHexStringFromRGB(RGB(m_Red, m_Green, m_Blue))
    End If
    
    m_suspendHexInput = False
    
End Sub

'This function *may modify the incoming string* so please review the comments thoroughly
Private Function DoesHexLookValid(ByRef hexStringToCheck As String) As Boolean

    'Before doing anything else, remove all invalid characters from the text box
    Dim validChars As String
    validChars = "0123456789abcdef"
    
    Dim curText As String
    curText = Trim$(hexStringToCheck)
    
    Dim newText As String, curChar As String
    
    Dim i As Long
    For i = 1 To Len(curText)
        curChar = Mid$(curText, i, 1)
        If InStr(1, validChars, curChar, vbTextCompare) > 0 Then newText = newText & curChar
    Next i
        
    newText = LCase$(newText)
    
    'Make sure the length is 1, 3, or 6.  Each case is handled specially; other lengths are not valid
    Select Case Len(newText)
    
        'One character is treated as a shade of gray; extend it to six characters.  (I don't know if this is actually
        ' valid CSS, but it doesn't hurt to support it... right?)
        Case 1
            newText = String$(6, newText)
            DoesHexLookValid = True
        
        'Three characters is standard shorthand hex; expand each character as a pair
        Case 3
            newText = Left$(newText, 1) & Left$(newText, 1) & Mid$(newText, 2, 1) & Mid$(newText, 2, 1) & Right$(newText, 1) & Right$(newText, 1)
            DoesHexLookValid = True
            
        'Six characters is already valid, so no need to screw with it further.
        Case 6
            DoesHexLookValid = True
        
        Case Else
            'We can't handle this character string, so reset it
            newText = Colors.GetHexStringFromRGB(RGB(m_Red, m_Green, m_Blue))
            DoesHexLookValid = False
    
    End Select
    
    If DoesHexLookValid Then hexStringToCheck = newText

End Function
