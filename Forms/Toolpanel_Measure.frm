VERSION 5.00
Begin VB.Form toolpanel_Measure 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   3855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14235
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   Icon            =   "Toolpanel_Measure.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   257
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   949
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.pdContainer cntrPopOut 
      Height          =   2055
      Index           =   0
      Left            =   0
      Top             =   1080
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   3625
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   375
         Index           =   0
         Left            =   0
         Top             =   90
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   661
         Caption         =   "modify measurement"
      End
      Begin PhotoDemon.pdCheckBox chkShare 
         Height          =   345
         Left            =   0
         TabIndex        =   3
         Top             =   1560
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   609
         Caption         =   "share measurements between images"
         Value           =   0   'False
      End
      Begin PhotoDemon.pdButton cmdAction 
         Height          =   450
         Index           =   0
         Left            =   2025
         TabIndex        =   4
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   794
         Caption         =   "swap points"
      End
      Begin PhotoDemon.pdButton cmdAction 
         Height          =   450
         Index           =   3
         Left            =   0
         TabIndex        =   5
         Top             =   960
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   794
         Caption         =   "rotate 90"
      End
      Begin PhotoDemon.pdButton cmdAction 
         Height          =   450
         Index           =   4
         Left            =   0
         TabIndex        =   6
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   794
         Caption         =   "clear points"
      End
      Begin PhotoDemon.pdButtonToolbox cmdFlyoutLock 
         Height          =   390
         Index           =   0
         Left            =   4080
         TabIndex        =   7
         Top             =   1530
         Width           =   390
         _ExtentX        =   1111
         _ExtentY        =   1111
         StickyToggle    =   -1  'True
      End
   End
   Begin PhotoDemon.pdLabel lblMeasure 
      Height          =   255
      Index           =   0
      Left            =   4080
      Top             =   30
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   450
      Alignment       =   1
      Caption         =   "distance:"
   End
   Begin PhotoDemon.pdLabel lblMeasure 
      Height          =   255
      Index           =   1
      Left            =   6600
      Top             =   30
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   450
      Alignment       =   1
      Caption         =   "angle:"
   End
   Begin PhotoDemon.pdLabel lblMeasure 
      Height          =   255
      Index           =   2
      Left            =   9120
      Top             =   30
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   450
      Alignment       =   1
      Caption         =   "width:"
   End
   Begin PhotoDemon.pdLabel lblMeasure 
      Height          =   255
      Index           =   3
      Left            =   11640
      Top             =   30
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   450
      Alignment       =   1
      Caption         =   "height:"
   End
   Begin PhotoDemon.pdLabel lblValue 
      Height          =   255
      Index           =   0
      Left            =   5640
      Top             =   30
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   450
      Caption         =   "0"
      FontBold        =   -1  'True
   End
   Begin PhotoDemon.pdLabel lblValue 
      Height          =   255
      Index           =   1
      Left            =   8040
      Top             =   30
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   450
      Caption         =   "0"
      FontBold        =   -1  'True
   End
   Begin PhotoDemon.pdLabel lblValue 
      Height          =   255
      Index           =   2
      Left            =   10680
      Top             =   30
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   450
      Caption         =   "0"
      FontBold        =   -1  'True
   End
   Begin PhotoDemon.pdLabel lblValue 
      Height          =   255
      Index           =   3
      Left            =   13200
      Top             =   30
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   450
      Caption         =   "0"
      FontBold        =   -1  'True
   End
   Begin PhotoDemon.pdLabel lblMeasure 
      Height          =   255
      Index           =   4
      Left            =   4080
      Top             =   480
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   450
      Alignment       =   1
      Caption         =   "distance:"
   End
   Begin PhotoDemon.pdLabel lblMeasure 
      Height          =   255
      Index           =   5
      Left            =   6600
      Top             =   480
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   450
      Alignment       =   1
      Caption         =   "angle:"
   End
   Begin PhotoDemon.pdLabel lblMeasure 
      Height          =   255
      Index           =   6
      Left            =   9120
      Top             =   480
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   450
      Alignment       =   1
      Caption         =   "width:"
   End
   Begin PhotoDemon.pdLabel lblMeasure 
      Height          =   255
      Index           =   7
      Left            =   11640
      Top             =   480
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   450
      Alignment       =   1
      Caption         =   "height:"
   End
   Begin PhotoDemon.pdLabel lblValue 
      Height          =   255
      Index           =   4
      Left            =   5640
      Top             =   480
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   450
      Caption         =   "0"
      FontBold        =   -1  'True
   End
   Begin PhotoDemon.pdLabel lblValue 
      Height          =   255
      Index           =   5
      Left            =   8040
      Top             =   480
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   450
      Caption         =   "0"
      FontBold        =   -1  'True
   End
   Begin PhotoDemon.pdLabel lblValue 
      Height          =   255
      Index           =   6
      Left            =   10680
      Top             =   480
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   450
      Caption         =   "0"
      FontBold        =   -1  'True
   End
   Begin PhotoDemon.pdLabel lblValue 
      Height          =   255
      Index           =   7
      Left            =   13200
      Top             =   480
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   450
      Caption         =   "0"
      FontBold        =   -1  'True
   End
   Begin PhotoDemon.pdButton cmdAction 
      Height          =   450
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   375
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   794
      Caption         =   "straighten image"
   End
   Begin PhotoDemon.pdButton cmdAction 
      Height          =   450
      Index           =   2
      Left            =   2025
      TabIndex        =   1
      Top             =   375
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   794
      Caption         =   "straighten layer"
   End
   Begin PhotoDemon.pdTitle ttlPanel 
      Height          =   360
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   635
      Caption         =   "use measurement to"
      Value           =   0   'False
   End
End
Attribute VB_Name = "toolpanel_Measure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Measurement Tool Panel
'Copyright 2013-2026 by Tanner Helland
'Created: 11/July/18
'Last updated: 28/May/24
'Last update: support for percent as a measurement unit
'
'PD's measurement tool is pretty straightforward: measure the distance and angle between two points,
' and relay those values to the user.  Can't beat that for simplicity!
'
'As an added convenience to the user, we also provide options for auto-straightening the image to
' match the current measurement angle.  This is great for visually aligning horizontal or vertical
' elements in a photo.  (And yes - it works for both horizontal *and* vertical lines, and it solves
' for this automagically.)
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Localized text is cached once, at theming time
Private m_NullTextString As String, m_StringsInitialized As Boolean

'Flyout manager
Private WithEvents m_Flyout As pdFlyout
Attribute m_Flyout.VB_VarHelpID = -1

'The value of all controls on this form are saved and loaded to file by this class
' (Normally this is declared WithEvents, but this dialog doesn't require custom settings behavior.)
Private m_lastUsedSettings As pdLastUsedSettings
Attribute m_lastUsedSettings.VB_VarHelpID = -1

'Call to update the current measurement text.
Public Sub UpdateUIText()
    
    If (Not m_StringsInitialized) Then Exit Sub
    
    Dim i As Long
    Dim firstPoint As PointFloat, secondPoint As PointFloat
    
    Dim measurementUnitText As String
    measurementUnitText = Units.GetNameOfUnit(mu_Pixels, True)
    
    If Tools_Measure.GetFirstPoint(firstPoint) And Tools_Measure.GetSecondPoint(secondPoint) Then
        
        'Save the current point positions to the active image.  (This lets us preserve measurements
        ' across images.)
        PDImages.GetActiveImage.ImgStorage.AddEntry "measure-tool-x1", firstPoint.x
        PDImages.GetActiveImage.ImgStorage.AddEntry "measure-tool-y1", firstPoint.y
        PDImages.GetActiveImage.ImgStorage.AddEntry "measure-tool-x2", secondPoint.x
        PDImages.GetActiveImage.ImgStorage.AddEntry "measure-tool-y2", secondPoint.y
        
        'Allow point clearing, swapping, and rotation.  (These index values are weird because
        ' this toolbox used a different UI layout in old versions.)
        cmdAction(0).Enabled = True
        cmdAction(3).Enabled = True
        cmdAction(4).Enabled = True
        
        'Next, we want to update all text.  After doing this, we will reflow the UI accordingly
        ' (based on the size of the new captions).
        
        'Distance
        Dim measureValue As Double
        If Tools_Measure.GetDistanceInPx(measureValue) Then
            lblValue(0).Caption = Format$(measureValue, "0.0") & " " & measurementUnitText
        Else
            lblValue(0).Caption = m_NullTextString
        End If
        
        'Angle
        If Tools_Measure.GetAngleInDegrees(measureValue) Then
            measureValue = Abs(measureValue)
            cmdAction(1).Enabled = (measureValue > 0.001)
            cmdAction(2).Enabled = (measureValue > 0.001)
            If (measureValue > 90#) Then measureValue = (180# - measureValue)
            lblValue(1).Caption = Format$(measureValue, "0.0#") & " " & ChrW$(&HB0)
        Else
            cmdAction(1).Enabled = False
            cmdAction(2).Enabled = False
            lblValue(1).Caption = m_NullTextString
        End If
        
        'Width
        lblValue(2).Caption = Format$(Abs(firstPoint.x - secondPoint.x), "0") & " " & measurementUnitText
        
        'Height
        lblValue(3).Caption = Format$(Abs(firstPoint.y - secondPoint.y), "0") & " " & measurementUnitText
        
        'If the current statusbar/ruler unit is something *other* than pixels, display a second set of
        ' measurement values, in said unit.
        If (FormMain.MainCanvas(0).GetRulerUnit <> mu_Pixels) Then
            
            Dim newUnit As PD_MeasurementUnit
            newUnit = FormMain.MainCanvas(0).GetRulerUnit()
            
            'Ensure the display elements are visible
            If (Not lblValue(4).Visible) Then
                For i = 4 To 7
                    lblMeasure(i).Visible = True
                    lblValue(i).Visible = True
                Next i
            End If
            
            'Repeat the same steps that we used for pixels, but this time, perform an additional conversion
            ' into the target unit space
            If Tools_Measure.GetDistanceInPx(measureValue) Then
                lblValue(4).Caption = Units.GetValueFormattedForUnit_FromPixel(newUnit, measureValue, PDImages.GetActiveImage.GetDPI, PDImages.GetActiveImage.Width, True)
            Else
                lblValue(4).Caption = m_NullTextString
            End If
            
            'Angle
            If Tools_Measure.GetAngleInDegrees(measureValue) Then
                measureValue = Abs(measureValue)
                If (measureValue > 90#) Then measureValue = (180# - measureValue)
                lblValue(5).Caption = Format$(measureValue, "0.0#") & " " & ChrW$(&HB0)
            Else
                lblValue(5).Caption = m_NullTextString
            End If
            
            'Width
            lblValue(6).Caption = Units.GetValueFormattedForUnit_FromPixel(newUnit, Abs(firstPoint.x - secondPoint.x), PDImages.GetActiveImage.GetDPI, PDImages.GetActiveImage.Width, True)
            
            'Height
            lblValue(7).Caption = Units.GetValueFormattedForUnit_FromPixel(newUnit, Abs(firstPoint.y - secondPoint.y), PDImages.GetActiveImage.GetDPI, PDImages.GetActiveImage.Height, True)
        
        'If the current unit is "pixels", hide the extra info area
        Else
            
            If lblValue(4).Visible Then
                For i = 4 To 7
                    lblMeasure(i).Visible = False
                    lblValue(i).Visible = False
                Next i
            End If
        
        End If
        
    'If a measurement isn't available, blank all labels and disable certain buttons
    Else
        
        For i = 0 To 7
            lblValue(i).Caption = m_NullTextString
        Next i
        
        For i = 4 To 7
            lblMeasure(i).Visible = False
            lblValue(i).Visible = False
        Next i
        
        For i = cmdAction.lBound To cmdAction.UBound
            cmdAction(i).Enabled = False
        Next i
        
    End If
    
    'With all captions and visibility set, we can now iterate through all labels and calculate ideal positioning.
    Dim maxWidth As Long, testWidth As Long
    
    'Use a pdFont object for string measuring
    Dim cFont As pdFont, cFontBold As pdFont
    Set cFont = New pdFont: Set cFontBold = New pdFont
    cFont.SetFontSize 10!: cFontBold.SetFontSize 10!
    cFont.SetFontBold False
    cFontBold.SetFontBold True
    
    'Calculate padding constants and initial offset (for the first label)
    Dim xPadding As Long, xOffset As Long
    xPadding = Interface.FixDPI(8)
    xOffset = ttlPanel(0).GetLeft + ttlPanel(0).GetWidth + xPadding * 2
    
    'Y position also needs to be set differently depending on whether one or two rows of measurements are visible.
    Dim yOffset As Long, yOffset2 As Long
    If (Not g_WindowManager Is Nothing) Then
        If lblValue(4).Visible Then
            yOffset = Interface.FixDPI(6)
            yOffset2 = yOffset + Interface.FixDPI(8) + cFont.GetHeightOfString("|y")
        Else
            yOffset = (g_WindowManager.GetClientHeight(Me.hWnd) - cFont.GetHeightOfString("|y")) \ 2
        End If
    End If
    
    For i = 0 To 3
        
        'Find the largest of the two visible title strings
        maxWidth = cFont.GetWidthOfString(lblMeasure(i).Caption)
        If lblMeasure(i + 4).Visible Then testWidth = cFont.GetWidthOfString(lblMeasure(i + 4).Caption)
        If (testWidth > maxWidth) Then maxWidth = testWidth
        maxWidth = maxWidth + 2 'safety margin for antialiasing and hinting
        
        'Move the labels into position and apply the max width (e.g. width of the larger caption)
        lblMeasure(i).SetPositionAndSize xOffset, yOffset, maxWidth, lblMeasure(i).GetHeight
        If lblMeasure(i + 4).Visible Then lblMeasure(i + 4).SetPositionAndSize xOffset, yOffset2, maxWidth, lblMeasure(i + 4).GetHeight
        
        'Increment offset and repeat for the value label(s), noting that we need to use the bolded
        ' font measurer (because those labels use a bold font).
        xOffset = xOffset + lblMeasure(i).GetWidth + xPadding
        
        maxWidth = cFontBold.GetWidthOfString(lblValue(i).Caption)
        If lblValue(i + 4).Visible Then testWidth = cFontBold.GetWidthOfString(lblValue(i + 4).Caption)
        If (testWidth > maxWidth) Then maxWidth = testWidth
        
        lblValue(i).SetPositionAndSize xOffset, yOffset, maxWidth, lblValue(i).GetHeight
        If lblValue(i + 4).Visible Then lblValue(i + 4).SetPositionAndSize xOffset, yOffset2, maxWidth, lblValue(i + 4).GetHeight
        
        xOffset = xOffset + lblValue(i).GetWidth + xPadding * 2
        
    Next i
    
    
End Sub

'The measurement tool has two settings: it can either share measurements across images
' (great for unifying measurements), or it can allow each image to have its own measurement.
' What we do when changing images depends on this setting.
Public Sub NotifyActiveImageChanged()
    
    'Measurements are shared between images
    If chkShare.Value Then
    
        'Simply redraw the screen; the current measurement points will be preserved
        Tools_Measure.RequestRedraw
    
    'Each image gets its *own* measurements
    Else
    
        'Relay this image's measurements (if any) to the measurement handler
        If PDImages.GetActiveImage.ImgStorage.DoesKeyExist("measure-tool-x1") Then
        
            'Send the updated points over
            With PDImages.GetActiveImage.ImgStorage
                Tools_Measure.SetPointsManually .GetEntry_Double("measure-tool-x1"), .GetEntry_Double("measure-tool-y1"), .GetEntry_Double("measure-tool-x2"), .GetEntry_Double("measure-tool-y2")
            End With
            
            Tools_Measure.RequestRedraw
            
        'This image doesn't have any stored measurements; clear 'em out
        Else
            Tools_Measure.ResetPoints True
        End If
    
    End If
    
End Sub

Private Sub chkShare_GotFocusAPI()
    UpdateFlyout 0, True
End Sub

Private Sub cmdAction_Click(Index As Integer)

    Select Case Index
    
        'Swap points
        Case 0
            Tools_Measure.SwapPoints
        
        'Straighten image to angle
        Case 1
            Tools_Measure.StraightenImageToMatch
            
        'Straighten layer to angle
        Case 2
            Tools_Measure.StraightenLayerToMatch
            
        'Rotate 90
        Case 3
            Tools_Measure.Rotate2ndPoint90Degrees
            
        'Clear points
        Case 4
            Tools_Measure.ResetPoints
        
    End Select

End Sub

Private Sub cmdAction_GotFocusAPI(Index As Integer)
    UpdateFlyout 0, True
End Sub

Private Sub cmdFlyoutLock_Click(Index As Integer, ByVal Shift As ShiftConstants)
    If (Not m_Flyout Is Nothing) Then m_Flyout.UpdateLockStatus Me.cntrPopOut(Index).hWnd, cmdFlyoutLock(Index).Value, cmdFlyoutLock(Index)
End Sub

Private Sub cmdFlyoutLock_GotFocusAPI(Index As Integer)
    UpdateFlyout Index, True
End Sub

Private Sub Form_Load()

    Tools.SetToolBusyState True
    
    'Load any last-used settings for this form
    Set m_lastUsedSettings = New pdLastUsedSettings
    m_lastUsedSettings.SetParentForm Me
    m_lastUsedSettings.LoadAllControlValues
    
    Tools.SetToolBusyState False
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'Save all last-used settings to file
    If Not (m_lastUsedSettings Is Nothing) Then
        m_lastUsedSettings.SaveAllControlValues
        m_lastUsedSettings.SetParentForm Nothing
    End If
    
    'Failsafe only
    If (Not m_Flyout Is Nothing) Then m_Flyout.HideFlyout
    Set m_Flyout = Nothing
    
End Sub

'Updating against the current theme accomplishes a number of things:
' 1) All user-drawn controls are redrawn according to the current g_Themer settings.
' 2) All tooltips and captions are translated according to the current language.
' 3) ApplyThemeAndTranslations is called, which redraws the form itself according to any theme and/or system settings.
'
'This function is called at least once, at Form_Load, but can be called again if the active language or theme changes.
Public Sub UpdateAgainstCurrentTheme()
    
    m_NullTextString = g_Language.TranslateMessage("n/a")
    m_StringsInitialized = True
    
    'Flyout lock controls use the same behavior across all instances
    UserControls.ThemeFlyoutControls cmdFlyoutLock
    
    'Redraw the form according to current theme and translation settings.  (This function also takes care of
    ' any common controls that may still exist in the program.)
    ApplyThemeAndTranslations Me
    
    'As language settings may have changed, we now need to redraw the current UI text
    Me.UpdateUIText

End Sub

'Whenever an active flyout panel is closed, we need to reset the matching titlebar to "closed" state
Private Sub m_Flyout_FlyoutClosed(origTriggerObject As Control)
    If (Not origTriggerObject Is Nothing) Then origTriggerObject.Value = False
End Sub

'Update the actively displayed flyout (if any).  Note that the flyout manager will automatically
' hide any other open flyouts, as necessary.
Private Sub UpdateFlyout(ByVal flyoutIndex As Long, Optional ByVal newState As Boolean = True)
    
    'Ensure we have a flyout manager
    If (m_Flyout Is Nothing) Then Set m_Flyout = New pdFlyout
    
    'Exit if we're already in the process of synchronizing
    If m_Flyout.GetFlyoutSyncState() Then Exit Sub
    m_Flyout.SetFlyoutSyncState True
    
    'Ensure we have a flyout manager, then raise the corresponding panel
    If newState Then
        If (flyoutIndex <> m_Flyout.GetFlyoutTrackerID()) Then m_Flyout.ShowFlyout Me, ttlPanel(flyoutIndex), cntrPopOut(flyoutIndex), flyoutIndex
    Else
        If (flyoutIndex = m_Flyout.GetFlyoutTrackerID()) Then m_Flyout.HideFlyout
    End If
    
    'Update titlebar state(s) to match
    Dim i As Long
    For i = ttlPanel.lBound To ttlPanel.UBound
        If (i = m_Flyout.GetFlyoutTrackerID()) Then
            If (Not ttlPanel(i).Value) Then ttlPanel(i).Value = True
        Else
            If ttlPanel(i).Value Then ttlPanel(i).Value = False
        End If
    Next i
    
    'Clear the synchronization flag before exiting
    m_Flyout.SetFlyoutSyncState False
    
End Sub

Private Sub ttlPanel_Click(Index As Integer, ByVal newState As Boolean)
    UpdateFlyout Index, newState
End Sub
