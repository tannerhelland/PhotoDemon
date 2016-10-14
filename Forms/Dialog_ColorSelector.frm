VERSION 5.00
Begin VB.Form dialog_ColorSelector 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Change color"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11535
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
   StartUpPosition =   3  'Windows Default
   Begin PhotoDemon.pdNewOld noColor 
      Height          =   1095
      Left            =   720
      TabIndex        =   17
      Top             =   4080
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   1931
   End
   Begin PhotoDemon.pdSlider sldHSV 
      Height          =   375
      Index           =   0
      Left            =   6480
      TabIndex        =   14
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
      TabIndex        =   11
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
      TabIndex        =   10
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
      Left            =   6480
      TabIndex        =   1
      Top             =   3735
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      Text            =   "abcdef"
   End
   Begin VB.PictureBox picRecColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   7
      Left            =   10680
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   2
      Top             =   4560
      Width           =   495
   End
   Begin VB.PictureBox picRecColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   6
      Left            =   10080
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   3
      Top             =   4560
      Width           =   495
   End
   Begin VB.PictureBox picRecColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   5
      Left            =   9480
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   4
      Top             =   4560
      Width           =   495
   End
   Begin VB.PictureBox picRecColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   4
      Left            =   8880
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   5
      Top             =   4560
      Width           =   495
   End
   Begin VB.PictureBox picRecColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   3
      Left            =   8280
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   6
      Top             =   4560
      Width           =   495
   End
   Begin VB.PictureBox picRecColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   2
      Left            =   7680
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   7
      Top             =   4560
      Width           =   495
   End
   Begin VB.PictureBox picRecColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   7080
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   8
      Top             =   4560
      Width           =   495
   End
   Begin VB.PictureBox picRecColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   6480
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   9
      Top             =   4560
      Width           =   495
   End
   Begin PhotoDemon.pdLabel lblColor 
      Height          =   600
      Index           =   9
      Left            =   5085
      Top             =   4680
      Width           =   1305
      _ExtentX        =   0
      _ExtentY        =   0
      Alignment       =   1
      Caption         =   "recent colors:"
      ForeColor       =   0
      Layout          =   1
   End
   Begin PhotoDemon.pdLabel lblColor 
      Height          =   720
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
      TabIndex        =   12
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
      TabIndex        =   13
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
      TabIndex        =   15
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
      TabIndex        =   16
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
'Copyright 2013-2016 by Tanner Helland
'Created: 11/November/13
'Last updated: 14/May/16
'Last update: improve real-time handling of hex input
'
'Basic color selection dialog.  At present, the dialog is heavily modeled after GIMP's color selection dialog.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
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

'A temporary DIB for drawing any other elements
Private m_tmpDIB As pdDIB

'Changing the various text boxes resyncs the dialog, unless this parameter is set.  (We use it to prevent
' infinite resyncs.)
Private m_suspendTextResync As Boolean, m_suspendHexInput As Boolean

Private Enum PD_COLOR_CHANGE
    ccRed = 0
    ccGreen = 1
    ccBlue = 2
    ccHue = 3
    ccSaturation = 4
    ccValue = 5
End Enum

#If False Then
    Private Const ccRed = 0, ccGreen = 1, ccBlue = 2, ccHue = 3, ccSaturation = 4, ccValue = 5
#End If

'Recently used colors are stored in XML format
Private m_XMLEngine As pdXML

'The file where recent color data is stored.  The filename is generated by the ShowDialog function, and the
' resulting file is saved in the /Data/Presets/ folder.
Private m_RecentColorsFilename As String

'The current list of recent colors.  Because we have to color-manage these, it's helpful to have them
' in a non-string format.
Private m_RecentColors() As Long

'If a user control spawned this dialog, it will pass itself as a reference.  We can then send color updates back
' to the control, allowing for real-time updates on the screen despite a modal dialog being raised!
Private m_ParentColorSelector As pdColorSelector

'pdInputMouse makes it easier to deal with a custom hand cursor for the many picture boxes on the form
Private WithEvents m_MouseEvents As pdInputMouse
Attribute m_MouseEvents.VB_VarHelpID = -1

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
        If (m_Saturation <> 0) Then m_Hue = tmpHue
        
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
    
    'Save the current list of recently used colors
    SaveRecentColorList
    
    'To prevent circular references, free our parent control reference immediately
    Set m_ParentColorSelector = Nothing
    
    m_DialogResult = vbOK
    Me.Hide
    
End Sub

'The ShowDialog routine presents the user with this form.
Public Sub ShowDialog(ByVal initialColor As Long, Optional ByRef callingControl As pdColorSelector = Nothing)

    'Store a reference to the calling control (if any)
    Set m_ParentColorSelector = callingControl

    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    m_DialogResult = vbCancel
    
    'The passed color may be an OLE constant rather than an actual RGB triplet, so convert it now.
    initialColor = ConvertSystemColor(initialColor)
    
    'Cache the currentColor parameter so we can access it later
    m_OriginalColor = initialColor
    
    'Sync all current color values to the initial color
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
    
    'Synchronize the interface to this new color
    SyncInterfaceToCurrentColor
    
    'Make sure that the proper cursor is set
    Screen.MousePointer = 0
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me
    
    'Manually assign a hand cursor to the various picture boxes.
    PrepSpecialMouseHandling True
    
    'Initialize an XML engine, which we will use to read/write recent color data to file
    Set m_XMLEngine = New pdXML
    
    'The XML file will be stored in the Preset path (/Data/Presets)
    m_RecentColorsFilename = g_UserPreferences.GetPresetPath & "Color_Selector.xml"
    
    'If an XML file exists, load its contents now
    LoadRecentColorList
    
    'Display the dialog
    ShowPDDialog vbModal, Me, True

End Sub

'Capture-from-screen mode requires special handling
Private Sub PrepSpecialMouseHandling(ByVal handleMode As Boolean)
    
    If g_IsProgramRunning And handleMode Then
    
        Set m_MouseEvents = New pdInputMouse
        
        With m_MouseEvents
        
            Dim i As Long
            For i = picRecColor.lBound To picRecColor.UBound
                .AddInputTracker picRecColor(i).hWnd, True, False, False, True, True
            Next i
            
            .SetSystemCursor IDC_HAND
            
        End With
        
    Else
        Set m_MouseEvents = Nothing
    End If
    
End Sub

'If the user has used the color selector before, their last-used colors will be stored to an XML file.  Use this function
' to load those colors.
Private Sub LoadRecentColorList()

    'Start by seeing if an XML file with previously saved color data exists
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    If cFile.FileExist(m_RecentColorsFilename) Then
        
        'Attempt to load and validate the current file; if we can't, create a new, blank XML object
        If Not m_XMLEngine.LoadXMLFile(m_RecentColorsFilename) Then
            Debug.Print "List of recent colors appears to be invalid.  A new recent color list has been created."
            ResetXMLData
        End If
        
    Else
        ResetXMLData
    End If
        
    'We are now ready to load the actual color data from file.
    
    'The XML engine will do most the heavy lifting for this task.  We pass it a String array, and it fills it with
    ' all values corresponding to the given tag name and attribute.  (We must do this dynamically, because we don't
    ' know how many recent colors are actually saved - it could be anywhere from 0 to picRecColor.Count.)
    Dim allm_RecentColors() As String
    Dim numColors As Long
    
    If m_XMLEngine.FindAllAttributeValues(allm_RecentColors, "colorEntry", "id") Then
        
        numColors = UBound(allm_RecentColors) + 1
        
        'Make sure the file does not contain more entries than are allowed (shouldn't theoretically be possible,
        ' but it doesn't hurt to check).
        If numColors > picRecColor.Count Then numColors = picRecColor.Count
        
    'No recent color entries were found.
    Else
        numColors = 0
    End If
    
    Dim i As Long
    
    'If one or more recent colors were found, load them now.
    If numColors > 0 Then
        
        ReDim m_RecentColors(0 To numColors - 1) As Long
        
        'Load the actual colors from the XML file
        Dim tmpColorString As String
        
        For i = 0 To numColors - 1
        
            'Retrieve the color, in string format
            tmpColorString = m_XMLEngine.GetUniqueTag_String("color", , , "colorEntry", "id", allm_RecentColors(i))
            
            'Translate the color into a long, and update the corresponding picture box
            If Len(tmpColorString) <> 0 Then m_RecentColors(i) = CLng(tmpColorString)
            
        Next i
    
    'No recent colors were found.  Populate the list with a few default values.
    Else
        
        ReDim m_RecentColors(0 To picRecColor.Count - 1)
        m_RecentColors(0) = RGB(0, 0, 255)
        m_RecentColors(1) = RGB(0, 255, 0)
        m_RecentColors(2) = RGB(255, 0, 0)
        m_RecentColors(3) = RGB(255, 0, 255)
        m_RecentColors(4) = RGB(0, 255, 255)
        m_RecentColors(5) = RGB(255, 255, 0)
        m_RecentColors(6) = 0
        m_RecentColors(7) = RGB(255, 255, 255)
    End If
    
    'For color management reasons, we must use DIBs to copy colors onto the recent color picture boxes
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    
    'Render the recent color list to their respective picture boxes
    For i = 0 To picRecColor.Count - 1
    
        If i <= UBound(m_RecentColors) Then
            tmpDIB.CreateBlank picRecColor(i).ScaleWidth, picRecColor(i).ScaleHeight, 24, m_RecentColors(i)
            tmpDIB.RenderToPictureBox picRecColor(i)
        End If
    
    Next i

End Sub

'Save the current list of last-used colors to an XML file, adding the color presently selected as the most-recent entry.
Private Sub SaveRecentColorList()
    
    'Reset whatever XML data we may have stored at present - we will be rewriting the full MRU file from scratch.
    ResetXMLData
    
    'We now need to update the colors array with the new color entry.  Start by seeing if this color is already in the
    ' array.  If it is, simply swap its order.
    Dim i As Long, j As Long
    
    Dim colorFound As Boolean
    colorFound = False
    
    For i = 0 To picRecColor.Count - 1
    
        'This color already exists in the list.  Move it to the top of the list, and shift everything else downward.
        If m_RecentColors(i) = m_NewColor Then
            
            colorFound = True
            
            For j = i To 1 Step -1
                m_RecentColors(j) = m_RecentColors(j - 1)
            Next j
            
            m_RecentColors(0) = m_NewColor
            Exit For
            
        End If
        
    Next i
    
    'If this color is not already in the list, add it now.
    If Not colorFound Then
        
        For i = picRecColor.Count - 1 To 1 Step -1
            m_RecentColors(i) = m_RecentColors(i - 1)
        Next i
        
        m_RecentColors(0) = m_NewColor
    
    End If
    
    'Add all color entries to the XML engine
    For i = 0 To UBound(m_RecentColors)
        m_XMLEngine.WriteTagWithAttribute "colorEntry", "id", Str(i), "", True
        m_XMLEngine.WriteTag "color", m_RecentColors(i)
        m_XMLEngine.CloseTag "colorEntry"
        m_XMLEngine.WriteBlankLine
    Next i
    
    'With the XML file now complete, write it out to file
    m_XMLEngine.WriteXMLToFile m_RecentColorsFilename
    
End Sub

'When creating a new recent coclors file, or overwriting a corrupt one, use this to initialize the new XML file.
Private Sub ResetXMLData()
    m_XMLEngine.PrepareNewXML "Recent colors"
    m_XMLEngine.WriteBlankLine
    m_XMLEngine.WriteComment "Everything past this point is recent color data.  Entries are sorted in reverse chronological order."
    m_XMLEngine.WriteBlankLine
End Sub

'Refresh the various color box cursors when the mouse enters
Private Sub m_MouseEvents_MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    m_MouseEvents.SetSystemCursor IDC_HAND, m_MouseEvents.GetLastHwnd
End Sub

Private Sub m_MouseEvents_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    m_MouseEvents.SetSystemCursor IDC_HAND, m_MouseEvents.GetLastHwnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'When *all* current color values are updated and valid, use this function to synchronize the interface to match
' their appearance.
Private Sub SyncInterfaceToCurrentColor()
    
    'The integrated color wheel is easy.  Just make it match our current RGB values!
    clrWheel.Color = RGB(m_Red, m_Green, m_Blue)
    
    'Render the "new" and "old" color boxes on the left
    noColor.RequestRedraw True
    
    'Synchronize all text boxes to their current values
    RedrawAllTextBoxes
    
    'If we have a reference to a parent color selection user control, notify that control that the user's color
    ' has changed.
    If Not (m_ParentColorSelector Is Nothing) Then
        m_ParentColorSelector.NotifyOfLiveColorChange RGB(m_Red, m_Green, m_Blue)
    End If
    
End Sub

'Use this sub to resync all text boxes to the current RGB/HSV values
Private Sub RedrawAllTextBoxes()

    'We don't want the _Change events for the text boxes firing while we resync them, so we disable any resyncing in advance
    m_suspendTextResync = True
    
    'Start by matching up the text values themselves
    sldRGB(0).Value = m_Red
    sldRGB(1).Value = m_Green
    sldRGB(2).Value = m_Blue
    
    sldHSV(0).Value = m_Hue * 359
    sldHSV(1).Value = m_Saturation * 100
    sldHSV(2).Value = m_Value * 100
    
    'Next, we need to prep all our color bar sliders
    
    'The RGB sliders support standard RGB gradients, so no special redraws are required
    sldRGB(0).GradientColorLeft = RGB(0, m_Green, m_Blue)
    sldRGB(0).GradientColorRight = RGB(255, m_Green, m_Blue)
    sldRGB(1).GradientColorLeft = RGB(m_Red, 0, m_Blue)
    sldRGB(1).GradientColorRight = RGB(m_Red, 255, m_Blue)
    sldRGB(2).GradientColorLeft = RGB(m_Red, m_Green, 0)
    sldRGB(2).GradientColorRight = RGB(m_Red, m_Green, 255)
    
    'The HSV sliders have their own redraw code.  They do not support RGB gradients (as their gradients must
    ' be calculated in the HSV space).
    sldHSV(0).RequestOwnerDrawChange
    sldHSV(1).RequestOwnerDrawChange
    sldHSV(2).RequestOwnerDrawChange
    
    'Update the hex representation box
    If (Not m_suspendHexInput) Then txtHex = Colors.GetHexStringFromRGB(RGB(m_Red, m_Green, m_Blue))
    
    'Re-enable syncing
    m_suspendTextResync = False
    
End Sub

Private Sub noColor_DrawNewItem(ByVal targetDC As Long, ByVal ptrToRectF As Long)
    
    Dim tmpRectF As RECTF
    CopyMemory ByVal VarPtr(tmpRectF), ByVal ptrToRectF, 16&
    
    If g_IsProgramRunning Then
        Dim cPainter As pd2DPainter, cSurface As pd2DSurface: Dim cBrush As pd2DBrush
        Drawing2D.QuickCreatePainter cPainter
        Drawing2D.QuickCreateSurfaceFromDC cSurface, targetDC
        Drawing2D.QuickCreateSolidBrush cBrush, RGB(m_Red, m_Green, m_Blue)
        cPainter.FillRectangleF_FromRectF cSurface, cBrush, tmpRectF
        Set cPainter = Nothing: Set cSurface = Nothing: Set cBrush = Nothing
    End If
    
End Sub

Private Sub noColor_DrawOldItem(ByVal targetDC As Long, ByVal ptrToRectF As Long)

    Dim tmpRectF As RECTF
    CopyMemory ByVal VarPtr(tmpRectF), ByVal ptrToRectF, 16&
    
    If g_IsProgramRunning Then
        Dim cPainter As pd2DPainter, cSurface As pd2DSurface: Dim cBrush As pd2DBrush
        Drawing2D.QuickCreatePainter cPainter
        Drawing2D.QuickCreateSurfaceFromDC cSurface, targetDC
        Drawing2D.QuickCreateSolidBrush cBrush, m_OriginalColor
        cPainter.FillRectangleF_FromRectF cSurface, cBrush, tmpRectF
        Set cPainter = Nothing: Set cSurface = Nothing: Set cBrush = Nothing
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

'When a recent color is clicked, update the screen with the new color
Private Sub picRecColor_Click(Index As Integer)

    'Update the current color values with the color of this box
    m_Red = Colors.ExtractRed(m_RecentColors(Index))
    m_Green = Colors.ExtractGreen(m_RecentColors(Index))
    m_Blue = Colors.ExtractBlue(m_RecentColors(Index))
    
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
    
    Dim targetColor As Long, targetHeight As Long, targetDC As Long
    targetDC = dstDIB.GetDIBDC
    targetHeight = dstDIB.GetDIBHeight
    
    'Simple gradient-ish code implementation of drawing any individual color component
    Dim x As Long
    For x = 0 To dstDIB.GetDIBWidth - 1
        
        If (x <= leftBoundary) Then
            targetColor = leftColor
        ElseIf (x >= rightBoundary) Then
            targetColor = rightColor
        Else
            gradientValue = (x - leftBoundary) / gradientMax
        
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
        GDI.DrawLineToDC targetDC, x, 0, x, targetHeight, targetColor
        
    Next x
    
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
        
    newText = LCase(newText)
    
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
