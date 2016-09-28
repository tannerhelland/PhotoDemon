VERSION 5.00
Begin VB.Form dialog_ColorSelector 
   AutoRedraw      =   -1  'True
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   403
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   769
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin PhotoDemon.pdColorWheel clrWheel 
      Height          =   3855
      Left            =   720
      TabIndex        =   24
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   6800
      WheelWidth      =   30
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
      TabIndex        =   3
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
      Index           =   6
      Left            =   10080
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
      Index           =   5
      Left            =   9480
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
      Index           =   4
      Left            =   8880
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
      Index           =   3
      Left            =   8280
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
      Index           =   2
      Left            =   7680
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   21
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
      TabIndex        =   22
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
      TabIndex        =   23
      Top             =   4560
      Width           =   495
   End
   Begin VB.PictureBox picSampleHSV 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   6480
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   247
      TabIndex        =   19
      Top             =   1320
      Width           =   3735
   End
   Begin VB.PictureBox picSampleHSV 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   6480
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   247
      TabIndex        =   17
      Top             =   720
      Width           =   3735
   End
   Begin VB.PictureBox picSampleHSV 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   6480
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   247
      TabIndex        =   15
      Top             =   120
      Width           =   3735
   End
   Begin VB.PictureBox picSampleRGB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   6480
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   247
      TabIndex        =   14
      Top             =   3120
      Width           =   3735
   End
   Begin VB.PictureBox picSampleRGB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   6480
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   247
      TabIndex        =   12
      Top             =   2520
      Width           =   3735
   End
   Begin VB.PictureBox picSampleRGB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   6480
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   247
      TabIndex        =   10
      Top             =   1920
      Width           =   3735
   End
   Begin VB.PictureBox picOriginal 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1200
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   222
      TabIndex        =   2
      Top             =   4560
      Width           =   3360
   End
   Begin VB.PictureBox picCurrent 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1200
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   222
      TabIndex        =   1
      Top             =   4080
      Width           =   3360
   End
   Begin PhotoDemon.pdSpinner tudRGB 
      Height          =   345
      Index           =   0
      Left            =   10320
      TabIndex        =   9
      Top             =   1905
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      Max             =   255
   End
   Begin PhotoDemon.pdSpinner tudRGB 
      Height          =   345
      Index           =   1
      Left            =   10320
      TabIndex        =   11
      Top             =   2505
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      Max             =   255
   End
   Begin PhotoDemon.pdSpinner tudRGB 
      Height          =   345
      Index           =   2
      Left            =   10320
      TabIndex        =   13
      Top             =   3105
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      Max             =   255
   End
   Begin PhotoDemon.pdSpinner tudHSV 
      Height          =   345
      Index           =   0
      Left            =   10320
      TabIndex        =   16
      Top             =   105
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      Max             =   359
   End
   Begin PhotoDemon.pdSpinner tudHSV 
      Height          =   345
      Index           =   1
      Left            =   10320
      TabIndex        =   18
      Top             =   705
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      Max             =   100
   End
   Begin PhotoDemon.pdSpinner tudHSV 
      Height          =   345
      Index           =   2
      Left            =   10320
      TabIndex        =   20
      Top             =   1305
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      Max             =   100
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
   Begin PhotoDemon.pdLabel lblColor 
      Height          =   525
      Index           =   1
      Left            =   30
      Top             =   4650
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   926
      Alignment       =   1
      Caption         =   "original:"
      FontSize        =   11
      ForeColor       =   0
      Layout          =   1
   End
   Begin PhotoDemon.pdLabel lblColor 
      Height          =   405
      Index           =   0
      Left            =   75
      Top             =   4170
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   714
      Alignment       =   1
      Caption         =   "current:"
      FontSize        =   11
      ForeColor       =   0
      Layout          =   1
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

'Backing DIB for the primary color box (luminance/saturation) on the left
Private primaryBox As pdDIB

'Backing DIB for the hue box to the right of the primary color box
Private hueBox As pdDIB

'To simplify color synchronization, the current color is parsed into RGB and HSV components, all of which
' are cached at module-level.  UI elements can grab these at any time to re-sync themselves.
Private m_CurrentColor As Long
Private m_Red As Long, m_Green As Long, m_Blue As Long
Private m_Hue As Double, m_Saturation As Double, m_Value As Double

'Backing DIBs are required for each individual color sample boxes
Private sRed As pdDIB, sGreen As pdDIB, sBlue As pdDIB
Private sHue As pdDIB, sSaturation As pdDIB, sValue As pdDIB

'Left/right/up arrows for the hue and color boxes; these are 7x13 (or 13x7) and loaded from the resource at run-time
Private leftSideArrow As pdDIB, rightSideArrow As pdDIB, upArrow As pdDIB

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
    
    'Load the left/right side hue box arrow images from the resource file
    Set leftSideArrow = New pdDIB
    Set rightSideArrow = New pdDIB
    Set upArrow = New pdDIB
    
    LoadResourceToDIB "CLR_ARROW_L", leftSideArrow
    LoadResourceToDIB "CLR_ARROW_R", rightSideArrow
    LoadResourceToDIB "CLR_ARROW_U", upArrow
    
    'The passed color may be an OLE constant rather than an actual RGB triplet, so convert it now.
    initialColor = ConvertSystemColor(initialColor)
    
    'Cache the currentColor parameter so we can access it later
    m_OriginalColor = initialColor
    
    'Render the old color to the screen.  Note that we must use a temporary DIB for this; otherwise, the color will
    ' not be properly color-managed.
    Dim tmpDIB As New pdDIB
    tmpDIB.CreateBlank picOriginal.ScaleWidth, picOriginal.ScaleHeight, 24, m_OriginalColor
    tmpDIB.RenderToPictureBox picOriginal
    
    'Sync all current color values to the initial color
    m_CurrentColor = initialColor
    m_Red = Colors.ExtractRed(initialColor)
    m_Green = Colors.ExtractGreen(initialColor)
    m_Blue = Colors.ExtractBlue(initialColor)
    
    RGBtoHSV m_Red, m_Green, m_Blue, m_Hue, m_Saturation, m_Value
    
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
        
            .AddInputTracker picOriginal.hWnd, True, False, False, True, True
            
            Dim i As Long
            For i = picRecColor.lBound To picRecColor.UBound
                .AddInputTracker picRecColor(i).hWnd, True, False, False, True, True
            Next i
            
            For i = picSampleHSV.lBound To picSampleHSV.UBound
                .AddInputTracker picSampleHSV(i).hWnd, True, False, False, True, True
            Next i
            
            For i = picSampleRGB.lBound To picSampleRGB.UBound
                .AddInputTracker picSampleRGB(i).hWnd, True, False, False, True, True
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
    
    'Render the current color box.  Note that we must use a temporary DIB for this; otherwise, the color will
    ' not be properly color managed.
    If (m_tmpDIB Is Nothing) Then Set m_tmpDIB = New pdDIB
    If (m_tmpDIB.GetDIBWidth <> picCurrent.ScaleWidth) Or (m_tmpDIB.GetDIBHeight <> picCurrent.ScaleHeight) Then
        m_tmpDIB.CreateBlank picCurrent.ScaleWidth, picCurrent.ScaleHeight, 24, RGB(m_Red, m_Green, m_Blue)
    Else
        GDI_Plus.GDIPlusFillDIBRect m_tmpDIB, 0, 0, m_tmpDIB.GetDIBWidth, m_tmpDIB.GetDIBHeight, RGB(m_Red, m_Green, m_Blue)
    End If
    
    m_tmpDIB.RenderToPictureBox picCurrent
    
    'Synchronize all text boxes to their current values
    RedrawAllTextBoxes
    
'    'Position the arrows along the hue box properly according to the current hue
'    Dim hueY As Long
'    hueY = picHue.Top + 1 + (m_Hue * picHue.ScaleHeight)
'
'    leftSideArrow.AlphaBlendToDC Me.hDC, , picHue.Left - leftSideArrow.GetDIBWidth, hueY - (leftSideArrow.GetDIBHeight \ 2)
'    rightSideArrow.AlphaBlendToDC Me.hDC, , picHue.Left + picHue.Width, hueY - (rightSideArrow.GetDIBHeight \ 2)
    Me.Picture = Me.Image
    Me.Refresh
    
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
    tudRGB(0) = m_Red
    tudRGB(1) = m_Green
    tudRGB(2) = m_Blue
    
    tudHSV(0) = m_Hue * 359
    tudHSV(1) = m_Saturation * 100
    tudHSV(2) = m_Value * 100
    
    'Next, prepare some universal values for the arrow image offsets
    Dim arrowOffset As Long
    arrowOffset = (upArrow.GetDIBWidth \ 2) - 1
    
    Dim leftOffset As Long
    leftOffset = picSampleRGB(0).Left
    
    Dim widthCheck As Long
    widthCheck = picSampleRGB(0).ScaleWidth - 1
    
    'Next, redraw all marker arrows
    upArrow.AlphaBlendToDC Me.hDC, , leftOffset + ((m_Red / 255) * widthCheck) - arrowOffset, picSampleRGB(0).Top + picSampleRGB(0).Height
    upArrow.AlphaBlendToDC Me.hDC, , leftOffset + ((m_Green / 255) * widthCheck) - arrowOffset, picSampleRGB(1).Top + picSampleRGB(1).Height
    upArrow.AlphaBlendToDC Me.hDC, , leftOffset + ((m_Blue / 255) * widthCheck) - arrowOffset, picSampleRGB(2).Top + picSampleRGB(2).Height
    
    upArrow.AlphaBlendToDC Me.hDC, , leftOffset + (m_Hue * widthCheck) - arrowOffset, picSampleHSV(0).Top + picSampleHSV(0).Height
    upArrow.AlphaBlendToDC Me.hDC, , leftOffset + (m_Saturation * widthCheck) - arrowOffset, picSampleHSV(1).Top + picSampleHSV(1).Height
    upArrow.AlphaBlendToDC Me.hDC, , leftOffset + (m_Value * widthCheck) - arrowOffset, picSampleHSV(2).Top + picSampleHSV(2).Height
    
    'Next, we need to prep all our color bar DIBs
    RenderSampleDIB sRed, ccRed
    RenderSampleDIB sGreen, ccGreen
    RenderSampleDIB sBlue, ccBlue
    
    RenderSampleDIB sHue, ccHue
    RenderSampleDIB sSaturation, ccSaturation
    RenderSampleDIB sValue, ccValue
    
    'Now we can render the bars to screen
    sRed.RenderToPictureBox picSampleRGB(0)
    sGreen.RenderToPictureBox picSampleRGB(1)
    sBlue.RenderToPictureBox picSampleRGB(2)
    
    sHue.RenderToPictureBox picSampleHSV(0)
    sSaturation.RenderToPictureBox picSampleHSV(1)
    sValue.RenderToPictureBox picSampleHSV(2)
    
    'Update the hex representation box
    If (Not m_suspendHexInput) Then txtHex = Colors.GetHexStringFromRGB(RGB(m_Red, m_Green, m_Blue))
    
    'Re-enable syncing
    m_suspendTextResync = False
    
End Sub

'This sub handles the preparation of the individual color sample boxes (one each for R/G/B/H/S/V)
' (Because we want these boxes to be color-managed, we must create them as DIBs.)
Private Sub RenderSampleDIB(ByRef dstDIB As pdDIB, ByVal dibColorType As PD_COLOR_CHANGE)

    If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
    If (dstDIB.GetDIBWidth <> picSampleRGB(0).ScaleWidth) Or (dstDIB.GetDIBHeight <> picSampleRGB(0).ScaleHeight) Then
        dstDIB.CreateBlank picSampleRGB(0).ScaleWidth, picSampleRGB(0).ScaleHeight
    End If
    
    Dim r As Long, g As Long, b As Long
    Dim h As Double, s As Double, v As Double
    
    'Initialize each component to its default type; only one parameter will be changed per dibColorType
    r = m_Red
    g = m_Green
    b = m_Blue
    h = m_Hue
    s = m_Saturation
    v = m_Value
    
    Dim gradientValue As Double, gradientMax As Double
    gradientMax = dstDIB.GetDIBWidth
    
    'Simple gradient-ish code implementation of drawing any individual color component
    Dim x As Long
    For x = 0 To dstDIB.GetDIBWidth
    
        gradientValue = x / gradientMax
    
        'We handle RGB separately from HSV
        If dibColorType <= ccBlue Then
            
            Select Case dibColorType
            
                Case ccRed
                    r = gradientValue * 255
                    
                Case ccGreen
                    g = gradientValue * 255
                    
                Case Else
                    b = gradientValue * 255
                    
            End Select
            
        Else
        
            Select Case dibColorType
            
                Case ccHue
                    h = gradientValue
                
                Case ccSaturation
                    s = gradientValue
                
                Case ccValue
                    v = gradientValue
            
            End Select
            
            HSVtoRGB h, s, v, r, g, b
        
        End If
        
        'Draw the color
        GDI.DrawLineToDC dstDIB.GetDIBDC, x, 0, x, dstDIB.GetDIBHeight, RGB(r, g, b)
        
    Next x
    
End Sub

Private Sub picOriginal_Click()

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

Private Sub picSampleHSV_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    HSVBoxClicked Index, x
End Sub

Private Sub picSampleHSV_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then HSVBoxClicked Index, x
End Sub

'Whenever one of the HSV sample boxes is clicked, this function is called; it calculates new RGB/HSV values, then redraws the interface
Private Sub HSVBoxClicked(ByVal boxIndex As Long, ByVal xPos As Long)

    Dim boxWidth As Long
    boxWidth = picSampleRGB(0).ScaleWidth
    
    'Restrict mouse clicks to the picture box area
    If xPos < 0 Then xPos = 0
    If xPos > boxWidth Then xPos = boxWidth

    Select Case (boxIndex + 3)
    
        Case ccHue
            m_Hue = (xPos / boxWidth)
        
        Case ccSaturation
            m_Saturation = (xPos / boxWidth)
        
        Case ccValue
            m_Value = (xPos / boxWidth)
    
    End Select
    
    'Recalculate RGB based on the new HSV values
    HSVtoRGB m_Hue, m_Saturation, m_Value, m_Red, m_Green, m_Blue
    
    'Redraw the interface
    SyncInterfaceToCurrentColor

End Sub

Private Sub picSampleRGB_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    RGBBoxClicked Index, x
End Sub

Private Sub picSampleRGB_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then RGBBoxClicked Index, x
End Sub

'Whenever one of the RGB sample boxes is clicked, this function is called; it calculates new RGB/HSV values, then redraws the interface
Private Sub RGBBoxClicked(ByVal boxIndex As Long, ByVal xPos As Long)

    Dim boxWidth As Long
    boxWidth = picSampleRGB(0).ScaleWidth
    
    'Restrict mouse clicks to the picture box area
    If xPos < 0 Then xPos = 0
    If xPos > boxWidth Then xPos = boxWidth

    Select Case boxIndex
    
        Case ccRed
            m_Red = (xPos / boxWidth) * 255
        
        Case ccGreen
            m_Green = (xPos / boxWidth) * 255
        
        Case ccBlue
            m_Blue = (xPos / boxWidth) * 255
    
    End Select
    
    'Recalculate HSV based on the new RGB values
    RGBtoHSV m_Red, m_Green, m_Blue, m_Hue, m_Saturation, m_Value
    
    'Redraw the interface
    SyncInterfaceToCurrentColor

End Sub

'Whenever a text box value is changed, sync only the relevant value, then redraw the interface
Private Sub tudHSV_Change(Index As Integer)

    If Not m_suspendTextResync Then

        Select Case (Index + 3)
        
            Case ccHue
                If tudHSV(Index).IsValid Then m_Hue = tudHSV(Index) / 359
            
            Case ccSaturation
                If tudHSV(Index).IsValid Then m_Saturation = tudHSV(Index) / 100
            
            Case ccValue
                If tudHSV(Index).IsValid Then m_Value = tudHSV(Index) / 100
        
        End Select
        
        'Recalculate RGB based on the new HSV values
        HSVtoRGB m_Hue, m_Saturation, m_Value, m_Red, m_Green, m_Blue
        
        'Redraw the interface
        SyncInterfaceToCurrentColor
        
    End If

End Sub

'Whenever a text box value is changed, sync only the relevant value, then redraw the interface
Private Sub tudRGB_Change(Index As Integer)

    If Not m_suspendTextResync Then

        Select Case Index
        
            Case ccRed
                If tudRGB(Index).IsValid Then m_Red = tudRGB(Index)
            
            Case ccGreen
                If tudRGB(Index).IsValid Then m_Green = tudRGB(Index)
        
            Case ccBlue
                If tudRGB(Index).IsValid Then m_Blue = tudRGB(Index)
        
        End Select
        
        'Recalculate HSV values based on the new RGB values
        RGBtoHSV m_Red, m_Green, m_Blue, m_Hue, m_Saturation, m_Value
        
        'Redraw the interface
        SyncInterfaceToCurrentColor
        
    End If

End Sub

'Full validation of hex input happens in its LostFocus event, but we also do a quick-and-dirty sync during change events
Private Sub txtHex_Change()
    
    If m_suspendHexInput Or m_suspendTextResync Then Exit Sub
    
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
