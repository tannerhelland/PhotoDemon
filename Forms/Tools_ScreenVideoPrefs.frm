VERSION 5.00
Begin VB.Form FormRecordAPNGPrefs 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   " Animated screen capture (APNG)"
   ClientHeight    =   5100
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5910
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
   ScaleHeight     =   340
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   394
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdCommandBarMini cmdBar 
      Align           =   2  'Align Bottom
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   4365
      Width           =   5910
      _ExtentX        =   10425
      _ExtentY        =   1296
   End
   Begin PhotoDemon.pdSlider sldFrameRate 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      _ExtentX        =   14631
      _ExtentY        =   1296
      Caption         =   "maximum frame rate (fps)"
      Min             =   0.1
      Max             =   30
      SigDigits       =   1
      Value           =   10
      NotchPosition   =   2
      NotchValueCustom=   10
   End
   Begin PhotoDemon.pdButtonStrip btsLoop 
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   1720
      Caption         =   "repeat final animation"
   End
   Begin PhotoDemon.pdSlider sldLoop 
      Height          =   735
      Left            =   480
      TabIndex        =   3
      Top             =   3240
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1296
      Caption         =   "repeat count"
      FontSizeCaption =   10
      Min             =   1
      Max             =   65535
      ScaleStyle      =   2
      Value           =   1
      DefaultValue    =   1
   End
   Begin PhotoDemon.pdButtonStrip btsMouse 
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   1720
      Caption         =   "record mouse actions"
   End
End
Attribute VB_Name = "FormRecordAPNGPrefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Animated screen capture dialog
'Copyright 2020-2020 by Tanner Helland
'Created: 01/July/20
'Last updated: 17/July/20
'Last update: expand recording options
'
'PD can write animated PNGs.  APNGs seem like a great fit for animated screen captures.
' Let's see if we can merge the two, eh?
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Destination filename, populated when the user hits OK
Private m_Filename As String

'Last-used settings must be stored manually, since we aren't using a dedicated command bar
Private WithEvents m_lastUsedSettings As pdLastUsedSettings
Attribute m_lastUsedSettings.VB_VarHelpID = -1

Private Sub btsLoop_Click(ByVal buttonIndex As Long)
    ReflowInterface
End Sub

Private Sub cmdBar_OKClick()
    
    'Before proceeding, we need to prompt the user for a destination filename.
    
    'Start by validating m_Filename, which will be filled with the user's past destination
    ' filename (if one exists), or a default capture filename in the user's current
    ' "Save image" folder
    If ((LenB(m_Filename) = 0) Or (Not Files.PathExists(Files.FileGetPath(m_Filename)))) Then
    
        'm_Filename is bad.  Attempt to populate it with default values.
        Dim tmpPath As String, tmpFilename As String
        tmpPath = UserPrefs.GetPref_String("Paths", "Save Image", vbNullString)
        tmpFilename = g_Language.TranslateMessage("capture")
        m_Filename = tmpPath & IncrementFilename(tmpPath, tmpFilename, "apng") & ".apng"
    
    End If
    
    'Use a standard common-dialog to prompt for filename
    Dim cSave As pdOpenSaveDialog
    Set cSave = New pdOpenSaveDialog
    
    Dim okToProceed As Boolean, sFile As String
    sFile = m_Filename
    okToProceed = cSave.GetSaveFileName(sFile, Files.FileGetName(m_Filename), True, "Animated PNG (.apng)|*.apng;*.png", 1, Files.FileGetPath(m_Filename), "Save image", ".apng", Me.hWnd)
    
    'The user can cancel the common-dialog - that's fine; it just means we don't save
    ' any of the current settings (or close the window).
    If okToProceed Then
        
        'Save the current export path as the latest "save image" path
        m_Filename = sFile
        UserPrefs.SetPref_String "Paths", "Save Image", Files.FileGetPath(m_Filename)
        
        'Before hiding this window, retrieve our current window position; the launched
        ' screen recording window will use this to position itself the first time it's invoked
        Dim myRect As winRect
        If (Not g_WindowManager Is Nothing) Then g_WindowManager.GetWindowRect_API Me.hWnd, myRect
        
        'Because this dialog is modal, it needs to be hidden before we invoke a modeless dialog
        Me.Hide
        
        'Also hide the main PhotoDemon window
        FormMain.WindowState = vbMinimized
        
        'The loop setting is a little weird.
        ' 0 = loop infinitely, 1 = loop once, 2+ = loop that many times exactly
        Dim loopCount As Long
        If (btsLoop.ListIndex = 0) Then
            loopCount = 1
        ElseIf (btsLoop.ListIndex = 1) Then
            loopCount = 0
        Else
            loopCount = CLng(sldLoop.Value + 1)
        End If
        
        'Launch the capture form, then note that the command bar will handle unloading this form
        FormRecordAPNG.ShowDialog VarPtr(myRect), m_Filename, sldFrameRate.Value, loopCount, (btsMouse.ListIndex >= 1), (btsMouse.ListIndex >= 2)
        
    Else
        cmdBar.DoNotUnloadForm
    End If
    
End Sub

Private Sub Form_Load()
    
    'If this dialog was previously used this session, we want to make sure the capture window
    ' has also been freed (as we need to reinitialize it)
    Set FormRecordAPNG = Nothing
    
    'Prep any UI elements
    btsMouse.AddItem "no", 0
    btsMouse.AddItem "cursor only", 1
    btsMouse.AddItem "cursor and clicks", 2
    btsMouse.ListIndex = 1
    
    btsLoop.AddItem "none", 0
    btsLoop.AddItem "forever", 1
    btsLoop.AddItem "custom", 2
    btsLoop.ListIndex = 0
    
    'Load any previously used settings
    Set m_lastUsedSettings = New pdLastUsedSettings
    m_lastUsedSettings.SetParentForm Me
    m_lastUsedSettings.LoadAllControlValues
    
    'The OK button uses custom text
    If (Not g_Language Is Nothing) Then
        cmdBar.SetCustomOKText g_Language.TranslateMessage("Continue"), 24
    End If
    
    'Apply custom themes
    Interface.ApplyThemeAndTranslations Me
    
    'With theming handled, reflow the interface one final time before displaying the window
    ReflowInterface
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    m_lastUsedSettings.SaveAllControlValues
    m_lastUsedSettings.SetParentForm Nothing
    Interface.ReleaseFormTheming Me
End Sub

Private Sub m_LastUsedSettings_AddCustomPresetData()
    If ((LenB(m_Filename) <> 0) And Files.PathExists(Files.FileGetPath(m_Filename))) Then m_lastUsedSettings.AddPresetData "dst-capture-filename", m_Filename
End Sub

Private Sub m_LastUsedSettings_ReadCustomPresetData()

    'Look for a previously saved destination filename.  If one does not exist,
    ' we want to populate the destination path with a good default suggestion.
    If (Not m_lastUsedSettings.DoesPresetExist("dst-capture-filename")) Then
    
        Dim tmpPath As String, tmpFilename As String, tmpCombined As String
        tmpPath = UserPrefs.GetPref_String("Paths", "Save Image", vbNullString)
        tmpFilename = g_Language.TranslateMessage("capture")
        tmpCombined = tmpPath & IncrementFilename(tmpPath, tmpFilename, "apng") & ".apng"
        
        m_Filename = tmpCombined
        
    Else
        m_Filename = m_lastUsedSettings.RetrievePresetData("dst-capture-filename", UserPrefs.GetPref_String("Paths", "Save Image", vbNullString) & "capture.apng")
    End If

End Sub

Private Sub ReflowInterface()

    Dim yPadding As Long, yPaddingTitle As Long
    yPadding = Interface.FixDPI(8)
    yPaddingTitle = Interface.FixDPI(12)
    
    Dim yOffset As Long
    yOffset = btsLoop.GetTop + btsLoop.GetHeight + yPadding
    sldLoop.Visible = (btsLoop.ListIndex = 2)
    If sldLoop.Visible Then
        sldLoop.SetTop yOffset
        yOffset = yOffset + sldLoop.GetHeight + yPaddingTitle
    Else
        yOffset = yOffset - yPadding + yPaddingTitle
    End If
    
End Sub
