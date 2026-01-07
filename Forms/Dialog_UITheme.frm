VERSION 5.00
Begin VB.Form dialog_UITheme 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Theme and language"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9045
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
   ScaleHeight     =   500
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   603
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdLabel lblExplanation 
      Height          =   975
      Index           =   0
      Left            =   240
      Top             =   5760
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   1720
      Alignment       =   2
      Caption         =   ""
      Layout          =   1
   End
   Begin PhotoDemon.pdButtonStrip btsIcons 
      Height          =   1335
      Left            =   120
      TabIndex        =   4
      Top             =   4200
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   2355
      Caption         =   "icons"
      FontSize        =   12
   End
   Begin PhotoDemon.pdStrip strAccents 
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   1931
      Caption         =   "interface accent color"
   End
   Begin PhotoDemon.pdButtonStrip btsInterface 
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   2143
      Caption         =   "interface theme"
      FontSize        =   12
   End
   Begin PhotoDemon.pdLabel lblLangAuthor 
      Height          =   375
      Left            =   240
      Top             =   1200
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   661
   End
   Begin PhotoDemon.pdDropDown cboLanguage 
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   1508
      Caption         =   "language"
      FontSize        =   12
   End
   Begin PhotoDemon.pdCommandBarMini cmdBar 
      Align           =   2  'Align Bottom
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   6765
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   1296
   End
End
Attribute VB_Name = "dialog_UITheme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'First-run Dialog
'Copyright 2012-2026 by Tanner Helland
'Created: 12/February/12
'Last updated: 14/February/17
'Last update: finally finish implementing this thing!
'
'At first-run, PhotoDemon now asks the user to confirm their choice of program language and UI theme.
' The dialog can be canceled (in which case default settings will be used), but my hope is that users
' will be able to configure everything the way they prefer, without needing to dive into PD's complicated
' menu system right off the bat.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'PD language files currently available on this system.
Private m_AvailableLanguages() As PDLanguageFile

'While language files are loaded, we suspend click events in the language drop-down box
Private m_SuspendUpdates As Boolean

'Theme accent colors are cached at startup, to improve rendering performance
Private m_AccentColors() As Long

'We default to a dark, blue-accent theme by default, but the dialog "live-updates" as the user toggles settings.
' These settings are cached at module-level so our parent function can retrieve them if the user hits "OK".
Private m_LangIndex As Long
Private m_ThemeClass As PD_THEME_CLASS, m_ThemeAccent As PD_THEME_ACCENT, m_MonoIcons As Boolean

'The user input from the dialog.  If the user cancels this dialog, default settings will be used.
Private m_CmdBarAnswer As VbMsgBoxResult

Public Property Get DialogResult() As VbMsgBoxResult
    DialogResult = m_CmdBarAnswer
End Property

'The ShowDialog routine presents the user with the form.
Public Sub ShowDialog()
    
    'Provide a default answer (in case the user closes the dialog via some means other than the command bar)
    m_CmdBarAnswer = vbCancel
    
    'Cache current theme settings
    m_ThemeClass = g_Themer.GetCurrentThemeClass()
    m_ThemeAccent = g_Themer.GetCurrentThemeAccent()
    If (m_ThemeAccent = PDTA_Undefined) Then m_ThemeAccent = PDTA_Blue
    
    'Retrieve a list of available languages from the translation engine
    m_LangIndex = g_Language.GetCurrentLanguageIndex
    g_Language.CopyListOfLanguages m_AvailableLanguages
    
    'Populate the language dropdown with any/all language files installed on this system
    m_SuspendUpdates = True
    cboLanguage.Clear
    
    Dim chrSpace As String
    chrSpace = " "
    
    Dim i As Long
    For i = 0 To UBound(m_AvailableLanguages)
        cboLanguage.AddItem chrSpace & m_AvailableLanguages(i).LangName
    Next i
    
    'TODO: high-contrast theme
    btsInterface.AddItem "dark", 0
    btsInterface.AddItem "light", 1
    If (m_ThemeClass = PDTC_Dark) Then btsInterface.ListIndex = 0 Else btsInterface.ListIndex = 1
    
    'Accent colors are currently listed by name, and these are manually mapped to colors in the rendering function.
    ' In the future, we may want to automate this process by manually iterating theme files and pulling in their
    ' significant accent colors.
    ReDim m_AccentColors(0 To 7) As Long
    strAccents.AddItem "blue", 0
    Colors.GetColorFromString "#2196f3", m_AccentColors(0), ColorHex
    
    strAccents.AddItem "brown", 1
    Colors.GetColorFromString "#795548", m_AccentColors(1), ColorHex
    
    strAccents.AddItem "green", 2
    Colors.GetColorFromString "#4caf50", m_AccentColors(2), ColorHex
    
    strAccents.AddItem "orange", 3
    Colors.GetColorFromString "#ff9800", m_AccentColors(3), ColorHex
    
    strAccents.AddItem "pink", 4
    Colors.GetColorFromString "#ec407a", m_AccentColors(4), ColorHex
    
    strAccents.AddItem "purple", 5
    Colors.GetColorFromString "#ab47bc", m_AccentColors(5), ColorHex
    
    strAccents.AddItem "red", 6
    Colors.GetColorFromString "#f44336", m_AccentColors(6), ColorHex
    
    strAccents.AddItem "teal", 7
    Colors.GetColorFromString "#26a69a", m_AccentColors(7), ColorHex
    
    strAccents.ListIndex = m_ThemeAccent
    
    btsIcons.AddItem "  default", 0
    btsIcons.AddItem "  monochromatic", 1
    btsIcons.AssignImageToItem 0, "generic_color", , 36, 36, True
    btsIcons.AssignImageToItem 1, "generic_grey", , 36, 36, True
    
    m_MonoIcons = g_Themer.GetMonochromeIconSetting()
    If m_MonoIcons Then btsIcons.ListIndex = 1 Else btsIcons.ListIndex = 0
    
    m_SuspendUpdates = False
    cboLanguage.ListIndex = m_LangIndex
    
    'Apply any custom styles to the form
    ApplyThemeAndTranslations Me

    'Display the form
    ShowPDDialog vbModal, Me, True

End Sub

Public Sub GetNewSettings(ByRef newLangIndex As Long, ByRef newThemeClass As PD_THEME_CLASS, ByRef newThemeAccent As PD_THEME_ACCENT, ByRef newMonoIcons As Boolean)
    newLangIndex = m_LangIndex
    newThemeClass = m_ThemeClass
    newThemeAccent = m_ThemeAccent
    newMonoIcons = m_MonoIcons
End Sub

Private Sub btsIcons_Click(ByVal buttonIndex As Long)
    LiveUpdateUITheme
End Sub

Private Sub btsInterface_Click(ByVal buttonIndex As Long)
    LiveUpdateUITheme
End Sub

Private Sub cboLanguage_Click()
    
    m_LangIndex = cboLanguage.ListIndex
    
    g_Language.UndoTranslations Me
    g_Language.ActivateNewLanguage m_LangIndex
    g_Language.ApplyLanguage False, False
    
    If (Not m_SuspendUpdates) Then
        If (LenB(m_AvailableLanguages(m_LangIndex).Author) <> 0) Then
            lblLangAuthor.Caption = g_Language.TranslateMessage("Translation by %1", m_AvailableLanguages(m_LangIndex).Author)
        Else
            lblLangAuthor.Caption = vbNullString
        End If
    End If
    
    Interface.ApplyThemeAndTranslations Me
    LiveUpdateUITheme
    
End Sub

Private Sub cmdBar_CancelClick()
    m_CmdBarAnswer = vbCancel
    Me.Hide
End Sub

Private Sub cmdBar_OKClick()
    m_CmdBarAnswer = vbOK
    Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub strAccents_Click(ByVal buttonIndex As Long)
    LiveUpdateUITheme
End Sub

'The accent color strip is owner-drawn, so we must respond to rendering events and paint the accent colors manually
Private Sub strAccents_DrawButton(ByVal btnIndex As Long, ByVal btnValue As String, ByVal targetDC As Long, ByVal ptrToRectF As Long)

    If ((LenB(btnValue) <> 0) And PDMain.IsProgramRunning()) Then
    
        Dim tmpRectF As RectF
        CopyMemoryStrict VarPtr(tmpRectF), ptrToRectF, 16&

        'Map the index to an actual color
        Dim targetColor As Long
        targetColor = m_AccentColors(btnIndex)
        
        'Note that accents colors *are* color-managed inside this dialog
        Dim cmResult As Long
        ColorManagement.ApplyDisplayColorManagement_SingleColor targetColor, cmResult
    
        Dim cSurface As pd2DSurface: Dim cBrush As pd2DBrush
        Drawing2D.QuickCreateSurfaceFromDC cSurface, targetDC
        Drawing2D.QuickCreateSolidBrush cBrush, cmResult
        PD2D.FillRectangleF_FromRectF cSurface, cBrush, tmpRectF
        
        Set cSurface = Nothing: Set cBrush = Nothing
        
    End If
    
End Sub

Private Sub LiveUpdateUITheme()
    
    If (Not m_SuspendUpdates) Then
        
        'Cache the currently selected theme settings at module level; these may need to be returned to our
        ' parent function if the dialog is closed after this request.
        If (btsInterface.ListIndex = 0) Then m_ThemeClass = PDTC_Dark Else m_ThemeClass = PDTC_Light
        m_ThemeAccent = strAccents.ListIndex
        m_MonoIcons = (btsIcons.ListIndex = 1)
        
        'Relay any changes to PD's central themer and load a new theme to match
        g_Themer.SetMonochromeIconSetting m_MonoIcons
        g_Themer.SetNewTheme m_ThemeClass, m_ThemeAccent
        g_Themer.LoadDefaultPDTheme
        
        'Let the user know these settings can be changed at any time, and there are many more
        ' customization options available to them!
        Dim cString As pdString
        Set cString = New pdString
        cString.Append g_Language.TranslateMessage("You can change language and theme settings at any time from the Tools menu.")
        
        Dim useLineBreak As Boolean: useLineBreak = True
        If (Not g_Language Is Nothing) Then useLineBreak = Not g_Language.TranslationActive()
        
        Dim tmpString As String
        tmpString = g_Language.TranslateMessage("Additional interface customizations are available in the Window menu.")
        If useLineBreak Then
            cString.AppendLineBreak
            cString.Append tmpString
        Else
            cString.Append "  "
            cString.Append tmpString
        End If
        
        lblExplanation(0).Caption = cString.ToString()
        
        'Normally, resources need to be reset after a theme change, but we deliberately suspend this
        ' inside this dialog (because we don't want the "color" representation icon to be forced to monochrome)
        'g_Resources.NotifyThemeChange
        
        'Re-theme this dialog
        Interface.ApplyThemeAndTranslations Me
        
    End If
    
End Sub
