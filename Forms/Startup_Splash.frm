VERSION 5.00
Begin VB.Form FormSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3300
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   11685
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   Enabled         =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   220
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   779
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
End
Attribute VB_Name = "FormSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Splash Screen
'Copyright 2001-2018 by Tanner Helland
'Created: 15/April/01
'Last updated: 02/September/17
'Last update: handle paint messages internally, rather than leaning on AutoRedraw
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RectL) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RectL) As Long

'A logo, drop shadow and screen backdrop are used to generate the splash.  These DIBs are released once m_splashDIB (below)
' has been successfully assembled.
Private m_logoDIB As pdDIB, m_shadowDIB As pdDIB
Private m_splashDIB As pdDIB

'AutoRedraw is set to FALSE for this form; we paint it ourselves using this DIB as a backbuffer
Private m_BackBuffer As pdDIB

'We skip the entire display process if any of the DIBs can't be created
Private m_dibsLoadedSuccessfully As Boolean

'Some information is custom-drawn onto the logo at run-time.  pdFont objects are used to render any text.
Private m_versionFont As pdFont

'On high-DPI monitors, some stretching is required.  In the future, I would like to replace this with a more
' elegant solution.
Private m_logoAspectRatio As Double

'The maximum progress count of the load operation is stored here.  The value is passed to the initial
' prepareSplashLogo function, and it is not modified once loaded.
Private m_maxProgress As Long, m_progressAtFirstNotify As Long

'To prevent flicker, we manually capture and ignore WM_ERASEBKGND messages
Private Const WM_ERASEBKGND As Long = &H14
Implements ISubclass

'Load any logo DIBs from the .exe's resource area, and precalculate some rendering values
Public Sub PrepareSplashLogo(ByVal maxProgressValue As Long)
    
    m_maxProgress = maxProgressValue
    m_progressAtFirstNotify = -1
    m_dibsLoadedSuccessfully = False
    
    Set m_logoDIB = New pdDIB
    Set m_shadowDIB = New pdDIB
    
    'Load the logo DIB, and calculate an aspect ratio (important if high-DPI settings are in use)
    Dim origLogoWidth As Long, origLogoHeight As Long
    origLogoWidth = FixDPI(779)
    origLogoHeight = FixDPI(220)
    m_dibsLoadedSuccessfully = LoadResourceToDIB("pd_logo_white", m_logoDIB, origLogoWidth, origLogoHeight, , , True)
    If m_dibsLoadedSuccessfully Then m_logoAspectRatio = CDbl(m_logoDIB.GetDIBWidth) / CDbl(m_logoDIB.GetDIBHeight)
    
    'Load the inverted logo DIB; this will be blurred and used as a shadow backdrop
    m_dibsLoadedSuccessfully = m_dibsLoadedSuccessfully And LoadResourceToDIB("pd_logo_black", m_shadowDIB, origLogoWidth, origLogoHeight, , , True)
    
End Sub

'Load the form backdrop.  Note that this CANNOT BE DONE until the global monitor classes are initialized.
Public Sub PrepareRestOfSplash()
    
    If m_dibsLoadedSuccessfully Then
    
        'Use the GetDesktopAsDIB function to retrieve a copy of the current screen.  We will use this to mimic window
        ' transparency.  (It's faster, and works more smoothly than attempting to use layered Windows, especially on XP.)
        Dim captureRect As RectL
        GetWindowRect Me.hWnd, captureRect
        ScreenCapture.GetPartialDesktopAsDIB m_splashDIB, captureRect
        
        Dim formLeft As Long, formTop As Long, formWidth As Long, formHeight As Long
        formLeft = captureRect.Left
        formTop = captureRect.Top
        GetClientRect Me.hWnd, captureRect
        formWidth = captureRect.Right - captureRect.Left
        formHeight = captureRect.Bottom - captureRect.Top
        
        'Copy the screen background, shadow, and logo onto a single composite DIB
        m_shadowDIB.AlphaBlendToDC m_splashDIB.GetDIBDC, , FixDPI(1), FixDPI(1), formWidth, formWidth / m_logoAspectRatio
        m_logoDIB.AlphaBlendToDC m_splashDIB.GetDIBDC, , 0, 0, formWidth, formWidth / m_logoAspectRatio
        
        'Free all intermediate DIBs
        Set m_shadowDIB = Nothing
        Set m_logoDIB = Nothing
        
        'Next, we need to figure out where the top and bottom of the "PHOTODEMON" logo lie.  These values may change
        ' depending on the current screen DPI.  (Their position is important, because other text is laid out proportional
        ' to these values.)
        Dim pdLogoTop As Long, pdLogoBottom As Long, pdLogoRight As Long
        
        'FYI: the hard-coded values are for 96 DPI
        pdLogoTop = FixDPI(60)
        pdLogoBottom = FixDPI(125)
        pdLogoRight = FixDPI(755)
        
        'Next, we need to prepare a font renderer for displaying the current program version
        Set m_versionFont = New pdFont
        m_versionFont.SetFontBold True
        m_versionFont.SetFontSize 14
        
        'Non-production builds are tagged RED; normal builds, BLUE.  In the future, this may be tied to the theming engine.
        ' (It's not easy to do it at present, because the themer is loaded late in the program intialization process.)
        If (PD_BUILD_QUALITY <> PD_PRODUCTION) Then
            m_versionFont.SetFontColor RGB(255, 50, 50)
        Else
            m_versionFont.SetFontColor RGB(50, 127, 255)
        End If
        
        m_versionFont.CreateFontObject
        
        'Assemble the current version and description strings
        Dim versionString As String
        Dim versionWidth As Long, versionHeight As Long
        
        versionString = g_Language.TranslateMessage("version %1", GetPhotoDemonVersion)
        
        'Render the version string just below the logo text
        If (m_splashDIB.GetDIBDC <> 0) Then
            m_versionFont.AttachToDC m_splashDIB.GetDIBDC
            versionWidth = m_versionFont.GetWidthOfString(versionString)
            versionHeight = m_versionFont.GetHeightOfString(versionString)
            m_versionFont.FastRenderText pdLogoRight - versionWidth, pdLogoBottom + FixDPI(8), versionString
            m_versionFont.ReleaseFromDC
        End If
        
        'Copy the composite image onto the underlying form
        If (m_BackBuffer Is Nothing) Then Set m_BackBuffer = New pdDIB
        m_BackBuffer.CreateBlank m_splashDIB.GetDIBWidth, m_splashDIB.GetDIBHeight, 24, 0
        GDI.BitBltWrapper m_BackBuffer.GetDIBDC, 0, 0, formWidth, formHeight, m_splashDIB.GetDIBDC, 0, 0, vbSrcCopy
        
        'Ensure the form has been painted at least once prior to display
        FlipBackBufferToScreen
        
    Else
        pdDebug.LogAction "WARNING!  Splash DIBs could not be loaded; something may be catastrophically wrong."
    End If
    
End Sub

'When the load function updates the current progress count, we refresh the splash screen to reflect the new progress.
Public Sub UpdateLoadProgress(ByVal newProgressMarker As Long)
    
    'If progress notifications arrived before the form was made visible, ignore them; this makes the loading bar appear
    ' more fluid, rather than magically jumping to the middle of the form when it's first loaded.
    If (m_progressAtFirstNotify = -1) Then m_progressAtFirstNotify = newProgressMarker - 1
    
    'Calculate the length of the progress line.  This is effectively arbitrary; I've made it the length of the
    ' logo image minus 10% for now.
    Dim lineLength As Long, lineOffset As Long
    lineLength = m_splashDIB.GetDIBWidth * 0.9
    lineOffset = (m_splashDIB.GetDIBWidth - lineLength) \ 2
    
    'Draw the current progress, if relevant
    If (m_maxProgress > 0) And Me.Visible Then
    
        'Copy the splash DIB to overwrite any old drawing
        GDI.BitBltWrapper Me.hDC, 0, 0, m_splashDIB.GetDIBWidth, m_splashDIB.GetDIBHeight, m_splashDIB.GetDIBDC, 0, 0, vbSrcCopy
        
        'Draw the progress line using GDI+
        Dim lineRadius As Long, lineY As Long
        lineRadius = FixDPI(6)
        lineY = m_splashDIB.GetDIBHeight - FixDPI(2) - lineRadius
        
        Dim cPainter As pd2DPainter, cSurface As pd2DSurface, cPen As pd2DPen
        Drawing2D.QuickCreatePainter cPainter
        
        Set cSurface = New pd2DSurface
        cSurface.WrapSurfaceAroundPDDIB m_BackBuffer
        cSurface.SetSurfaceAntialiasing P2_AA_HighQuality
        cSurface.SetSurfacePixelOffset P2_PO_Half
        
        Drawing2D.QuickCreateSolidPen cPen, lineRadius, g_Themer.GetGenericUIColor(UI_Accent), 100#, , P2_LC_Round
        cPainter.DrawLineF cSurface, cPen, lineOffset, lineY, (m_splashDIB.GetDIBWidth - lineOffset) * ((newProgressMarker - m_progressAtFirstNotify) / (m_maxProgress - m_progressAtFirstNotify)), lineY
        
        Set cSurface = Nothing
        
        'Manually refresh the form
        FlipBackBufferToScreen
        
    End If

End Sub

Private Sub Form_Load()

    'Unfortunately, we have to subclass to prevent obnoxious flickering when the form is first displayed
    If pdMain.IsProgramRunning() And OS.IsProgramCompiled() Then VBHacks.StartSubclassing Me.hWnd, Me
    
End Sub

'Painting is easy - just flip the backbuffer to the screen and call it a day!
Private Sub Form_Paint()
    FlipBackBufferToScreen
End Sub

Private Sub FlipBackBufferToScreen()
    GDI.BitBltWrapper Me.hDC, 0, 0, m_BackBuffer.GetDIBWidth, m_BackBuffer.GetDIBHeight, m_BackBuffer.GetDIBDC, 0, 0, vbSrcCopy
End Sub

Private Function ISubclass_WindowMsg(ByVal hWnd As Long, ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long

    If (uiMsg = WM_ERASEBKGND) Then
         ISubclass_WindowMsg = 1
    ElseIf (uiMsg = WM_NCDESTROY) Then
        
        VBHacks.StopSubclassing hWnd, Me
        
        'Allow VB to continue with its own internal teardown process
        ISubclass_WindowMsg = VBHacks.DefaultSubclassProc(hWnd, uiMsg, wParam, lParam)
        
    Else
        ISubclass_WindowMsg = VBHacks.DefaultSubclassProc(hWnd, uiMsg, wParam, lParam)
    End If

End Function
