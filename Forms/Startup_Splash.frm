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
   HasDC           =   0   'False
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
'Copyright 2001-2026 by Tanner Helland
'Created: 15/April/01
'Last updated: 18/December/20
'Last update: remove GDI+ font code from splash rendering; GDI+ font creation functions have
'             unpredictable perf impacts, and it's reliably faster to use GDI on 24-bpp surfaces
'             (then produce a 32-bpp copy ourselves)
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RectL) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal hWnd As Long, ByVal hdcDst As Long, ByVal pptDst As Long, ByVal psize As Long, ByVal hdcSrc As Long, ByVal pptSrc As Long, ByVal crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long

'A logo, drop shadow and screen backdrop are used to generate the splash.
' These DIBs are released once m_splashDIB (below) has been successfully assembled.
Private m_logoDIB As pdDIB, m_shadowDIB As pdDIB
Private m_splashDIB As pdDIB

'AutoRedraw is set to FALSE for this form; we paint it ourselves using this DIB as a backbuffer
Private m_BackBuffer As pdDIB

'We skip the entire display process if any of the DIBs can't be created
Private m_dibsLoadedSuccessfully As Boolean

'On high-DPI monitors, some stretching is required.
Private m_logoAspectRatio As Double

'The maximum progress count of the load operation is stored here.  The value is passed to the initial
' prepareSplashLogo function, and it is not modified once loaded.
Private m_maxProgress As Long, m_progressAtFirstNotify As Long

'To prevent flicker, we manually capture and ignore WM_ERASEBKGND messages
Private Const WM_ERASEBKGND As Long = &H14
Implements ISubclass

'Load any logo DIBs from the .exe's resource area, and pre-calculate some rendering values
Public Sub PrepareSplashLogo(ByVal maxProgressValue As Long)
    
    m_maxProgress = maxProgressValue
    m_progressAtFirstNotify = -1
    m_dibsLoadedSuccessfully = False
    
    Set m_logoDIB = New pdDIB
    Set m_shadowDIB = New pdDIB
    
    'Load the logo DIB, and calculate an aspect ratio (important if high-DPI settings are in use)
    Dim origLogoWidth As Long, origLogoHeight As Long
    origLogoWidth = Interface.FixDPI(779)
    origLogoHeight = Interface.FixDPI(220)
    m_dibsLoadedSuccessfully = LoadResourceToDIB("pd_logo_white", m_logoDIB, origLogoWidth, origLogoHeight, suspendMonochrome:=True)
    If m_dibsLoadedSuccessfully Then m_logoAspectRatio = CDbl(m_logoDIB.GetDIBWidth) / CDbl(m_logoDIB.GetDIBHeight)
    
    'Load the inverted logo DIB; this will be blurred and used as a shadow backdrop
    m_dibsLoadedSuccessfully = m_dibsLoadedSuccessfully And LoadResourceToDIB("pd_logo_black", m_shadowDIB, origLogoWidth, origLogoHeight, suspendMonochrome:=True)
    
End Sub

'Load the form backdrop.  Note that this CANNOT BE DONE until the global monitor classes are initialized.
Public Sub PrepareRestOfSplash()
    
    If m_dibsLoadedSuccessfully Then
        
        'Create a blank DIB at the size of the current splash window
        Dim captureRect As RectL
        GetWindowRect Me.hWnd, captureRect
        
        Dim formWidth As Long, formHeight As Long
        formWidth = captureRect.Right - captureRect.Left
        formHeight = captureRect.Bottom - captureRect.Top
        
        Set m_splashDIB = New pdDIB
        m_splashDIB.CreateBlank formWidth, formHeight, 32, 0, 0
        m_splashDIB.SetInitialAlphaPremultiplicationState True
        
        'Paint the drop shadow and logo onto the newly created composite DIB
        m_shadowDIB.AlphaBlendToDC m_splashDIB.GetDIBDC, , Interface.FixDPI(1), Interface.FixDPI(1), formWidth, formWidth / m_logoAspectRatio
        m_logoDIB.AlphaBlendToDC m_splashDIB.GetDIBDC, , 0, 0, formWidth, formWidth / m_logoAspectRatio
        
        'Free all intermediate DIBs; they are no longer needed
        Set m_shadowDIB = Nothing
        Set m_logoDIB = Nothing
        
        'Next, we need to figure out where the top and bottom of the "PHOTODEMON" logo lie.
        ' These values may change depending on screen DPI.  (Their position is important,
        ' because other text - like PD's version number - gets laid out proportional to
        ' these values.)
        Dim pdLogoTop As Long, pdLogoBottom As Long, pdLogoRight As Long
        
        'FYI: these hard-coded values are for 96 DPI; the FixDPI() function scales as needed
        pdLogoTop = Interface.FixDPI(60)
        pdLogoBottom = Interface.FixDPI(125)
        pdLogoRight = Interface.FixDPI(760)
        
        'Next, we need to prepare a font renderer for displaying the current program version.
        ' Because we're rendering to a 32-bpp surface, we can't use simple GDI fonts -
        ' GDI+ is required.
        Dim logoFontSize As Single
        logoFontSize = 14
        
        'Font availability varies by OS
        Dim logoFontName As String
        If OS.IsVistaOrLater Then logoFontName = "Segoe UI" Else logoFontName = "Tahoma"
        
        'Font color varies by build version.
        ' (As a convenience, non-production builds are tagged RED; normal builds, BLUE.)
        Dim logoFontColor As Long
        If (PD_BUILD_QUALITY <> PD_PRODUCTION) Then logoFontColor = RGB(255, 50, 50) Else logoFontColor = RGB(50, 127, 255)
        
        'Create a GDI font with the desired settings
        Dim tmpFont As pdFont
        Set tmpFont = New pdFont
        tmpFont.SetFontFace logoFontName
        tmpFont.SetFontSize logoFontSize
        tmpFont.SetFontBold True
        tmpFont.CreateFontObject
        
        'Assemble the current version and description strings
        Dim versionString As String
        versionString = Trim$(g_Language.TranslateMessage("version %1", Updates.GetPhotoDemonVersion()))
        
        'Create a dummy 24-bpp DIB and paint the version to it.  (This is required for reliable
        ' antialiasing behavior; GDI can't render text to 32-bpp surfaces if antialiasing is active,
        ' so we need to paint to a 24-bpp surface, then upsample to 32-bpp and fill in alpha manually.)
        Dim fntWidth As Long, fntHeight As Long
        fntWidth = tmpFont.GetWidthOfString(versionString)
        fntHeight = tmpFont.GetHeightOfString(versionString)
        
        'Create a temporary 24-bpp target for the text.  (We'll use white-on-black text to
        ' simplify the process of manually upsampling to 32-bpp.)
        Dim tmpDIB As pdDIB
        Set tmpDIB = New pdDIB
        tmpDIB.CreateBlank fntWidth, fntHeight, 24, vbBlack
        
        'Paint the version string
        tmpFont.SetFontColor vbWhite
        tmpFont.AttachToDC tmpDIB.GetDIBDC
        tmpFont.FastRenderText 0, 0, versionString
        tmpFont.ReleaseFromDC
        Set tmpFont = Nothing
        
        'Convert the temporary DIB to 32-bpp, using the grayscale channel as the alpha guide
        tmpDIB.ConvertTo32bpp
        Dim grayArray() As Byte
        ReDim grayArray(0 To fntWidth - 1, 0 To fntHeight - 1) As Byte
        DIBs.GetDIBGrayscaleMap tmpDIB, grayArray, False
        
        tmpDIB.CreateBlank fntWidth, fntHeight, 32, logoFontColor, 255
        DIBs.ApplyTransparencyTable tmpDIB, grayArray
        tmpDIB.SetAlphaPremultiplication True
        
        'Paint the final 32-bpp version image onto the splash screen
        tmpDIB.AlphaBlendToDC m_splashDIB.GetDIBDC, dstX:=pdLogoRight - fntWidth, dstY:=pdLogoBottom + Interface.FixDPI(8)
        
        'We now have a back buffer with everything the splash screen requires
        ' (except the progress bar, which will be drawn later).  Create a front buffer
        ' and copy the back buffer over to it.
        If (m_BackBuffer Is Nothing) Then Set m_BackBuffer = New pdDIB
        m_BackBuffer.CreateBlank m_splashDIB.GetDIBWidth, m_splashDIB.GetDIBHeight, 32, 0, 0
        m_BackBuffer.SetInitialAlphaPremultiplicationState True
        m_splashDIB.AlphaBlendToDC m_BackBuffer.GetDIBDC
        
        'Ensure the splash gets painted at least once prior to display
        UpdateLayeredWindowAPI
        
    Else
        PDDebug.LogAction "WARNING!  Splash DIBs could not be loaded; something may be catastrophically wrong."
    End If
    
End Sub

'When the load function updates the current progress count, we refresh the splash screen to reflect the new progress.
Public Sub UpdateLoadProgress(ByVal newProgressMarker As Long)
    
    If (m_splashDIB Is Nothing) Then Exit Sub
    
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
        
        'Erase any previous rendering by copying the static backbuffer over the front buffer
        GDI.BitBltWrapper m_BackBuffer.GetDIBDC, 0, 0, m_splashDIB.GetDIBWidth, m_splashDIB.GetDIBHeight, m_splashDIB.GetDIBDC, 0, 0, vbSrcCopy
        
        'Draw a progress line using GDI+
        Dim lineRadius As Long, lineY As Long
        lineRadius = Interface.FixDPI(6)
        lineY = m_splashDIB.GetDIBHeight - Interface.FixDPI(2) - lineRadius
        
        Dim cSurface As pd2DSurface
        Set cSurface = New pd2DSurface
        cSurface.WrapSurfaceAroundPDDIB m_BackBuffer
        cSurface.SetSurfaceAntialiasing P2_AA_HighQuality
        cSurface.SetSurfacePixelOffset P2_PO_Half
        
        Dim cPen As pd2DPen
        Drawing2D.QuickCreateSolidPen cPen, lineRadius, g_Themer.GetGenericUIColor(UI_Accent), 100#, , P2_LC_Round
        PD2D.DrawLineF cSurface, cPen, lineOffset, lineY, (m_splashDIB.GetDIBWidth - lineOffset) * ((newProgressMarker - m_progressAtFirstNotify) / (m_maxProgress - m_progressAtFirstNotify)), lineY
        
        Set cSurface = Nothing
        
        'Manually refresh the form
        UpdateLayeredWindowAPI
        
    End If

End Sub

Private Sub Form_Load()

    'Unfortunately, we have to subclass to prevent obnoxious flickering when the form is first displayed
    If PDMain.IsProgramRunning() And OS.IsProgramCompiled() Then VBHacks.StartSubclassing Me.hWnd, Me
    
    'Immediately after load, we need to change this to a layered window (for alpha handling)
    Const GWL_EXSTYLE As Long = -20
    Const WS_EX_LAYERED As Long = &H80000
    SetWindowLong Me.hWnd, GWL_EXSTYLE, GetWindowLong(Me.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    
End Sub

Private Sub UpdateLayeredWindowAPI()
    
    'Create a temporary blend function parameter; see https://docs.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-blendfunction
    ' (This is the equivalent of AlphaFormat = AC_SRC_ALPHA (1) and SourceConstantAlpha = 255)
    Dim bfParams As Long
    bfParams = &H1FF0000
    
    'MSDN says size is optional, but it really isn't; if this isn't supplied, the update request
    ' gets ignored on Win 10 (haven't tested other OSes).
    Dim srcSize(0 To 1) As Long
    srcSize(0) = m_BackBuffer.GetDIBWidth
    srcSize(1) = m_BackBuffer.GetDIBHeight
    
    'Also create a dummy (0, 0) point
    Dim srcPoint(0 To 1) As Long
    
    'Request full 32-bpp alpha as part of the update
    Const ULW_ALPHA As Long = &H2&
    UpdateLayeredWindow Me.hWnd, 0&, 0&, VarPtr(srcSize(0)), m_BackBuffer.GetDIBDC, VarPtr(srcPoint(0)), 0&, VarPtr(bfParams), ULW_ALPHA
    
End Sub

Private Function ISubclass_WindowMsg(ByVal hWnd As Long, ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long

    If (uiMsg = WM_ERASEBKGND) Then
        ISubclass_WindowMsg = 1
    ElseIf (uiMsg = WM_NCDESTROY) Then
        VBHacks.StopSubclassing hWnd, Me
        ISubclass_WindowMsg = VBHacks.DefaultSubclassProc(hWnd, uiMsg, wParam, lParam)
    Else
        ISubclass_WindowMsg = VBHacks.DefaultSubclassProc(hWnd, uiMsg, wParam, lParam)
    End If

End Function
