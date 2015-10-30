VERSION 5.00
Begin VB.Form FormAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " About PhotoDemon"
   ClientHeight    =   8925
   ClientLeft      =   2340
   ClientTop       =   1875
   ClientWidth     =   11685
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   595
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   779
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.pdButton cmdOK 
      Height          =   495
      Left            =   9840
      TabIndex        =   4
      Top             =   8280
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      Caption         =   "&OK"
   End
   Begin PhotoDemon.pdButtonToolbox cmdSpeed 
      Height          =   510
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   8280
      Width           =   510
      _ExtentX        =   1296
      _ExtentY        =   661
      BackColor       =   0
      AutoToggle      =   -1  'True
   End
   Begin VB.PictureBox picBuffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   600
      ScaleHeight     =   369
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   697
      TabIndex        =   0
      Top             =   600
      Width           =   10455
   End
   Begin VB.Timer tmrText 
      Enabled         =   0   'False
      Interval        =   17
      Left            =   2520
      Top             =   8280
   End
   Begin PhotoDemon.pdButtonToolbox cmdSpeed 
      Height          =   510
      Index           =   1
      Left            =   720
      TabIndex        =   2
      Top             =   8280
      Width           =   510
      _ExtentX        =   1296
      _ExtentY        =   661
      BackColor       =   0
      AutoToggle      =   -1  'True
   End
   Begin PhotoDemon.pdButtonToolbox cmdSpeed 
      Height          =   510
      Index           =   2
      Left            =   1320
      TabIndex        =   3
      Top             =   8280
      Width           =   510
      _ExtentX        =   1296
      _ExtentY        =   661
      BackColor       =   0
      StickyToggle    =   -1  'True
   End
End
Attribute VB_Name = "FormAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon About Dialog
'Copyright 2001-2015 by Tanner Helland
'Created: 6/12/01
'Last updated: 31/May/14
'Last update: convert extra mouse handling code to pdInput
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private Type pdCredit
    Name As String
    URL As String
    Clickable As Boolean
End Type

Private creditList() As pdCredit
Private numOfCredits As Long

'Number of pixels to scroll the list on each timer tick
Private m_pxScroll As Long

'The offset is incremented upward; this controls the credit scroll distance
Private scrollOffset As Double

'Height of each credit content block
Private Const BLOCKHEIGHT As Long = 54

Private backDIB As pdDIB
Private bufferDIB As pdDIB
Private m_BufferWidth As Long, m_BufferHeight As Long
Private m_FormWidth As Long

Private logoDIB As pdDIB, maskDIB As pdDIB

'Two font objects; one for names and one for URLs.  (Two are needed because they have different sizes and colors.)
Private firstFont As pdFont, secondFont As pdFont

'...and another font object for highlighted text (when URLs are hovered)
Private highlightFont As pdFont

'Current mouse position; to make the URLs clickable, we track the current mouse position and highlight the relevant credit
Private mouseX As Long, mouseY As Long

'Currently hovered credit (if any)
Private curHoveredCredit As Long
Private inHoverState As Boolean

'As the credit list is now clickable, we display "click to visit" with the current entry
Private clickToVisitText As String

'An outside class provides access to specialized mouse events (mouse enter/leave, in this case)
Private WithEvents cMouseEvents As pdInputMouse
Attribute cMouseEvents.VB_VarHelpID = -1

'When the mouse moves over something clickable, update the pointer and stop the timer
Private Sub updateHoverState(ByVal isSomethingUsefulHovered As Boolean)

    If isSomethingUsefulHovered Then
        
        'If we are already in hover state, disregard this command
        If Not inHoverState Then
            
            'Slow the scrolling (to simplify clicking)
            tmrText.Interval = 50
            
            'Display a hand cursor
            cMouseEvents.setSystemCursor IDC_HAND
            
            'Mark the new hover state
            inHoverState = True
            
        Else
            cMouseEvents.setSystemCursor IDC_HAND
        End If
        
    Else
        
        If inHoverState Then
        
            'Return scrolling to normal speed
            tmrText.Interval = 17
            
            'Restore an arrow cursor
            cMouseEvents.setSystemCursor IDC_ARROW
            
            'Mark the new hover state
            inHoverState = False
        
        Else
            If Not (cMouseEvents Is Nothing) Then cMouseEvents.setSystemCursor IDC_ARROW
        End If
        
    End If
    
End Sub

Private Sub CmdOK_Click()
    tmrText.Enabled = False
    Unload Me
End Sub

Private Sub cmdSpeed_Click(Index As Integer)
    
    'Reduce speed
    If Index = 0 Then
    
        m_pxScroll = m_pxScroll - 1
        If m_pxScroll = 1 Then cmdSpeed(0).Enabled = False
        If (m_pxScroll < 5) And (Not cmdSpeed(1).Enabled) Then cmdSpeed(1).Enabled = True
    
    'Increase speed
    ElseIf Index = 1 Then
    
        m_pxScroll = m_pxScroll + 1
        If m_pxScroll = 5 Then cmdSpeed(1).Enabled = False
        If (m_pxScroll > 1) And (Not cmdSpeed(0).Enabled) Then cmdSpeed(0).Enabled = True
    
    'Pause
    ElseIf Index = 2 Then
    
        tmrText.Enabled = Not cmdSpeed(2).Value
    
    End If
    
End Sub

Private Sub cMouseEvents_ClickCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    If curHoveredCredit >= 0 Then OpenURL creditList(curHoveredCredit).URL
End Sub

Private Sub cMouseEvents_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    mouseX = -1
    mouseY = -1
    curHoveredCredit = -1
    updateHoverState False
End Sub

Private Sub cMouseEvents_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
        
    mouseX = x
    mouseY = y
    
    If Not tmrText.Enabled Then renderFullCreditList
    
End Sub

Private Sub Form_Load()

    'By default, the image will scroll by 1 px each timer interval
    m_pxScroll = 1
    cmdSpeed(0).Enabled = False
    cmdSpeed(1).Enabled = True
    cmdSpeed(2).Value = False
    
    'Load button images
    cmdSpeed(0).AssignImage "ABOUT_MINUS"
    cmdSpeed(1).AssignImage "ABOUT_PLUS"
    cmdSpeed(2).AssignImage "ABOUT_PAUSE"
    
    'Load button tooltips
    cmdSpeed(0).AssignTooltip "Decrease speed"
    cmdSpeed(1).AssignTooltip "Increase speed"
    cmdSpeed(2).AssignTooltip "Pause scrolling"
    
    'Reset the mouse coordinates and currently hovered entry
    mouseX = -1
    mouseY = -1
    curHoveredCredit = -1
    updateHoverState False
    
    'Translate "click to visit" and cache it to improve performance
    clickToVisitText = "(" & g_Language.TranslateMessage("click to visit") & ") "
    
    'Enable mouse subclassing for the main buffer box, which allows us to track when the mouse leaves
    Set cMouseEvents = New pdInputMouse
    cMouseEvents.addInputTracker picBuffer.hWnd, True, True, , True
    cMouseEvents.setSystemCursor IDC_ARROW

    'Load the logo from the resource file
    Set logoDIB = New pdDIB
    loadResourceToDIB "PDLOGONOTEXT", logoDIB
    
    'Load the logo mask from the resource file into a temporary DIB
    Dim tmpMaskDIB As pdDIB
    Set tmpMaskDIB = New pdDIB
    loadResourceToDIB "PDLOGOMASK", tmpMaskDIB
    
    scrollOffset = 0

    ReDim creditList(0) As pdCredit

    numOfCredits = 0
    
    'Shout-outs to other designers, programmers, testers and sponsors who provided various resources
    GenerateThankyou ""
    GenerateThankyou ""
    GenerateThankyou ""
    GenerateThankyou ""
    GenerateThankyou ""
    GenerateThankyou ""
    GenerateThankyou ""
    GenerateThankyou ""
    GenerateThankyou ""
    GenerateThankyou ""
    GenerateThankyou getPhotoDemonNameAndVersion(), "Copyright 2015 Tanner Helland and contributors"
    GenerateThankyou g_Language.TranslateMessage("the fast, free, portable photo editor"), ""
    GenerateThankyou ""
    GenerateThankyou g_Language.TranslateMessage("PhotoDemon is the product of many talented contributors, including:"), ""
    GenerateThankyou "Abhijit Mhapsekar"
    GenerateThankyou "A.G. Violette", "http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=55938&lngWId=1", True
    GenerateThankyou "Allan Lima"
    GenerateThankyou "Andrew Yeoman"
    GenerateThankyou "Ari Sohandri Putra", "http://arisohandrip.indonesiaz.com/", True
    GenerateThankyou "Avery", "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=37541&lngWId=1", True
    GenerateThankyou "Audioglider", "https://github.com/audioglider", True
    GenerateThankyou "Bernhard Stockmann", "http://www.gimpusers.com/tutorials/colorful-light-particle-stream-splash-screen-gimp.html", True
    GenerateThankyou "Boban Gjerasimoski", "https://www.behance.net/Boban_Gjerasimoski", True
    GenerateThankyou "Bonnie West", "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=74264&lngWId=1", True
    GenerateThankyou "Carles P.V.", "http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=42376&lngWId=1", True
    GenerateThankyou "chrfb @ deviantart.com", "http://chrfb.deviantart.com/art/quot-ecqlipse-2-quot-PNG-59941546", True
    GenerateThankyou "Dana Seaman", "http://www.cyberactivex.com/", True
    GenerateThankyou "dilettante", "http://www.vbforums.com/showthread.php?660014-VB6-ShellPipe-quot-Shell-with-I-O-Redirection-quot-control", True
    GenerateThankyou "Djordje Djoric", "https://www.odesk.com/o/profiles/users/_~0181c1599705edab79/", True
    GenerateThankyou "Dirk Hartmann", "http://www.taichi-zentrum-heidelberg.de", True
    GenerateThankyou "Dosadi", "http://eztwain.com/eztwain1.htm", True
    GenerateThankyou "Easy RGB", "http://www.easyrgb.com/", True
    GenerateThankyou "Everaldo Coelho", "http://www.everaldo.com/", True
    GenerateThankyou "Frank Donckers", "http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?lngWId=1&txtCriteria=donckers", True
    GenerateThankyou "Frans van Beers", "https://plus.google.com/+FransvanBeers/", True
    GenerateThankyou "FreeImage Project", "http://freeimage.sourceforge.net/", True
    GenerateThankyou "Gilles Vollant", "http://www.winimage.com/zLibDll/index.html", True
    GenerateThankyou "GioRock", "http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?lngWId=1&txtCriteria=giorock", True
    GenerateThankyou "Google Translate", "http://translate.google.com", True
    GenerateThankyou "Hans Nolte", "https://github.com/hansnolte", True
    GenerateThankyou "Helmut Kuerbiss"
    GenerateThankyou "Jason Bullen", "http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=11488&lngWId=1", True
    GenerateThankyou "Jerry Huxtable", "http://www.jhlabs.com/ie/index.html", True
    GenerateThankyou "Kroc Camen", "http://camendesign.com", True
    GenerateThankyou "LaVolpe", "http://www.vbforums.com/showthread.php?t=606736", True
    GenerateThankyou "Leandro Ascierto", "http://leandroascierto.com/blog/clsmenuimage/", True
    GenerateThankyou "Manuel Augusto Santos", "http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=26303&lngWId=1", True
    GenerateThankyou "Mark James", "http://www.famfamfam.com/lab/icons/silk/", True
    GenerateThankyou "Mohammad Reza Karimi"
    GenerateThankyou "Nguyen Van Hung"
    GenerateThankyou "Olaf Schmidt", "http://www.vbrichclient.com/#/en/About/", True
    GenerateThankyou "Paul Bourke", "http://paulbourke.net/miscellaneous/", True
    GenerateThankyou "Peter Burn"
    GenerateThankyou "Phil Fresle", "http://www.frez.co.uk/vb6.aspx", True
    GenerateThankyou "Phil Harvey", "http://www.sno.phy.queensu.ca/~phil/exiftool/", True
    GenerateThankyou "Plinio C Garcia"
    GenerateThankyou "PortableFreeware.com team", "http://www.portablefreeware.com/forums/viewtopic.php?t=21652", True
    GenerateThankyou "Roy (rk)"
    GenerateThankyou "Raj Chaudhuri", "https://github.com/rajch", True
    GenerateThankyou "Robert Rayment", "http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=66991&lngWId=1", True
    GenerateThankyou "Steve McMahon", "http://www.vbaccelerator.com/home/VB/index.asp", True
    GenerateThankyou "Tango Icon Library", "http://tango.freedesktop.org/", True
    GenerateThankyou "Tom Loos", "http://www.designedbyinstinct.com", True
    GenerateThankyou "Will Stampfer", "https://github.com/epmatsw", True
    GenerateThankyou "Yusuke Kamiyamane", "http://p.yusukekamiyamane.com/", True
    GenerateThankyou "Zhu JinYong", "http://www.planetsourcecode.com/vb/authors/ShowBio.asp?lngAuthorId=2211529461&lngWId=1", True
    GenerateThankyou ""
    
    Dim extraString1 As String, extraString2 As String
    extraString1 = g_Language.TranslateMessage("PhotoDemon is released under an open-source BSD license")
    GenerateThankyou extraString1
    extraString1 = g_Language.TranslateMessage("For more information on licensing, please visit")
    GenerateThankyou extraString1, "http://photodemon.org/about/license/", True
    GenerateThankyou ""
    extraString1 = g_Language.TranslateMessage("Please note that PhotoDemon uses several third-party plugins")
    GenerateThankyou extraString1
    GenerateThankyou ""
    extraString1 = g_Language.TranslateMessage("These plugins are also free and open source...")
    extraString2 = g_Language.TranslateMessage("...but they are governed by their own licenses, separate from PhotoDemon")
    GenerateThankyou extraString1, extraString2
    GenerateThankyou ""
    extraString1 = g_Language.TranslateMessage("For more information on plugin licensing, please visit:")
    GenerateThankyou extraString1
    GenerateThankyou "ExifTool", "http://dev.perl.org/licenses/", True
    GenerateThankyou "EZTwain", "http://eztwain.com/ezt1faq.htm", True
    GenerateThankyou "FreeImage", "http://freeimage.sourceforge.net/license.html", True
    GenerateThankyou "PNGQuant", "http://pngquant.org/#source", True
    GenerateThankyou "zLib", "http://www.zlib.net/zlib_license.html", True
    GenerateThankyou ""
    GenerateThankyou g_Language.TranslateMessage("Thank you for using PhotoDemon"), "http://photodemon.org", True
    
    'Apply translations and visual themes
    MakeFormPretty Me
    
    'Initialize the background DIB (this allows for faster blitting than a picture box)
    ' Note that this DIB is dynamically resized; this solves issues with high-DPI screens
    Set backDIB = New pdDIB
    Dim logoAspectRatio As Double
    logoAspectRatio = CDbl(logoDIB.getDIBWidth) / CDbl(logoDIB.getDIBHeight)
    backDIB.createFromExistingDIB logoDIB, Me.ScaleWidth, Me.ScaleWidth / logoAspectRatio
    
    'Copy the resized logo into the logo DIB.  (We don't want to resize it every time we need it.)
    logoDIB.eraseDIB
    logoDIB.createFromExistingDIB backDIB
    
    'Create a mask DIB at the same size.
    Set maskDIB = New pdDIB
    maskDIB.createFromExistingDIB tmpMaskDIB, backDIB.getDIBWidth, backDIB.getDIBHeight, False
    tmpMaskDIB.eraseDIB
    Set tmpMaskDIB = Nothing
    
    'In order to fix high-DPI screen issues, resize the buffer at run-time.  (Why not blit directly to the form?  Because
    ' the OK command button will flicker.  Instead, we just draw to a picture box sized to match the form.)
    picBuffer.Move 0, 0, backDIB.getDIBWidth, backDIB.getDIBHeight
    
    'Remember that the PicBuffer picture box is used only as a placeholder.  We render everything manually to an
    ' off-screen buffer, then flip that buffer to the picture box after all rendering is complete.
    Set bufferDIB = New pdDIB
    bufferDIB.createBlank backDIB.getDIBWidth, backDIB.getDIBHeight, 24, 0
    
    'Initialize a few other variables for speed reasons
    m_BufferWidth = backDIB.getDIBWidth
    m_BufferHeight = backDIB.getDIBHeight
    m_FormWidth = Me.ScaleWidth
    
    'Initialize a custom font objects for names
    Set firstFont = New pdFont
    firstFont.SetFontColor RGB(255, 255, 255)
    firstFont.SetFontBold True
    firstFont.SetFontSize 14
    firstFont.CreateFontObject
    firstFont.SetTextAlignment vbRightJustify
    
    '...and a second custom font object for URLs
    Set secondFont = New pdFont
    secondFont.SetFontColor RGB(192, 192, 192)
    secondFont.SetFontBold False
    secondFont.SetFontSize 10
    secondFont.CreateFontObject
    secondFont.SetTextAlignment vbRightJustify
    
    '...and a third custom font object for highlighted text
    Set highlightFont = New pdFont
    highlightFont.SetFontColor ConvertSystemColor(vbHighlight)
    highlightFont.SetFontBold False
    highlightFont.SetFontSize 10
    highlightFont.SetFontUnderline True
    highlightFont.CreateFontObject
    highlightFont.SetTextAlignment vbRightJustify
    
    'Render the primary background image to the form
    BitBlt picBuffer.hDC, 0, 0, picBuffer.ScaleWidth, picBuffer.ScaleHeight, logoDIB.getDIBDC, 0, 0, vbSrcCopy
    picBuffer.Picture = picBuffer.Image
    picBuffer.Refresh
    
    'Start the credit scroll timer
    tmrText.Enabled = True
    
End Sub

'Generate a label with the specified "thank you" text, and link it to the specified URL
Private Sub GenerateThankyou(ByVal thxText As String, Optional ByVal creditURL As String = "", Optional ByVal isClickable As Boolean = False)
    
    creditList(numOfCredits).Name = thxText
    creditList(numOfCredits).URL = creditURL
    creditList(numOfCredits).Clickable = isClickable
    
    numOfCredits = numOfCredits + 1
    ReDim Preserve creditList(0 To numOfCredits) As pdCredit
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Scroll the credit list; nothing fancy here, just a basic credit scroller, using a modified version of the
' scrolling code I wrote for the metadata browser.
Private Sub tmrText_Timer()
    
    scrollOffset = scrollOffset + FixDPIFloat(m_pxScroll)
    If scrollOffset > (numOfCredits * BLOCKHEIGHT) Then scrollOffset = 0
    
    renderFullCreditList
    
End Sub

Private Sub renderFullCreditList()

    'Erase the back DIB by copying over the logo (onto which we will render the text)
    BitBlt backDIB.getDIBDC, 0, 0, m_BufferWidth, m_BufferHeight, logoDIB.getDIBDC, 0, 0, vbSrcCopy
        
    'Render all text
    Dim i As Long
    For i = 0 To numOfCredits - 1
        renderCredit i, FixDPI(8), FixDPI(i * BLOCKHEIGHT) - scrollOffset - FixDPIFloat(2)
    Next i
    
    'The back DIB now contains the credit text drawn over the program logo.
    
    'Black out the section of the back DIB where the base text appears - we don't want text rendering over
    ' the top of this section.
    BitBlt backDIB.getDIBDC, 0, 0, m_BufferWidth, m_BufferHeight, maskDIB.getDIBDC, 0, 0, vbMergePaint
    
    'Blit a blank copy of the logo to the buffer DIB
    BitBlt bufferDIB.getDIBDC, 0, 0, m_BufferWidth, m_BufferHeight, logoDIB.getDIBDC, 0, 0, vbSrcCopy
    
    'Blit the logo mask over the top
    BitBlt bufferDIB.getDIBDC, 0, 0, m_BufferWidth, m_BufferHeight, maskDIB.getDIBDC, 0, 0, vbSrcPaint
    
    'Blit the back DIB, with the text, over the top of the buffer
    BitBlt bufferDIB.getDIBDC, 0, 0, m_BufferWidth, m_BufferHeight, backDIB.getDIBDC, 0, 0, vbSrcAnd
    
    'Copy the buffer to the main form and refresh it
    BitBlt picBuffer.hDC, 0, 0, m_BufferWidth, m_BufferHeight, bufferDIB.getDIBDC, 0, 0, vbSrcCopy
    picBuffer.Picture = picBuffer.Image
    picBuffer.Refresh

End Sub

'Render the given metadata index onto the background picture box at the specified offset.  Custom font objects are used for better performance.
Private Sub renderCredit(ByVal blockIndex As Long, ByVal offsetX As Long, ByVal offsetY As Long)

    'Only draw the current block if it will be visible
    If ((offsetY + FixDPI(BLOCKHEIGHT)) > 0) And (offsetY < m_BufferHeight - 40) Then
    
        'Check to see if the current credit block is highlighted
        Dim isHovered As Boolean
        
        'If this entry is clickable, compare it to the current mouse position
        If (mouseX >= 0) And (mouseX < m_BufferWidth) And (mouseY >= offsetY) And (mouseY < offsetY + BLOCKHEIGHT) Then
            
            'Ignore unclickable entries
            If creditList(blockIndex).Clickable Then
                isHovered = True
                curHoveredCredit = blockIndex
                updateHoverState True
            Else
                isHovered = False
                curHoveredCredit = -1
                updateHoverState False
            End If
            
        Else
            isHovered = False
        End If
                
        Dim linePadding As Long
        linePadding = 1
    
        Dim mHeight As Single
        
        Dim drawString As String
        drawString = creditList(blockIndex).Name
        
        'If this entry is hovered, append "click to visit" to the name
        If isHovered Then drawString = clickToVisitText & drawString
        
        'Render the "name" field
        firstFont.AttachToDC backDIB.getDIBDC
        firstFont.FastRenderText m_BufferWidth - offsetX, offsetY, drawString
                
        'Below the name, add the URL (or other description)
        mHeight = firstFont.GetHeightOfString(drawString) + linePadding
        drawString = creditList(blockIndex).URL
        
        'Detach the first font from the DC
        firstFont.ReleaseFromDC
        
        If isHovered Then
            highlightFont.AttachToDC backDIB.getDIBDC
            highlightFont.FastRenderText m_BufferWidth - offsetX, offsetY + mHeight, drawString
            highlightFont.ReleaseFromDC
        Else
            secondFont.AttachToDC backDIB.getDIBDC
            secondFont.FastRenderText m_BufferWidth - offsetX, offsetY + mHeight, drawString
            secondFont.ReleaseFromDC
        End If
        
        'If the user's mouse is over the current block, highlight the block
        If isHovered Then
        
            Dim tmpRect As RECTL, hBrush As Long
            SetRect tmpRect, offsetX, offsetY, m_BufferWidth, offsetY + FixDPI(BLOCKHEIGHT)
            hBrush = CreateSolidBrush(ConvertSystemColor(vbHighlight))
            FrameRect backDIB.getDIBDC, tmpRect, hBrush
            DeleteObject hBrush
        
        End If
        
    End If

End Sub
