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
   Icon            =   "VBP_FormAbout.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   595
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   779
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.Timer tmrText 
      Enabled         =   0   'False
      Interval        =   17
      Left            =   360
      Top             =   8400
   End
   Begin VB.PictureBox picBuffer 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   4455
      Left            =   600
      ScaleHeight     =   297
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   721
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   10815
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   10095
      TabIndex        =   0
      Top             =   8280
      Width           =   1365
   End
   Begin VB.PictureBox picBackground 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8265
      Left            =   0
      Picture         =   "VBP_FormAbout.frx":000C
      ScaleHeight     =   551
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   779
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   11685
   End
End
Attribute VB_Name = "FormAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'About Form
'Copyright ©2001-2013 by Tanner Helland
'Created: 6/12/01
'Last updated: 31/May/13
'Last update: rewrote the scrolling credits against my new pdText object.  Performance is waaaay better!
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

Private Type pdCredit
    Name As String
    URL As String
End Type

Dim creditList() As pdCredit
Dim numOfCredits As Long

'The offset is incremented upward; this controls the credit scroll distance
Dim scrollOffset As Long

'Height of each credit content block
Private Const BLOCKHEIGHT As Long = 54

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

Dim backLayer As pdLayer
Dim bufferLayer As pdLayer
Dim m_BufferLeft As Long, m_BufferTop As Long
Dim m_BufferWidth As Long, m_BufferHeight As Long
Dim m_FormWidth As Long

'Two font objects; one for names and one for URLs.  (Two are needed because they have different sizes and colors.)
Dim firstFont As pdFont, secondFont As pdFont

Private Sub CmdOK_Click()
    tmrText.Enabled = False
    Unload Me
End Sub

Private Sub Form_Load()

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
    GenerateThankyou "PhotoDemon " & App.Major & "." & App.Minor & "." & App.Revision, "©2013 Tanner Helland"
    GenerateThankyou g_Language.TranslateMessage("a free, portable, powerful photo editor"), ""
    GenerateThankyou ""
    GenerateThankyou g_Language.TranslateMessage("PhotoDemon is the product of many talented contributors, including:"), ""
    GenerateThankyou "Adrian Pellas-Rice", "http://sourceforge.net/projects/pngnqs9/"
    GenerateThankyou "Alfred Hellmueller"
    GenerateThankyou "Andrew Yeoman"
    GenerateThankyou "Avery", "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=37541&lngWId=1"
    GenerateThankyou "Bernhard Stockmann", "http://www.gimpusers.com/tutorials/colorful-light-particle-stream-splash-screen-gimp.html"
    GenerateThankyou "Brad Martinez", "http://btmtz.mvps.org/gfxfromfrx/"
    GenerateThankyou "Carles P.V.", "http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=42376&lngWId=1"
    GenerateThankyou "chrfb @ deviantart.com", "http://chrfb.deviantart.com/art/quot-ecqlipse-2-quot-PNG-59941546"
    GenerateThankyou "Dave Jamison", "http://www.modeltrainsoftware.com/"
    GenerateThankyou "Dosadi", "http://eztwain.com/eztwain1.htm"
    GenerateThankyou "Everaldo Coelho", "http://www.everaldo.com/"
    GenerateThankyou "Frank Donckers"
    GenerateThankyou "FreeImage Project", "http://freeimage.sourceforge.net/"
    GenerateThankyou "Gilles Vollant", "http://www.winimage.com/zLibDll/index.html"
    GenerateThankyou "GioRock", ""
    GenerateThankyou "Jason Bullen", "http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=11488&lngWId=1"
    GenerateThankyou "Jerry Huxtable", "http://www.jhlabs.com/ie/index.html"
    GenerateThankyou "Juned Chhipa", "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=71482&lngWId=1"
    GenerateThankyou "Kroc Camen", "http://camendesign.com"
    GenerateThankyou "LaVolpe", "http://www.vbforums.com/showthread.php?t=606736"
    GenerateThankyou "Leandro Ascierto", "http://leandroascierto.com/blog/clsmenuimage/"
    GenerateThankyou "Manuel Augusto Santos", "http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=26303&lngWId=1"
    GenerateThankyou "Mark James", "http://www.famfamfam.com/lab/icons/silk/"
    GenerateThankyou "Phil Fresle", "http://www.frez.co.uk/vb6.aspx"
    GenerateThankyou "Phil Harvey", "http://www.sno.phy.queensu.ca/~phil/exiftool/"
    GenerateThankyou "Robert Rayment", "http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=66991&lngWId=1"
    GenerateThankyou "Rod Stephens", "http://www.vb-helper.com"
    GenerateThankyou "Steve McMahon", "http://www.vbaccelerator.com/home/VB/index.asp"
    GenerateThankyou "Tango Icon Library", "http://tango.freedesktop.org/"
    GenerateThankyou "Waty Thierry", "http://www.ppreview.net/"
    GenerateThankyou "Yusuke Kamiyamane", "http://p.yusukekamiyamane.com/"
    GenerateThankyou "Zhu JinYong"
    GenerateThankyou ""
    
    Dim extraString1 As String, extraString2 As String
    extraString1 = g_Language.TranslateMessage("PhotoDemon is released under an open-source BSD license")
    extraString2 = "for more information on licensing, visit tannerhelland.com/photodemon/#license"
    GenerateThankyou extraString1, extraString2
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
    GenerateThankyou "ExifTool", "http://dev.perl.org/licenses/"
    GenerateThankyou "EZTwain", "http://eztwain.com/ezt1faq.htm"
    GenerateThankyou "FreeImage", "http://freeimage.sourceforge.net/license.html"
    GenerateThankyou "pngnq-s9", "http://sourceforge.net/projects/pngnqs9/"
    GenerateThankyou "zLib", "http://www.zlib.net/zlib_license.html"
    GenerateThankyou "", ""
    GenerateThankyou "Thank you for using PhotoDemon"
    GenerateThankyou "tannerhelland.com/photodemon"
    
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
    'Initialize the background layer (this allows for faster blitting than a picture box)
    Set backLayer = New pdLayer
    backLayer.CreateFromPicture picBackground.Picture
    
    'The PicBuffer picture box is used only as a reference.  No drawing ever occurs on it - instead, we use our own buffer layer.
    Set bufferLayer = New pdLayer
    bufferLayer.createBlank picBuffer.ScaleWidth, picBuffer.ScaleHeight
    
    'Initialize a few other variables for speed reasons
    m_BufferLeft = picBuffer.Left
    m_BufferTop = picBuffer.Top
    m_BufferWidth = picBuffer.ScaleWidth
    m_BufferHeight = picBuffer.ScaleHeight
    m_FormWidth = Me.ScaleWidth
    
    'Initialize a custom font objects for names
    Set firstFont = New pdFont
    firstFont.setFontColor RGB(255, 255, 255)
    firstFont.setFontBold True
    firstFont.setFontSize 14
    firstFont.createFontObject
    firstFont.setTextAlignment vbRightJustify
    
    '...and a second custom font object for URLs
    Set secondFont = New pdFont
    secondFont.setFontColor RGB(192, 192, 192)
    secondFont.setFontBold False
    secondFont.setFontSize 10
    secondFont.createFontObject
    secondFont.setTextAlignment vbRightJustify
    
    'Render the primary background image to the form
    StretchBlt Me.hDC, 0, 0, m_FormWidth, backLayer.getLayerHeight, backLayer.getLayerDC, 0, 0, backLayer.getLayerWidth, backLayer.getLayerHeight, vbSrcCopy
    Me.Picture = Me.Image
    
    'Start the credit scroll timer
    tmrText.Enabled = True
    
End Sub

'Generate a label with the specified "thank you" text, and link it to the specified URL
Private Sub GenerateThankyou(ByVal thxText As String, Optional ByVal creditURL As String = "")
    
    creditList(numOfCredits).Name = thxText
    creditList(numOfCredits).URL = creditURL
    
    numOfCredits = numOfCredits + 1
    ReDim Preserve creditList(0 To numOfCredits) As pdCredit
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Scroll the credit list; nothing fancy here, just a basic credit scroller, using a modified version of the
' scrolling code I wrote for the metadata browser.
Private Sub tmrText_Timer()
    
    scrollOffset = scrollOffset + 1
    If scrollOffset > (numOfCredits * BLOCKHEIGHT) Then scrollOffset = 0
    
    'Erase the buffer by copying over a chunk of the background image (onto which we will render the text)
    BitBlt bufferLayer.getLayerDC, 0, 0, m_BufferWidth, m_BufferHeight, backLayer.getLayerDC, m_BufferLeft, m_BufferTop, vbSrcCopy
    
    'Render all text
    Dim i As Long
    For i = 0 To numOfCredits - 1
        renderCredit i, 8, i * BLOCKHEIGHT - scrollOffset - 2
    Next i
    
    'Copy the buffer to the main form and refresh it
    BitBlt Me.hDC, m_BufferLeft, m_BufferTop, m_BufferWidth, m_BufferHeight, bufferLayer.getLayerDC, 0, 0, vbSrcCopy
    Me.Picture = Me.Image
    
End Sub

'Render the given metadata index onto the background picture box at the specified offset.  Custom font objects are used for better performance.
Private Sub renderCredit(ByVal blockIndex As Long, ByVal offsetX As Long, ByVal offsetY As Long)

    'Only draw the current block if it will be visible
    If ((offsetY + BLOCKHEIGHT) > 0) And (offsetY < m_BufferHeight) Then
    
        Dim linePadding As Long
        linePadding = 1
    
        Dim mWidth As Single, mHeight As Single
        
        Dim drawString As String
        drawString = creditList(blockIndex).Name
        
        'Render the "name" field
        firstFont.attachToDC bufferLayer.getLayerDC
        firstFont.fastRenderText m_BufferWidth - offsetX, offsetY, drawString
                
        'Below the name, add the URL (or other description)
        mHeight = firstFont.getHeightOfString(drawString) + linePadding
        drawString = creditList(blockIndex).URL
        
        secondFont.attachToDC bufferLayer.getLayerDC
        secondFont.fastRenderText m_BufferWidth - offsetX, offsetY + mHeight, drawString
        
    End If

End Sub
