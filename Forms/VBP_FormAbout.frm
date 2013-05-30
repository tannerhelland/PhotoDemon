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
      AutoRedraw      =   -1  'True
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
'Last updated: 30/May/13
'Last update: redesigned the dialog from the ground up.  Also, scrolling credits!
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

'Height of each credit content block
Private Const BLOCKHEIGHT As Long = 54

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

Private Sub CmdOK_Click()
    tmrText.Enabled = False
    Unload Me
End Sub

Private Sub Form_Load()

    ReDim creditList(0) As pdCredit

    'Make the invisible buffer's font match the rest of PD
    If g_UseFancyFonts Then
        picBuffer.FontName = "Segoe UI"
    Else
        picBuffer.FontName = "Tahoma"
    End If

    numOfCredits = 0
    
    'Shout-outs to other designers, programmers, testers and sponsors who provided various resources
    GenerateThankyou "", ""
    GenerateThankyou "", ""
    GenerateThankyou "", ""
    GenerateThankyou "", ""
    GenerateThankyou "", ""
    GenerateThankyou "", ""
    GenerateThankyou "PhotoDemon " & App.Major & "." & App.Minor & "." & App.Revision, "©2013 Tanner Helland"
    GenerateThankyou "", ""
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
    GenerateThankyou "", ""
    
    Dim extraString1 As String, extraString2 As String
    extraString1 = g_Language.TranslateMessage("PhotoDemon uses several third-party plugins")
    extraString2 = g_Language.TranslateMessage("These plugins may be governed by additional licenses, specifically:")
    GenerateThankyou extraString1, extraString2
    GenerateThankyou "ExifTool", "http://dev.perl.org/licenses/"
    GenerateThankyou "EZTwain", "http://eztwain.com/ezt1faq.htm"
    GenerateThankyou "FreeImage", "http://freeimage.sourceforge.net/license.html"
    GenerateThankyou "pngnq-s9", "http://sourceforge.net/projects/pngnqs9/"
    GenerateThankyou "zLib", "http://www.zlib.net/zlib_license.html"
    GenerateThankyou "", ""
    GenerateThankyou "Thank you for using PhotoDemon!"
    GenerateThankyou "tannerhelland.com/photodemon"
    
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
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
        
    Static scrollOffset As Long
    scrollOffset = scrollOffset + 1
    If scrollOffset > (numOfCredits * BLOCKHEIGHT) Then scrollOffset = 0
        
    picBuffer.PaintPicture picBackground.Picture, 0, 0, picBuffer.ScaleWidth, picBuffer.ScaleHeight, picBuffer.Left, picBuffer.Top, picBuffer.ScaleWidth, picBuffer.ScaleHeight, vbSrcCopy
    
    Dim i As Long
    For i = 0 To numOfCredits - 1
        renderCredit i, 8, i * BLOCKHEIGHT - scrollOffset - 2
    Next i
    
    'Copy the buffer to the main form
    Me.PaintPicture picBackground.Picture, 0, 0, Me.ScaleWidth, picBackground.ScaleHeight, 0, 0, picBackground.ScaleWidth, picBackground.ScaleHeight, vbSrcCopy
    picBuffer.Picture = picBuffer.Image
    Me.PaintPicture picBuffer.Picture, picBuffer.Left, picBuffer.Top, picBuffer.ScaleWidth, picBuffer.ScaleHeight, 0, 0, picBuffer.ScaleWidth, picBuffer.ScaleHeight, vbSrcCopy
    Me.Picture = Me.Image
    
End Sub

'Render the given metadata index onto the background picture box at the specified offset
Private Sub renderCredit(ByVal blockIndex As Long, ByVal offsetX As Long, ByVal offsetY As Long)

    'Only draw the current block if it will be visible
    If ((offsetY + BLOCKHEIGHT) > 0) And (offsetY < picBuffer.Height) Then
    
        Dim primaryColor As Long, secondaryColor As Long, tertiaryColor As Long
        primaryColor = RGB(255, 255, 255)
        secondaryColor = RGB(192, 192, 192)
        tertiaryColor = RGB(255, 255, 255)
    
        Dim linePadding As Long
        linePadding = 1
    
        Dim mWidth As Single, mHeight As Single
        
        Dim drawString As String
        drawString = creditList(blockIndex).Name
        picBuffer.FontSize = 14
        picBuffer.FontBold = True
        
        'Render the "name" field
        drawTextOnObject picBuffer, drawString, picBuffer.ScaleWidth - picBuffer.TextWidth(drawString) - offsetX, offsetY + 0, 14, primaryColor, True, False
                
        'Below the name, add the URL (or other description)
        mHeight = picBuffer.TextHeight(drawString) + linePadding
        
        drawString = creditList(blockIndex).URL
        
        picBuffer.FontSize = 10
        picBuffer.FontBold = False
        
        drawTextOnObject picBuffer, drawString, picBuffer.ScaleWidth - picBuffer.TextWidth(drawString) - offsetX, offsetY + mHeight, 10, secondaryColor, False
        
        'Draw a divider line near the bottom of the block
        'Dim lineY As Long
        'If blockIndex < mdCategories(blockCategory).Count - 1 Then
        '    lineY = offsetY + BLOCKHEIGHT - 8
        '    picBuffer.Line (4, lineY)-(picBuffer.ScaleWidth - 8, lineY), tertiaryColor
        'End If
        
    End If

End Sub
