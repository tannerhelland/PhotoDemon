VERSION 5.00
Begin VB.Form FormAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " About PhotoDemon"
   ClientHeight    =   7980
   ClientLeft      =   2340
   ClientTop       =   1875
   ClientWidth     =   9870
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   532
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   658
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.pdCommandBarMini cmdBar 
      Align           =   2  'Align Bottom
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   7365
      Width           =   9870
      _ExtentX        =   17410
      _ExtentY        =   1085
   End
   Begin PhotoDemon.pdButtonStrip btsPanel 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   1085
   End
   Begin PhotoDemon.pdContainer pnlAbout 
      Height          =   6375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   11245
      Begin PhotoDemon.pdHyperlink hypAbout 
         Height          =   375
         Index           =   0
         Left            =   240
         Top             =   1440
         Width           =   9300
         _ExtentX        =   16404
         _ExtentY        =   661
         Caption         =   "Donate to PhotoDemon development"
         URL             =   "http://photodemon.org/donate/"
      End
      Begin PhotoDemon.pdLabel lblAbout 
         Height          =   375
         Index           =   0
         Left            =   240
         Top             =   5880
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   661
         Caption         =   ""
      End
      Begin PhotoDemon.pdLabel lblAbout 
         Height          =   495
         Index           =   1
         Left            =   120
         Top             =   240
         Width           =   9300
         _ExtentX        =   16404
         _ExtentY        =   873
         Caption         =   "PhotoDemon"
         FontBold        =   -1  'True
         FontSize        =   14
      End
      Begin PhotoDemon.pdLabel lblAbout 
         Height          =   495
         Index           =   2
         Left            =   120
         Top             =   720
         Width           =   9300
         _ExtentX        =   16404
         _ExtentY        =   873
         Caption         =   "the fast, free, portable photo editor"
         FontSize        =   12
      End
      Begin PhotoDemon.pdHyperlink hypAbout 
         Height          =   375
         Index           =   1
         Left            =   240
         Top             =   1920
         Width           =   9300
         _ExtentX        =   16404
         _ExtentY        =   661
         Caption         =   "Participate in development, design, or translation work"
         URL             =   "http://photodemon.org/get-involved/"
      End
      Begin PhotoDemon.pdHyperlink hypAbout 
         Height          =   375
         Index           =   2
         Left            =   240
         Top             =   3840
         Width           =   9300
         _ExtentX        =   16404
         _ExtentY        =   661
         Caption         =   "Contact the author"
         URL             =   "http://photodemon.org/about/contact/"
      End
      Begin PhotoDemon.pdHyperlink hypAbout 
         Height          =   375
         Index           =   3
         Left            =   240
         Top             =   2400
         Width           =   9300
         _ExtentX        =   16404
         _ExtentY        =   661
         Caption         =   "Download program source code"
         URL             =   "https://github.com/tannerhelland/PhotoDemon"
      End
      Begin PhotoDemon.pdHyperlink hypAbout 
         Height          =   375
         Index           =   4
         Left            =   240
         Top             =   2880
         Width           =   9300
         _ExtentX        =   16404
         _ExtentY        =   661
         Caption         =   "Review program and plugin licenses"
         URL             =   "http://photodemon.org/about/license/"
      End
      Begin PhotoDemon.pdHyperlink hypAbout 
         Height          =   375
         Index           =   5
         Left            =   240
         Top             =   3360
         Width           =   9300
         _ExtentX        =   16404
         _ExtentY        =   661
         Caption         =   "Latest news"
         URL             =   "http://photodemon.org/blog/"
      End
   End
   Begin PhotoDemon.pdContainer pnlAbout 
      Height          =   6375
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   11245
      Begin PhotoDemon.pdListBoxOD lstContributors 
         Height          =   6135
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   9405
         _ExtentX        =   16589
         _ExtentY        =   10821
         BorderlessMode  =   -1  'True
      End
   End
End
Attribute VB_Name = "FormAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon About Dialog
'Copyright 2001-2018 by Tanner Helland
'Created: 6/12/01
'Last updated: 14/June/17
'Last update: update contributor list
'
'PhotoDemon would not be possible without the help of many, many amazing people.  THANK YOU!
'
'If you contributed to PhotoDemon's development in some way, but your name isn't listed here,
' please let me know!  I can always be reached via http://photodemon.org/about/contact/
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private Type PD_Contributor
    ctbName As String
    ctbURL As String
End Type

Private m_contributorList() As PD_Contributor
Private m_numOfContributors As Long

'Height of each credit content block
Private Const BLOCKHEIGHT As Long = 24

Private Sub btsPanel_Click(ByVal buttonIndex As Long)
    UpdateVisiblePanel
End Sub

Private Sub UpdateVisiblePanel()
    Dim i As Long
    For i = 0 To btsPanel.ListCount - 1
        pnlAbout(i).Visible = (i = btsPanel.ListIndex)
    Next i
End Sub

Private Sub Form_Load()
    
    lstContributors.ListItemHeight = FixDPI(BLOCKHEIGHT)
    
    btsPanel.AddItem "About", 0
    btsPanel.AddItem "Contributors", 1
    btsPanel.ListIndex = 0
    UpdateVisiblePanel
    
    'Fill any custom "About" panel text
    lblAbout(0).Caption = g_Language.TranslateMessage("PhotoDemon is Copyright %1 2000-%2 by Tanner Helland and Contributors", ChrW$(169), Year(Now))
    lblAbout(1).Caption = Updates.GetPhotoDemonNameAndVersion()
    
    'Fill the "Contributor" panel text
    ReDim m_contributorList(0 To 31) As PD_Contributor
    m_numOfContributors = 0
    
    'Shout-outs to designers, programmers, testers and sponsors
    GenerateContributor "Abhijit Mhapsekar"
    GenerateContributor "A.G. Violette", "http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=55938&lngWId=1"
    GenerateContributor "Alexander Dorn"
    GenerateContributor "Allan Lima"
    GenerateContributor "Andrew Yeoman"
    GenerateContributor "Ari Sohandri Putra", "http://arisohandrip.indonesiaz.com/"
    GenerateContributor "Avery", "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=37541&lngWId=1"
    GenerateContributor "Audioglider", "https://github.com/audioglider"
    GenerateContributor "Bernhard Stockmann", "http://www.gimpusers.com/tutorials/colorful-light-particle-stream-splash-screen-gimp.html"
    GenerateContributor "Boban Gjerasimoski", "https://www.behance.net/Boban_Gjerasimoski"
    GenerateContributor "Bonnie West", "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=74264&lngWId=1"
    GenerateContributor "Carles P.V.", "http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=42376&lngWId=1"
    GenerateContributor "ChenLin"
    GenerateContributor "Chiahong Hong", "https://github.com/ChiahongHong"
    GenerateContributor "chrfb @ deviantart.com", "http://chrfb.deviantart.com/art/quot-ecqlipse-2-quot-PNG-59941546"
    GenerateContributor "Cody Robertson"
    GenerateContributor "Dana Seaman", "http://www.cyberactivex.com/"
    GenerateContributor "Davor Sikic"
    GenerateContributor "dilettante", "http://www.vbforums.com/showthread.php?660014-VB6-ShellPipe-quot-Shell-with-I-O-Redirection-quot-control"
    GenerateContributor "Dirk Hartmann", "http://www.taichi-zentrum-heidelberg.de"
    GenerateContributor "Djordje Djoric", "https://www.odesk.com/o/profiles/users/_~0181c1599705edab79/"
    GenerateContributor "Dosadi", "http://eztwain.com/eztwain1.htm"
    GenerateContributor "Easy RGB", "http://www.easyrgb.com/"
    GenerateContributor "FlatIcons.net", "http://flaticons.net/"
    GenerateContributor "Francis DC"
    GenerateContributor "Frank Donckers", "http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?lngWId=1&txtCriteria=donckers"
    GenerateContributor "Frans van Beers", "https://plus.google.com/+FransvanBeers/"
    GenerateContributor "FreeImage Project", "http://freeimage.sourceforge.net/"
    GenerateContributor "Giorgio ""Gibra"" Brausi", "http://nuke.vbcorner.net"
    GenerateContributor "GioRock", "http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?lngWId=1&txtCriteria=giorock"
    GenerateContributor "Google Translate", "http://translate.google.com"
    GenerateContributor "Hans Nolte", "https://github.com/hansnolte"
    GenerateContributor "Helmut Kuerbiss"
    GenerateContributor "Jason Bullen", "http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=11488&lngWId=1"
    GenerateContributor "Jason Peter Brown", "https://github.com/jpbro"
    GenerateContributor "Jerry Huxtable", "http://www.jhlabs.com/ie/index.html"
    GenerateContributor "Johannes Nendel"
    GenerateContributor "Joseph Greco"
    GenerateContributor "Kroc Camen", "http://camendesign.com"
    GenerateContributor "LaVolpe", "http://www.vbforums.com/showthread.php?t=606736"
    GenerateContributor "Leandro Ascierto", "http://leandroascierto.com/blog/clsmenuimage/"
    GenerateContributor "Lemuel Cushing", "https://github.com/LemuelCushing"
    GenerateContributor "Leonid Blyakher"
    GenerateContributor "Manuel Augusto Santos", "http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=26303&lngWId=1"
    GenerateContributor "Mohammad Reza Karimi"
    GenerateContributor "Nguyen Van Hung"
    GenerateContributor "Olaf Schmidt", "http://www.vbrichclient.com/#/en/About/"
    GenerateContributor "Old Abiquiu Photographic"
    GenerateContributor "Paul Bourke", "http://paulbourke.net/miscellaneous/"
    GenerateContributor "Peter Burn"
    GenerateContributor "Phil Harvey", "http://www.sno.phy.queensu.ca/~phil/exiftool/"
    GenerateContributor "Plinio C Garcia"
    GenerateContributor "PortableFreeware.com team", "http://www.portablefreeware.com/forums/viewtopic.php?t=21652"
    GenerateContributor "Raj Chaudhuri", "https://github.com/rajch"
    GenerateContributor "Robert Rayment", "http://rrprogs.com/"
    GenerateContributor "Roy (rk)"
    GenerateContributor "Shishi"
    GenerateContributor "Steve McMahon", "http://www.vbaccelerator.com/home/VB/index.asp"
    GenerateContributor "Tom Loos", "http://www.designedbyinstinct.com"
    GenerateContributor "Vatterspun", "https://github.com/vatterspun"
    GenerateContributor "Vladimir Vissoultchev", "https://github.com/wqweto"
    GenerateContributor "Will Stampfer", "https://github.com/epmatsw"
    GenerateContributor "Zhu JinYong", "http://www.planetsourcecode.com/vb/authors/ShowBio.asp?lngAuthorId=55292624&lngWId=1"
    
    'Add dummy entries to the owner-drawn list box
    lstContributors.SetAutomaticRedraws False, False
    Dim i As Long
    For i = 0 To m_numOfContributors - 1
        lstContributors.AddItem vbNullString
    Next i
    lstContributors.SetAutomaticRedraws True, True
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me
     
End Sub

Private Sub GenerateContributor(ByVal contributorName As String, Optional ByVal contributorURL As String = vbNullString)
    If (m_numOfContributors > UBound(m_contributorList)) Then ReDim Preserve m_contributorList(0 To m_numOfContributors * 2 - 1) As PD_Contributor
    m_contributorList(m_numOfContributors).ctbName = contributorName
    m_contributorList(m_numOfContributors).ctbURL = contributorURL
    m_numOfContributors = m_numOfContributors + 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub lstContributors_Click()
    If ((lstContributors.ListIndex < m_numOfContributors) And (lstContributors.ListIndex >= 0)) Then
        If (LenB(m_contributorList(lstContributors.ListIndex).ctbURL) <> 0) Then Web.OpenURL m_contributorList(lstContributors.ListIndex).ctbURL
    End If
End Sub

Private Sub lstContributors_DrawListEntry(ByVal bufferDC As Long, ByVal itemIndex As Long, itemTextEn As String, ByVal itemIsSelected As Boolean, ByVal itemIsHovered As Boolean, ByVal ptrToRectF As Long)

    If (bufferDC = 0) Then Exit Sub
    
    'Calculate text colors (which vary depending on hover state and URL availability)
    Dim itemIsClickable As Boolean
    itemIsClickable = (LenB(m_contributorList(itemIndex).ctbURL) <> 0)
    
    If pdMain.IsProgramRunning() Then
    
        Dim txtColor As Long
        If itemIsClickable Then
            txtColor = g_Themer.GetGenericUIColor(UI_TextClickable, , , itemIsHovered)
        Else
            txtColor = g_Themer.GetGenericUIColor(UI_TextReadOnly)
        End If
        
        'Prep various default rendering values (including retrieval of the boundary rect from the list box manager)
        Dim tmpRectF As RectF
        CopyMemory ByVal VarPtr(tmpRectF), ByVal ptrToRectF, 16&
        
        'Manually paint an uncolored background
        Dim cPainter As pd2DPainter, cSurface As pd2DSurface, cBrush As pd2DBrush
        Drawing2D.QuickCreatePainter cPainter
        Drawing2D.QuickCreateSurfaceFromDC cSurface, bufferDC
        Drawing2D.QuickCreateSolidBrush cBrush, g_Themer.GetGenericUIColor(UI_Background, Me.Enabled)
        cPainter.FillRectangleF_FromRectF cSurface, cBrush, tmpRectF
        Set cBrush = Nothing: Set cSurface = Nothing: Set cPainter = Nothing
        
        'Prepare and render the contributor's name
        Dim drawString As String
        drawString = m_contributorList(itemIndex).ctbName
        
        Dim tmpFont As pdFont, txtFontSize As Single
        txtFontSize = 10#
        
        If itemIsClickable Then
            Set tmpFont = Fonts.GetMatchingUIFont(txtFontSize, False, False, itemIsHovered)
        Else
            Set tmpFont = Fonts.GetMatchingUIFont(txtFontSize, False, False, False)
        End If
        
        tmpFont.AttachToDC bufferDC
        tmpFont.SetFontColor txtColor
        tmpFont.SetTextAlignment vbLeftJustify
        
        Dim targetRect As RECT
        With targetRect
            .Left = tmpRectF.Left
            .Top = tmpRectF.Top
            .Right = tmpRectF.Left + tmpRectF.Width
            .Bottom = tmpRectF.Top + tmpRectF.Height
        End With
        
        tmpFont.DrawCenteredTextToRect drawString, targetRect, True
        tmpFont.ReleaseFromDC
        
    End If
    
End Sub

Private Sub lstContributors_MouseLeave()
    lstContributors.AssignTooltip vbNullString
End Sub

Private Sub lstContributors_MouseOver(ByVal itemIndex As Long, itemTextEn As String)
    If ((itemIndex < m_numOfContributors) And (itemIndex >= 0)) Then
        lstContributors.AssignTooltip m_contributorList(itemIndex).ctbURL
    Else
        lstContributors.AssignTooltip vbNullString
    End If
End Sub
