VERSION 5.00
Begin VB.Form FormAbout 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " About PhotoDemon"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   390
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
      Top             =   840
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   11245
      Begin PhotoDemon.pdHyperlink hypAbout 
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   1920
         Width           =   9300
         _ExtentX        =   16404
         _ExtentY        =   661
         Caption         =   ""
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
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   2400
         Width           =   9300
         _ExtentX        =   16404
         _ExtentY        =   661
         Caption         =   "Participate in development, design, or translation work"
         URL             =   "http://photodemon.org/get-involved/"
      End
      Begin PhotoDemon.pdHyperlink hypAbout 
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   9300
         _ExtentX        =   16404
         _ExtentY        =   661
         Caption         =   ""
         URL             =   "https://www.patreon.com/photodemon/overview"
      End
      Begin PhotoDemon.pdHyperlink hypAbout 
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   2880
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
         TabIndex        =   8
         Top             =   3360
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
         TabIndex        =   9
         Top             =   3840
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
      Top             =   840
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   11245
      Begin PhotoDemon.pdListBoxOD lstContributors 
         Height          =   6135
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   9405
         _ExtentX        =   16589
         _ExtentY        =   10821
         BorderlessMode  =   -1  'True
      End
   End
   Begin PhotoDemon.pdContainer pnlAbout 
      Height          =   6375
      Index           =   2
      Left            =   120
      Top             =   840
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   11245
      Begin PhotoDemon.pdListBoxOD lstPatrons 
         Height          =   6135
         Left            =   120
         TabIndex        =   2
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
'Copyright 2001-2026 by Tanner Helland
'Created: 6/12/01
'Last updated: 07/January/25
'Last update: update contributor list
'
'PhotoDemon would not be possible without the help of many, many amazing people.  THANK YOU!
'
'If you contributed to PhotoDemon's development in some way, but your name isn't listed here,
' please let me know!  I can always be reached via https://photodemon.org/about/contact/
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Type PD_Contributor
    ctbName As String
    ctbURL As String
    isInternal As Boolean
End Type

Private m_contributorList() As PD_Contributor, m_numOfContributors As Long
Private m_patronList() As PD_Contributor, m_numOfPatrons As Long
Private m_superPatronEndIndex As Long

'Height of each credit content block (at 96 DPI)
Private Const BLOCKHEIGHT_CONTRIBUTOR As Long = 24
Private Const BLOCKHEIGHT_PATRON As Long = 28

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
    
    lstContributors.ListItemHeight = Interface.FixDPI(BLOCKHEIGHT_CONTRIBUTOR)
    lstPatrons.ListItemHeight = Interface.FixDPI(BLOCKHEIGHT_PATRON)
    
    btsPanel.AddItem "About", 0
    btsPanel.AddItem "Contributors", 1
    btsPanel.AddItem "Patrons", 2
    btsPanel.ListIndex = 0
    UpdateVisiblePanel
    
    'Fill any custom "About" panel text
    lblAbout(0).Caption = g_Language.TranslateMessage("PhotoDemon is Copyright %1 2000-%2 by Tanner Helland and Contributors", ChrW$(169), Year(Now))
    lblAbout(1).Caption = Updates.GetPhotoDemonNameAndVersion()
    
    'Fill some specialty links
    Dim actualText As String
    actualText = g_Language.TranslateMessage("Support us on Patreon")
    
    'Win 7+ supports some helpful symbolic unicode chars
    If OS.IsWin7OrLater Then actualText = ChrW$(&H2665) & Space$(2) & actualText
    hypAbout(0).Caption = actualText
    
    actualText = g_Language.TranslateMessage("Donate to PhotoDemon development")
    If OS.IsWin7OrLater Then actualText = ChrW$(&H2665) & Space$(2) & actualText
    hypAbout(1).Caption = actualText
    
    'Fill the "Patron" panel text
    ReDim m_patronList(0 To 15) As PD_Contributor
    m_numOfPatrons = 0
    
    'On certain OSes, we can use Unicode symbols to make the text a little more fun
    Dim uncSpecial As String, setSpaces As String
    setSpaces = Space$(4)
    
    'Super patrons
    actualText = g_Language.TranslateMessage("SUPER PATRONS")
    If OS.IsWin7OrLater Then
        uncSpecial = ChrW$(&H2605) & ChrW$(&H2605)    'Stars
        GeneratePatron uncSpecial & setSpaces & actualText & setSpaces & uncSpecial, "header", True
    Else
        GeneratePatron actualText, , True
    End If
    
    GeneratePatron "DeltaVenue"
    GeneratePatron "Johannes Nendel"
    GeneratePatron "Refael Ackermann"
    
    m_superPatronEndIndex = m_numOfPatrons
    GeneratePatron vbNullString
    
    'Regular patrons
    actualText = g_Language.TranslateMessage("PATRONS")
    If OS.IsWin7OrLater Then
        uncSpecial = ChrW$(&H2605) & ChrW$(&H2605)    'Stars
        GeneratePatron uncSpecial & setSpaces & actualText & setSpaces & uncSpecial, "header", True
    Else
        GeneratePatron actualText, , True
    End If
    
    GeneratePatron "Frank Reibold"
    GeneratePatron "Heiner Dietz"
    GeneratePatron "Jim Schmidt"
    
    GeneratePatron vbNullString
    
    'Thank you text
    actualText = g_Language.TranslateMessage("Thank you to our wonderful Patreon supporters!")
    If OS.IsWin7OrLater Then
        uncSpecial = ChrW$(&H2665)    'Hearts
        GeneratePatron uncSpecial & setSpaces & actualText, , True
    Else
        GeneratePatron actualText
    End If
    
    actualText = g_Language.TranslateMessage("(You can become a patron, too!  Click here to learn more.)")
    GeneratePatron actualText, "https://www.patreon.com/photodemon/overview"
    
    'Fill the "Contributor" panel text
    ReDim m_contributorList(0 To 31) As PD_Contributor
    m_numOfContributors = 0
    
    GenerateContributor "Abhijit Mhapsekar"
    GenerateContributor "Allan Lima"
    GenerateContributor "Alric Rahl", "https://t.me/Alricrahl"
    GenerateContributor "Andrew Yeoman"
    GenerateContributor "Ari Sohandri Putra", "https://github.com/arisohandriputra"
    GenerateContributor Strings.StringFromUtf8Base64("QsOhbnN6a2kgSXN0dsOhbg")
    GenerateContributor "Boban Gjerasimoski", "https://www.behance.net/Boban_Gjerasimoski"
    GenerateContributor "Bonnie West", "https://github.com/Planet-Source-Code/bonnie-west-the-optimum-fileexists-function__1-74264"
    GenerateContributor "Carles P.V.", "https://github.com/Planet-Source-Code/carles-p-v-ibmp-1-2__1-42376"
    GenerateContributor "CharLS Project, including Jan de Vaan and Victor Derks", "https://github.com/team-charls/charls"
    GenerateContributor "Charltsing", "https://www.cnblogs.com/Charltsing/"
    GenerateContributor "ChenLin"
    GenerateContributor "Chiahong Hong", "https://github.com/ChiahongHong"
    GenerateContributor Strings.StringFromUtf8Base64("Q2zDqW1lbnQgTWFyaWFnZQ")
    GenerateContributor "Cody Robertson"
    GenerateContributor "Dana Seaman"
    GenerateContributor "DarkAlchy"
    GenerateContributor Strings.StringFromUtf8Base64("RGF2b3IgxaBpa2lj")
    GenerateContributor "dilettante", "http://www.vbforums.com/showthread.php?660014-VB6-ShellPipe-quot-Shell-with-I-O-Redirection-quot-control"
    GenerateContributor "Dirk Hartmann", "https://taijidaoyin.com/"
    GenerateContributor "Djordje Djoric"
    GenerateContributor "Dosadi", "https://eztwain.com/eztwain1.htm"
    GenerateContributor "Easy RGB", "https://www.easyrgb.com/"
    GenerateContributor "FlatIcons.net", "https://flaticons.net/"
    GenerateContributor "Francis DC"
    GenerateContributor "Frank Donckers", "https://github.com/Planet-Source-Code?q=donckers&type=&language="
    GenerateContributor "Frans van Beers", "https://www.flickr.com/photos/stillicht/"
    GenerateContributor "FreeImage Project", "https://freeimage.sourceforge.net/"
    GenerateContributor "Gerry Busch"
    GenerateContributor "Giorgio ""Gibra"" Brausi", "http://nuke.vbcorner.net"
    GenerateContributor "GioRock", "https://github.com/Planet-Source-Code?q=giorock&type=&language="
    GenerateContributor "Graeme Gill", "https://argyllcms.com/"
    GenerateContributor "Hans Nolte", "https://github.com/hansnolte"
    GenerateContributor "Heiner Dietz"
    GenerateContributor "Helmut Kuerbiss"
    GenerateContributor "Hidayat Suriapermana"
    GenerateContributor "J. Scott Elblein", "https://geekdrop.com"
    GenerateContributor "Jason Bullen", "https://github.com/Planet-Source-Code/jason-bullen-simple-cubic-spline-curve-plot__1-11488"
    GenerateContributor "Jason Peter Brown", "https://github.com/jpbro"
    GenerateContributor "Jean Jacques Piedfort"
    GenerateContributor "Jerry Huxtable", "http://www.jhlabs.com/"
    GenerateContributor "Jian Ma", "https://www.cnblogs.com/stronghorse/"
    GenerateContributor "Johannes Nendel"
    GenerateContributor "John Desrosiers", "https://johndesrosiers.com"
    GenerateContributor "Jose de TECNORAMA", "https://www.tecnorama.es/"
    GenerateContributor "Joseph Greco"
    GenerateContributor "Kenji Hoshimoto (Hosiken)", "http://hosiken.jp/"
    GenerateContributor "Kroc Camen", "https://camendesign.com"
    GenerateContributor "LaVolpe", "https://www.vbforums.com/showthread.php?t=606736"
    GenerateContributor "Leandro Ascierto", "https://leandroascierto.com/blog/clsmenuimage/"
    GenerateContributor "Lemuel Cushing", "https://github.com/LemuelCushing"
    GenerateContributor "Leonid Blyakher"
    GenerateContributor "libavif Project", "https://github.com/AOMediaCodec/libavif"
    GenerateContributor "libwebp Project", "https://developers.google.com/speed/webp"
    GenerateContributor "Liviu Ivanov"
    GenerateContributor "LsGeorge", "https://github.com/LsGeorge"
    GenerateContributor "Manfredi Marceca"
    GenerateContributor "Manuel Augusto Santos", "https://github.com/Planet-Source-Code/manuel-augusto-santos-fast-graphics-filters__1-26303"
    GenerateContributor Strings.StringFromUtf8Base64("TWFyY29zIFZlbnR1cmEgKMaUzp7QmNCixrHGps6UKQ")
    GenerateContributor "Mariozo"
    GenerateContributor "martin19", "https://github.com/martin19"
    GenerateContributor "Miguel Chamorro"
    GenerateContributor "Ming", "https://ufoym.com/"
    GenerateContributor "Mohammad Reza Karimi"
    GenerateContributor "Need74", "https://github.com/Need74"
    GenerateContributor "Nguyen Van Hung", "https://github.com/vhreal1302"
    GenerateContributor "Olaf Schmidt", "http://www.vbrichclient.com/#/en/About/"
    GenerateContributor "Old Abiquiu Photographic"
    GenerateContributor "OpenJPEG", "https://www.openjpeg.org/"
    GenerateContributor "Paul Bourke", "https://paulbourke.net/miscellaneous/"
    GenerateContributor "Peter Burn"
    GenerateContributor "Phil Harvey", "https://exiftool.org/"
    GenerateContributor "pixman library", "https://pixman.org/"
    GenerateContributor "Plinio C Garcia"
    GenerateContributor "PortableFreeware.com team", "http://www.portablefreeware.com/forums/viewtopic.php?t=21652"
    GenerateContributor "Raj Chaudhuri", "https://github.com/rajch"
    GenerateContributor "Robert Rayment"
    GenerateContributor "Ron van Tilburg", "https://github.com/Planet-Source-Code/ron-van-tilburg-rvtvbimg__1-14210"
    GenerateContributor "Roy K (rk)"
    GenerateContributor "Ryszard Cwenar"
    GenerateContributor "Shishi"
    GenerateContributor "Sinisa Petric", "https://github.com/spetric/Photoshop-Plugin-Host"
    GenerateContributor "Steve McMahon", "http://www.vbaccelerator.com/home/VB/index.asp"
    GenerateContributor "twinBASIC", "https://github.com/twinbasic/twinbasic"
    GenerateContributor "TyraVex", "https://github.com/TyraVex"
    GenerateContributor "Vatterspun", "https://github.com/vatterspun"
    GenerateContributor "Vladimir Vissoultchev", "https://github.com/wqweto"
    GenerateContributor "Will Stampfer", "https://github.com/epmatsw"
    GenerateContributor "Yuriy Balyuk", "https://github.com/veksha"
    GenerateContributor "Zhu JinYong", "https://github.com/Planet-Source-Code?q=jinyong&type=&language="
    GenerateContributor Strings.StringFromUtf8Base64("5piH"), "https://github.com/love80312"
    
    'Add dummy entries to the owner-drawn list boxes
    Dim i As Long
    
    lstPatrons.SetAutomaticRedraws False, False
    For i = 0 To m_numOfPatrons - 1
        lstPatrons.AddItem vbNullString
    Next i
    lstPatrons.SetAutomaticRedraws True, True
    
    lstContributors.SetAutomaticRedraws False, False
    For i = 0 To m_numOfContributors - 1
        lstContributors.AddItem vbNullString
    Next i
    lstContributors.SetAutomaticRedraws True, True
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me
     
End Sub

Private Sub GenerateContributor(ByRef contributorName As String, Optional ByRef contributorURL As String = vbNullString)
    If (m_numOfContributors > UBound(m_contributorList)) Then ReDim Preserve m_contributorList(0 To m_numOfContributors * 2 - 1) As PD_Contributor
    m_contributorList(m_numOfContributors).ctbName = contributorName
    m_contributorList(m_numOfContributors).ctbURL = contributorURL
    m_numOfContributors = m_numOfContributors + 1
End Sub

Private Sub GeneratePatron(ByRef patronName As String, Optional ByRef patronURL As String = vbNullString, Optional ByVal patronIsInternalText As Boolean = False)
    If (m_numOfPatrons > UBound(m_patronList)) Then ReDim Preserve m_patronList(0 To m_numOfPatrons * 2 - 1) As PD_Contributor
    With m_patronList(m_numOfPatrons)
        .ctbName = patronName
        .ctbURL = patronURL
        .isInternal = patronIsInternalText
    End With
    m_numOfPatrons = m_numOfPatrons + 1
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
    
    If PDMain.IsProgramRunning() Then
    
        Dim txtColor As Long
        If itemIsClickable Then
            txtColor = g_Themer.GetGenericUIColor(UI_TextClickable, , , itemIsHovered)
        Else
            txtColor = g_Themer.GetGenericUIColor(UI_TextReadOnly)
        End If
        
        'Prep various default rendering values (including retrieval of the boundary rect from the list box manager)
        Dim tmpRectF As RectF
        CopyMemoryStrict VarPtr(tmpRectF), ptrToRectF, 16&
        
        'Manually paint an uncolored background
        Dim cSurface As pd2DSurface, cBrush As pd2DBrush
        Drawing2D.QuickCreateSurfaceFromDC cSurface, bufferDC
        Drawing2D.QuickCreateSolidBrush cBrush, g_Themer.GetGenericUIColor(UI_Background, Me.Enabled)
        PD2D.FillRectangleF_FromRectF cSurface, cBrush, tmpRectF
        Set cBrush = Nothing: Set cSurface = Nothing
        
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
        
        'Only clickable entries should get a hand icon
        If (LenB(m_contributorList(itemIndex).ctbURL) <> 0) Then
            lstContributors.RequestCursor IDC_HAND
        Else
            lstContributors.RequestCursor IDC_DEFAULT
        End If
        
    Else
        lstContributors.AssignTooltip vbNullString
    End If
    
End Sub

Private Sub lstPatrons_Click()
    If ((lstPatrons.ListIndex < m_numOfPatrons) And (lstPatrons.ListIndex >= 0)) Then
        If (LenB(m_patronList(lstPatrons.ListIndex).ctbURL) <> 0) Then Web.OpenURL m_patronList(lstPatrons.ListIndex).ctbURL
    End If
End Sub

Private Sub lstPatrons_DrawListEntry(ByVal bufferDC As Long, ByVal itemIndex As Long, itemTextEn As String, ByVal itemIsSelected As Boolean, ByVal itemIsHovered As Boolean, ByVal ptrToRectF As Long)

    If (bufferDC = 0) Then Exit Sub
    
    'Calculate text colors (which vary depending on hover state and URL availability)
    Dim itemIsClickable As Boolean
    itemIsClickable = (LenB(m_patronList(itemIndex).ctbURL) <> 0) And (Not m_patronList(itemIndex).isInternal)
    
    If PDMain.IsProgramRunning() Then
    
        Dim txtColor As Long
        If itemIsClickable Then
            txtColor = g_Themer.GetGenericUIColor(UI_TextClickable, , , itemIsHovered)
        Else
            txtColor = g_Themer.GetGenericUIColor(UI_TextReadOnly)
        End If
        
        'Prep various default rendering values (including retrieval of the boundary rect from the list box manager)
        Dim tmpRectF As RectF
        CopyMemoryStrict VarPtr(tmpRectF), ptrToRectF, 16&
        
        'Manually paint an uncolored background
        Dim cSurface As pd2DSurface, cBrush As pd2DBrush
        Drawing2D.QuickCreateSurfaceFromDC cSurface, bufferDC
        Drawing2D.QuickCreateSolidBrush cBrush, g_Themer.GetGenericUIColor(UI_Background, Me.Enabled)
        PD2D.FillRectangleF_FromRectF cSurface, cBrush, tmpRectF
        Set cBrush = Nothing: Set cSurface = Nothing
        
        'Prepare and render the patron's name
        Dim drawString As String
        drawString = m_patronList(itemIndex).ctbName
        
        Dim tmpFont As pdFont, txtFontSize As Single
        txtFontSize = 10!
        If (itemIndex <= m_superPatronEndIndex) Then txtFontSize = 12!
        
        If itemIsClickable Then
            Set tmpFont = Fonts.GetMatchingUIFont(txtFontSize, False, False, itemIsHovered)
        Else
            
            If m_patronList(itemIndex).isInternal Then
                
                'Win 7+ supports some symbols that earlier OSes do not
                If OS.IsWin7OrLater Then
                
                    With m_patronList(itemIndex)
                        If Strings.StringsEqual(.ctbURL, "header", True) Then
                            Set tmpFont = New pdFont
                            tmpFont.SetFontFace "Segoe UI Symbol"
                            tmpFont.SetFontBold True
                            tmpFont.SetFontSize txtFontSize
                            tmpFont.CreateFontObject
                        Else
                            Set tmpFont = Fonts.GetMatchingUIFont(txtFontSize, (itemIndex <= m_superPatronEndIndex), False, False)
                        End If
                    End With
                    
                Else
                    Set tmpFont = Fonts.GetMatchingUIFont(txtFontSize, (itemIndex <= m_superPatronEndIndex), False, False)
                End If
            
            Else
                Set tmpFont = Fonts.GetMatchingUIFont(txtFontSize, False, False, False)
            End If
            
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

Private Sub lstPatrons_MouseLeave()
    lstPatrons.AssignTooltip vbNullString
End Sub

Private Sub lstPatrons_MouseOver(ByVal itemIndex As Long, itemTextEn As String)

    If ((itemIndex < m_numOfPatrons) And (itemIndex >= 0)) Then
        
        'Only clickable entries should get a hand icon
        Dim isClickable As Boolean
        isClickable = (LenB(m_patronList(itemIndex).ctbURL) <> 0) And (Not m_patronList(itemIndex).isInternal)
        
        If isClickable Then
            lstPatrons.RequestCursor IDC_HAND
            lstPatrons.AssignTooltip m_patronList(itemIndex).ctbURL
        Else
            lstPatrons.RequestCursor IDC_DEFAULT
            lstPatrons.AssignTooltip vbNullString
        End If
        
    Else
        lstPatrons.AssignTooltip vbNullString
    End If
    
End Sub
