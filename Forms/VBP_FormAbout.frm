VERSION 5.00
Begin VB.Form FormAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " About PhotoDemon"
   ClientHeight    =   8055
   ClientLeft      =   2340
   ClientTop       =   1875
   ClientWidth     =   9000
   ClipControls    =   0   'False
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
   ScaleHeight     =   537
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   7440
      TabIndex        =   0
      Top             =   7440
      Width           =   1365
   End
   Begin VB.Label lblzLib 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "zLib license page"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4650
      MouseIcon       =   "VBP_FormAbout.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   6960
      Width           =   1200
   End
   Begin VB.Label lblEZTW32 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EZTwain license page"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2535
      MouseIcon       =   "VBP_FormAbout.frx":015E
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   6960
      Width           =   1530
   End
   Begin VB.Label lblPngnq 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "pngnq-s9 homepage"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   6420
      MouseIcon       =   "VBP_FormAbout.frx":02B0
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   6960
      Width           =   1470
   End
   Begin VB.Label lblFreeImage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FreeImage license page"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   270
      MouseIcon       =   "VBP_FormAbout.frx":0402
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   6960
      Width           =   1710
   End
   Begin VB.Label lblLinks 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   240
      TabIndex        =   4
      Top             =   6480
      Width           =   8580
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000002&
      X1              =   8
      X2              =   592
      Y1              =   184
      Y2              =   184
   End
   Begin VB.Label lblThanks 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "PhotoDemon would not be possible without contributions from the following individuals and organizations:"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   0
      Left            =   240
      MouseIcon       =   "VBP_FormAbout.frx":0554
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2880
      Width           =   7560
   End
   Begin VB.Label lblDisclaimer 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (auto-populated at run-time)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   2910
      TabIndex        =   2
      Top             =   2400
      Width           =   5985
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   210
      TabIndex        =   1
      Top             =   2400
      Width           =   795
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
'Last updated: 23/April/13
'Last update: added additional names
'
'A simple "about"/credits form.  Contains credits, copyright, and the program logo.
'
'***************************************************************************

Option Explicit

Dim creditList() As String
Dim curCredit As Long

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    'Automatic generation of version & copyright information
    lblVersion.Caption = g_Language.TranslateMessage("Version") & " " & App.Major & "." & App.Minor & "." & App.Revision
    lblDisclaimer.Caption = App.LegalCopyright & " "
    lblLinks.Caption = g_Language.TranslateMessage("PhotoDemon uses several third-party plugins.  These plugins may be governed by additional licenses. For more information, please visit:")
    
    curCredit = 1
    
    'Shout-outs to other designers, programmers, testers and sponsors who provided various resources
    GenerateThankyou "Adrian Pellas-Rice", "http://sourceforge.net/projects/pngnqs9/"
    GenerateThankyou "Alfred Hellmueller"
    GenerateThankyou "Andrew Yeoman"
    GenerateThankyou "Avery", "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=37541&lngWId=1"
    GenerateThankyou "Brad Martinez", "http://btmtz.mvps.org/gfxfromfrx/"
    GenerateThankyou "Carles P.V.", "http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=42376&lngWId=1"
    GenerateThankyou "chrfb @ deviantart.com", "http://chrfb.deviantart.com/art/quot-ecqlipse-2-quot-PNG-59941546"
    GenerateThankyou "Dave Jamison", "http://www.modeltrainsoftware.com/"
    GenerateThankyou "Dosadi", "http://eztwain.com/eztwain1.htm"
    GenerateThankyou "Everaldo Coelho", "http://www.everaldo.com/"
    GenerateThankyou "Frank Donckers"
    GenerateThankyou "FreeImage Project", "http://freeimage.sourceforge.net/"
    GenerateThankyou "Gilles Vollant", "http://www.winimage.com/zLibDll/index.html"
    GenerateThankyou "GioRock", "http://www.planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?lngWId=1&blnAuthorSearch=TRUE&lngAuthorId=77440558266&strAuthorName=GioRock&txtMaxNumberOfEntriesPerPage=25"
    GenerateThankyou "Jason Bullen", "http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=11488&lngWId=1"
    GenerateThankyou "Jerry Huxtable", "http://www.jhlabs.com/ie/index.html"
    GenerateThankyou "Juned Chhipa", "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=71482&lngWId=1"
    GenerateThankyou "Herman Liu"
    GenerateThankyou "Kroc Camen", "http://camendesign.com"
    GenerateThankyou "LaVolpe", "http://www.vbforums.com/showthread.php?t=606736"
    GenerateThankyou "Leandro Ascierto", "http://leandroascierto.com/blog/clsmenuimage/"
    GenerateThankyou "Manuel Augusto Santos", "http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=26303&lngWId=1"
    GenerateThankyou "Mark James", "http://www.famfamfam.com/lab/icons/silk/"
    GenerateThankyou "Phil Fresle", "http://www.frez.co.uk/vb6.aspx"
    GenerateThankyou "Robert Rayment", "http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=66991&lngWId=1"
    GenerateThankyou "Rod Stephens", "http://www.vb-helper.com"
    GenerateThankyou "Ron van Tilburg", "http://www.planet-source-code.com/vb/scripts/showcode.asp?txtCodeId=71370&lngWid=1"
    GenerateThankyou "Steve McMahon", "http://www.vbaccelerator.com/home/VB/index.asp"
    GenerateThankyou "Tango Icon Library", "http://tango.freedesktop.org/"
    GenerateThankyou "Waty Thierry", "http://www.ppreview.net/"
    GenerateThankyou "Yusuke Kamiyamane", "http://p.yusukekamiyamane.com/"
    GenerateThankyou "Zhu JinYong"
    
    lblThanks(0).MousePointer = vbDefault
    
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
    
End Sub

'Generate a label with the specified "thank you" text, and link it to the specified URL
Private Sub GenerateThankyou(ByVal thxText As String, Optional ByVal creditURL As String = "")
    
    'Generate a new label
    Load lblThanks(curCredit)
    
    'Because I now have too many people to thank, it's necessary to split the list into multiple columns
    Dim columnLimit As Long
    columnLimit = 12
    
    Dim thxOffset As Long
    thxOffset = 50
    
    If curCredit = 1 Then
        lblThanks(curCredit).Top = lblThanks(curCredit - 1).Top + lblThanks(curCredit - 1).Height + 20
        lblThanks(curCredit).Left = lblThanks(0).Left + 2 + thxOffset
    ElseIf curCredit < columnLimit Then
        lblThanks(curCredit).Top = lblThanks(curCredit - 1).Top + lblThanks(curCredit - 1).Height + 4
        lblThanks(curCredit).Left = lblThanks(0).Left + 2 + thxOffset
    ElseIf curCredit = columnLimit Then
        lblThanks(curCredit).Top = lblThanks(curCredit - 1).Top + lblThanks(curCredit - 1).Height + 20 - (lblThanks(columnLimit - 1).Top - lblThanks(0).Top)
        lblThanks(curCredit).Left = lblThanks(0).Left + 180 + thxOffset
    ElseIf curCredit < columnLimit * 2 - 1 Then
        lblThanks(curCredit).Top = lblThanks(curCredit - 1).Top + lblThanks(curCredit - 1).Height + 4
        lblThanks(curCredit).Left = lblThanks(0).Left + 180 + thxOffset
    ElseIf curCredit = columnLimit * 2 - 1 Then
        lblThanks(curCredit).Top = lblThanks(curCredit - 1).Top + lblThanks(curCredit - 1).Height + 20 - (lblThanks(columnLimit * 2 - 2).Top - lblThanks(0).Top)
        lblThanks(curCredit).Left = lblThanks(0).Left + 360 + thxOffset
    Else
        lblThanks(curCredit).Top = lblThanks(curCredit - 1).Top + lblThanks(curCredit - 1).Height + 4
        lblThanks(curCredit).Left = lblThanks(0).Left + 360 + thxOffset
    End If
    
    lblThanks(curCredit).Caption = thxText
    If creditURL = "" Then
        lblThanks(curCredit).MousePointer = vbDefault
    Else
        lblThanks(curCredit).FontUnderline = True
        lblThanks(curCredit).ForeColor = vbBlue
        lblThanks(curCredit).ToolTipText = g_Language.TranslateMessage("Click to open") & " " & creditURL
    End If
    lblThanks(curCredit).Visible = True
    
    ReDim Preserve creditList(0 To curCredit) As String
    creditList(curCredit) = creditURL
    
    curCredit = curCredit + 1

End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub lblEZTW32_Click()
    OpenURL "http://eztwain.com/ezt1faq.htm"
End Sub

Private Sub lblFreeImage_Click()
    OpenURL "http://freeimage.sourceforge.net/license.html"
End Sub

Private Sub lblPngnq_Click()
    OpenURL "http://sourceforge.net/projects/pngnqs9/"
End Sub

'When a thank-you credit is clicked, launch the corresponding website
Private Sub lblThanks_Click(Index As Integer)

    If creditList(Index) <> "" Then OpenURL creditList(Index)

End Sub

Private Sub lblzLib_Click()
    OpenURL "http://www.zlib.net/zlib_license.html"
End Sub
