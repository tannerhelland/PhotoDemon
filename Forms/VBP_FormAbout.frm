VERSION 5.00
Begin VB.Form FormAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " About PhotoDemon"
   ClientHeight    =   9735
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
   ScaleHeight     =   649
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   465
      Left            =   7320
      TabIndex        =   0
      Top             =   9120
      Width           =   1485
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
      Left            =   4050
      MouseIcon       =   "VBP_FormAbout.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   8640
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
      Left            =   2235
      MouseIcon       =   "VBP_FormAbout.frx":015E
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   8640
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
      Left            =   5520
      MouseIcon       =   "VBP_FormAbout.frx":02B0
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   8640
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
      Top             =   8640
      Width           =   1710
   End
   Begin VB.Label lblLinks 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"VBP_FormAbout.frx":0554
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   240
      TabIndex        =   4
      Top             =   8160
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
      MouseIcon       =   "VBP_FormAbout.frx":05DF
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
      Caption         =   "Version XX.XX.XXXX"
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
      Width           =   2100
   End
End
Attribute VB_Name = "FormAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'About Form
'Copyright ©2001-2012 by Tanner Helland
'Created: 6/12/01
'Last updated: 23/November/12
'Last update: updated list to reflect recent changes to the codebase.
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
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblDisclaimer.Caption = App.LegalCopyright & " "
    
    curCredit = 1
    
    'Shout-outs to other designers, programmers, testers and sponsors who provided various resources
    GenerateThankyou "Kroc of camendesign.com (UI design and organization)", "http://camendesign.com"
    GenerateThankyou "Leandro Ascierto (embedding PNGs as menu icons)", "http://leandroascierto.com/blog/clsmenuimage/"
    GenerateThankyou "Mark James (Silk icon set, CC-BY-2.5)", "http://www.famfamfam.com/lab/icons/silk/"
    GenerateThankyou "LaVolpe (automated VB6 Manifest Creator tool)", "http://www.vbforums.com/showthread.php?t=606736"
    GenerateThankyou "Jason Bullen (knot-based cubic spline interpolation)", "http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=11488&lngWId=1"
    GenerateThankyou "Waty Thierry (printer interfacing in VB)", "http://www.ppreview.net/"
    GenerateThankyou "Dosadi (EZTW32 scanner/digital camera library)", "http://eztwain.com/eztwain1.htm"
    GenerateThankyou "Carles P.V., Avery, Dana Seaman (GDI+ references)", "http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=42376&lngWId=1"
    GenerateThankyou "Brad Martinez (VB binary file extraction)", "http://btmtz.mvps.org/gfxfromfrx/"
    GenerateThankyou "Paul Turcksin (dynamic MDI child icons)", "http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=60600&lngWId=1"
    GenerateThankyou "Manuel Augusto Santos ('Artistic Contour')", "http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=26303&lngWId=1"
    GenerateThankyou "Ron van Tilburg (implementation of Xiaolin Wu line antialiasing)", "http://www.planet-source-code.com/vb/scripts/showcode.asp?txtCodeId=71370&lngWid=1"
    GenerateThankyou "Floris van de Berg, Hervé Drolon, Carsten Klein (FreeImage library, GPLv2)", "http://freeimage.sourceforge.net/"
    GenerateThankyou "Jean-Loup Gailly, Mark Adler, Gilles Vollant (zLib library and wrapper)", "http://www.winimage.com/zLibDll/index.html"
    GenerateThankyou "Juned Chhipa ('jcButton 1.7' customizable command button replacement control)", "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=71482&lngWId=1"
    GenerateThankyou "Steve McMahon (CommonDialog interface, accelerator key handler, progress bar)", "http://www.vbaccelerator.com/home/VB/index.asp"
    GenerateThankyou "chrfb of deviantart.com (PhotoDemon's icon, 'Ecqlipse 2', CC-BY-NC-SA-3.0)", "http://chrfb.deviantart.com/art/quot-ecqlipse-2-quot-PNG-59941546"
    GenerateThankyou "Everaldo and The Crystal Project (Crystal icons, LGPL-licensed, click for details)", "http://www.everaldo.com/crystal/"
    GenerateThankyou "Yusuke Kamiyamane (Fugue icon set, CC-BY-3.0)", "http://p.yusukekamiyamane.com/"
    GenerateThankyou "The Tango Icon Library (public-domain)", "http://tango.freedesktop.org/"
    GenerateThankyou "Phil Fresle (SHA-256 hash implementation)", "http://www.frez.co.uk/vb6.aspx"
    GenerateThankyou "Johannes B ('Fog')", "http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=42642&lngWId=1"
    GenerateThankyou "Dave Jamison", "http://www.modeltrainsoftware.com/"
    GenerateThankyou "Herman Liu"
    GenerateThankyou "Robert Rayment"
    GenerateThankyou "Alfred Hellmueller"
    GenerateThankyou "Zhu JinYong"
    
    lblThanks(0).MousePointer = vbDefault
    
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
    
End Sub

'Generate a label with the specified "thank you" text, and link it to the specified URL
Private Sub GenerateThankyou(ByVal thxText As String, Optional ByVal creditURL As String = "")
    
    'Generate a new label
    Load lblThanks(curCredit)
    
    'Because I now have too many people to thank, it's necessary to split the list into two columns
    Dim columnLimit As Long
    columnLimit = 19
    
    If curCredit = 1 Then
        lblThanks(curCredit).Top = lblThanks(curCredit - 1).Top + lblThanks(curCredit - 1).Height + 12
        lblThanks(curCredit).Left = lblThanks(0).Left + 2
    ElseIf curCredit < columnLimit Then
        lblThanks(curCredit).Top = lblThanks(curCredit - 1).Top + lblThanks(curCredit - 1).Height + 4
        lblThanks(curCredit).Left = lblThanks(0).Left + 2
    ElseIf curCredit = columnLimit Then
        lblThanks(curCredit).Top = lblThanks(curCredit - 1).Top + lblThanks(curCredit - 1).Height + 12 - (lblThanks(columnLimit - 1).Top - lblThanks(0).Top)
        lblThanks(curCredit).Left = lblThanks(0).Left + 325
    Else
        lblThanks(curCredit).Top = lblThanks(curCredit - 1).Top + lblThanks(curCredit - 1).Height + 4
        lblThanks(curCredit).Left = lblThanks(0).Left + 325
    End If
    
    lblThanks(curCredit).Caption = thxText
    If creditURL = "" Then
        lblThanks(curCredit).MousePointer = vbDefault
    Else
        lblThanks(curCredit).FontUnderline = True
        lblThanks(curCredit).ForeColor = vbBlue
        lblThanks(curCredit).ToolTipText = "Click to open " & creditURL
    End If
    lblThanks(curCredit).Visible = True
    
    ReDim Preserve creditList(0 To curCredit) As String
    creditList(curCredit) = creditURL
    
    curCredit = curCredit + 1

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
