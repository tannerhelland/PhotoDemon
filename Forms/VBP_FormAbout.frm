VERSION 5.00
Begin VB.Form FormAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " About PhotoDemon"
   ClientHeight    =   8115
   ClientLeft      =   2340
   ClientTop       =   1875
   ClientWidth     =   9030
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
   Picture         =   "VBP_FormAbout.frx":058A
   ScaleHeight     =   541
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   602
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   465
      Left            =   7635
      TabIndex        =   0
      Top             =   7560
      Width           =   1245
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
      Caption         =   "PhotoDemon would not be possible without the following individuals. My sincerest thanks goes out to: "
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   0
      Left            =   240
      MouseIcon       =   "VBP_FormAbout.frx":A6C2
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2880
      Width           =   7335
   End
   Begin VB.Label lblDisclaimer 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " Copyright (automatically populated at run-time)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   3000
      TabIndex        =   2
      Top             =   2400
      Width           =   5985
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version (automatically populated at run-time)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   3900
   End
End
Attribute VB_Name = "FormAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'About Form
'Copyright ©2000-2012 by Tanner Helland
'Created: 6/12/01
'Last updated: 25/June/12
'Last update: display the linked URL as the tooltip text as well
'
'A simple "about"/credits form.  Contains credits, copyright, and the program logo.
'
'***************************************************************************

Option Explicit

Dim creditList() As String
Dim curCredit As Long

Private Sub CmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    'Automatic generation of version & copyright information
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblDisclaimer.Caption = App.LegalCopyright & "   "
    
    curCredit = 1
    
    'Shout-outs to other programmers who provided various resources
    GenerateThankyou "Kroc of camendesign.com for many suggestions regarding UI design and organization", "http://camendesign.com"
    GenerateThankyou "chrfb of deviantart.com for PhotoDemon's icon ('Ecqlipse 2,' CC-BY-NC-SA-3.0)", "http://chrfb.deviantart.com/art/quot-ecqlipse-2-quot-PNG-59941546"
    GenerateThankyou "Juned Chhipa for the 'jcButton 1.7' customizable command button replacement control used on the left-hand toolbar", "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=71482&lngWId=1"
    GenerateThankyou "Steve McMahon for an excellent CommonDialog interface, accelerator key handler, and progress bar replacement", "http://www.vbaccelerator.com/home/VB/index.asp"
    GenerateThankyou "Floris van de Berg and Hervé Drolon for the FreeImage library, and Carsten Klein for the VB interface", "http://freeimage.sourceforge.net/"
    GenerateThankyou "Alfred Koppold for native-VB PCX import/export and PNG import interfaces", "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=56537&lngWId=1"
    GenerateThankyou "John Korejwa for his native-VB JPEG encoding class", "http://planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=50065&lngWId=1"
    GenerateThankyou "Brad Martinez for the original implementation of VB binary file extraction", "http://btmtz.mvps.org/gfxfromfrx/"
    GenerateThankyou "Ron van Tilburg for a native-VB implementation of Xiaolin Wu's line antialiasing routine", "http://www.planet-source-code.com/vb/scripts/showcode.asp?txtCodeId=71370&lngWid=1"
    GenerateThankyou "Jason Bullen for a native-VB implementation of knot-based cubic spline interpolation", "http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=11488&lngWId=1"
    GenerateThankyou "Dosadi for the EZTW32 scanner/digital camera library", "http://eztwain.com/eztwain1.htm"
    GenerateThankyou "Jean-Loup Gailly and Mark Adler for the zLib compression library, and Gilles Vollant for the WAPI wrapper", "http://www.winimage.com/zLibDll/index.html"
    GenerateThankyou "Waty Thierry for many insights regarding printer interfacing in VB", "http://www.ppreview.net/"
    GenerateThankyou "Manuel Augusto Santos for original versions of the 'Enhanced 2-bit Color Reduction' and 'Artistic Contour' algorithms", "http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=26303&lngWId=1"
    GenerateThankyou "Johannes B for the original version of the 'Fog' algorithm", "http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=42642&lngWId=1"
    GenerateThankyou "LaVolpe for his automated VB6 Manifest Creator tool", "http://www.vbforums.com/showthread.php?t=606736"
    GenerateThankyou "Leandro Ascierto for a clean, lightweight class that adds PNGs to menu items", "http://leandroascierto.com/blog/clsmenuimage/"
    GenerateThankyou "Everaldo and The Crystal Project for menu and button icons (LGPL-licensed, click for details)", "http://www.everaldo.com/crystal/"
    GenerateThankyou "The public-domain Tango Icon Library for menu and button icons", "http://tango.freedesktop.org/Tango_Icon_Library"
    
    lblThanks(0).MousePointer = vbDefault
    
    'Assign the system hand cursor to all relevant objects
    setHandCursorForAll Me
    
End Sub

'Generate a label with the specified "thank you" text, and link it to the specified URL
Private Sub GenerateThankyou(ByVal thxText As String, Optional ByVal creditURL As String = "")
    
    'Generate a new label
    Load lblThanks(curCredit)
    
    lblThanks(curCredit).Top = lblThanks(curCredit - 1).Top + lblThanks(curCredit - 1).Height + 4
    lblThanks(curCredit).Left = lblThanks(0).Left + 2
    lblThanks(curCredit).Caption = thxText
    If creditURL = "" Then
        lblThanks(curCredit).MousePointer = vbDefault
    Else
        lblThanks(curCredit).FontUnderline = True
        lblThanks(curCredit).ForeColor = vbBlue
        lblThanks(curCredit).ToolTipText = creditURL
    End If
    lblThanks(curCredit).Visible = True
    
    ReDim Preserve creditList(0 To curCredit) As String
    creditList(curCredit) = creditURL
    
    curCredit = curCredit + 1

End Sub

'When a thank-you credit is clicked, launch the corresponding website
Private Sub lblThanks_Click(Index As Integer)

    If creditList(Index) <> "" Then ShellExecute FormMain.HWnd, "Open", creditList(Index), "", 0, SW_SHOWNORMAL

End Sub
