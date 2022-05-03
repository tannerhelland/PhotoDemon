VERSION 5.00
Begin VB.Form FormFillContentAware 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Content-aware fill"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9705
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   420
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   647
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdSlider sldOptions 
      Height          =   975
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   1720
      Caption         =   "patch test size"
      Min             =   4
      Max             =   50
      Value           =   20
      NotchPosition   =   2
      NotchValueCustom=   20
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5550
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdSlider sldOptions 
      Height          =   975
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   1720
      Caption         =   "random patch candidates"
      Min             =   5
      Max             =   200
      Value           =   60
      NotchPosition   =   2
      NotchValueCustom=   60
   End
   Begin PhotoDemon.pdSlider sldOptions 
      Height          =   975
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   3360
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   1720
      Caption         =   "refinement (percent)"
      Max             =   99
      Value           =   50
      NotchPosition   =   2
      NotchValueCustom=   50
   End
   Begin PhotoDemon.pdSlider sldOptions 
      Height          =   975
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Top             =   4440
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   1720
      Caption         =   "match quality threshold"
      Min             =   1
      Max             =   100
      Value           =   15
      NotchPosition   =   2
      NotchValueCustom=   15
   End
   Begin PhotoDemon.pdSlider sldOptions 
      Height          =   975
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   1720
      Caption         =   "search radius"
      Min             =   5
      Max             =   500
      Value           =   200
      NotchPosition   =   2
      NotchValueCustom=   200
   End
End
Attribute VB_Name = "FormFillContentAware"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Content-Aware Fill (aka "Heal Selection" in some software) Settings Dialog
'Copyright 2022-2022 by Tanner Helland
'Created: 03/May/22
'Last updated: 03/May/22
'Last update: initial build
'
'Content-aware fill was added in PhotoDemon 9.0.  This simple dialog serves a simple purpose:
' allowing the user to modify various content-aware settings.  When OK is pressed, those settings
' are forwarded to an instance of the pdInpaint class, which performs the actual content-aware fill.
' Please review that class for further details on the algorithm and how it works.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'These constants are copied directly from pdInpaint
Private Const MAX_NEIGHBORS_DEFAULT As Long = 20
Private Const COMPARE_RADIUS_DEFAULT As Long = 200
Private Const RANDOM_CANDIDATES_DEFAULT As Long = 60
Private Const REFINEMENT_DEFAULT As Double = 0.5
Private Const ALLOW_OUTLIERS_DEFAULT As Double = 0.15

Private Sub cmdBar_OKClick()
    
    'Place all settings in an XML string
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    With cParams
        .AddParam "search-radius", sldOptions(0).Value
        .AddParam "patch-size", sldOptions(1).Value
        .AddParam "random-candidates", sldOptions(2).Value
        .AddParam "refinement", sldOptions(3).Value / 100#
        .AddParam "allow-outliers", sldOptions(4).Value / 100#
    End With
    
    Processor.Process "Content-aware fill", False, cParams.GetParamString(), UNDO_Layer
    
End Sub

'I'm not sure that randomize serves any purpose on this dialog, but as I don't have a way to hide that button at
' present, simply randomize the width/height to +/- the current image's width/height divided by two.
Private Sub cmdBar_RandomizeClick()
    'TODO
End Sub

Private Sub cmdBar_ResetClick()
    sldOptions(0).Value = COMPARE_RADIUS_DEFAULT
    sldOptions(1).Value = MAX_NEIGHBORS_DEFAULT
    sldOptions(2).Value = RANDOM_CANDIDATES_DEFAULT
    sldOptions(3).Value = REFINEMENT_DEFAULT * 100#
    sldOptions(4).Value = ALLOW_OUTLIERS_DEFAULT * 100#
End Sub

Private Sub Form_Load()
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub
