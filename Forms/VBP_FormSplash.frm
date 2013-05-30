VERSION 5.00
Begin VB.Form FormSplash 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8265
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   11685
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "VBP_FormSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   551
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   779
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picBackground 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8265
      Left            =   0
      Picture         =   "VBP_FormSplash.frx":000C
      ScaleHeight     =   551
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   779
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   11685
   End
   Begin VB.Label lblMessage 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Live updates..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   330
      Left            =   285
      TabIndex        =   1
      Top             =   7320
      Width           =   11205
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "FormSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Splash Screen
'Copyright ©2001-2013 by Tanner Helland
'Created: 15/April/01
'Last updated: 30/May/13
'Last update: removed all code from here and placed it in the Loading module; also, new splash screen!
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

Public Sub prepareSplash()

    'Before we can display the splash screen, we need to paint the program logo to it.  (This is done to
    ' guarantee proper painting regardless of screen DPI.)
    Dim logoWidth As Long, logoHeight As Long
    Dim logoAspectRatio As Double
    
    logoWidth = picBackground.ScaleWidth
    logoHeight = picBackground.ScaleHeight
    logoAspectRatio = CDbl(logoWidth) / CDbl(logoHeight)
    
    SetStretchBltMode Me.hDC, STRETCHBLT_HALFTONE
    StretchBlt Me.hDC, 0, 0, Me.ScaleWidth, Me.ScaleWidth / logoAspectRatio, picBackground.hDC, 0, 0, logoWidth, logoHeight, vbSrcCopy
    Me.Picture = Me.Image
    
End Sub
