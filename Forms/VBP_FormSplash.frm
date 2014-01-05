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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   551
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   779
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Live updates..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   330
      Left            =   285
      TabIndex        =   0
      Top             =   7920
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
'Copyright ©2001-2014 by Tanner Helland
'Created: 15/April/01
'Last updated: 13/September/13
'Last update: logos are now stored in the resource file.  No more picture box placeholders!
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Public Sub prepareSplash()

    'Before we can display the splash screen, we need to paint the program logo to it.  (This is done to
    ' guarantee proper painting regardless of screen DPI.)
    Dim logoLayer As pdLayer
    Set logoLayer = New pdLayer
    If loadResourceToLayer("PDLOGO", logoLayer) Then
    
        Dim logoAspectRatio As Double
        logoAspectRatio = CDbl(logoLayer.getLayerWidth) / CDbl(logoLayer.getLayerHeight)
        
        SetStretchBltMode Me.hDC, STRETCHBLT_HALFTONE
        StretchBlt Me.hDC, 0, 0, Me.ScaleWidth, Me.ScaleWidth / logoAspectRatio, logoLayer.getLayerDC, 0, 0, logoLayer.getLayerWidth, logoLayer.getLayerHeight, vbSrcCopy
        Me.Picture = Me.Image
        
    End If
    
End Sub

