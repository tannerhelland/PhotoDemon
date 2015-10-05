VERSION 5.00
Begin VB.Form FormWhiteBalance 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " White balance"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12120
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   434
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   808
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5760
      Width           =   12120
      _ExtentX        =   21378
      _ExtentY        =   1323
      BackColor       =   14802140
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.sliderTextCombo sltStrength 
      Height          =   720
      Left            =   6000
      TabIndex        =   2
      Top             =   2520
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   1270
      Caption         =   "strength"
      Min             =   0.01
      Max             =   5
      SigDigits       =   2
      Value           =   0.05
   End
End
Attribute VB_Name = "FormWhiteBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'White Balance Handler
'Copyright 2012-2015 by Tanner Helland
'Created: 03/July/12
'Last updated: 24/August/13
'Last update: added command bar
'
'White balance handler.  Unlike other programs, which shove this under the Levels dialog as an "auto levels"
' function, I consider it worthy of its own interface.  The reason is - white balance is an important function.
' It's arguably more useful than the Levels dialog, especially to a casual user, because it automatically
' calculates levels according to a reliable, often-accurate algorithm.  Rather than forcing the user through the
' Levels dialog (because really, how many people know that Auto Levels is actually White Balance in photography
' parlance?), PhotoDemon provides a full implementation of custom white balance handling.
'
'The value box on the form is the percentage of pixels ignored at the top and bottom of the histogram.
' 0.05 is the recommended default.  I've specified 5.0 as the maximum, but there's no reason it couldn't be set
' higher... just be forewarned that higher values (obviously) blow out the picture with increasing strength.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'OK button
Private Sub cmdBar_OKClick()
    Process "White balance", , buildParams(sltStrength), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltStrength.Value = 0.05
End Sub

Private Sub Form_Activate()
        
    'Apply translations and visual themes
    MakeFormPretty Me
        
End Sub

'Correct white balance by stretching the histogram and ignoring pixels above or below the 0.05% threshold
Public Sub AutoWhiteBalance(Optional ByVal percentIgnore As Double = 0.05, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)

    If Not toPreview Then Message "Adjusting image white balance..."
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    
    WhiteBalanceDIB percentIgnore, workingDIB, toPreview
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingDIB
    finalizeImageData toPreview, dstPic
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub sltStrength_Change()
    updatePreview
End Sub

Private Sub updatePreview()
    If cmdBar.previewsAllowed Then AutoWhiteBalance sltStrength, True, fxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

