VERSION 5.00
Begin VB.Form FormTransparency_Basic 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Add basic transparency"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11820
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
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   788
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   1323
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.smartOptionButton optAlpha 
      Height          =   375
      Index           =   0
      Left            =   6240
      TabIndex        =   3
      Top             =   1920
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   661
      Caption         =   "fully opaque"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
   Begin PhotoDemon.smartOptionButton optAlpha 
      Height          =   375
      Index           =   1
      Left            =   6240
      TabIndex        =   4
      Top             =   2400
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   661
      Caption         =   "fully transparent"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.smartOptionButton optAlpha 
      Height          =   375
      Index           =   2
      Left            =   6240
      TabIndex        =   5
      Top             =   2880
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   661
      Caption         =   "partially transparent:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.sliderTextCombo sltConstant 
      Height          =   495
      Left            =   6480
      TabIndex        =   6
      Top             =   3360
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Min             =   1
      Max             =   254
      Value           =   127
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "make image:"
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
      Index           =   0
      Left            =   6000
      TabIndex        =   2
      Top             =   1440
      Width           =   1380
   End
End
Attribute VB_Name = "FormTransparency_Basic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'"Add basic transparency" (e.g. constant alpha channel) interface
'Copyright ©2013-2014 by Tanner Helland
'Created: 13/August/13
'Last updated: 21/August/13
'Last update: moved "make color transparent" to its own form.  This dialog is now much simpler.
'
'PhotoDemon has long provided the ability to convert a 24bpp image to 32bpp, but the lack of an interface meant it could
' only add a fully opaque alpha channel.  Now the user can select from one of several conversion methods.
'
'This dialog deals with the most obvious conversion method: setting a constant alpha value for the entire image.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'OK button
Private Sub cmdBar_OKClick()
    Process "Add alpha channel", , buildParams(getRelevantAlpha()), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltConstant.Value = 127
End Sub

Private Sub Form_Activate()
    
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
    'Render a preview of the alpha effect
    updatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub optAlpha_Click(Index As Integer)
    updatePreview
End Sub

'Convert a DIB from 24bpp to 32bpp, using a constant alpha channel (specified by the user)
Public Sub simpleConvert32bpp(Optional ByVal convertConstant As Long = 255, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)

    If Not toPreview Then Message "Adding new alpha channel to image..."
    
    'Call prepImageData, which will prepare a temporary copy of the image
    Dim tmpSA As SAFEARRAY2D
    prepImageData tmpSA, toPreview, dstPic
    
    'Pretty simple - ask pdDIB to apply a constant alpha channel to the image, and we're done!
    workingDIB.convertTo32bpp convertConstant
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
    
End Sub

'Note that if the user is moving this slider, they presumably want the corresponding option button selected
Private Sub sltConstant_Change()
    If Not optAlpha(2) Then optAlpha(2).Value = True
    updatePreview
End Sub

'Translate the current option button selection into a relevant alpha value
Private Function getRelevantAlpha() As Long

    Dim convertConstant As Long
    If optAlpha(0) Then
        convertConstant = 255
    ElseIf optAlpha(1) Then
        convertConstant = 0
    Else
        convertConstant = sltConstant.Value
    End If
    
    getRelevantAlpha = convertConstant

End Function

'Render a new preview
Private Sub updatePreview()
    If cmdBar.previewsAllowed Then simpleConvert32bpp getRelevantAlpha(), True, fxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

