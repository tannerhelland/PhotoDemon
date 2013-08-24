VERSION 5.00
Begin VB.Form FormSharpen 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Sharpen"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12030
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
   ScaleWidth      =   802
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   3
      Top             =   5790
      Width           =   12030
      _ExtentX        =   21220
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
   Begin PhotoDemon.sliderTextCombo sltStrength 
      Height          =   495
      Left            =   6000
      TabIndex        =   2
      Top             =   2760
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Min             =   0.1
      SigDigits       =   1
      Value           =   0.1
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
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "strength:"
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
      TabIndex        =   0
      Top             =   2400
      Width           =   960
   End
End
Attribute VB_Name = "FormSharpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Sharpen Tool
'Copyright ©2012-2013 by Tanner Helland
'Created: 09/August/13 (actually, a naive version was built years ago, but didn't offer variable strength)
'Last updated: 22/August/13
'Last update: rewrote the DoFilter call against the new paramString implementation
'
'Basic sharpening tool.  A 3x3 convolution kernel is used to apply the sharpening, so the results will
' be inferior to Unsharp Masking - but the tool is much simpler, and for light sharpening, the results are
' often acceptable.
'
'The bulk of the work happens in the DoFilter routine that handles all of PhotoDemon's generic convolution
' work.  All this dialog does is set up the kernel, then pass it on to DoFilter.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

'When previewing, we need to modify the strength to be representative of the final filter.  This means dividing by the
' original image dimensions in order to establish the right ratio.
Dim iWidth As Long, iHeight As Long

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'Convolve an image using a gaussian kernel (separable implementation!)
'Input: radius of the blur (min 1, no real max - but the scroll bar is maxed at 200 presently)
Public Sub ApplySharpenFilter(ByVal sStrength As Double, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    'Sharpening uses a basic 3x3 convolution filter, which we generate dynamically based on the requested strength
    Dim tmpString As String
    
    'Start with a filter name
    tmpString = g_Language.TranslateMessage("sharpen") & "|"
    
    'Next comes an invert parameter (not used for sharpening)
    tmpString = tmpString & "0|"
    
    'Next is the divisor and offset
    tmpString = tmpString & "1|0|"
    
    'And finally, the convolution array itself
    tmpString = tmpString & "0|0|0|0|0|"
    tmpString = tmpString & "0|0|" & CStr(-sStrength) & "|0|0|"
    tmpString = tmpString & "0|" & CStr(-sStrength) & "|" & CStr(sStrength * 4 + 1) & "|" & CStr(-sStrength) & "|0|"
    tmpString = tmpString & "0|0|" & CStr(-sStrength) & "|0|0|"
    tmpString = tmpString & "0|0|0|0|0"
    
    'Pass our new parameter string to the main convolution filter function
    DoFilter tmpString, toPreview, dstPic
                
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Sharpen", , CStr(sltStrength)
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub Form_Activate()

    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
    'Draw a preview of the effect
    cmdBar.markPreviewStatus True
    updatePreview
    
End Sub

Private Sub Form_Load()
    
    'Disable previews until the dialog is completely ready
    cmdBar.markPreviewStatus False
    
    'Note the current image's width and height, which will be needed to adjust the preview effect
    If pdImages(CurrentImage).selectionActive Then
        iWidth = pdImages(CurrentImage).mainSelection.boundWidth
        iHeight = pdImages(CurrentImage).mainSelection.boundHeight
    Else
        iWidth = pdImages(CurrentImage).Width
        iHeight = pdImages(CurrentImage).Height
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub updatePreview()
    If cmdBar.previewsAllowed Then ApplySharpenFilter sltStrength.Value, True, fxPreview
End Sub

Private Sub sltStrength_Change()
    updatePreview
End Sub
