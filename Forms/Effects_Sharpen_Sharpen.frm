VERSION 5.00
Begin VB.Form FormSharpen 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Sharpen"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11580
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
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   772
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   11580
      _ExtentX        =   20426
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdSlider sltStrength 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   2520
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1244
      Caption         =   "strength"
      Min             =   0.01
      SigDigits       =   2
      Value           =   0.01
      DefaultValue    =   0.01
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
End
Attribute VB_Name = "FormSharpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Sharpen Tool
'Copyright 2013-2026 by Tanner Helland
'Created: 09/August/13 (actually, a naive version was built years ago, but didn't offer variable strength)
'Last updated: 28/July/17
'Last update: performance improvements, migrate to XML params
'
'Basic sharpening tool.  A 3x3 convolution kernel is used to apply the sharpening, so the results will
' be inferior to Unsharp Masking - but the tool is much simpler, and for minor sharpening, the results
' are often acceptable.
'
'The bulk of the work happens in the ApplyConvolutionFilter routine that handles all of PhotoDemon's
' generic convolution tasks.  All this dialog does is set up the kernel, then pass it to that function.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Convolve an image using a gaussian kernel (separable implementation!)
'Input: radius of the blur (min 1, no real max - but the scroll bar is maxed at 200 presently)
Public Sub ApplySharpenFilter(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim sStrength As Double
    sStrength = cParams.GetDouble("strength", 0.01)
    
    'Sharpening uses a basic 3x3 convolution filter, which we generate dynamically based on the requested strength
    Dim cParamsOut As pdSerialize
    Set cParamsOut = New pdSerialize
    
    With cParamsOut
    
        .AddParam "name", g_Language.TranslateMessage("sharpen")
        .AddParam "invert", False
        .AddParam "weight", 1#
        .AddParam "bias", 0#
        
        'And finally, the convolution array itself.  This is just a pipe-delimited string with a 5x5 array
        ' of weights.
        Dim tmpString As String
        tmpString = tmpString & "0|0|0|0|0|"
        tmpString = tmpString & "0|0|" & Trim$(Str$(-sStrength)) & "|0|0|"
        tmpString = tmpString & "0|" & Trim$(Str$(-sStrength)) & "|" & Trim$(Str$(sStrength * 4# + 1#)) & "|" & Trim$(Str$(-sStrength)) & "|0|"
        tmpString = tmpString & "0|0|" & Trim$(Str$(-sStrength)) & "|0|0|"
        tmpString = tmpString & "0|0|0|0|0"
    
        .AddParam "matrix", tmpString
        
    End With
    
    'Pass our new parameter string to the main convolution filter function
    Filters_Area.ApplyConvolutionFilter_XML cParamsOut.GetParamString(), toPreview, dstPic
    
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Sharpen", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub Form_Load()
    cmdBar.SetPreviewStatus False
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then ApplySharpenFilter GetLocalParamString(), True, pdFxPreview
End Sub

Private Sub sltStrength_Change()
    UpdatePreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "strength", sltStrength.Value
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
