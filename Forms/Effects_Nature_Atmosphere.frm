VERSION 5.00
Begin VB.Form FormAtmosphere 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Atmosphere"
   ClientHeight    =   6540
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
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   808
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdDropDown cboBlendMode 
      Height          =   975
      Left            =   6000
      TabIndex        =   4
      Top             =   3480
      Width           =   5895
      _extentx        =   10398
      _extenty        =   1720
      caption         =   "blend mode"
   End
   Begin PhotoDemon.pdButtonStrip btsStyle 
      Height          =   975
      Left            =   6000
      TabIndex        =   3
      Top             =   1320
      Width           =   5895
      _extentx        =   10398
      _extenty        =   1720
      caption         =   "style"
   End
   Begin PhotoDemon.pdSlider sltIntensity 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   2520
      Width           =   5880
      _extentx        =   10372
      _extenty        =   1270
      caption         =   "intensity"
      max             =   100
      value           =   50
      defaultvalue    =   50
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _extentx        =   9922
      _extenty        =   9922
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   12120
      _extentx        =   21378
      _extenty        =   1323
   End
End
Attribute VB_Name = "FormAtmosphere"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Blacklight Form
'Copyright 2001-2017 by Tanner Helland
'Created: some time 2001
'Last updated: 01/October/13
'Last update: use a floating-point slider for more precise results
'
'I found this effect on accident, and it has gradually become one of my favorite effects.
' Visually stunning on many photographs.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'To improve preview performance, a persistent preview DIB is cached locally
Private m_EffectDIB As pdDIB

'Apply a hazy, cool color transformation I call an "atmospheric" transform.
Public Sub ApplyAtmosphereEffect(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)

    If (Not toPreview) Then Message "Creating artificial atmosphere..."
    
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.SetParamString effectParams
    
    Dim atmIntensity As Double, atmStyle As Long, atmBlend As LAYER_BLENDMODE
    
    With cParams
        atmBlend = .GetLong("blendmode", BL_OVERLAY)
        atmStyle = .GetLong("style", 0)
        atmIntensity = .GetDouble("intensity", 50#)
    End With
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte
    Dim tmpSA As SAFEARRAY2D, tmpSA1D As SAFEARRAY1D
    EffectPrep.PrepImageData tmpSA, toPreview, dstPic
    
    'Create a copy of the working data
    If (m_EffectDIB Is Nothing) Then Set m_EffectDIB = New pdDIB
    m_EffectDIB.CreateFromExistingDIB workingDIB
    
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim quickVal As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'Color and grayscale variables
    Dim r As Long, g As Long, b As Long, grayVal As Long
    Dim newR As Long, newG As Long, newB As Long
    
    'Loop through each pixel in the image, converting values as we go
    initX = initX * qvDepth
    finalX = finalX * qvDepth
    
    For y = initY To finalY
        m_EffectDIB.WrapArrayAroundScanline imageData, tmpSA1D, y
    For x = initX To finalX Step qvDepth
        
        'Get the source pixel color values
        b = imageData(x)
        g = imageData(x + 1)
        r = imageData(x + 2)
        
        If (atmStyle = 0) Then
            newR = (g + b) * 0.5
            newG = (r + b) * 0.5
            newB = (r + g) * 0.5
        ElseIf (atmStyle = 1) Then
            
            grayVal = Colors.GetHQLuminance(r, g, b)
        
            newR = g + b - grayVal
            newG = newR + b - grayVal
            newB = newR + newG - grayVal
        
            If (newR > 255) Then newR = 255
            If (newR < 0) Then newR = 0
            If (newG > 255) Then newG = 255
            If (newG < 0) Then newG = 0
            If (newB > 255) Then newB = 255
            If (newB < 0) Then newB = 0
            
        End If
        
        'Assign that gray value to each color channel.  (Yes, channel order is reversed - that's deliberate!)
        imageData(x) = newB
        imageData(x + 1) = newG
        imageData(x + 2) = newR
        
    Next x
        If (Not toPreview) Then
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal y
            End If
        End If
    Next y
    
    'Safely deallocate imageData()
    m_EffectDIB.UnwrapArrayFromDIB imageData
    
    'Merge the results
    m_EffectDIB.SetAlphaPremultiplication True
    workingDIB.SetAlphaPremultiplication True
    
    Dim cCompositor As pdCompositor
    Set cCompositor = New pdCompositor
    cCompositor.QuickMergeTwoDibsOfEqualSize workingDIB, m_EffectDIB, atmBlend, atmIntensity
    
    'On non-previews, free our intermediate copy
    If (Not toPreview) Then Set m_EffectDIB = Nothing
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic, True
    
End Sub

Private Sub btsStyle_Click(ByVal buttonIndex As Long)
    UpdatePreview
End Sub

Private Sub cboBlendMode_Click()
    UpdatePreview
End Sub

Private Sub cmdBar_OKClick()
    Process "Atmosphere", , GetLocalParamString(), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    cboBlendMode.ListIndex = BL_OVERLAY
End Sub

Private Sub Form_Load()
    
    cmdBar.MarkPreviewStatus False
    
    btsStyle.AddItem "global", 0
    btsStyle.AddItem "local", 1
    btsStyle.ListIndex = 0
    
    Interface.PopulateBlendModeDropDown cboBlendMode, BL_OVERLAY
    
    ApplyThemeAndTranslations Me
    cmdBar.MarkPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

'Update the preview whenever the combination slider/text control has its value changed
Private Sub sltIntensity_Change()
    UpdatePreview
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then Me.ApplyAtmosphereEffect GetLocalParamString(), True, pdFxPreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    
    With cParams
        .AddParam "blendmode", cboBlendMode.ListIndex
        .AddParam "intensity", sltIntensity.Value
        .AddParam "style", btsStyle.ListIndex
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
