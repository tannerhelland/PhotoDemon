VERSION 5.00
Begin VB.Form FormChannelMixer 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Channel Mixer"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12090
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
   ScaleWidth      =   806
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbChannel 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      ItemData        =   "VBP_FormChannelMixer.frx":0000
      Left            =   6120
      List            =   "VBP_FormChannelMixer.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   480
      Width           =   5820
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   9120
      TabIndex        =   0
      Top             =   5910
      Width           =   1365
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   10590
      TabIndex        =   1
      Top             =   5910
      Width           =   1365
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.sliderTextCombo sltRed 
      Height          =   495
      Left            =   6000
      TabIndex        =   7
      Top             =   1260
      Width           =   6015
      _ExtentX        =   10398
      _ExtentY        =   873
      Min             =   -200
      Max             =   200
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
   Begin PhotoDemon.sliderTextCombo sltGreen 
      Height          =   495
      Left            =   6000
      TabIndex        =   8
      Top             =   2160
      Width           =   6015
      _ExtentX        =   10398
      _ExtentY        =   873
      Min             =   -200
      Max             =   200
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
   Begin PhotoDemon.sliderTextCombo sltBlue 
      Height          =   495
      Left            =   6000
      TabIndex        =   9
      Top             =   3060
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   873
      Min             =   -200
      Max             =   200
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
   Begin PhotoDemon.smartCheckBox chkMonochrome 
      Height          =   570
      Left            =   6120
      TabIndex        =   10
      Top             =   5040
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   1005
      Caption         =   "monochrome"
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
      Left            =   6000
      TabIndex        =   11
      Top             =   3960
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   873
      Min             =   -255
      Max             =   255
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
   Begin PhotoDemon.smartCheckBox chkLuminance 
      Height          =   570
      Left            =   8640
      TabIndex        =   13
      Top             =   5040
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   1005
      Caption         =   "preserve luminance"
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "other options:"
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
      Left            =   6000
      TabIndex        =   16
      Top             =   4680
      Width           =   1500
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      X1              =   392
      X2              =   800
      Y1              =   304
      Y2              =   304
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "output channel:"
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
      TabIndex        =   15
      Top             =   120
      Width           =   1665
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "constant:"
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
      Index           =   4
      Left            =   6000
      TabIndex        =   12
      Top             =   3630
      Width           =   975
   End
   Begin VB.Label lblBackground 
      Height          =   855
      Left            =   0
      TabIndex        =   5
      Top             =   5760
      Width           =   12135
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "blue:"
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
      Index           =   3
      Left            =   6000
      TabIndex        =   4
      Top             =   2730
      Width           =   540
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "green:"
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
      Index           =   2
      Left            =   6000
      TabIndex        =   3
      Top             =   1830
      Width           =   690
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "red:"
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
      Index           =   1
      Left            =   6000
      TabIndex        =   2
      Top             =   930
      Width           =   435
   End
End
Attribute VB_Name = "FormChannelMixer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Channel Mixer Form
'Copyright ©2012-2013 by audioglider and Tanner Helland
'Created: 08/June/13
'Last updated: 09/June/13
'Last update: added "preserve luminance" and applied a few final fixes
'
'Many thanks to talented contributer audioglider for creating this tool.
'
'Standard channel mixer dialog.  Layout and feature set derived from comparable tools in Photoshop and GIMP.
' Per convention, all channels can be modified simultaneously.  For convenience, a "constant" slider is also
' provided, allowing for simple uniform adjustments.
'
'A "monochrome" option is provided for outputting a grayscale image.  Monochrome values are stored separately, so
' any changes made while in monochrome mode will not overwrite existing color channel values.
'
'A "preserve luminance" option is provided for applying color changes without changing the overall composition
' of the photo.  This is disabled when "monochrome" is active (obviously, as otherwise the gray values would
' never change!)
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

Private Enum OutputChannel
    RedOutput = 0
    GreenOutput = 1
    BlueOutput = 2
    GrayOutput = 3
End Enum

Private Enum InputChannel
    RedInput = 0
    GreenInput = 1
    BlueInput = 2
    ConstantInput = 3
End Enum

'Because all channels can be modified independently, we need to store the settings of each channel.
' First dim: output channel (red/green/blue/gray)
' Second dim: input channel (red/green/blue/constant value)
Dim curSliderValues(0 To 3, 0 To 3) As Long

Dim forbidUpdate As Boolean

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

Private Sub chkLuminance_Click()
    updatePreview
End Sub

'To match GIMP's behavior (which is actually well-designed in this case), disable the output combo box
Private Sub chkMonochrome_Click()
    
    If CBool(chkMonochrome) Then
        
        chkLuminance.Enabled = False
        cmbChannel.Enabled = False
        
        'Populate the sliders with any previously saved values
        forbidUpdate = True
        sltRed.Value = curSliderValues(GrayOutput, RedInput)
        sltGreen.Value = curSliderValues(GrayOutput, GreenInput)
        sltBlue.Value = curSliderValues(GrayOutput, BlueInput)
        sltConstant.Value = curSliderValues(GrayOutput, ConstantInput)
        forbidUpdate = False
                
    Else
    
        chkLuminance.Enabled = True
        cmbChannel.Enabled = True
        
        'Populate the sliders with any previously saved values
        forbidUpdate = True
        sltRed.Value = curSliderValues(cmbChannel.ListIndex, RedInput)
        sltGreen.Value = curSliderValues(cmbChannel.ListIndex, GreenInput)
        sltBlue.Value = curSliderValues(cmbChannel.ListIndex, BlueInput)
        sltConstant.Value = curSliderValues(cmbChannel.ListIndex, ConstantInput)
        forbidUpdate = False
        
    End If
    
    updatePreview
    
End Sub

Private Sub cmbChannel_Click()
    
    'Populate the sliders with any previously saved values
    forbidUpdate = True
    sltRed.Value = curSliderValues(cmbChannel.ListIndex, RedInput)
    sltGreen.Value = curSliderValues(cmbChannel.ListIndex, GreenInput)
    sltBlue.Value = curSliderValues(cmbChannel.ListIndex, BlueInput)
    sltConstant.Value = curSliderValues(cmbChannel.ListIndex, ConstantInput)
    forbidUpdate = False
    
    updatePreview
    
End Sub

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub CmdOK_Click()
    
    'Validate all textbox entries
    If sltRed.IsValid And sltGreen.IsValid And sltBlue.IsValid And sltConstant.IsValid Then
        Me.Visible = False
        updateStoredValues
        Process "Channel mixer", , createChannelParamString()
        Unload Me
    End If
    
End Sub

'Apply a new channel mixer to the image
' Inputs:
'  - all modifiers as one long string; see "createChannelParamString" for how this string is assembled
Public Sub ApplyChannelMixer(ByVal channelMixerParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If toPreview = False Then Message "Mixing color channels..."
    
    'Because this tool has so many parameters, they are condensed into a single string and passed here.  We need to
    ' parse out individual values before continuing.
    Dim cParams As pdParamString
    Set cParams = New pdParamString
    cParams.setParamString channelMixerParams
    
    Dim channelModifiers(0 To 3, 0 To 3) As Double
    Dim x As Long, y As Long
    For x = 0 To 3
        For y = 0 To 3
            'The "constant" modifier is added to the final channel value as a whole number, but the other values are
            ' used as multiplication factors - so divide them by 100.
            If y = 3 Then
                channelModifiers(x, y) = cParams.GetLong((x * 4) + y + 1)
            Else
                channelModifiers(x, y) = CDbl(cParams.GetLong((x * 4) + y + 1)) / 100
            End If
        Next y
    Next x
    
    Dim isMonochrome As Boolean
    isMonochrome = cParams.GetBool(17)
    
    Dim preserveLuminance As Boolean
    preserveLuminance = cParams.GetBool(18)
        
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    
    prepImageData tmpSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Finally, a bunch of variables used in color calculation
    Dim r As Long, g As Long, b As Long
    Dim newR As Long, newG As Long, newB As Long, newGray As Long
    Dim h As Double, s As Double, l As Double
    Dim originalLuminance As Double
        
    'Apply the filter
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
        
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Create a new value for each color based on the input parameters
        If isMonochrome Then
            newGray = r * channelModifiers(3, 0) + g * channelModifiers(3, 1) + b * channelModifiers(3, 2) + channelModifiers(3, 3)
            
            If newGray > 255 Then newGray = 255
            If newGray < 0 Then newGray = 0
            
            'Note: luminance preservation serves no purpose when monochrome is selected, so I do not process it here
            
            ImageData(QuickVal + 2, y) = newGray
            ImageData(QuickVal + 1, y) = newGray
            ImageData(QuickVal, y) = newGray
            
        Else
        
            'If luminance is being preserved, we need to determine the initial luminance value
            If preserveLuminance Then originalLuminance = (getLuminance(r, g, b) / 255)
        
            newR = r * channelModifiers(0, 0) + g * channelModifiers(0, 1) + b * channelModifiers(0, 2) + channelModifiers(0, 3)
            newG = r * channelModifiers(1, 0) + g * channelModifiers(1, 1) + b * channelModifiers(1, 2) + channelModifiers(1, 3)
            newB = r * channelModifiers(2, 0) + g * channelModifiers(2, 1) + b * channelModifiers(2, 2) + channelModifiers(3, 3)
            
            'Fit everything in the [0, 255] range
            If newR > 255 Then newR = 255
            If newR < 0 Then newR = 0
            If newG > 255 Then newG = 255
            If newG < 0 Then newG = 0
            If newB > 255 Then newB = 255
            If newB < 0 Then newB = 0
            
            'If the user wants us to preserve luminance, determine the hue and saturation of the new color, then replace the luminance
            ' value with the original
            If preserveLuminance Then
                tRGBToHSL newR, newG, newB, h, s, l
                tHSLToRGB h, s, originalLuminance, newR, newG, newB
            End If
            
            ImageData(QuickVal + 2, y) = newR
            ImageData(QuickVal + 1, y) = newG
            ImageData(QuickVal, y) = newB
            
        End If
                
    Next y
        If toPreview = False Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x
            End If
        End If
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
    
End Sub

Private Sub Form_Activate()
        
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
    'Fill the "stored value" array with default settings appropriate to each channel; basically, set each channel
    ' to their current value (e.g. red = "red 100%", "0%" for all other channels)
    Dim i As Long
    For i = 0 To 3
    
        Select Case i
        
            Case RedOutput
                curSliderValues(RedOutput, RedInput) = 100
                curSliderValues(RedOutput, GreenInput) = 0
                curSliderValues(RedOutput, BlueInput) = 0
                curSliderValues(RedOutput, ConstantInput) = 0
            
            Case GreenOutput
                curSliderValues(GreenOutput, RedInput) = 0
                curSliderValues(GreenOutput, GreenInput) = 100
                curSliderValues(GreenOutput, BlueInput) = 0
                curSliderValues(GreenOutput, ConstantInput) = 0
            
            Case BlueOutput
                curSliderValues(BlueOutput, RedInput) = 0
                curSliderValues(BlueOutput, GreenInput) = 0
                curSliderValues(BlueOutput, BlueInput) = 100
                curSliderValues(BlueOutput, ConstantInput) = 0
                
            'I'm not sure the best preset values to suggest for gray; for now, I'm defaulting to the ITU standard
            ' conversion formula - that should provide a good starting point for user modifications.
            Case GrayOutput
                curSliderValues(GrayOutput, RedInput) = 21
                curSliderValues(GrayOutput, GreenInput) = 72
                curSliderValues(GrayOutput, BlueInput) = 7
                curSliderValues(GrayOutput, ConstantInput) = 0
        
        End Select
    
    Next i
    
    'Per convention, monochrome mode is handled via a separate checkbox.  This is also an easier solution for us, as
    ' it's difficult to apply changes to an imaginary "gray channel" (we'd have to divvy up any "gray channel"
    ' changes to each of red, green, and blue, and without a consistent way to do that the results would be
    ' unpredictable - I'm fairly certain this is why Photoshop etc. provide a separate "monochrome" checkbox)
    cmbChannel.Clear
    cmbChannel.AddItem " red", 0
    cmbChannel.AddItem " green", 1
    cmbChannel.AddItem " blue", 2
    cmbChannel.ListIndex = 0
    
    'To account for translation width possibilities, align the monochrome and luminance check boxes manually
    chkLuminance.Left = chkMonochrome.Left + chkMonochrome.Width + 24
    
    'Display the previewed effect in the neighboring window
    updatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub sltBlue_Change()
    If Not forbidUpdate Then
        updateStoredValues
        updatePreview
    End If
End Sub

Private Sub sltConstant_Change()
    If Not forbidUpdate Then
        updateStoredValues
        updatePreview
    End If
End Sub

Private Sub sltGreen_Change()
    If Not forbidUpdate Then
        updateStoredValues
        updatePreview
    End If
End Sub

Private Sub sltRed_Change()
    If Not forbidUpdate Then
        updateStoredValues
        updatePreview
    End If
End Sub

Private Sub updatePreview()
    ApplyChannelMixer createChannelParamString(), True, fxPreview
End Sub

'Because the user can change multiple channels at once, we need to store all current channel values in memory.
Private Sub updateStoredValues()

    'Store values according to the current combo box or monochrome setting
    If CBool(chkMonochrome) Then
        curSliderValues(GrayOutput, RedInput) = sltRed.Value
        curSliderValues(GrayOutput, GreenInput) = sltGreen.Value
        curSliderValues(GrayOutput, BlueInput) = sltBlue.Value
        curSliderValues(GrayOutput, ConstantInput) = sltConstant.Value
    Else
        curSliderValues(cmbChannel.ListIndex, RedInput) = sltRed.Value
        curSliderValues(cmbChannel.ListIndex, GreenInput) = sltGreen.Value
        curSliderValues(cmbChannel.ListIndex, BlueInput) = sltBlue.Value
        curSliderValues(cmbChannel.ListIndex, ConstantInput) = sltConstant.Value
    End If

End Sub

'Because this tool has a complex set of input values, we need to condense them all into a single string.
' This function handles the creation of that string for both previews and full-image applications.
Private Function createChannelParamString() As String

    Dim paramString As String
    paramString = ""
    
    'Start by adding all channel input values to the string
    Dim i As Long, j As Long
    For i = 0 To 3
        For j = 0 To 3
            paramString = paramString & CStr(curSliderValues(i, j)) & "|"
        Next j
    Next i
    
    'Next, add the monochrome checkbox value
    paramString = paramString & CStr(CBool(chkMonochrome)) & "|"
    
    'Finally, add the preserve luminance checkbox value
    paramString = paramString & CStr(CBool(chkLuminance))
    
    createChannelParamString = paramString

End Function

