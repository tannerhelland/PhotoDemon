VERSION 5.00
Begin VB.Form FormChannelMixer 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Channel mixer"
   ClientHeight    =   7215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12150
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
   ScaleHeight     =   481
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   810
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   6255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   11033
   End
   Begin PhotoDemon.pdSlider sltRed 
      Height          =   705
      Left            =   6120
      TabIndex        =   2
      Top             =   1680
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "red"
      Min             =   -200
      Max             =   200
      SliderTrackStyle=   3
      GradientColorMiddle=   255
   End
   Begin PhotoDemon.pdSlider sltGreen 
      Height          =   705
      Left            =   6120
      TabIndex        =   3
      Top             =   2520
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "green"
      Min             =   -200
      Max             =   200
      SliderTrackStyle=   3
      GradientColorMiddle=   65280
   End
   Begin PhotoDemon.pdSlider sltBlue 
      Height          =   705
      Left            =   6120
      TabIndex        =   4
      Top             =   3360
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "blue"
      Min             =   -200
      Max             =   200
      SliderTrackStyle=   3
      GradientColorMiddle=   16711680
   End
   Begin PhotoDemon.pdCheckBox chkMonochrome 
      Height          =   375
      Left            =   6120
      TabIndex        =   5
      Top             =   5520
      Width           =   5760
      _ExtentX        =   10160
      _ExtentY        =   582
      Caption         =   "monochrome"
   End
   Begin PhotoDemon.pdSlider sltConstant 
      Height          =   705
      Left            =   6120
      TabIndex        =   6
      Top             =   4200
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "constant"
      Min             =   -255
      Max             =   255
      SliderTrackStyle=   2
   End
   Begin PhotoDemon.pdCheckBox chkLuminance 
      Height          =   375
      Left            =   6120
      TabIndex        =   7
      Top             =   5880
      Width           =   5760
      _ExtentX        =   10160
      _ExtentY        =   582
      Caption         =   "preserve luminance"
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   6465
      Width           =   12150
      _ExtentX        =   21431
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdButtonStrip btsChannel 
      Height          =   960
      Left            =   6000
      TabIndex        =   8
      Top             =   120
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   1058
      Caption         =   "output channel"
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   0
      Left            =   6000
      Top             =   1200
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   503
      Caption         =   "input channel(s)"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   1
      Left            =   6000
      Top             =   5160
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   503
      Caption         =   "options for all channels"
      FontSize        =   12
      ForeColor       =   4210752
   End
End
Attribute VB_Name = "FormChannelMixer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Channel Mixer Form
'Copyright 2014-2026 by Tanner Helland
'Created: 08/June/13
'Last updated: 17/April/20
'Last update: minor perf improvements
'
'Standard channel mixer dialog.  Layout and feature set derived from comparable tools
' in Photoshop and GIMP.
'
'Per convention, all channels can be modified simultaneously.  For convenience,
' a "constant" slider is also provided, allowing for uniform adjustments.
'
'A "monochrome" option is provided for outputting a grayscale image.  Monochrome values
' are stored separately, so any changes made while in monochrome mode do *not* overwrite
' separate color channel modifications.
'
'A "preserve luminance" option is provided for applying color changes without changing
' the overall composition of the photo.  This is disabled when "monochrome" is active
' (obviously, because otherwise the gray values wouldn't change!)
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
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

#If False Then
    Private Const RedOutput = 0, GreenOutput = 1, BlueOutput = 2, GrayOutput = 3
    Private Const RedInput = 0, GreenInput = 1, BlueInput = 2, ConstantInput = 3
#End If

'Because all channels can be modified independently, we need to store the settings of each channel.
' First dim: output channel (red/green/blue/gray)
' Second dim: input channel (red/green/blue/constant value)
Private m_curSliderValues(0 To 3, 0 To 3) As Long

'Sometimes, we need to update all slider values at once (like loading presets from file).  To prevent
' a bazillion redraws as we set individual slider values, we instead use this variable to forcibly
' suspend auto-updates until all UI elements are synched.
Private m_forbidUpdate As Boolean

Private Sub btsChannel_Click(ByVal buttonIndex As Long)

    'Populate the sliders with any previously saved values
    m_forbidUpdate = True
    sltRed.Value = m_curSliderValues(btsChannel.ListIndex, RedInput)
    sltGreen.Value = m_curSliderValues(btsChannel.ListIndex, GreenInput)
    sltBlue.Value = m_curSliderValues(btsChannel.ListIndex, BlueInput)
    sltConstant.Value = m_curSliderValues(btsChannel.ListIndex, ConstantInput)
    m_forbidUpdate = False
    
    UpdatePreview

End Sub

Private Sub chkLuminance_Click()
    UpdatePreview
End Sub

'To match GIMP's behavior (which is actually well-designed in this case), disable the output combo box
Private Sub chkMonochrome_Click()
    
    If chkMonochrome.Value Then
        
        chkLuminance.Enabled = False
        btsChannel.Enabled = False
        
        'Populate the sliders with any previously saved values
        m_forbidUpdate = True
        sltRed.Value = m_curSliderValues(GrayOutput, RedInput)
        sltGreen.Value = m_curSliderValues(GrayOutput, GreenInput)
        sltBlue.Value = m_curSliderValues(GrayOutput, BlueInput)
        sltConstant.Value = m_curSliderValues(GrayOutput, ConstantInput)
        m_forbidUpdate = False
                
    Else
    
        chkLuminance.Enabled = True
        btsChannel.Enabled = True
        
        'Populate the sliders with any previously saved values
        m_forbidUpdate = True
        sltRed.Value = m_curSliderValues(btsChannel.ListIndex, RedInput)
        sltGreen.Value = m_curSliderValues(btsChannel.ListIndex, GreenInput)
        sltBlue.Value = m_curSliderValues(btsChannel.ListIndex, BlueInput)
        sltConstant.Value = m_curSliderValues(btsChannel.ListIndex, ConstantInput)
        m_forbidUpdate = False
        
    End If
    
    UpdatePreview
    
End Sub

'Apply a new channel mixer to the image
' Inputs:
'  - all modifiers as one long string; see "createChannelParamString" for how this string is assembled
Public Sub ApplyChannelMixer(ByVal channelMixerParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    If (Not toPreview) Then Message "Mixing color channels..."
    
    'Parse all relevant parameters from the input XML string
    Dim isMonochrome As Boolean, preserveLuminance As Boolean
    Dim channelModifiers(0 To 3, 0 To 3) As Double
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString channelMixerParams
    
    With cParams
    
        'Start by grabbing the two simple parameters from the list
        isMonochrome = .GetBool("monochrome", chkMonochrome.Value)
        preserveLuminance = .GetBool("preserveluminance", chkLuminance.Value)
        
        'Next, we need to retrieve the 4x4 "grid" of values: four inputs (RGB/Constant) for each of
        ' four possible output channels (RGB/Gray).  For reference, you may want to refer to the
        ' named enums at the top of this module.
        channelModifiers(0, 0) = .GetDouble("RedOutRedIn", 100) / 100#
        channelModifiers(0, 1) = .GetDouble("RedOutGreenIn", 0#) / 100#
        channelModifiers(0, 2) = .GetDouble("RedOutBlueIn", 0#) / 100#
        channelModifiers(0, 3) = .GetDouble("RedOutConstantIn", 0#)
        
        channelModifiers(1, 0) = .GetDouble("GreenOutRedIn", 0#) / 100#
        channelModifiers(1, 1) = .GetDouble("GreenOutGreenIn", 100#) / 100#
        channelModifiers(1, 2) = .GetDouble("GreenOutBlueIn", 0#) / 100#
        channelModifiers(1, 3) = .GetDouble("GreenOutConstantIn", 0#)
        
        channelModifiers(2, 0) = .GetDouble("BlueOutRedIn", 0#) / 100#
        channelModifiers(2, 1) = .GetDouble("BlueOutGreenIn", 0#) / 100#
        channelModifiers(2, 2) = .GetDouble("BlueOutBlueIn", 100#) / 100#
        channelModifiers(2, 3) = .GetDouble("BlueOutConstantIn", 0#)
        
        channelModifiers(3, 0) = .GetDouble("GrayOutRedIn", 0#) / 100#
        channelModifiers(3, 1) = .GetDouble("GrayOutGreenIn", 0#) / 100#
        channelModifiers(3, 2) = .GetDouble("GrayOutBlueIn", 0#) / 100#
        channelModifiers(3, 3) = .GetDouble("GrayOutConstantIn", 100#)
        
    End With
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte, tmpSA As SafeArray2D, tmpSA1D As SafeArray1D
    EffectPrep.PrepImageData tmpSA, toPreview, dstPic
    
    Dim initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left * 4
    initY = curDIBValues.Top
    finalX = curDIBValues.Right * 4
    finalY = curDIBValues.Bottom
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not toPreview) Then ProgressBars.SetProgBarMax finalY
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'Finally, a bunch of variables used in color calculation
    Dim r As Long, g As Long, b As Long
    Dim rFloat As Double, gFloat As Double, bFloat As Double
    Dim newR As Long, newG As Long, newB As Long, newGray As Long
    Dim h As Double, s As Double, l As Double
    Dim originalLuminance As Double
    Const ONE_DIV_255 As Double = 1# / 255#
    
    'Apply the filter
    Dim x As Long, y As Long
    
    For y = initY To finalY
        workingDIB.WrapArrayAroundScanline imageData, tmpSA1D, y
    For x = initX To finalX Step 4
        
        b = imageData(x)
        g = imageData(x + 1)
        r = imageData(x + 2)
        
        'Create a new value for each color based on the input parameters
        If isMonochrome Then
        
            newGray = r * channelModifiers(3, 0) + g * channelModifiers(3, 1) + b * channelModifiers(3, 2) + channelModifiers(3, 3)
            
            If (newGray > 255) Then newGray = 255
            If (newGray < 0) Then newGray = 0
            
            'Note: luminance preservation serves no purpose when monochrome is selected, so I do not process it here
            imageData(x) = newGray
            imageData(x + 1) = newGray
            imageData(x + 2) = newGray
            
        Else
        
            'If luminance is being preserved, we need to determine the initial luminance value
            If preserveLuminance Then originalLuminance = (Colors.GetLuminance(r, g, b) * ONE_DIV_255)
        
            newR = r * channelModifiers(0, 0) + g * channelModifiers(0, 1) + b * channelModifiers(0, 2) + channelModifiers(0, 3)
            newG = r * channelModifiers(1, 0) + g * channelModifiers(1, 1) + b * channelModifiers(1, 2) + channelModifiers(1, 3)
            newB = r * channelModifiers(2, 0) + g * channelModifiers(2, 1) + b * channelModifiers(2, 2) + channelModifiers(2, 3)
            
            'Fit everything in the [0, 255] range
            If (newR > 255) Then newR = 255
            If (newR < 0) Then newR = 0
            If (newG > 255) Then newG = 255
            If (newG < 0) Then newG = 0
            If (newB > 255) Then newB = 255
            If (newB < 0) Then newB = 0
            
            'If the user wants us to preserve luminance, determine the hue and saturation of the new color, then replace the luminance
            ' value with the original
            If preserveLuminance Then
                
                Colors.PreciseRGBtoHSL CDbl(newR) * ONE_DIV_255, CDbl(newG) * ONE_DIV_255, CDbl(newB) * ONE_DIV_255, h, s, l
                Colors.PreciseHSLtoRGB h, s, originalLuminance, rFloat, gFloat, bFloat
                
                imageData(x) = bFloat * 255
                imageData(x + 1) = gFloat * 255
                imageData(x + 2) = rFloat * 255
                
            Else
                imageData(x) = newB
                imageData(x + 1) = newG
                imageData(x + 2) = newR
            End If
            
            
        End If
        
    Next x
        If (Not toPreview) Then
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal y
            End If
        End If
    Next y
    
    'Safely deallocate imageData()
    workingDIB.UnwrapArrayFromDIB imageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic
    
End Sub

Private Sub cmdBar_AddCustomPresetData()

    'Because this control encompasses a bunch of "invisible" settings, e.g. channel values for channels other
    ' than the selected one, we must write out the ENTIRE CHANNEL ARRAY to the preset file
    cmdBar.AddPresetData "channelArray", GetLocalParamString()

End Sub

'OK button
Private Sub cmdBar_OKClick()
    UpdateStoredValues
    Process "Channel mixer", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RandomizeClick()
    
    'We actually want to randomize the entire stored value array, including channels that are not current visible
    Dim x As Long, y As Long
    For x = 0 To 3
        For y = 0 To 3
            If (x < 3) Then
                m_curSliderValues(x, y) = -200 + Int(Rnd * 401)
            Else
                m_curSliderValues(x, y) = -255 + Int(Rnd * 511)
            End If
        Next y
    Next x
    
    UpdateStoredValues
    
End Sub

Private Sub cmdBar_ReadCustomPresetData()

    'Because this control encompasses a bunch of "invisible" settings (e.g. the same sliders are reused
    ' against multiple channels, and we cache those settings independent of UI objects), we must read out
    ' a custom preset string that contains the ENTIRE CHANNEL ARRAY - not just the active one.
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString cmdBar.RetrievePresetData("channelArray")
    
    'We can now parse that string to retrieve the values for each individual channel
    With cParams
    
        'Next, we need to retrieve the 4x4 "grid" of values: four inputs (RGB/Constant) for each of
        ' four possible output channels (RGB/Gray).  For reference, you may want to refer to the
        ' named enums at the top of this module.
        m_curSliderValues(0, 0) = .GetLong("RedOutRedIn", 100)
        m_curSliderValues(0, 1) = .GetLong("RedOutGreenIn", 0)
        m_curSliderValues(0, 2) = .GetLong("RedOutBlueIn", 0)
        m_curSliderValues(0, 3) = .GetLong("RedOutConstantIn", 0)
        
        m_curSliderValues(1, 0) = .GetLong("GreenOutRedIn", 0)
        m_curSliderValues(1, 1) = .GetLong("GreenOutGreenIn", 100)
        m_curSliderValues(1, 2) = .GetLong("GreenOutBlueIn", 0)
        m_curSliderValues(1, 3) = .GetLong("GreenOutConstantIn", 0)
        
        m_curSliderValues(2, 0) = .GetLong("BlueOutRedIn", 0)
        m_curSliderValues(2, 1) = .GetLong("BlueOutGreenIn", 0)
        m_curSliderValues(2, 2) = .GetLong("BlueOutBlueIn", 100)
        m_curSliderValues(2, 3) = .GetLong("BlueOutConstantIn", 0)
        
        m_curSliderValues(3, 0) = .GetLong("GrayOutRedIn", 0)
        m_curSliderValues(3, 1) = .GetLong("GrayOutGreenIn", 0)
        m_curSliderValues(3, 2) = .GetLong("GrayOutBlueIn", 0)
        m_curSliderValues(3, 3) = .GetLong("GrayOutConstantIn", 100)
        
    End With
    
    'Sync the on-screen controls with whatever slider values are relevant
    m_forbidUpdate = True
    If chkMonochrome.Value Then
        btsChannel.Enabled = False
        sltRed.Value = m_curSliderValues(GrayOutput, RedInput)
        sltGreen.Value = m_curSliderValues(GrayOutput, GreenInput)
        sltBlue.Value = m_curSliderValues(GrayOutput, BlueInput)
        sltConstant.Value = m_curSliderValues(GrayOutput, ConstantInput)
    Else
        btsChannel.Enabled = True
        sltRed.Value = m_curSliderValues(btsChannel.ListIndex, RedInput)
        sltGreen.Value = m_curSliderValues(btsChannel.ListIndex, GreenInput)
        sltBlue.Value = m_curSliderValues(btsChannel.ListIndex, BlueInput)
        sltConstant.Value = m_curSliderValues(btsChannel.ListIndex, ConstantInput)
    End If
    m_forbidUpdate = False
    
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdateStoredValues
    UpdatePreview
End Sub

'RESET button
Private Sub cmdBar_ResetClick()

    'Fill the "stored value" array with default settings appropriate to each channel; basically, set each channel
    ' to their current value (e.g. red = "red 100%", "0%" for all other channels)
    Dim i As Long
    For i = 0 To 3
    
        Select Case i
        
            Case RedOutput
                m_curSliderValues(RedOutput, RedInput) = 100
                m_curSliderValues(RedOutput, GreenInput) = 0
                m_curSliderValues(RedOutput, BlueInput) = 0
                m_curSliderValues(RedOutput, ConstantInput) = 0
            
            Case GreenOutput
                m_curSliderValues(GreenOutput, RedInput) = 0
                m_curSliderValues(GreenOutput, GreenInput) = 100
                m_curSliderValues(GreenOutput, BlueInput) = 0
                m_curSliderValues(GreenOutput, ConstantInput) = 0
            
            Case BlueOutput
                m_curSliderValues(BlueOutput, RedInput) = 0
                m_curSliderValues(BlueOutput, GreenInput) = 0
                m_curSliderValues(BlueOutput, BlueInput) = 100
                m_curSliderValues(BlueOutput, ConstantInput) = 0
                
            'I'm not sure the best preset values to suggest for gray; for now, I'm defaulting to the ITU standard
            ' conversion formula - that should provide a good starting point for user modifications.
            Case GrayOutput
                m_curSliderValues(GrayOutput, RedInput) = 21
                m_curSliderValues(GrayOutput, GreenInput) = 72
                m_curSliderValues(GrayOutput, BlueInput) = 7
                m_curSliderValues(GrayOutput, ConstantInput) = 0
        
        End Select
    
    Next i
    
    'Reset the combo box and sliders on this page to default values
    btsChannel.ListIndex = 0
    sltRed.Value = 100
    sltGreen.Value = 0
    sltBlue.Value = 0
    sltConstant.Value = 0
    chkMonochrome.Value = False
    chkLuminance.Value = True
    
End Sub

Private Sub Form_Load()
    
    cmdBar.SetPreviewStatus False
    
    'Per convention, monochrome mode is handled via a separate checkbox.  This is also an easier solution for us, as
    ' it's difficult to apply changes to an imaginary "gray channel" (we'd have to divvy up any "gray channel"
    ' changes to each of red, green, and blue, and without a consistent way to do that the results would be
    ' unpredictable - I'm fairly certain this is why Photoshop etc. provide a separate "monochrome" checkbox)
    
    'Populate the channel selector
    btsChannel.AddItem "red", 0
    btsChannel.AddItem "green", 1
    btsChannel.AddItem "blue", 2
    
    Dim btnImageSize As Long
    btnImageSize = Interface.FixDPI(16)
    btsChannel.AssignImageToItem 0, vbNullString, Interface.GetRuntimeUIDIB(pdri_ChannelRed, btnImageSize, 2), btnImageSize, btnImageSize
    btsChannel.AssignImageToItem 1, vbNullString, Interface.GetRuntimeUIDIB(pdri_ChannelGreen, btnImageSize, 2), btnImageSize, btnImageSize
    btsChannel.AssignImageToItem 2, vbNullString, Interface.GetRuntimeUIDIB(pdri_ChannelBlue, btnImageSize, 2), btnImageSize, btnImageSize
    
    btsChannel.ListIndex = 0
            
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True
    
    'If the last-used settings involve the monochrome check box, the luminance check box may not be deactivated properly
    ' (due to no Click event being fired).  Forcibly check this state in advance.
    chkLuminance.Enabled = Not chkMonochrome.Value
    
    'Display the previewed effect in the neighboring window
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub sltBlue_Change()
    If (Not m_forbidUpdate) Then
        UpdateStoredValues
        UpdatePreview
    End If
End Sub

Private Sub sltBlue_ResetClick()
    If chkMonochrome.Value Then
        sltBlue.Value = 7
    Else
        If (btsChannel.ListIndex = BlueOutput) Then sltBlue.Value = 100 Else sltBlue.Value = 0
    End If
End Sub

Private Sub sltConstant_Change()
    If (Not m_forbidUpdate) Then
        UpdateStoredValues
        UpdatePreview
    End If
End Sub

Private Sub sltGreen_Change()
    If (Not m_forbidUpdate) Then
        UpdateStoredValues
        UpdatePreview
    End If
End Sub

Private Sub sltGreen_ResetClick()
    If chkMonochrome.Value Then
        sltGreen.Value = 72
    Else
        If (btsChannel.ListIndex = GreenOutput) Then sltGreen.Value = 100 Else sltGreen.Value = 0
    End If
End Sub

Private Sub sltRed_Change()
    If (Not m_forbidUpdate) Then
        UpdateStoredValues
        UpdatePreview
    End If
End Sub

Private Sub sltRed_ResetClick()
    If chkMonochrome.Value Then
        sltRed.Value = 21
    Else
        If (btsChannel.ListIndex = RedOutput) Then sltRed.Value = 100 Else sltRed.Value = 0
    End If
End Sub

'Because the user can change multiple channels at once, we need to store all current channel values in memory.
Private Sub UpdateStoredValues()

    'Store values according to the current combo box or monochrome setting
    If chkMonochrome.Value Then
        m_curSliderValues(GrayOutput, RedInput) = sltRed.Value
        m_curSliderValues(GrayOutput, GreenInput) = sltGreen.Value
        m_curSliderValues(GrayOutput, BlueInput) = sltBlue.Value
        m_curSliderValues(GrayOutput, ConstantInput) = sltConstant.Value
    Else
        m_curSliderValues(btsChannel.ListIndex, RedInput) = sltRed.Value
        m_curSliderValues(btsChannel.ListIndex, GreenInput) = sltGreen.Value
        m_curSliderValues(btsChannel.ListIndex, BlueInput) = sltBlue.Value
        m_curSliderValues(btsChannel.ListIndex, ConstantInput) = sltConstant.Value
    End If

End Sub

'Because this tool has a complex set of input values, we need to condense them all into a single string.
' This function handles the creation of that string for both previews and full-image applications.
Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
    
        'Start by adding the two simple parameters to the list
        cParams.AddParam "monochrome", chkMonochrome.Value
        cParams.AddParam "preserveluminance", chkLuminance.Value
        
        'Next, we have a 4x4 "grid" of values that needs to be added: four inputs (RGB/Constant) for each of
        ' four possible output channels (RGB/Gray).  For reference, you may want to refer to the named enums
        ' at the top of this module.
        
        'The order here isn't important; what matters is matching up the correct named parameter to
        ' our internal tracking array of slider values.  (A mirror version of this occurs in the
        ' actual channel mixer, where these values are mapped back into an adjustment array.)
        cParams.AddParam "RedOutRedIn", m_curSliderValues(0, 0)
        cParams.AddParam "RedOutGreenIn", m_curSliderValues(0, 1)
        cParams.AddParam "RedOutBlueIn", m_curSliderValues(0, 2)
        cParams.AddParam "RedOutConstantIn", m_curSliderValues(0, 3)
        
        cParams.AddParam "GreenOutRedIn", m_curSliderValues(1, 0)
        cParams.AddParam "GreenOutGreenIn", m_curSliderValues(1, 1)
        cParams.AddParam "GreenOutBlueIn", m_curSliderValues(1, 2)
        cParams.AddParam "GreenOutConstantIn", m_curSliderValues(1, 3)
        
        cParams.AddParam "BlueOutRedIn", m_curSliderValues(2, 0)
        cParams.AddParam "BlueOutGreenIn", m_curSliderValues(2, 1)
        cParams.AddParam "BlueOutBlueIn", m_curSliderValues(2, 2)
        cParams.AddParam "BlueOutConstantIn", m_curSliderValues(2, 3)
        
        cParams.AddParam "GrayOutRedIn", m_curSliderValues(3, 0)
        cParams.AddParam "GrayOutGreenIn", m_curSliderValues(3, 1)
        cParams.AddParam "GrayOutBlueIn", m_curSliderValues(3, 2)
        cParams.AddParam "GrayOutConstantIn", m_curSliderValues(3, 3)
        
    End With

    GetLocalParamString = cParams.GetParamString()

End Function

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then ApplyChannelMixer GetLocalParamString(), True, pdFxPreview
End Sub
