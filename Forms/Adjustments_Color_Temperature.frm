VERSION 5.00
Begin VB.Form FormColorTemp 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Color temperature"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12330
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
   ScaleWidth      =   822
   Begin PhotoDemon.pdButtonStrip btsMethod 
      Height          =   975
      Left            =   5880
      TabIndex        =   2
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   1720
      Caption         =   "method"
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   12330
      _ExtentX        =   21749
      _ExtentY        =   1323
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
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   4455
      Index           =   0
      Left            =   5880
      Top             =   1200
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   7858
      Begin PhotoDemon.pdSlider sldTempBasic 
         Height          =   705
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   1270
         Caption         =   "temperature"
         Min             =   -100
         Max             =   100
         SliderTrackStyle=   3
      End
      Begin PhotoDemon.pdLabel lblCool 
         Height          =   435
         Index           =   1
         Left            =   2760
         Top             =   960
         Width           =   2055
         _ExtentX        =   0
         _ExtentY        =   0
         Alignment       =   1
         Caption         =   "warmer"
         FontItalic      =   -1  'True
         FontSize        =   9
         ForeColor       =   4210752
         Layout          =   1
      End
      Begin PhotoDemon.pdLabel lblWarm 
         Height          =   435
         Index           =   1
         Left            =   240
         Top             =   960
         Width           =   2280
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "cooler"
         FontItalic      =   -1  'True
         FontSize        =   9
         ForeColor       =   4210752
         Layout          =   1
      End
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   4455
      Index           =   1
      Left            =   5880
      Top             =   1200
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   7858
      Begin PhotoDemon.pdSlider sltStrength 
         Height          =   705
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   1270
         Caption         =   "strength"
         Min             =   1
         Max             =   100
         Value           =   50
         NotchPosition   =   2
         NotchValueCustom=   50
      End
      Begin PhotoDemon.pdSlider sltTemperature 
         Height          =   705
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   1270
         Caption         =   "new temperature (K)"
         Min             =   1000
         Max             =   15000
         SliderTrackStyle=   3
         Value           =   5500
         DefaultValue    =   5500
      End
      Begin PhotoDemon.pdLabel lblCool 
         Height          =   435
         Index           =   0
         Left            =   2760
         Top             =   960
         Width           =   2055
         _ExtentX        =   0
         _ExtentY        =   0
         Alignment       =   1
         Caption         =   "cool tones"
         FontItalic      =   -1  'True
         FontSize        =   9
         ForeColor       =   4210752
         Layout          =   1
      End
      Begin PhotoDemon.pdLabel lblWarm 
         Height          =   435
         Index           =   0
         Left            =   360
         Top             =   960
         Width           =   2280
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "warm tones"
         FontItalic      =   -1  'True
         FontSize        =   9
         ForeColor       =   4210752
         Layout          =   1
      End
   End
End
Attribute VB_Name = "FormColorTemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Color Temperature Adjustment Form
'Copyright 2012-2026 by Tanner Helland
'Created: 16/September/12
'Last updated: 27/April/20
'Last update: minor perf improvements
'
'Color temperature adjustment form.  A full discussion of color temperature and how it works is available at this wikipedia article:
' https://en.wikipedia.org/wiki/Color_temperature
'
'Many cameras provide the ability to compensate for shooting under various types of lights by reversing the color casts
' caused by certain wavelengths.  For example, you may have seen options like "tungsten" or "fluorescent" or "overcast"
' when using a point-and-shoot camera.
'
'This form provides a similar effect, but more powerful.  It can be used to:
' 1) Automatically correct certain lighting conditions.  For example, a picture taken under fluorescent lights can be
'     adjusted to attempt to make it look more natural.
' 2) Convert image lighting from one type to another.  For example, a picture taken under overcast conditions can be made
'     to look like it was taken on a sunny day.
' 3) Manually apply color temperature changes.  Warning: this involves some ridiculous math.  Basically, I manually calculated
'     best-fit curves for established blackbody radiance values (taken from http://www.vendian.org/mncharity/dir3/blackbody/UnstableURLs/bbr_color.html).
'     Then I wrote a function to return values from these best-fit curves based on a supplied color temperature.  It's not perfect
'     but I've never found a function capable of doing this - especially not in VB - so it's better than anything out there right now.
'
'For a detailed explanation of how I reverse-engineered the math, please see this article:
' https://tannerhelland.com/2012/09/18/convert-temperature-rgb-algorithm-code.html
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Enum PD_TempAdjustment
    PDTA_Basic = 0
    PDTA_ExactTemp = 1
End Enum

#If False Then
    Private Const PDTA_Basic = 0, PDTA_ExactTemp = 1
#End If

'Cast an image with a new temperature value
' Input: desired temperature, whether to preserve luminance or not, and a blend ratio between 1 and 100
Public Sub ApplyTemperatureToImage(ByVal parameterList As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    'Parse out the parameter list
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString parameterList
    
    Dim temperatureMethod As PD_TempAdjustment, basicTempAdjustment As Long
    temperatureMethod = cParams.GetLong("temp-method", 0)
    basicTempAdjustment = cParams.GetLong("basic-temp", 0) \ 2
    
    Dim newTemperature As Long, preserveLuminance As Boolean, tempStrength As Double
    newTemperature = cParams.GetLong("advanced-temp", 5500)
    preserveLuminance = cParams.GetBool("advanced-preserve-luminance", True)
    tempStrength = cParams.GetDouble("advanced-strength", 25#)
    
    If (Not toPreview) Then Message "Applying new temperature to image..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte, tmpSA As SafeArray2D, tmpSA1D As SafeArray1D
    EffectPrep.PrepImageData tmpSA, toPreview, dstPic
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left * 4
    initY = curDIBValues.Top
    finalX = curDIBValues.Right * 4
    finalY = curDIBValues.Bottom
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not toPreview) Then ProgressBars.SetProgBarMax finalY
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'Color variables
    Dim r As Long, g As Long, b As Long
    Dim h As Double, s As Double, l As Double
    Dim originalLuminance As Double
    Dim tmpR As Long, tmpG As Long, tmpB As Long
    Dim rLookup() As Long, bLookup() As Long
    ReDim rLookup(0 To 255) As Long
    ReDim bLookup(0 To 255) As Long
    
    If (temperatureMethod = PDTA_Basic) Then
        
        'Build a look-up table of new temperature values.  (Temperature adjustments only affect the red and blue channels)
        For x = 0 To 255
            b = x - basicTempAdjustment
            If (b > 255) Then b = 255
            If (b < 0) Then b = 0
            r = x + basicTempAdjustment
            If (r > 255) Then r = 255
            If (r < 0) Then r = 0
            bLookup(x) = b
            rLookup(x) = r
        Next x
        
    Else
    
        'Get the corresponding RGB values for this temperature
        GetRGBfromTemperature tmpR, tmpG, tmpB, newTemperature
                
        'tempStrength needs to be on the range [0, 1], but it's presented to the user as [0, 100]
        tempStrength = tempStrength / 100#
        
    End If
    
    Const ONE_DIV_255 As Double = 1# / 255#
    
    'Loop through each pixel in the image, converting values as we go
    For y = initY To finalY
        workingDIB.WrapArrayAroundScanline imageData, tmpSA1D, y
    For x = initX To finalX Step 4
        
        'Get the source pixel color values
        b = imageData(x)
        g = imageData(x + 1)
        r = imageData(x + 2)
        
        'If luminance is being preserved, we need to determine the initial luminance value
        originalLuminance = Colors.GetLuminance(r, g, b) * ONE_DIV_255
        
        If (temperatureMethod = PDTA_Basic) Then
            Colors.ImpreciseRGBtoHSL rLookup(r), g, bLookup(b), h, s, l
            Colors.ImpreciseHSLtoRGB h, s, originalLuminance, r, g, b
        Else
            
            'Blend the original and new RGB values using the specified strength
            r = Colors.BlendColors(r, tmpR, tempStrength)
            g = Colors.BlendColors(g, tmpG, tempStrength)
            b = Colors.BlendColors(b, tmpB, tempStrength)
            
            'If the user wants us to preserve luminance, determine the hue and saturation of the new color, then replace the luminance
            ' value with the original
            If preserveLuminance Then
                Colors.ImpreciseRGBtoHSL r, g, b, h, s, l
                Colors.ImpreciseHSLtoRGB h, s, originalLuminance, r, g, b
            End If
        
        End If
        
        'Assign the new values to each color channel
        imageData(x) = b
        imageData(x + 1) = g
        imageData(x + 2) = r
        
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

Private Sub btsMethod_Click(ByVal buttonIndex As Long)
    UpdatePanelVisibility
    UpdatePreview
End Sub

Private Sub cmdBar_OKClick()
    Process "Temperature", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub Form_Load()
    
    cmdBar.SetPreviewStatus False
    
    btsMethod.AddItem "basic", 0
    btsMethod.AddItem "advanced", 1
    btsMethod.ListIndex = 0
    UpdatePanelVisibility
    
    'Calculate gradient colors for the temperature slider, using the built-in Kelvin to RGB converter
    Dim r As Long, g As Long, b As Long
    
    'Simple gradient-ish code implementation of drawing temperature between 1000 and 12000 Kelvin
    GetRGBfromTemperature r, g, b, sltTemperature.Min
    sltTemperature.GradientColorLeft = RGB(r, g, b)
    sldTempBasic.GradientColorRight = RGB(r, g, b)
    
    GetRGBfromTemperature r, g, b, sltTemperature.Max
    sltTemperature.GradientColorRight = RGB(r, g, b)
    sldTempBasic.GradientColorLeft = RGB(r, g, b)
    
    sltTemperature.GradientMiddleValue = 6500
    GetRGBfromTemperature r, g, b, sltTemperature.GradientMiddleValue
    sltTemperature.GradientColorMiddle = RGB(r, g, b)
    sldTempBasic.GradientColorMiddle = RGB(r, g, b)
    
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview

End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub UpdatePanelVisibility()
    Dim i As Long
    For i = 0 To btsMethod.ListCount - 1
        picContainer(i).Visible = (i = btsMethod.ListIndex)
    Next i
End Sub

'Given a temperature (in Kelvin), generate the RGB equivalent of an ideal black body
' NOTE: the mathematical formula used in this routine is NOT STANDARD.  I wrote it myself using self-calculated regression equations based
'        off the raw data on blackbody radiation provided at http://www.vendian.org/mncharity/dir3/blackbody/UnstableURLs/bbr_color.html
'        Because of that, I can't guarantee great precision - but the function works well enough for photo-manipulation purposes.
Private Sub GetRGBfromTemperature(ByRef r As Long, ByRef g As Long, ByRef b As Long, ByVal tmpKelvin As Long)

    Dim tmpCalc As Double

    'Temperature must fall between 1000 and 40000 degrees
    If (tmpKelvin < 1000) Then tmpKelvin = 1000
    If (tmpKelvin > 40000) Then tmpKelvin = 40000
    
    'All calculations require tmpKelvin \ 100, so only do the conversion once
    tmpKelvin = tmpKelvin \ 100
    
    'Calculate each color in turn
    
    'First: red
    If (tmpKelvin <= 66) Then
        r = 255
    Else
        tmpCalc = tmpKelvin - 55
        r = 351.976905668057 + 0.114206453784165 * tmpCalc + -40.2536630933213 * Log(tmpCalc)
        If (r < 0) Then r = 0
        If (r > 255) Then r = 255
    End If
    
    'Second: green
    If (tmpKelvin <= 66) Then
        tmpCalc = tmpKelvin - 2
        g = -155.254855627092 + -0.445969504695791 * tmpCalc + 104.492161993939 * Log(tmpCalc)
        If (g < 0) Then g = 0
        If (g > 255) Then g = 255
    Else
        tmpCalc = tmpKelvin - 50
        g = 325.449412571197 + 7.94345653666234E-02 * tmpCalc + -28.0852963507957 * Log(tmpCalc)
        If (g < 0) Then g = 0
        If (g > 255) Then g = 255
    End If
    
    'Third: blue
    If (tmpKelvin >= 66) Then
        b = 255
    ElseIf (tmpKelvin <= 19) Then
        b = 0
    Else
        tmpCalc = tmpKelvin - 10
        b = -254.769351841209 + 0.827409606400739 * tmpCalc + 115.679944010661 * Log(tmpCalc)
        If (b < 0) Then b = 0
        If (b > 255) Then b = 255
    End If
    
End Sub

Private Sub sldTempBasic_Change()
    UpdatePreview
End Sub

Private Sub sltStrength_Change()
    UpdatePreview
End Sub

Private Sub sltTemperature_Change()
    UpdatePreview
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then
        ApplyTemperatureToImage GetLocalParamString(), True, pdFxPreview
    End If
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    GetLocalParamString = BuildParamList("temp-method", btsMethod.ListIndex, "basic-temp", sldTempBasic.Value, "advanced-temp", sltTemperature.Value, "advanced-preserve-luminance", True, "advanced-strength", sltStrength.Value / 2)
End Function
