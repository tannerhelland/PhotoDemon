VERSION 5.00
Begin VB.Form FormColorTemp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Adjust Color Temperature"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6255
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
   ScaleHeight     =   449
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   417
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.HScrollBar hsStrength 
      Height          =   255
      Left            =   240
      Max             =   100
      Min             =   1
      TabIndex        =   11
      Top             =   5400
      Value           =   55
      Width           =   5055
   End
   Begin VB.TextBox txtStrength 
      Alignment       =   2  'Center
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
      Left            =   5400
      MaxLength       =   3
      TabIndex        =   10
      Text            =   "50"
      Top             =   5355
      Width           =   735
   End
   Begin VB.TextBox txtTemperature 
      Alignment       =   2  'Center
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
      Left            =   5400
      MaxLength       =   5
      TabIndex        =   9
      Text            =   "5500"
      Top             =   3675
      Width           =   735
   End
   Begin VB.PictureBox picEffect 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2730
      Left            =   3240
      ScaleHeight     =   180
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   191
      TabIndex        =   5
      Top             =   120
      Width           =   2895
   End
   Begin VB.PictureBox picPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2730
      Left            =   120
      ScaleHeight     =   180
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   191
      TabIndex        =   4
      Top             =   120
      Width           =   2895
   End
   Begin VB.HScrollBar hsTemperature 
      Height          =   255
      Left            =   240
      Max             =   150
      Min             =   10
      TabIndex        =   2
      Top             =   3720
      Value           =   55
      Width           =   5055
   End
   Begin VB.PictureBox picTempDemo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   480
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   303
      TabIndex        =   3
      Top             =   4080
      Width           =   4575
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   4800
      TabIndex        =   1
      Top             =   6120
      Width           =   1245
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   3480
      TabIndex        =   0
      Top             =   6120
      Width           =   1245
   End
   Begin VB.Label lblCool 
      AutoSize        =   -1  'True
      Caption         =   "cool tones"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   4245
      TabIndex        =   14
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label lblWarm 
      AutoSize        =   -1  'True
      Caption         =   "warm tones"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   480
      TabIndex        =   13
      Top             =   4560
      Width           =   840
   End
   Begin VB.Label lblStrength 
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
      Left            =   120
      TabIndex        =   12
      Top             =   5040
      Width           =   960
   End
   Begin VB.Label lblTemperature 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "new temperature:"
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
      Left            =   120
      TabIndex        =   8
      Top             =   3360
      Width           =   1890
   End
   Begin VB.Label lblAfter 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "after"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   3360
      TabIndex        =   7
      Top             =   2880
      Width           =   360
   End
   Begin VB.Label lblBefore 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "before"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   2880
      Width           =   480
   End
End
Attribute VB_Name = "FormColorTemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Color Temperature Adjustment Form
'Copyright ©2006-2013 by Tanner Helland
'Created: 16/September/12
'Last updated: 18/September/12
'Last update: remove "preserve luminance" checkbox.  There was never any reason to uncheck it.
'
'Color temperature adjustment form.  A full discussion of color temperature and how it works is available at this wikipedia article:
' http://en.wikipedia.org/wiki/Color_temperature
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
'***************************************************************************

Option Explicit

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub cmdOK_Click()
    
    'The scroll bar max and min values are used to check the temperature input for validity
    If EntryValid(txtTemperature, hsTemperature.Min * 100, hsTemperature.Max * 100) Then
        
        'Same goes for the "strength" value
        If EntryValid(txtStrength, hsStrength.Min, hsStrength.Max) Then
            
            Me.Visible = False
            Process AdjustTemperature, CLng(hsTemperature.Value) * 100, True, CSng(hsStrength.Value) / 2
            Unload Me
            
        Else
            AutoSelectText txtStrength
        End If
        
    Else
        AutoSelectText txtTemperature
    End If
    
End Sub

'Cast an image with a new temperature value
' Input: desired temperature, whether to preserve luminance or not, and a blend ratio between 1 and 100
Public Sub ApplyTemperatureToImage(ByVal newTemperature As Long, Optional ByVal preserveLuminance As Boolean = True, Optional ByVal tempStrength As Single = 25, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As PictureBox)
    
    If toPreview = False Then Message "Applying new temperature to image..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    
    prepImageData tmpSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
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
    
    'Color variables
    Dim r As Long, g As Long, b As Long
    Dim h As Single, s As Single, l As Single
    Dim originalLuminance As Single
    Dim tmpR As Long, tmpG As Long, tmpB As Long
            
    'Get the corresponding RGB values for this temperature
    getRGBfromTemperature tmpR, tmpG, tmpB, newTemperature
            
    'Divide tempStrength by 100 to yield a value between 0 and 1
    tempStrength = tempStrength / 100
            
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'If luminance is being preserved, we need to determine the initial luminance value
        originalLuminance = (getLuminance(r, g, b) / 255)
        
        'Blend the original and new RGB values using the specified strength
        r = BlendColors(r, tmpR, tempStrength)
        g = BlendColors(g, tmpG, tempStrength)
        b = BlendColors(b, tmpB, tempStrength)
        
        'If the user wants us to preserve luminance, determine the hue and saturation of the new color, then replace the luminance
        ' value with the original
        If preserveLuminance Then
            tRGBToHSL r, g, b, h, s, l
            tHSLToRGB h, s, originalLuminance, r, g, b
        End If
        
        'Assign the new values to each color channel
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
        
    Next y
        If toPreview = False Then
            If (x And progBarCheck) = 0 Then SetProgBarVal x
        End If
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
    
End Sub

'When the form is activated (e.g. made visible and receives focus),
Private Sub Form_Activate()

    'This short routine is for drawing the picture box below the temperature slider
    Dim temperatureVal As Double
    Dim r As Long, g As Long, b As Long
    
    'Simple gradient-ish code implementation of drawing temperature between 1000 and 12000 Kelvin
    Dim x As Long
    For x = 0 To picTempDemo.ScaleWidth
    
        'Based on our x-position, gradient a value between 1000 and 12000
        temperatureVal = x / picTempDemo.ScaleWidth
        temperatureVal = temperatureVal * (CLng(hsTemperature.Max) * 100)
        temperatureVal = temperatureVal + (CLng(hsTemperature.Min) * 100)
        
        'Generate an RGB equivalent for this temperature
        getRGBfromTemperature r, g, b, temperatureVal
        
        'Draw the color
        picTempDemo.Line (x, 0)-(x, picTempDemo.ScaleHeight), RGB(r, g, b)
        
    Next x
    
    picTempDemo.Picture = picTempDemo.Image
    
    'Create a copy of the image on the preview window
    DrawPreviewImage picPreview
    
    'Display the previewed effect in the neighboring window
    ApplyTemperatureToImage CLng(hsTemperature.Value) * 100, True, CSng(hsStrength.Value) / 2, True, picEffect
    
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub hsStrength_Change()
    copyToTextBoxI txtStrength, hsStrength
    ApplyTemperatureToImage CLng(hsTemperature.Value) * 100, True, CSng(hsStrength.Value) / 2, True, picEffect
End Sub

Private Sub hsStrength_Scroll()
    copyToTextBoxI txtStrength, hsStrength
    ApplyTemperatureToImage CLng(hsTemperature.Value) * 100, True, CSng(hsStrength.Value) / 2, True, picEffect
End Sub

'When the hue scroll bar is changed, redraw the preview
Private Sub hsTemperature_Change()
    copyToTextBoxI txtTemperature, hsTemperature * 100
    ApplyTemperatureToImage CLng(hsTemperature.Value) * 100, True, CSng(hsStrength.Value) / 2, True, picEffect
End Sub

Private Sub hsTemperature_Scroll()
    copyToTextBoxI txtTemperature, hsTemperature * 100
    ApplyTemperatureToImage CLng(hsTemperature.Value) * 100, True, CSng(hsStrength.Value) / 2, True, picEffect
End Sub

'Given a temperature (in Kelvin), generate the RGB equivalent of an ideal black body
' NOTE: the mathematical formula used in this routine is NOT STANDARD.  I wrote it myself using self-calculated regression equations based
'        off the raw data on blackbody radiation provided at http://www.vendian.org/mncharity/dir3/blackbody/UnstableURLs/bbr_color.html
'        Because of that, I can't guarantee great precision - but the function works well enough for photo-manipulation purposes.
Private Sub getRGBfromTemperature(ByRef r As Long, ByRef g As Long, ByRef b As Long, ByVal tmpKelvin As Long)

    Static tmpCalc As Double

    'Temperature must fall between 1000 and 40000 degrees
    If tmpKelvin < 1000 Then tmpKelvin = 1000
    If tmpKelvin > 40000 Then tmpKelvin = 40000
    
    'All calculations require tmpKelvin \ 100, so only do the conversion once
    tmpKelvin = tmpKelvin \ 100
    
    'Calculate each color in turn
    
    'First: red
    If tmpKelvin <= 66 Then
        r = 255
    Else
        'Note: the R-squared value for this approximation is .988
        tmpCalc = tmpKelvin - 60
        tmpCalc = 329.698727446 * (tmpCalc ^ -0.1332047592)
        r = tmpCalc
        If r < 0 Then r = 0
        If r > 255 Then r = 255
    End If
    
    'Second: green
    If tmpKelvin <= 66 Then
        'Note: the R-squared value for this approximation is .996
        tmpCalc = tmpKelvin
        tmpCalc = 99.4708025861 * Log(tmpCalc) - 161.1195681661
        g = tmpCalc
        If g < 0 Then g = 0
        If g > 255 Then g = 255
    Else
        'Note: the R-squared value for this approximation is .987
        tmpCalc = tmpKelvin - 60
        tmpCalc = 288.1221695283 * (tmpCalc ^ -0.0755148492)
        g = tmpCalc
        If g < 0 Then g = 0
        If g > 255 Then g = 255
    End If
    
    'Third: blue
    If tmpKelvin >= 66 Then
        b = 255
    ElseIf tmpKelvin <= 19 Then
        b = 0
    Else
        'Note: the R-squared value for this approximation is .998
        tmpCalc = tmpKelvin - 10
        tmpCalc = 138.5177312231 * Log(tmpCalc) - 305.0447927307
        
        b = tmpCalc
        If b < 0 Then b = 0
        If b > 255 Then b = 255
    End If
    
End Sub

'Keep the "Strength" scroll bar and text box in sync
Private Sub txtStrength_GotFocus()
    AutoSelectText txtStrength
End Sub

Private Sub txtStrength_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtStrength
    If EntryValid(txtStrength, 1, 100, False, False) Then hsStrength.Value = Val(txtStrength)
End Sub

'Keep the "Temperature" scroll bar and text box in sync
Private Sub txtTemperature_GotFocus()
    AutoSelectText txtTemperature
End Sub

Private Sub txtTemperature_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtTemperature
    If EntryValid(txtTemperature, hsTemperature.Min * 100, hsTemperature.Max * 100, False, False) Then hsTemperature.Value = Val(txtTemperature) \ 100
End Sub
