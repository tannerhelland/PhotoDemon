VERSION 5.00
Begin VB.Form FormExposure 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Exposure"
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
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   12030
      _ExtentX        =   21220
      _ExtentY        =   1323
      BackColor       =   14802140
   End
   Begin VB.PictureBox picChart 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   8280
      ScaleHeight     =   159
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   231
      TabIndex        =   3
      Top             =   480
      Width           =   3495
   End
   Begin PhotoDemon.sliderTextCombo sltExposure 
      Height          =   720
      Left            =   6000
      TabIndex        =   2
      Top             =   3720
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "exposure compensation (stops)"
      Min             =   -5
      Max             =   5
      SigDigits       =   2
      SliderTrackStyle=   2
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
      BackStyle       =   0  'Transparent
      Caption         =   "new exposure curve:"
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
      Height          =   1005
      Index           =   2
      Left            =   5880
      TabIndex        =   4
      Top             =   1530
      Width           =   2280
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "FormExposure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Exposure Dialog
'Copyright 2013-2015 by audioglider and Tanner Helland
'Created: 13/July/13
'Last updated: 09/August/13
'Last update: rewrote the exposure calculation to operate on a "stops" (power-of-2) scale
'
'Many thanks to talented contributer audioglider for creating this tool.
'
'Basic image exposure adjustment dialog.  Exposure is a complex topic in photography, and (obviously) the best way to
' adjust it is at image capture time.  This is because true exposure relies on a number of variables (see
' http://en.wikipedia.org/wiki/Exposure_%28photography%29) inherent in the scene itself, with a technical definition
' of "the accumulated physical quantity of visible light energy applied to a surface during a given exposure time."
' Once a set amount of light energy has been applied to a digital sensor and the resulting photo is captured, actual
' exposure can never fully be corrected or adjusted in post-production.
'
'That said, in the event that a poor choice is made at time of photography, certain approximate adjustments can be
' applied in post-production, with the understanding that missing shadows and highlights cannot be "magically"
' recreated out of thin air.  This is done by approximating an EV adjustment using a simple power-of-two formula.
' For more information on exposure compensation, see
' http://en.wikipedia.org/wiki/Exposure_value#Exposure_compensation_in_EV
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Adjust an image's exposure.
' PRIMARY INPUT: exposureAdjust represents the number of stops to correct the image.  Each stop corresponds to a power-of-2
'                 increase (+values) or decrease (-values) in luminance.  Thus an EV of -1 will cut the amount of light in
'                 half, while an EV of +1 will double the amount of light.
Public Sub Exposure(ByVal exposureAdjust As Double, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If Not toPreview Then Message "Adjusting image exposure..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    
    prepImageData tmpSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    Dim r As Long, g As Long, b As Long
    
    'Exposure can be easily applied using a look-up table
    Dim gLookUp(0 To 255) As Byte
    Dim tmpVal As Double
    
    For x = 0 To 255
        tmpVal = x / 255
        tmpVal = tmpVal * 2 ^ (exposureAdjust)
        tmpVal = tmpVal * 255
        
        If tmpVal > 255 Then tmpVal = 255
        If tmpVal < 0 Then tmpVal = 0
        
        gLookUp(x) = tmpVal
    Next x
    
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Apply a new value based on the lookup table
        ImageData(QuickVal + 2, y) = gLookUp(r)
        ImageData(QuickVal + 1, y) = gLookUp(g)
        ImageData(QuickVal, y) = gLookUp(b)
        
    Next y
        If Not toPreview Then
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

Private Sub cmdBar_OKClick()
    Process "Exposure", , buildParams(sltExposure), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub Form_Activate()
    
    'Apply translations and visual themes
    MakeFormPretty Me
    
    'Draw a preview of the effect
    updatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Update the preview whenever the combination slider/text control has its value changed
Private Sub sltExposure_Change()
    updatePreview
End Sub

'Redrawing a preview of the exposure effect also redraws the exposure curve (which isn't really a curve, but oh well)
Private Sub updatePreview()
    
    If cmdBar.previewsAllowed And sltExposure.IsValid Then
    
        Dim prevX As Double, prevY As Double
        Dim curX As Double, curY As Double
        Dim x As Long
        
        Dim xWidth As Long, yHeight As Long
        xWidth = picChart.ScaleWidth
        yHeight = picChart.ScaleHeight
            
        'Clear out the old chart and draw a gray line across the diagonal for reference
        picChart.Picture = LoadPicture("")
        picChart.ForeColor = RGB(127, 127, 127)
        GDIPlusDrawLineToDC picChart.hDC, 0, yHeight, xWidth, 0, RGB(127, 127, 127)
        
        'Draw the corresponding exposure curve (line, actually) for this EV
        Dim expVal As Double, tmpVal As Double
        expVal = sltExposure
        
        picChart.ForeColor = RGB(0, 0, 255)
        
        prevX = 0
        prevY = yHeight
        curX = 0
        curY = yHeight
        
        For x = 0 To xWidth
            tmpVal = x / xWidth
            tmpVal = tmpVal * 2 ^ (expVal)
            tmpVal = yHeight - (tmpVal * yHeight)
            curY = tmpVal
            curX = x
            GDIPlusDrawLineToDC picChart.hDC, prevX, prevY, curX, curY, picChart.ForeColor
            prevX = curX
            prevY = curY
        Next x
        
        picChart.Picture = picChart.Image
        picChart.Refresh
    
        'Finally, apply the exposure correction to the preview image
        Exposure sltExposure, True, fxPreview
        
    End If
    
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub


