VERSION 5.00
Begin VB.Form FormKuwahara 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Kuwahara"
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
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   12090
      _ExtentX        =   21325
      _ExtentY        =   1323
      BackColor       =   14802140
   End
   Begin PhotoDemon.sliderTextCombo sltRadius 
      Height          =   720
      Left            =   6000
      TabIndex        =   2
      Top             =   2280
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "radius"
      Min             =   1
      Max             =   20
      Value           =   5
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
End
Attribute VB_Name = "FormKuwahara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Kuwahara Blur Dialog
'Copyright 2014 by Audioglider
'Created: 22/June/14
'Last updated: 22/June/14
'Last update: Initial build
'TODO: adopt Median filter optimization, where instead of rebuilding each quadrant from scratch for each pixel,
'       we simply add and remove a single horizontal line to each (or a single vertical line when moving to a new
'       row).  Similar to the Median filter, I expect this to provide an exponential performance improvement
'       relative to filter radius.
'
'Kuwahara is a non-linear smoothing filter that preserves edges.
'
' It works as follows:
'
' For each pixel, divide up the region around it into four overlapping
' blocks where each block has the center pixel as a corner pixel.
' For each of the blocks calculate the mean and the variance. Set the
' middle pixel equal to the mean of the block with the smallest variance.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Apply the Kuwahara filter to an image.
Public Sub Kuwahara(ByVal filterSize As Long, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If Not toPreview Then Message "Applying Kuwahara smoothing..."
    
    'Indicies for each of the 4 quadrants
    Dim meansR(0 To 3) As Double
    Dim meansG(0 To 3) As Double
    Dim meansB(0 To 3) As Double
    Dim stdDevsR(0 To 3) As Double
    Dim stdDevsG(0 To 3) As Double
    Dim stdDevsB(0 To 3) As Double
    Dim pixels(0 To 3) As RGBQUAD

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
    
    'If this is a preview, we need to adjust the kernal
    If toPreview Then filterSize = filterSize * curDIBValues.previewModifier
    If filterSize < 1 Then filterSize = 1
    
    Dim dx As Long, dy As Long
    Dim xdx(0 To 1) As Long, ydy(0 To 1) As Long
    Dim i As Long
    
    Dim radius As Long
    Dim scaler As Double
    Dim lowest As Double
    Dim lowestIndex As Long
    Dim rgbSum As Double
        
    radius = filterSize / 2
    scaler = 1# / ((radius + 1) * (radius + 1))

    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Clear the means and standard deviations
        For i = 0 To 3
            meansR(i) = 0
            meansG(i) = 0
            meansB(i) = 0
            stdDevsR(i) = 0
            stdDevsG(i) = 0
            stdDevsB(i) = 0
        Next i
        
        'Calculate means
        For dx = 0 To radius
            xdx(0) = x - dx
            If xdx(0) < initX Then
                xdx(0) = initX + initX - xdx(0)
            End If
            xdx(1) = x + dx
            If xdx(1) >= finalX Then
                xdx(1) = finalX - 2 - xdx(1) + finalX
            End If
            
            For dy = 0 To radius
                ydy(0) = y - dy
                If ydy(0) < initY Then
                    ydy(0) = initY + initY - ydy(0)
                End If
                ydy(1) = y + dy
                If ydy(1) >= finalY Then
                    ydy(1) = finalY - 2 - ydy(1) + finalY
                End If
                
                'Get the source pixel color values for each direction
                pixels(0).Red = ImageData(xdx(0) * qvDepth + 2, ydy(0))
                pixels(0).Green = ImageData(xdx(0) * qvDepth + 1, ydy(0))
                pixels(0).Blue = ImageData(xdx(0) * qvDepth, ydy(0))
                
                pixels(1).Red = ImageData(xdx(1) * qvDepth + 2, ydy(0))
                pixels(1).Green = ImageData(xdx(1) * qvDepth + 1, ydy(0))
                pixels(1).Blue = ImageData(xdx(1) * qvDepth, ydy(0))
                
                pixels(2).Red = ImageData(xdx(0) * qvDepth + 2, ydy(1))
                pixels(2).Green = ImageData(xdx(0) * qvDepth + 1, ydy(1))
                pixels(2).Blue = ImageData(xdx(0) * qvDepth, ydy(1))
                
                pixels(3).Red = ImageData(xdx(1) * qvDepth + 2, ydy(1))
                pixels(3).Green = ImageData(xdx(1) * qvDepth + 1, ydy(1))
                pixels(3).Blue = ImageData(xdx(1) * qvDepth, ydy(1))
                
                For i = 0 To 3
                    meansR(i) = meansR(i) + CDbl(pixels(i).Red)
                    meansG(i) = meansG(i) + CDbl(pixels(i).Green)
                    meansB(i) = meansB(i) + CDbl(pixels(i).Blue)
                Next i
                
            Next dy
        Next dx
        
        For i = 0 To 3
            meansR(i) = meansR(i) * scaler
            meansG(i) = meansG(i) * scaler
            meansB(i) = meansB(i) * scaler
        Next i
        
        'Calculate standard deviations
        For dx = 0 To radius
            xdx(0) = x - dx
            If xdx(0) < initX Then
                xdx(0) = initX + initX - xdx(0)
            End If
            xdx(1) = x + dx
            If xdx(1) >= finalX Then
                xdx(1) = finalX - 2 - xdx(1) + finalX
            End If
            
            For dy = 0 To radius
                ydy(0) = y - dy
                If ydy(0) < initY Then
                    ydy(0) = initY + initY - ydy(0)
                End If
                ydy(1) = y + dy
                If ydy(1) >= finalY Then
                    ydy(1) = finalY - 2 - ydy(1) + finalY
                End If
                
                'Get the source pixel color values for each quadrant
                pixels(0).Red = ImageData(xdx(0) * qvDepth + 2, ydy(0))
                pixels(0).Green = ImageData(xdx(0) * qvDepth + 1, ydy(0))
                pixels(0).Blue = ImageData(xdx(0) * qvDepth, ydy(0))
                
                pixels(1).Red = ImageData(xdx(1) * qvDepth + 2, ydy(0))
                pixels(1).Green = ImageData(xdx(1) * qvDepth + 1, ydy(0))
                pixels(1).Blue = ImageData(xdx(1) * qvDepth, ydy(0))
                
                pixels(2).Red = ImageData(xdx(0) * qvDepth + 2, ydy(1))
                pixels(2).Green = ImageData(xdx(0) * qvDepth + 1, ydy(1))
                pixels(2).Blue = ImageData(xdx(0) * qvDepth, ydy(1))
                
                pixels(3).Red = ImageData(xdx(1) * qvDepth + 2, ydy(1))
                pixels(3).Green = ImageData(xdx(1) * qvDepth + 1, ydy(1))
                pixels(3).Blue = ImageData(xdx(1) * qvDepth, ydy(1))
                
                For i = 0 To 3
                    stdDevsR(i) = stdDevsR(i) + (meansR(i) - CDbl(pixels(i).Red)) * (meansR(i) - CDbl(pixels(i).Red))
                    stdDevsG(i) = stdDevsR(i) + (meansG(i) - CDbl(pixels(i).Green)) * (meansG(i) - CDbl(pixels(i).Green))
                    stdDevsB(i) = stdDevsB(i) + (meansB(i) - CDbl(pixels(i).Blue)) * (meansB(i) - CDbl(pixels(i).Blue))
                Next i
                
            Next dy
        Next dx
        
        lowest = DOUBLE_MAX
        lowestIndex = 0
        
        'Work out the lowest standard deviation
        For i = 0 To 3
            rgbSum = (stdDevsR(i) + stdDevsG(i) + stdDevsB(i))
            If rgbSum < lowest Then
                lowest = rgbSum
                lowestIndex = i
            End If
        Next i
        
        'Assign the new values to each color channel
        ImageData(QuickVal + 2, y) = CLng(meansR(lowestIndex))
        ImageData(QuickVal + 1, y) = CLng(meansG(lowestIndex))
        ImageData(QuickVal, y) = CLng(meansB(lowestIndex))
        
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
    Process "Kuwahara filter", , buildParams(sltRadius.Value), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub Form_Activate()
        
    'Apply translations and visual themes
    MakeFormPretty Me
    
    'Display the previewed effect in the neighboring window
    updatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub sltRadius_Change()
    updatePreview
End Sub

Private Sub updatePreview()
    If cmdBar.previewsAllowed Then Kuwahara sltRadius.Value, True, fxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub
