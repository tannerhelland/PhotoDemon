VERSION 5.00
Begin VB.Form FormMetal 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Metal"
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14802140
   End
   Begin PhotoDemon.sliderTextCombo sltRadius 
      Height          =   495
      Left            =   6000
      TabIndex        =   3
      Top             =   3000
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Max             =   200
      SigDigits       =   1
      Value           =   20
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.sliderTextCombo sltDetail 
      Height          =   495
      Left            =   6000
      TabIndex        =   5
      Top             =   2040
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Max             =   16
      Value           =   4
      NotchPosition   =   2
      NotchValueCustom=   4
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "detail:"
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
      TabIndex        =   4
      Top             =   1650
      Width           =   660
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "smoothness:"
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
      TabIndex        =   1
      Top             =   2640
      Width           =   1350
   End
End
Attribute VB_Name = "FormMetal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'"Metal" or "Chrome" Image effect
'Copyright 2002-2015 by Tanner Helland
'Created: sometime 2002
'Last updated: 02/April/15
'Last update: rewrite function from scratch
'
'PhotoDemon's "Metal" filter is the rough equivalent of "Chrome" in Photoshop.  Our implementation is relatively
' straightforward; a normalized graymap is created for the image, then remapped according to a sinusoidal-like
' lookup table (created using the pdFilterLUT class).
'
'The user currently has control over two parameters: "smoothness", which determines a pre-effect blur radius,
' and "detail" which controls the number of octaves in the lookup table.
'
'Still TODO: allow the user to set a highlight and shadow color, instead of using boring ol' gray
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Apply a metallic "shimmer" to an image
Public Sub ApplyMetalFilter(ByVal steelDetail As Long, ByVal steelSmoothness As Double, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If Not toPreview Then Message "Pouring smoldering metal onto image..."
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    
    'If this is a preview, we need to adjust the smoothness (kernel radius) to match the size of the preview box
    If toPreview Then steelSmoothness = steelSmoothness * curDIBValues.previewModifier
    
    'Retrieve a normalized luminance map of the current image
    Dim grayMap() As Byte
    DIB_Handler.getDIBGrayscaleMap workingDIB, grayMap, True
    
    'If the user specified a non-zero smoothness, apply it now
    If steelSmoothness > 0 Then Filters_ByteArray.GaussianBlur_IIR_ByteArray grayMap, workingDIB.getDIBWidth, workingDIB.getDIBHeight, steelSmoothness, 3
    
    'Re-normalize the data
    'Filters_ByteArray.normalizeByteArray grayMap, workingDIB.getDIBWidth, workingDIB.getDIBHeight
    
    'Next, we need to generate a sinusoidal octave lookup table for the graymap.  This causes the luminance of the map to
    ' vary evently between the number of detail points requested by the user.
    
    'Detail cannot be lower than 2, but it is presented to the user as [0, (arbitrary upper bound)], so add two to the total now
    steelDetail = steelDetail + 2
    
    'We will be using pdFilterLUT to generate the corresponding lookup table, which means we need to use a POINTFLOAT array
    Dim curvePoints() As POINTFLOAT
    ReDim curvePoints(0 To steelDetail) As POINTFLOAT
    
    'X values are evenly distributed from 0 to 255
    Dim i As Long
    For i = 0 To steelDetail
        curvePoints(i).x = CDbl(i / steelDetail) * 255
    Next i
    
    'Y values alternate between 0 and 255
    For i = 0 To steelDetail
        
        If i Mod 2 = 0 Then
            
            If i = 0 Then
                curvePoints(i).y = 0
            Else
                curvePoints(i).y = 25
            End If
            
        Else
        
            If i = steelDetail Then
                curvePoints(i).y = 255
            Else
                curvePoints(i).y = 230
            End If
            
        End If
        
    Next i
    
    'Convert our point array into a luminance curve
    Dim luminanceLookup() As Byte
    
    Dim cLUT As pdFilterLUT
    Set cLUT = New pdFilterLUT
    cLUT.fillLUT_Curve luminanceLookup, curvePoints
    
    'We are now ready to apply the final curve to the image!
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(dstSA), 4
    
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
    
    Dim grayVal As Long
    
    'Apply the filter
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
        
        grayVal = luminanceLookup(grayMap(x, y))
        
        ImageData(QuickVal + 2, y) = grayVal
        ImageData(QuickVal + 1, y) = grayVal
        ImageData(QuickVal, y) = grayVal
        
    Next y
        If (x And progBarCheck) = 0 Then
            If userPressedESC() Then Exit For
            SetProgBarVal x
        End If
    Next x
        
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingDIB
    finalizeImageData toPreview, dstPic
            
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Metal", , buildParams(sltDetail, sltRadius), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltRadius.Value = 20
    sltDetail.Value = 4
End Sub

Private Sub Form_Activate()
    
    'Apply translations and visual themes
    makeFormPretty Me
    
    'Draw an initial preview of the effect
    updatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub updatePreview()
    If cmdBar.previewsAllowed Then ApplyMetalFilter sltDetail.Value, sltRadius.Value, True, fxPreview
End Sub

Private Sub sltDetail_Change()
    updatePreview
End Sub

Private Sub sltRadius_Change()
    updatePreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub


