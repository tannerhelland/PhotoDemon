VERSION 5.00
Begin VB.Form FormTwins 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Twins"
   ClientHeight    =   6525
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11310
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
   ScaleHeight     =   435
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   754
   Begin PhotoDemon.pdButtonStrip btsOrientation 
      Height          =   1095
      Left            =   6000
      TabIndex        =   2
      Top             =   2040
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   1931
      Caption         =   "orientation"
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5775
      Width           =   11310
      _ExtentX        =   19950
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
      DisableZoomPan  =   -1  'True
   End
End
Attribute VB_Name = "FormTwins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'"Twin" Filter Interface
'Copyright 2001-2026 by Tanner Helland
'Created: 6/12/01
'Last updated: 08/August/17
'Last update: migrate to XML params, minor performance improvements
'
'Unoptimized "twin" generator.  Simple 50% alpha blending combined with a flip.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'This routine mirrors and alphablends an image, making it "tilable" or symmetrical
Public Sub GenerateTwins(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
   
    If (Not toPreview) Then Message "Generating image twin..."
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim tType As Long
    
    With cParams
        tType = .GetLong("orientation", btsOrientation.ListIndex)
    End With
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte, dstSA As SafeArray2D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic, doNotUnPremultiplyAlpha:=True
    workingDIB.WrapArrayAroundDIB dstImageData, dstSA
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent already-processed pixels from affecting the results of later pixels.)
    Dim srcImageData() As Byte, srcSA As SafeArray2D
    
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.CreateFromExistingDIB workingDIB
    srcDIB.WrapArrayAroundDIB srcImageData, srcSA
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    Dim xStride As Long
    
    'Pre-calculate the largest possible processed x-value
    Dim maxX As Long
    maxX = finalX * 4
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = ProgressBars.FindBestProgBarValue()
            
    'This look-up table will be used for alpha-blending.  It contains the equivalent of any two color values [0,255] added
    ' together and divided by 2.
    Dim hLookup(0 To 510) As Byte
    For x = 0 To 510
        hLookup(x) = x \ 2
    Next x
    
    'Color variables
    Dim r As Long, g As Long, b As Long, a As Long
    Dim r2 As Long, g2 As Long, b2 As Long, a2 As Long
    
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        xStride = x * 4
    For y = initY To finalY
    
        'Grab the current pixel values
        b = srcImageData(xStride, y)
        g = srcImageData(xStride + 1, y)
        r = srcImageData(xStride + 2, y)
        a = srcImageData(xStride + 3, y)
        
        'Grab the value of the "second" pixel, whose position will vary depending on the method (vertical or horizontal)
        If (tType = 0) Then
            b2 = srcImageData(maxX - xStride, y)
            g2 = srcImageData(maxX - xStride + 1, y)
            r2 = srcImageData(maxX - xStride + 2, y)
            a2 = srcImageData(maxX - xStride + 3, y)
        Else
            b2 = srcImageData(xStride, finalY - y)
            g2 = srcImageData(xStride + 1, finalY - y)
            r2 = srcImageData(xStride + 2, finalY - y)
            a2 = srcImageData(xStride + 3, finalY - y)
        End If
        
        'Alpha-blend the two pixels using our shortcut look-up table
        dstImageData(xStride, y) = hLookup(b + b2)
        dstImageData(xStride + 1, y) = hLookup(g + g2)
        dstImageData(xStride + 2, y) = hLookup(r + r2)
        dstImageData(xStride + 3, y) = hLookup(a + a2)
        
    Next y
        If (Not toPreview) Then
            If (x And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal x
            End If
        End If
    Next x
    
    'Safely deallocate all image arrays
    workingDIB.UnwrapArrayFromDIB dstImageData
    srcDIB.UnwrapArrayFromDIB srcImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic, True
        
End Sub

Private Sub btsOrientation_Click(ByVal buttonIndex As Long)
    UpdatePreview
End Sub

Private Sub cmdBar_OKClick()
    Process "Twins", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub Form_Load()
    
    cmdBar.SetPreviewStatus False
    
    btsOrientation.AddItem "horizontal", 0
    btsOrientation.AddItem "vertical", 1
    btsOrientation.ListIndex = 0
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then Me.GenerateTwins GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "orientation", btsOrientation.ListIndex
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
