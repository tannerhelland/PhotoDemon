VERSION 5.00
Begin VB.Form FormTwins 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Generate Twins"
   ClientHeight    =   6525
   ClientLeft      =   -15
   ClientTop       =   225
   ClientWidth     =   9555
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
   ScaleHeight     =   435
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   637
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5775
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   1323
      BackColor       =   14802140
   End
   Begin PhotoDemon.smartOptionButton optTwins 
      Height          =   375
      Index           =   0
      Left            =   6120
      TabIndex        =   3
      Top             =   2640
      Width           =   3240
      _ExtentX        =   5715
      _ExtentY        =   582
      Caption         =   "horizontal"
      Value           =   -1  'True
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
      DisableZoomPan  =   -1  'True
   End
   Begin PhotoDemon.smartOptionButton optTwins 
      Height          =   375
      Index           =   1
      Left            =   6120
      TabIndex        =   4
      Top             =   3090
      Width           =   3240
      _ExtentX        =   5715
      _ExtentY        =   582
      Caption         =   "vertical"
   End
   Begin VB.Label lblAlgorithm 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "orientation"
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
      TabIndex        =   2
      Top             =   2160
      Width           =   1140
   End
End
Attribute VB_Name = "FormTwins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'"Twin" Filter Interface
'Copyright 2001-2015 by Tanner Helland
'Created: 6/12/01
'Last updated: 24/August/13
'Last update: added command bar
'
'Unoptimized "twin" generator.  Simple 50% alpha blending combined with a flip.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'This routine mirrors and alphablends an image, making it "tilable" or symmetrical
Public Sub GenerateTwins(ByVal tType As Long, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
   
    If Not toPreview Then Message "Generating image twin..."
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic, , , True
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent already-processed pixels from affecting the results of later pixels.)
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.createFromExistingDIB workingDIB
    
    prepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
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
    
    'Pre-calculate the largest possible processed x-value
    Dim maxX As Long
    maxX = finalX * qvDepth
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
            
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
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Grab the current pixel values
        r = srcImageData(QuickVal + 2, y)
        g = srcImageData(QuickVal + 1, y)
        b = srcImageData(QuickVal, y)
        If qvDepth = 4 Then a = srcImageData(QuickVal + 3, y)
        
        'Grab the value of the "second" pixel, whose position will vary depending on the method (vertical or horizontal)
        If tType = 0 Then
            r2 = srcImageData(maxX - QuickVal + 2, y)
            g2 = srcImageData(maxX - QuickVal + 1, y)
            b2 = srcImageData(maxX - QuickVal, y)
            If qvDepth = 4 Then a2 = srcImageData(maxX - QuickVal + 3, y)
        Else
            r2 = srcImageData(QuickVal + 2, finalY - y)
            g2 = srcImageData(QuickVal + 1, finalY - y)
            b2 = srcImageData(QuickVal, finalY - y)
            If qvDepth = 4 Then a2 = srcImageData(QuickVal + 3, finalY - y)
        End If
        
        'Alpha-blend the two pixels using our shortcut look-up table
        dstImageData(QuickVal + 2, y) = hLookup(r + r2)
        dstImageData(QuickVal + 1, y) = hLookup(g + g2)
        dstImageData(QuickVal, y) = hLookup(b + b2)
        
        If qvDepth = 4 Then dstImageData(QuickVal + 3, y) = hLookup(a + a2)
        
    Next y
        If Not toPreview Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x
            End If
        End If
    Next x
    
    'With our work complete, point both ImageData() arrays away from their DIBs and deallocate them
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    Erase srcImageData
    
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    Erase dstImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic, True
        
End Sub

Private Sub cmdBar_OKClick()
    Process "Twins", , buildParams(CLng(optTwins(1))), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub Form_Activate()
        
    'Apply translations and visual themes
    MakeFormPretty Me
    
    'Render an image preview
    updatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub optTwins_Click(Index As Integer)
    updatePreview
End Sub

Private Sub updatePreview()
    If cmdBar.previewsAllowed Then GenerateTwins optTwins(1), True, fxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

