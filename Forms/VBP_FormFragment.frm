VERSION 5.00
Begin VB.Form FormFragment 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Fragment"
   ClientHeight    =   6525
   ClientLeft      =   -15
   ClientTop       =   225
   ClientWidth     =   11895
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
   ScaleWidth      =   793
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5775
      Width           =   11895
      _ExtentX        =   20981
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
   Begin PhotoDemon.sliderTextCombo sltDistance 
      Height          =   495
      Left            =   6000
      TabIndex        =   3
      Top             =   2520
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   873
      Max             =   50
      Value           =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
   End
   Begin VB.Label lblAlgorithm 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "distance:"
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
      Width           =   945
   End
End
Attribute VB_Name = "FormFragment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Fragment Filter Dialog
'Copyright ©2014 by Audioglider
'Created: 09/May/14
'Last updated: 09/May/14
'Last update: added handling for 32bpp images (e.g. alpha channels)
'
'Similar to the Fragment effect from Photoshop except adjustable.
' We create 4 layers and offset them the same distance from the origin,
' but at different positions (top, bottom, left and right), then merge them
' all together.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'Apply a fragment filter to the active layer
Public Sub Fragment(ByVal Distance As Long, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
   
    If Not toPreview Then Message "Applying beer goggles..."
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
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
    Dim newR As Long, newG As Long, newB As Long, newA As Long
    
    Dim r2 As Long, g2 As Long, b2 As Long, a2 As Long
    Dim r3 As Long, g3 As Long, b3 As Long, a3 As Long
    Dim r4 As Long, g4 As Long, b4 As Long, a4 As Long
    Dim r5 As Long, g5 As Long, b5 As Long, a5 As Long
    
    Dim xOffset As Long, yOffset As Long
        
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Grab the current pixel values
        r = srcImageData(QuickVal + 2, y)
        g = srcImageData(QuickVal + 1, y)
        b = srcImageData(QuickVal, y)
        If qvDepth = 4 Then a = srcImageData(QuickVal + 3, y)
        
        'Bottom
        If y + Distance >= finalY Then
            yOffset = finalY - 1
        Else
            yOffset = y + Distance
        End If
        r2 = srcImageData(QuickVal + 2, yOffset)
        g2 = srcImageData(QuickVal + 1, yOffset)
        b2 = srcImageData(QuickVal, yOffset)
        If qvDepth = 4 Then a2 = srcImageData(QuickVal + 3, yOffset)
        
        'Top
        If y - Distance < initY Then
            yOffset = initY
        Else
            yOffset = y - Distance
        End If
        r3 = srcImageData(QuickVal + 2, yOffset)
        g3 = srcImageData(QuickVal + 1, yOffset)
        b3 = srcImageData(QuickVal, yOffset)
        If qvDepth = 4 Then a3 = srcImageData(QuickVal + 3, yOffset)
        
        'Right
        If x + Distance >= finalX Then
            xOffset = maxX
        Else
            xOffset = (x + Distance) * qvDepth
        End If
        r4 = srcImageData(xOffset + 2, y)
        g4 = srcImageData(xOffset + 1, y)
        b4 = srcImageData(xOffset, y)
        If qvDepth = 4 Then a4 = srcImageData(xOffset + 3, y)
                
        'Left
        If x - Distance < 0 Then
            xOffset = 0
        Else
            xOffset = Abs(x - Distance) * qvDepth
        End If
        r5 = srcImageData(xOffset + 2, y)
        g5 = srcImageData(xOffset + 1, y)
        b5 = srcImageData(xOffset, y)
        If qvDepth = 4 Then a5 = srcImageData(xOffset + 3, y)
        
        'Alpha-blend the the four layers using our shortcut look-up table
        newR = (CLng(hLookup(r + r2)) + CLng(hLookup(r2 + r3)) + CLng(hLookup(r3 + r4)) + CLng(hLookup(r4 + r5))) \ 4
        newG = (CLng(hLookup(g + g2)) + CLng(hLookup(g2 + g3)) + CLng(hLookup(g3 + g4)) + CLng(hLookup(g4 + g5))) \ 4
        newB = (CLng(hLookup(b + b2)) + CLng(hLookup(b2 + b3)) + CLng(hLookup(b3 + b4)) + CLng(hLookup(b4 + b5))) \ 4
        If qvDepth = 4 Then newA = (CLng(hLookup(a + a2)) + CLng(hLookup(a2 + a3)) + CLng(hLookup(a3 + a4)) + CLng(hLookup(a4 + a5))) \ 4
      
        dstImageData(QuickVal + 2, y) = newR
        dstImageData(QuickVal + 1, y) = newG
        dstImageData(QuickVal, y) = newB
        If qvDepth = 4 Then dstImageData(QuickVal + 3, y) = newA
        
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
    finalizeImageData toPreview, dstPic
        
End Sub

Private Sub cmdBar_OKClick()
    Process "Fragment", False, buildParams(sltDistance.Value)
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltDistance.Value = 4
End Sub

Private Sub Form_Activate()
        
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
    'Render an image preview
    updatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub updatePreview()
    If cmdBar.previewsAllowed Then Fragment sltDistance, True, fxPreview
End Sub

Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

Private Sub sltDistance_Change()
    updatePreview
End Sub
