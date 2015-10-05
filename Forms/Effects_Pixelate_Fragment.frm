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
   Begin VB.ComboBox cmbEdges 
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
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3975
      Width           =   5700
   End
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5775
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   1323
      BackColor       =   14802140
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
      Height          =   720
      Left            =   6000
      TabIndex        =   2
      Top             =   1440
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1270
      Caption         =   "distance"
      Max             =   200
      SigDigits       =   1
      Value           =   8
   End
   Begin PhotoDemon.sliderTextCombo sltFragments 
      Height          =   720
      Left            =   6000
      TabIndex        =   3
      Top             =   360
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1270
      Caption         =   "number of fragments"
      Min             =   1
      Max             =   25
      Value           =   4
   End
   Begin PhotoDemon.sliderTextCombo sltAngle 
      Height          =   720
      Left            =   6000
      TabIndex        =   4
      Top             =   2520
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1270
      Caption         =   "angle"
      Max             =   360
      SigDigits       =   2
   End
   Begin PhotoDemon.smartOptionButton OptInterpolate 
      Height          =   360
      Index           =   0
      Left            =   6120
      TabIndex        =   6
      Top             =   4920
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   582
      Caption         =   "quality"
      Value           =   -1  'True
   End
   Begin PhotoDemon.smartOptionButton OptInterpolate 
      Height          =   360
      Index           =   1
      Left            =   6120
      TabIndex        =   7
      Top             =   5280
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   582
      Caption         =   "speed"
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "render emphasis"
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
      TabIndex        =   9
      Top             =   4530
      Width           =   1755
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "if pixels lie outside the image..."
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
      Index           =   5
      Left            =   6000
      TabIndex        =   8
      Top             =   3540
      Width           =   3315
   End
End
Attribute VB_Name = "FormFragment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Fragment Filter Dialog
'Copyright 2014 by Audioglider and Tanner Helland
'Created: 10/May/14
'Last updated: 10/May/14
'Last update: lots of minor updates, fixes, and optimizations
'
'PD's Fragment effect is similar in concept to Paint.NET's Fragment tool, which is in turn much more customizable
' than Photoshop's version.
'
'Specifically, PD's Fragment tool allows the user to specify any number of fragments, their distance from the
' original pixel position, and the angle at which the fragments appear.  (Note that angle is more relevant when
' the number of fragments is low; as the fragment count increases, angle becomes less important.)  Like other
' coordinate-transform tools in the project, additional options are provided for edge handling and interpolation
' of pixel positions.
'
'Look-up tables are used to improve performance, which is especially important when the fragment count is high.
'
'Many thanks to pro developer Audioglider for contributing this great tool to PhotoDemon.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Apply a fragment filter to the active layer
Public Sub Fragment(ByVal fragCount As Long, ByVal fragDistance As Double, ByVal rotationAngle As Double, ByVal edgeHandling As Long, ByVal useBilinear As Boolean, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
   
    If Not toPreview Then Message "Applying beer goggles..."
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Create a second local array. This will contain the a copy of the current image, and we will use it as our source reference
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
    
    'If this is a preview, we need to adjust the distance values to match the size of the preview box
    If toPreview Then
        fragDistance = fragDistance * curDIBValues.previewModifier
        If fragDistance = 0 Then fragDistance = 1
    End If
    
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'Create a filter support class, which will aid with edge handling and interpolation
    Dim fSupport As pdFilterSupport
    Set fSupport = New pdFilterSupport
    fSupport.setDistortParameters qvDepth, edgeHandling, useBilinear, curDIBValues.maxX, curDIBValues.MaxY
    
    'To keep processing quick, only update the progress bar when absolutely necessary. This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Rather than recalculating fragment positions for each pixel, calculate them as offsets from (0, 0)
    ' and store those values in x/y lookup tables.
    Dim n As Long
    Dim Num As Double, num2 As Double, num3 As Double
    
    Num = PI_DOUBLE / CDbl(fragCount)
    num2 = ((rotationAngle - 90) * PI) / 180
    
    Dim xOffsetLookup() As Double
    Dim yOffsetLookup() As Double
    ReDim xOffsetLookup(0 To fragCount - 1) As Double
    ReDim yOffsetLookup(0 To fragCount - 1) As Double
    
    For n = 0 To UBound(xOffsetLookup)
        num3 = num2 + (Num * n)
        xOffsetLookup(n) = CDbl(fragDistance * -Sin(num3))
        yOffsetLookup(n) = CDbl(fragDistance * -Cos(num3))
    Next n
    
    'numPoints is the loop termination value for the fragment array (which is one less than the number of fragments,
    ' because they are stored in a zero-based array.
    Dim numPoints As Long
    numPoints = fragCount - 1
    
    'numPointsCalc is the number of fragments actually being processed.  This is the number of fragments + 1, because
    ' the original pixel is also considered a fragment.
    Dim numPointsCalc As Long
    numPointsCalc = fragCount + 1
    
    'Color variables
    Dim r As Long, g As Long, b As Long, a As Long
    Dim newR As Long, newG As Long, newB As Long, newA As Long
    
    'Pixel offsets
    Dim xOffset As Double, yOffset As Double
    
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Grab the current pixel values
        newR = srcImageData(QuickVal + 2, y)
        newG = srcImageData(QuickVal + 1, y)
        newB = srcImageData(QuickVal, y)
        If qvDepth = 4 Then newA = srcImageData(QuickVal + 3, y)
        
        'Iterate through each fragment in turn, adding together their values as we go
        For n = 0 To numPoints
        
            'Calculate an offset for this fragment.
            xOffset = x - xOffsetLookup(n)
            yOffset = y - yOffsetLookup(n)
                        
            'Use the filter support class to interpolate and edge-wrap pixels as necessary
            fSupport.getColorsFromSource r, g, b, a, xOffset, yOffset, srcImageData
            
            'Add the retrieved values to our running average
            newR = newR + r
            newG = newG + g
            newB = newB + b
            If qvDepth = 4 Then newA = newA + a
            
        Next n
        
        'Take the average of all fragments, and apply them to the image
        newR = newR \ numPointsCalc
        newG = newG \ numPointsCalc
        newB = newB \ numPointsCalc
                
        dstImageData(QuickVal + 2, y) = newR
        dstImageData(QuickVal + 1, y) = newG
        dstImageData(QuickVal, y) = newB
        
        'If the image has an alpha channel, repeat the calculation there too
        If qvDepth = 4 Then
            newA = newA \ numPointsCalc
            dstImageData(QuickVal + 3, y) = newA
        End If
        
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

Private Sub cmbEdges_Click()
    updatePreview
End Sub

Private Sub cmdBar_OKClick()
    Process "Fragment", , buildParams(sltFragments.Value, sltDistance.Value, sltAngle.Value, CLng(cmbEdges.ListIndex), OptInterpolate(0).Value), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltFragments.Value = 4
    sltDistance.Value = 8
    sltAngle.Value = 0
    cmbEdges.ListIndex = EDGE_CLAMP
    OptInterpolate(1).Value = True
End Sub

Private Sub Form_Activate()
        
    'Apply translations and visual themes
    MakeFormPretty Me
        
    'Create the preview
    cmdBar.markPreviewStatus True
    updatePreview
    
End Sub

Private Sub Form_Load()

    'Disable previews until the dialog has been fully initialized
    cmdBar.markPreviewStatus False
    
    'I use a central function to populate the edge handling combo box; this way, I can add new methods and have
    ' them immediately available to all distort functions.
    PopDistortEdgeBox cmbEdges, EDGE_CLAMP

End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub updatePreview()
    If cmdBar.previewsAllowed Then Fragment sltFragments, sltDistance, sltAngle, CLng(cmbEdges.ListIndex), OptInterpolate(0).Value, True, fxPreview
End Sub

Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

Private Sub OptInterpolate_Click(Index As Integer)
    updatePreview
End Sub

Private Sub sltAngle_Change()
    updatePreview
End Sub

Private Sub sltDistance_Change()
    updatePreview
End Sub

Private Sub sltFragments_Change()
    updatePreview
End Sub
