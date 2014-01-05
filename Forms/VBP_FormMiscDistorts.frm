VERSION 5.00
Begin VB.Form FormMiscDistorts 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Miscellaneous Distort Tools"
   ClientHeight    =   6540
   ClientLeft      =   -15
   ClientTop       =   225
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
   Begin VB.ListBox lstDistorts 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   2460
      Left            =   6120
      TabIndex        =   8
      Top             =   960
      Width           =   5655
   End
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
      TabIndex        =   4
      Top             =   4005
      Width           =   5700
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.smartOptionButton OptInterpolate 
      Height          =   330
      Index           =   0
      Left            =   6120
      TabIndex        =   6
      Top             =   4920
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   635
      Caption         =   "quality"
      Value           =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.smartOptionButton OptInterpolate 
      Height          =   330
      Index           =   1
      Left            =   7920
      TabIndex        =   7
      Top             =   4920
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   635
      Caption         =   "speed"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "if pixels lie outside the corrected area..."
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
      TabIndex        =   5
      Top             =   3600
      Width           =   4170
   End
   Begin VB.Label lblInterpolation 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "render emphasis:"
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
      Top             =   4530
      Width           =   1845
   End
   Begin VB.Label lblType 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "distortions:"
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
      TabIndex        =   1
      Top             =   570
      Width           =   1200
   End
End
Attribute VB_Name = "FormMiscDistorts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Miscellaneous Distort Tools
'Copyright ©2013-2014 by Tanner Helland
'Created: 07/June/13
'Last updated: 07/June/13
'Last update: initial build
'
'Some one-off distorts (e.g. no tunable parameters) are useful under very specific circumstances.  However, it is
' impractical to give every such tool its own menu entry, so all non-tunable distorts are being placed here from
' now on.
'
'Bilinear interpolation is available to improve output quality.
'
'Certain transformations aer modified versions of basic math originally shared by Paul Bourke. You can see Paul's
' original (and very helpful article) at the following link, good as of 07 June '13:
' http://paulbourke.net/miscellaneous/imagewarp/
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'Correct lens distortion in an image
Public Sub ApplyMiscDistort(ByVal distortName As String, ByVal distortStyle As Long, ByVal edgeHandling As Long, ByVal useBilinear As Boolean, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If toPreview = False Then Message "Applying %1 distortion...", distortName
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent diffused pixels from spreading across the image as we go.)
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    
    Dim srcLayer As pdLayer
    Set srcLayer = New pdLayer
    srcLayer.createFromExistingLayer workingLayer
    
    prepSafeArray srcSA, srcLayer
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
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
    
    'Create a filter support class, which will aid with edge handling and interpolation
    Dim fSupport As pdFilterSupport
    Set fSupport = New pdFilterSupport
    fSupport.setDistortParameters qvDepth, edgeHandling, useBilinear, curLayerValues.maxX, curLayerValues.MaxY
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Calculate the center of the image
    Dim midX As Double, midY As Double
    midX = CDbl(finalX - initX) / 2
    midX = midX + initX
    midY = CDbl(finalY - initY) / 2
    midY = midY + initY
    
    'Rotation values
    Dim theta As Double, r As Double
    
    'X and Y values, remapped around a center point of (0, 0)
    Dim nX As Double, nY As Double
    
    'Source X and Y values, which may or may not be used as part of a bilinear interpolation function
    Dim srcX As Double, srcY As Double
        
    'Max radius is calculated as the distance from the center of the image to a corner
    Dim tWidth As Long, tHeight As Long
    tWidth = curLayerValues.Width
    tHeight = curLayerValues.Height
    
    'Because coordinates will be mapped identically for each x-coord and y-coord, we can calculate them in advance
    ' and store them in lookup tables to improve performance.
    Dim xCoords() As Double, yCoords() As Double
    ReDim xCoords(initX To finalX) As Double
    ReDim yCoords(initY To finalY) As Double
    
    'Basically, we want to remap coordinates around a center point of (0, 0), and normalize them to (-1, 1).
    ' This makes distort strength uniform regardless of image size.
    For x = initX To finalX
        xCoords(x) = (2 * x) / tWidth - 1
    Next x
    
    For y = initY To finalY
        yCoords(y) = (2 * y) / tHeight - 1
    Next y
    
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
                            
        'Pull coordinates from the lookup table
        nX = xCoords(x)
        nY = yCoords(y)
        
        'Next, map them to polar coordinates
        r = Sqr(nX * nX + nY * nY)
        theta = Atan2(nY, nX)
        
        Select Case distortStyle
            
            'Emphasize center
            Case 0
                nX = 2 * Asin(nX) / PI
                nY = 2 * Asin(nY) / PI
            
            'Flatten corners
            Case 1
                nX = Sin(nX)
                nY = Sin(nY)
                
            'Inside-out
            Case 2
                If r > 0 Then r = 1 - r Else r = -1 - r
                nX = r * Cos(theta)
                nY = r * Sin(theta)
                
            'Pull in
            Case 3
                r = Sqr(r)
                nX = r * Cos(theta)
                nY = r * Sin(theta)
            
            'Push out
            Case 4
                r = r * r
                nX = r * Cos(theta)
                nY = r * Sin(theta)
            
            'Rounding
            Case 5
                If nX < 0 Then
                    nX = -1 * nX * nX
                Else
                    nX = nX * nX
                End If
                If nY < 0 Then
                    nY = -1 * nY * nY
                Else
                    nY = nY * nY
                End If
                
            'Twist edges
            Case 6
                r = Sin(PI * r / 2)
                nX = r * Cos(theta)
                nY = r * Sin(theta)
                
            'Wormhole
            Case 7
                If r = 0 Then r = 0 Else r = Sin(1 / r)
                nX = r * Cos(theta)
                nY = r * Sin(theta)
            
        End Select
        
        'Convert the recalculated coordinates back to the Cartesian plane
        srcX = (tWidth * (nX + 1)) / 2
        srcY = (tHeight * (nY + 1)) / 2
        
        'The lovely .setPixels routine will handle edge detection and interpolation for us as necessary
        fSupport.setPixels x, y, srcX, srcY, srcImageData, dstImageData
                
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
    Process "Miscellaneous distort", , buildParams(lstDistorts.List(lstDistorts.ListIndex), lstDistorts.ListIndex, CLng(cmbEdges.ListIndex), OptInterpolate(0).Value)
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    cmbEdges.ListIndex = EDGE_WRAP
End Sub

Private Sub Form_Activate()
        
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
    'Draw a preview of the effect
    cmdBar.markPreviewStatus True
    updatePreview
            
End Sub

Private Sub Form_Load()
    
    'Disable previews while we populate various dialog controls
    cmdBar.markPreviewStatus False
    
    'Populate a list of available distort operations
    lstDistorts.Clear
    lstDistorts.AddItem g_Language.TranslateMessage("emphasize center"), 0
    lstDistorts.AddItem g_Language.TranslateMessage("flatten corners"), 1
    lstDistorts.AddItem g_Language.TranslateMessage("inside-out"), 2
    lstDistorts.AddItem g_Language.TranslateMessage("pull in"), 3
    lstDistorts.AddItem g_Language.TranslateMessage("push out"), 4
    lstDistorts.AddItem g_Language.TranslateMessage("ring"), 5
    lstDistorts.AddItem g_Language.TranslateMessage("twist edges"), 6
    lstDistorts.AddItem g_Language.TranslateMessage("wormhole"), 7
    
    'I use a central function to populate the edge handling combo box; this way, I can add new methods and have
    ' them immediately available to all distort functions.
    popDistortEdgeBox cmbEdges, EDGE_WRAP
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub lstDistorts_Click()
    updatePreview
End Sub

Private Sub OptInterpolate_Click(Index As Integer)
    updatePreview
End Sub

'Redraw the on-screen preview of the transformed image
Private Sub updatePreview()
    If cmdBar.previewsAllowed Then ApplyMiscDistort "", lstDistorts.ListIndex, CLng(cmbEdges.ListIndex), OptInterpolate(0).Value, True, fxPreview
End Sub
