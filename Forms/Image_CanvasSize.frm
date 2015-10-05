VERSION 5.00
Begin VB.Form FormCanvasSize 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Resize Canvas"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   9705
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
   ScaleHeight     =   512
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   647
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdButton cmdAnchor 
      Height          =   570
      Index           =   0
      Left            =   840
      TabIndex        =   5
      Top             =   3720
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1005
   End
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   6930
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   1323
      BackColor       =   14802140
      AutoloadLastPreset=   -1  'True
   End
   Begin PhotoDemon.colorSelector colorPicker 
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   6120
      Width           =   7935
      _ExtentX        =   10398
      _ExtentY        =   873
   End
   Begin PhotoDemon.smartResize ucResize 
      Height          =   2850
      Left            =   360
      TabIndex        =   4
      Top             =   360
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   5027
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
   Begin PhotoDemon.pdButton cmdAnchor 
      Height          =   570
      Index           =   1
      Left            =   1680
      TabIndex        =   6
      Top             =   3720
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1005
   End
   Begin PhotoDemon.pdButton cmdAnchor 
      Height          =   570
      Index           =   2
      Left            =   2520
      TabIndex        =   7
      Top             =   3720
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1005
   End
   Begin PhotoDemon.pdButton cmdAnchor 
      Height          =   570
      Index           =   3
      Left            =   840
      TabIndex        =   8
      Top             =   4320
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1005
   End
   Begin PhotoDemon.pdButton cmdAnchor 
      Height          =   570
      Index           =   4
      Left            =   1680
      TabIndex        =   9
      Top             =   4320
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1005
   End
   Begin PhotoDemon.pdButton cmdAnchor 
      Height          =   570
      Index           =   5
      Left            =   2520
      TabIndex        =   10
      Top             =   4320
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1005
   End
   Begin PhotoDemon.pdButton cmdAnchor 
      Height          =   570
      Index           =   6
      Left            =   840
      TabIndex        =   11
      Top             =   4920
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1005
   End
   Begin PhotoDemon.pdButton cmdAnchor 
      Height          =   570
      Index           =   7
      Left            =   1680
      TabIndex        =   12
      Top             =   4920
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1005
   End
   Begin PhotoDemon.pdButton cmdAnchor 
      Height          =   570
      Index           =   8
      Left            =   2520
      TabIndex        =   13
      Top             =   4920
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1005
   End
   Begin VB.Label lblAnchor 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "anchor position"
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
      Left            =   360
      TabIndex        =   1
      Top             =   3360
      Width           =   1635
   End
   Begin VB.Label lblFill 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "fill empty areas with"
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
      Left            =   360
      TabIndex        =   2
      Top             =   5760
      Width           =   2145
   End
End
Attribute VB_Name = "FormCanvasSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Canvas Size Handler
'Copyright 2013-2015 by Tanner Helland
'Created: 13/June/13
'Last updated: 14/April/14
'Last update: rewrite everything against layers
'
'This form handles canvas resizing.  You may wonder why it took me over a decade to implement this tool, when it's such a
' trivial one algorithmically.  The answer is that a number of user-interface support functions are necessary to build
' this tool correctly, primarily the command buttons used to select an anchor location.  These require the ability to
' apply 32bpp images to command buttons at run-time, which I lacked for many years.
'
'But now I have such tools at my disposal, so no excuses!  :)  The resulting tool should be self-explanatory.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Current anchor position; used to render the anchor selection command buttons, among other things
Private m_CurrentAnchor As Long

'We must also track which arrows are drawn where on the command button array
Private arrowLocations() As String

Private Sub fillArrowLocations(ByRef aLocations() As String)

    'Start with the current position.  It's the easiest one to fill
    aLocations(m_CurrentAnchor) = "IMGMEDIUM"
    
    'Next, fill in upward arrows as necessary
    If m_CurrentAnchor > 2 Then
        aLocations(m_CurrentAnchor - 3) = "MARROWUP"
        If (m_CurrentAnchor Mod 3) <> 0 Then aLocations(m_CurrentAnchor - 4) = "MARROWUPL"
        If ((m_CurrentAnchor + 1) Mod 3) <> 0 Then aLocations(m_CurrentAnchor - 2) = "MARROWUPR"
    End If
    
    'Next, fill in left/right arrows as necessary
    If ((m_CurrentAnchor + 1) Mod 3) <> 0 Then aLocations(m_CurrentAnchor + 1) = "MARROWRIGHT"
    If (m_CurrentAnchor Mod 3) <> 0 Then aLocations(m_CurrentAnchor - 1) = "MARROWLEFT"
    
    'Finally, fill in downward arrows as necessary
    If m_CurrentAnchor < 6 Then
        aLocations(m_CurrentAnchor + 3) = "MARROWDOWN"
        If (m_CurrentAnchor Mod 3) <> 0 Then aLocations(m_CurrentAnchor + 2) = "MARROWDOWNL"
        If ((m_CurrentAnchor + 1) Mod 3) <> 0 Then aLocations(m_CurrentAnchor + 4) = "MARROWDOWNR"
    End If
    
End Sub

'The user can use an array of command buttons to specify the image's anchor position on the new canvas.  I adopted this
' model from comparable tools in Photoshop and Paint.NET, among others.  Images are loaded from the resource section
' of the EXE and applied to the command buttons as necessary.
Private Sub updateAnchorButtons()
    
    Dim i As Long
    
    'Build an array that contains the arrow to appear in each location.
    ReDim arrowLocations(0 To 8) As String
    fillArrowLocations arrowLocations
    
    'Next, extract relevant icons from the resource file, and render them onto the buttons at run-time.
    For i = 0 To 8
        If Len(arrowLocations(i)) <> 0 Then
            cmdAnchor(i).AssignImage arrowLocations(i)
        Else
            cmdAnchor(i).AssignImage "", Nothing
        End If
    Next i
    
End Sub

Private Sub cmdAnchor_Click(Index As Integer)
    m_CurrentAnchor = Index
    updateAnchorButtons
End Sub

'The current anchor must be manually saved as part of preset data
Private Sub cmdBar_AddCustomPresetData()
    cmdBar.addPresetData "currentAnchor", Str(m_CurrentAnchor)
End Sub

Private Sub cmdBar_ExtraValidations()
    If Not ucResize.IsValid(True) Then cmdBar.validationFailed
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Canvas size", , buildParams(ucResize.imgWidth, ucResize.imgHeight, m_CurrentAnchor, colorPicker.Color, ucResize.unitOfMeasurement, ucResize.imgDPIAsPPI), UNDO_IMAGEHEADER
End Sub

'I'm not sure that randomize serves any purpose on this dialog, but as I don't have a way to hide that button at
' present, simply randomize the width/height to +/- the current image's width/height divided by two.
Private Sub cmdBar_RandomizeClick()
    
    ucResize.lockAspectRatio = False
    ucResize.imgWidthInPixels = (pdImages(g_CurrentImage).Width / 2) + (Rnd * pdImages(g_CurrentImage).Width)
    ucResize.imgHeightInPixels = (pdImages(g_CurrentImage).Height / 2) + (Rnd * pdImages(g_CurrentImage).Height)
    
End Sub

'The saved anchor must be custom-loaded, as the command bar won't handle it automatically
Private Sub cmdBar_ReadCustomPresetData()
    m_CurrentAnchor = CLng(cmdBar.retrievePresetData("currentAnchor"))
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updateAnchorButtons
End Sub

Private Sub cmdBar_ResetClick()

    'Automatically set the width and height text boxes to match the image's current dimensions
    ucResize.unitOfMeasurement = MU_PIXELS
    ucResize.setInitialDimensions pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height, pdImages(g_CurrentImage).getDPI
    ucResize.lockAspectRatio = True
    
    'Make borders fill with black by default
    colorPicker.Color = RGB(0, 0, 0)
    
    'Set the middle position as the anchor
    m_CurrentAnchor = 4

End Sub

'Upon form activation, determine the ratio between the width and height of the image
Private Sub Form_Activate()
    
    'Apply translations and visual themes
    MakeFormPretty Me
        
End Sub

'Certain actions are done at LOAD time instead of ACTIVATE time to minimize visible flickering
Private Sub Form_Load()
    
    'Automatically set the width and height text boxes to match the image's current dimensions
    ucResize.setInitialDimensions pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height, pdImages(g_CurrentImage).getDPI
    
    'Start with a default top-left position for the anchor
    updateAnchorButtons
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Resize an image using any one of several resampling algorithms.  (Some algorithms are provided by FreeImage.)
Public Sub ResizeCanvas(ByVal iWidth As Long, ByVal iHeight As Long, ByVal anchorPosition As Long, Optional ByVal newBackColor As Long = vbWhite, Optional ByVal unitOfMeasurement As MeasurementUnit = MU_PIXELS, Optional ByVal iDPI As Long)

    Dim srcWidth As Long, srcHeight As Long
    srcWidth = pdImages(g_CurrentImage).Width
    srcHeight = pdImages(g_CurrentImage).Height
    
    'In past versions of the software, we could assume the passed measurements were always in pixels,
    ' but that is no longer the case!  Using the supplied "unit of measurement", convert the passed
    ' width and height values to pixel measurements.
    iWidth = convertOtherUnitToPixels(unitOfMeasurement, iWidth, iDPI, srcWidth)
    iHeight = convertOtherUnitToPixels(unitOfMeasurement, iHeight, iDPI, srcHeight)
    
    'If the image contains an active selection, disable it before transforming the canvas
    If pdImages(g_CurrentImage).selectionActive Then
        pdImages(g_CurrentImage).selectionActive = False
        pdImages(g_CurrentImage).mainSelection.lockRelease
    End If
    
    'Based on the anchor position, determine x and y locations for the image on the new canvas
    Dim dstX As Long, dstY As Long
    
    Select Case anchorPosition
    
        'Top-left
        Case 0
            dstX = 0
            dstY = 0
        
        'Top-center
        Case 1
            dstX = (iWidth - srcWidth) \ 2
            dstY = 0
        
        'Top-right
        Case 2
            dstX = (iWidth - srcWidth)
            dstY = 0
        
        'Middle-left
        Case 3
            dstX = 0
            dstY = (iHeight - srcHeight) \ 2
        
        'Middle-center
        Case 4
            dstX = (iWidth - srcWidth) \ 2
            dstY = (iHeight - srcHeight) \ 2
        
        'Middle-right
        Case 5
            dstX = (iWidth - srcWidth)
            dstY = (iHeight - srcHeight) \ 2
        
        'Bottom-left
        Case 6
            dstX = 0
            dstY = (iHeight - srcHeight)
        
        'Bottom-center
        Case 7
            dstX = (iWidth - srcWidth) \ 2
            dstY = (iHeight - srcHeight)
        
        'Bottom right
        Case 8
            dstX = (iWidth - srcWidth)
            dstY = (iHeight - srcHeight)
    
    End Select
    
    'Now that we have our new top-left corner coordinates (and new width/height values), resizing the canvas
    ' is actually very easy.  In PhotoDemon, there is no such thing as "image data"; an image is just an
    ' imaginary bounding box around the layers collection.  Because of this, we don't actually need to
    ' resize any pixel data - we just need to modify all layer offsets to account for the new top-left corner!
    Dim i As Long
    For i = 0 To pdImages(g_CurrentImage).getNumOfLayers - 1
    
        With pdImages(g_CurrentImage).getLayerByIndex(i)
            .setLayerOffsetX .getLayerOffsetX + dstX
            .setLayerOffsetY .getLayerOffsetY + dstY
        End With
    
    Next i
    
    'Finally, update the parent image's size and DPI values
    pdImages(g_CurrentImage).updateSize False, iWidth, iHeight
    pdImages(g_CurrentImage).setDPI iDPI, iDPI
    DisplaySize pdImages(g_CurrentImage)
    
    'In other functions, we would refresh the layer box here; however, because we haven't actually changed the
    ' appearance of any of the layers, we can leave it as-is!
    
    'Fit the new image on-screen and redraw its viewport
    Viewport_Engine.Stage1_InitializeBuffer pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
    Message "Finished."
    
End Sub
