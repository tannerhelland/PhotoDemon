VERSION 5.00
Begin VB.Form FormCanvasSize 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Canvas size"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9705
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
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   647
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdButton cmdAnchor 
      Height          =   570
      Index           =   0
      Left            =   840
      TabIndex        =   4
      Top             =   3720
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1005
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   1323
      AutoloadLastPreset=   -1  'True
   End
   Begin PhotoDemon.pdResize ucResize 
      Height          =   2850
      Left            =   360
      TabIndex        =   3
      Top             =   360
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   5027
   End
   Begin PhotoDemon.pdButton cmdAnchor 
      Height          =   570
      Index           =   1
      Left            =   1680
      TabIndex        =   5
      Top             =   3720
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1005
   End
   Begin PhotoDemon.pdButton cmdAnchor 
      Height          =   570
      Index           =   2
      Left            =   2520
      TabIndex        =   6
      Top             =   3720
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1005
   End
   Begin PhotoDemon.pdButton cmdAnchor 
      Height          =   570
      Index           =   3
      Left            =   840
      TabIndex        =   7
      Top             =   4320
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1005
   End
   Begin PhotoDemon.pdButton cmdAnchor 
      Height          =   570
      Index           =   4
      Left            =   1680
      TabIndex        =   8
      Top             =   4320
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1005
   End
   Begin PhotoDemon.pdButton cmdAnchor 
      Height          =   570
      Index           =   5
      Left            =   2520
      TabIndex        =   9
      Top             =   4320
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1005
   End
   Begin PhotoDemon.pdButton cmdAnchor 
      Height          =   570
      Index           =   6
      Left            =   840
      TabIndex        =   10
      Top             =   4920
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1005
   End
   Begin PhotoDemon.pdButton cmdAnchor 
      Height          =   570
      Index           =   7
      Left            =   1680
      TabIndex        =   1
      Top             =   4920
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1005
   End
   Begin PhotoDemon.pdButton cmdAnchor 
      Height          =   570
      Index           =   8
      Left            =   2520
      TabIndex        =   2
      Top             =   4920
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1005
   End
   Begin PhotoDemon.pdLabel lblAnchor 
      Height          =   285
      Left            =   360
      Top             =   3360
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   503
      Caption         =   "anchor position"
      FontSize        =   12
      ForeColor       =   4210752
   End
End
Attribute VB_Name = "FormCanvasSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Canvas Size Handler
'Copyright 2013-2026 by Tanner Helland
'Created: 13/June/13
'Last updated: 09/January/17
'Last update: overhaul anchor point code to use arrows rendered at run-time (instead of fixed resources)
'
'This form handles canvas resizing.  You may wonder why it took me over a decade to implement this tool, when it's such a
' trivial one algorithmically.  The answer is that a number of user-interface support functions are necessary to build
' this tool correctly, primarily the command buttons used to select an anchor location.  These require the ability to
' apply 32bpp images to command buttons at run-time, which I lacked for many years.
'
'But now I have such tools at my disposal, so no excuses!  :)  The resulting tool should be self-explanatory.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Current anchor position; used to render the anchor selection command buttons, among other things
Private m_CurrentAnchor As Long

'We must also track which arrows are drawn where on the command button array
Private m_ArrowLocations() As String

Private Sub FillArrowLocations(ByRef aLocations() As String)

    'Start with the current position.  It's the easiest one to fill
    aLocations(m_CurrentAnchor) = "generic_image"
    
    'Next, fill in upward arrows as necessary
    If (m_CurrentAnchor > 2) Then
        aLocations(m_CurrentAnchor - 3) = "arrow_up"
        If ((m_CurrentAnchor Mod 3) <> 0) Then aLocations(m_CurrentAnchor - 4) = "arrow_upl"
        If (((m_CurrentAnchor + 1) Mod 3) <> 0) Then aLocations(m_CurrentAnchor - 2) = "arrow_upr"
    End If
    
    'Next, fill in left/right arrows as necessary
    If (((m_CurrentAnchor + 1) Mod 3) <> 0) Then aLocations(m_CurrentAnchor + 1) = "arrow_right"
    If ((m_CurrentAnchor Mod 3) <> 0) Then aLocations(m_CurrentAnchor - 1) = "arrow_left"
    
    'Finally, fill in downward arrows as necessary
    If (m_CurrentAnchor < 6) Then
        aLocations(m_CurrentAnchor + 3) = "arrow_down"
        If ((m_CurrentAnchor Mod 3) <> 0) Then aLocations(m_CurrentAnchor + 2) = "arrow_downl"
        If (((m_CurrentAnchor + 1) Mod 3) <> 0) Then aLocations(m_CurrentAnchor + 4) = "arrow_downr"
    End If
    
End Sub

'The user can use an array of command buttons to specify the image's anchor position on the new canvas.  I adopted this
' model from comparable tools in Photoshop and Paint.NET, among others.  Images are loaded from the resource section
' of the EXE and applied to the command buttons as necessary.
Private Sub UpdateAnchorButtons()
    
    Dim i As Long
    
    'Build an array that contains the arrow to appear in each location.
    ReDim m_ArrowLocations(0 To 8) As String
    FillArrowLocations m_ArrowLocations
    
    Dim dibSize As Long
    dibSize = Interface.FixDPI(24)
                
    'Next, extract relevant icons from the resource file, and render them onto the buttons at run-time.
    For i = 0 To 8
    
        If (LenB(m_ArrowLocations(i)) <> 0) Then
            If (StrComp(m_ArrowLocations(i), "generic_image", vbBinaryCompare) = 0) Then
                cmdAnchor(i).AssignImage m_ArrowLocations(i), , dibSize, dibSize
            Else
                
                Dim tmpDIB As pdDIB
                
                If (StrComp(m_ArrowLocations(i), "arrow_up", vbBinaryCompare) = 0) Then
                    Set tmpDIB = Interface.GetRuntimeUIDIB(pdri_ArrowUp, dibSize)
                ElseIf (StrComp(m_ArrowLocations(i), "arrow_upr", vbBinaryCompare) = 0) Then
                    Set tmpDIB = Interface.GetRuntimeUIDIB(pdri_ArrowUpR, dibSize)
                ElseIf (StrComp(m_ArrowLocations(i), "arrow_right", vbBinaryCompare) = 0) Then
                    Set tmpDIB = Interface.GetRuntimeUIDIB(pdri_ArrowRight, dibSize)
                ElseIf (StrComp(m_ArrowLocations(i), "arrow_downr", vbBinaryCompare) = 0) Then
                    Set tmpDIB = Interface.GetRuntimeUIDIB(pdri_ArrowDownR, dibSize)
                ElseIf (StrComp(m_ArrowLocations(i), "arrow_down", vbBinaryCompare) = 0) Then
                    Set tmpDIB = Interface.GetRuntimeUIDIB(pdri_ArrowDown, dibSize)
                ElseIf (StrComp(m_ArrowLocations(i), "arrow_downl", vbBinaryCompare) = 0) Then
                    Set tmpDIB = Interface.GetRuntimeUIDIB(pdri_ArrowDownL, dibSize)
                ElseIf (StrComp(m_ArrowLocations(i), "arrow_left", vbBinaryCompare) = 0) Then
                    Set tmpDIB = Interface.GetRuntimeUIDIB(pdri_ArrowLeft, dibSize)
                ElseIf (StrComp(m_ArrowLocations(i), "arrow_upl", vbBinaryCompare) = 0) Then
                    Set tmpDIB = Interface.GetRuntimeUIDIB(pdri_ArrowUpL, dibSize)
                End If
                
                cmdAnchor(i).AssignImage vbNullString, tmpDIB, dibSize, dibSize
                    
            End If
            
        Else
            cmdAnchor(i).AssignImage vbNullString, Nothing
        End If
        
    Next i
    
End Sub

Private Sub cmdAnchor_Click(Index As Integer)
    m_CurrentAnchor = Index
    UpdateAnchorButtons
End Sub

'The current anchor must be manually saved as part of preset data
Private Sub cmdBar_AddCustomPresetData()
    cmdBar.AddPresetData "currentAnchor", Trim$(Str$(m_CurrentAnchor))
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Canvas size", , GetCurrentParams, UNDO_ImageHeader
End Sub

Private Function GetCurrentParams() As String

    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    With cParams
        .AddParam "width", ucResize.ResizeWidth
        .AddParam "height", ucResize.ResizeHeight
        .AddParam "unit", ucResize.UnitOfMeasurement
        .AddParam "anchor", m_CurrentAnchor
        .AddParam "dpi", ucResize.ResizeDPIAsPPI
    End With
    
    GetCurrentParams = cParams.GetParamString

End Function

'I'm not sure that randomize serves any purpose on this dialog, but as I don't have a way to hide that button (at
' present), simply randomize the width/height to +/- the current image's width/height divided by two.
Private Sub cmdBar_RandomizeClick()
    ucResize.AspectRatioLock = False
    ucResize.ResizeWidthInPixels = (PDImages.GetActiveImage.Width / 2) + (Rnd * PDImages.GetActiveImage.Width)
    ucResize.ResizeHeightInPixels = (PDImages.GetActiveImage.Height / 2) + (Rnd * PDImages.GetActiveImage.Height)
End Sub

'The saved anchor must be custom-loaded, as the command bar won't handle it automatically
Private Sub cmdBar_ReadCustomPresetData()
    ucResize.AspectRatioLock = False
    m_CurrentAnchor = CLng(cmdBar.RetrievePresetData("currentAnchor"))
    UpdateAnchorButtons
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdateAnchorButtons
End Sub

Private Sub cmdBar_ResetClick()

    'Automatically set the width and height text boxes to match the image's current dimensions
    ucResize.UnitOfMeasurement = mu_Pixels
    ucResize.SetInitialDimensions PDImages.GetActiveImage.Width, PDImages.GetActiveImage.Height, PDImages.GetActiveImage.GetDPI
    ucResize.AspectRatioLock = False
    
    'Set the middle position as the anchor
    m_CurrentAnchor = 4

End Sub

'Certain actions are done at LOAD time instead of ACTIVATE time to minimize visible flickering
Private Sub Form_Load()
    
    'Automatically set the width and height text boxes to match the image's current dimensions
    ucResize.SetInitialDimensions PDImages.GetActiveImage.Width, PDImages.GetActiveImage.Height, PDImages.GetActiveImage.GetDPI
    
    'Update the anchor button layout
    UpdateAnchorButtons
    
    Interface.ApplyThemeAndTranslations Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Resize an image using any one of several resampling algorithms.  (Some algorithms are provided by FreeImage.)
Public Sub ResizeCanvas(ByVal functionParams As String)
    
    'TODO: retrieve new measurements as float ( to enable cm/in measurements)
    Dim dstWidthF As Double, dstHeightF As Double
    Dim anchorPosition As Long, curUnit As PD_MeasurementUnit, iDPI As Double
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString functionParams
    
    With cParams
        dstWidthF = .GetDouble("width", PDImages.GetActiveImage.Width)
        dstHeightF = .GetDouble("height", PDImages.GetActiveImage.Height)
        anchorPosition = .GetLong("anchor", 0&)
        curUnit = .GetLong("unit", mu_Pixels)
        iDPI = .GetDouble("dpi", PDImages.GetActiveImage.GetDPI)
    End With
    
    Dim srcWidth As Long, srcHeight As Long
    srcWidth = PDImages.GetActiveImage.Width
    srcHeight = PDImages.GetActiveImage.Height
    
    'In past versions of the software, we could assume the passed measurements were always in pixels,
    ' but that is no longer the case!  Using the supplied "unit of measurement", convert the passed
    ' width and height values to pixel measurements.
    Dim pxWidth As Long, pxHeight As Long
    pxWidth = Units.ConvertOtherUnitToPixels(curUnit, dstWidthF, iDPI, srcWidth)
    pxHeight = Units.ConvertOtherUnitToPixels(curUnit, dstHeightF, iDPI, srcHeight)
    
    'If the image contains an active selection, disable it before transforming the canvas
    If PDImages.GetActiveImage.IsSelectionActive Then
        PDImages.GetActiveImage.SetSelectionActive False
        PDImages.GetActiveImage.MainSelection.LockRelease
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
            dstX = (pxWidth - srcWidth) \ 2
            dstY = 0
        
        'Top-right
        Case 2
            dstX = (pxWidth - srcWidth)
            dstY = 0
        
        'Middle-left
        Case 3
            dstX = 0
            dstY = (pxHeight - srcHeight) \ 2
        
        'Middle-center
        Case 4
            dstX = (pxWidth - srcWidth) \ 2
            dstY = (pxHeight - srcHeight) \ 2
        
        'Middle-right
        Case 5
            dstX = (pxWidth - srcWidth)
            dstY = (pxHeight - srcHeight) \ 2
        
        'Bottom-left
        Case 6
            dstX = 0
            dstY = (pxHeight - srcHeight)
        
        'Bottom-center
        Case 7
            dstX = (pxWidth - srcWidth) \ 2
            dstY = (pxHeight - srcHeight)
        
        'Bottom right
        Case 8
            dstX = (pxWidth - srcWidth)
            dstY = (pxHeight - srcHeight)
    
    End Select
    
    'Now that we have our new top-left corner coordinates (and new width/height values), resizing the canvas
    ' is actually very easy.  In PhotoDemon, there is no such thing as "image data"; an image is just an
    ' imaginary bounding box around the layers collection.  Because of this, we don't actually need to
    ' resize any pixel data - we just need to modify all layer offsets to account for the new top-left corner!
    Dim i As Long
    For i = 0 To PDImages.GetActiveImage.GetNumOfLayers - 1
        With PDImages.GetActiveImage.GetLayerByIndex(i)
            .SetLayerOffsetX .GetLayerOffsetX + dstX
            .SetLayerOffsetY .GetLayerOffsetY + dstY
        End With
    Next i
    
    'Finally, update the parent image's size and DPI values
    PDImages.GetActiveImage.UpdateSize False, pxWidth, pxHeight
    PDImages.GetActiveImage.SetDPI iDPI, iDPI
    Interface.DisplaySize PDImages.GetActiveImage()
    Tools.NotifyImageSizeChanged
    
    'In other functions, we would refresh the layer box here; however, because we haven't actually changed the
    ' appearance of any of the layers, we can leave it as-is!
    
    'Fit the new image on-screen and redraw its viewport
    Viewport.Stage1_InitializeBuffer PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
    Message "Finished."
    
End Sub
