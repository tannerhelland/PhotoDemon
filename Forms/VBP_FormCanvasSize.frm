VERSION 5.00
Begin VB.Form FormCanvasSize 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Resize Canvas"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   6285
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
   ScaleHeight     =   452
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   419
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAnchor 
      Height          =   570
      Index           =   8
      Left            =   1800
      TabIndex        =   22
      Top             =   4080
      Width           =   570
   End
   Begin VB.CommandButton cmdAnchor 
      Height          =   570
      Index           =   7
      Left            =   1200
      TabIndex        =   21
      Top             =   4080
      Width           =   570
   End
   Begin VB.CommandButton cmdAnchor 
      Height          =   570
      Index           =   6
      Left            =   600
      TabIndex        =   20
      Top             =   4080
      Width           =   570
   End
   Begin VB.CommandButton cmdAnchor 
      Height          =   570
      Index           =   5
      Left            =   1800
      TabIndex        =   19
      Top             =   3480
      Width           =   570
   End
   Begin VB.CommandButton cmdAnchor 
      Height          =   570
      Index           =   4
      Left            =   1200
      TabIndex        =   18
      Top             =   3480
      Width           =   570
   End
   Begin VB.CommandButton cmdAnchor 
      Height          =   570
      Index           =   3
      Left            =   600
      TabIndex        =   17
      Top             =   3480
      Width           =   570
   End
   Begin VB.CommandButton cmdAnchor 
      Height          =   570
      Index           =   2
      Left            =   1800
      TabIndex        =   16
      Top             =   2880
      Width           =   570
   End
   Begin VB.CommandButton cmdAnchor 
      Height          =   570
      Index           =   1
      Left            =   1200
      TabIndex        =   15
      Top             =   2880
      Width           =   570
   End
   Begin VB.CommandButton cmdAnchor 
      Height          =   570
      Index           =   0
      Left            =   600
      TabIndex        =   14
      Top             =   2880
      Width           =   570
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   6150
      Width           =   1365
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   4800
      TabIndex        =   1
      Top             =   6150
      Width           =   1365
   End
   Begin PhotoDemon.smartCheckBox chkRatio 
      Height          =   480
      Left            =   4005
      TabIndex        =   4
      Top             =   975
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   847
      Caption         =   "lock aspect ratio"
      Value           =   1
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
   Begin PhotoDemon.textUpDown tudWidth 
      Height          =   405
      Left            =   1440
      TabIndex        =   2
      Top             =   705
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   714
      Max             =   32767
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
   Begin PhotoDemon.textUpDown tudHeight 
      Height          =   405
      Left            =   1440
      TabIndex        =   3
      Top             =   1335
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   714
      Max             =   32767
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
   Begin PhotoDemon.colorSelector colorPicker 
      Height          =   495
      Left            =   600
      TabIndex        =   23
      Top             =   5280
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   873
   End
   Begin VB.Label lblAnchor 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "anchor position:"
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
      Left            =   240
      TabIndex        =   5
      Top             =   2520
      Width           =   1725
   End
   Begin VB.Label lblFill 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "fill empty areas with:"
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
      Left            =   240
      TabIndex        =   13
      Top             =   4920
      Width           =   2235
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "new size:"
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
      Left            =   240
      TabIndex        =   12
      Top             =   240
      Width           =   990
   End
   Begin VB.Label lblAspectRatio 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "new aspect ratio will be"
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
      Left            =   615
      TabIndex        =   11
      Top             =   1950
      Width           =   2490
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   262
      X2              =   262
      Y1              =   57
      Y2              =   105
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   232
      X2              =   262
      Y1              =   56
      Y2              =   56
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   232
      X2              =   262
      Y1              =   105
      Y2              =   105
   End
   Begin VB.Label lblBackground 
      Height          =   855
      Left            =   -2280
      TabIndex        =   10
      Top             =   6000
      Width           =   9975
   End
   Begin VB.Label lblHeightUnit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "pixels"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   2730
      TabIndex        =   9
      Top             =   1365
      Width           =   855
   End
   Begin VB.Label lblWidthUnit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "pixels"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   480
      Left            =   2730
      TabIndex        =   8
      Top             =   735
      Width           =   855
   End
   Begin VB.Label lblHeight 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "height:"
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
      Left            =   600
      TabIndex        =   7
      Top             =   1365
      Width           =   750
   End
   Begin VB.Label lblWidth 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "width:"
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
      Left            =   600
      TabIndex        =   6
      Top             =   735
      Width           =   675
   End
End
Attribute VB_Name = "FormCanvasSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Canvas Size Handler
'Copyright ©2013-2013 by Tanner Helland
'Created: 13/June/13
'Last updated: 13/June/13
'Last update: initial build
'
'This form handles canvas resizing.  You may wonder why it took me over a decade to implement this tool, when it's such a
' trivial one algorithmically.  The answer is that a number of user-interface support functions are necessary to build
' this tool correctly, primarily the command buttons used to select an anchor location.  These require the ability to
' apply 32bpp images to command buttons at run-time, which I lacked for many years.
'
'But now I have such tools at my disposal, so no excuses!  :)  The resulting tool should be self-explanatory.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

'Used to render images onto the tool buttons at run-time
' NOTE: TOOLBOX IMAGES WILL NOT APPEAR IN THE IDE.  YOU MUST COMPILE FIRST.
Private cImgCtl As clsControlImage

'Used for maintaining ratios when the check box is clicked
Private wRatio As Double, hRatio As Double
Dim allowedToUpdateWidth As Boolean, allowedToUpdateHeight As Boolean

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'Current anchor position; used to render the anchor selection command buttons, among other things
Dim m_CurrentAnchor As Long

'We must also track which arrows are drawn where on the command button array
Dim arrowLocations() As String

'If the ratio button is checked, update the height box to reflect the image's current aspect ratio
Private Sub ChkRatio_Click()
    If CBool(chkRatio.Value) Then tudHeight = Int((tudWidth * hRatio) + 0.5)
End Sub

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
    
    'If the buttons already have images, remove them first
    If Not cImgCtl Is Nothing Then
        For i = 0 To 8
            If Len(arrowLocations(i)) > 0 Then cImgCtl.RemoveImage cmdAnchor(i).hWnd
        Next i
        Set cImgCtl = Nothing
    End If
    
    'Build an array that contains the arrow to appear in each location.
    ReDim arrowLocations(0 To 8) As String
    fillArrowLocations arrowLocations
    
    If g_IsVistaOrLater And g_IsThemingEnabled Then
    
        'Next, extract relevant icons from the resource file, and render them onto the buttons at run-time.
        ' (NOTE: because the icons require manifest theming, they will not appear in the IDE.)
        Set cImgCtl = New clsControlImage
        If g_IsProgramCompiled Then
            
            For i = 0 To 8
                If Len(arrowLocations(i)) > 0 Then
                    With cImgCtl
                        .LoadImageFromStream cmdAnchor(i).hWnd, LoadResData(arrowLocations(i), "CUSTOM"), 16, 16
                        .SetMargins cmdAnchor(i).hWnd, 0
                        .Align(cmdAnchor(i).hWnd) = Icon_Center
                    End With
                    cmdAnchor(i).Refresh
                    DoEvents
                End If
            Next i
            
        End If
        
    Else
        For i = 0 To 8
            If arrowLocations(i) = "IMGMEDIUM" Then
                cmdAnchor(i).Caption = "*"
            Else
                cmdAnchor(i).Caption = ""
            End If
        Next i
    End If

End Sub

Private Sub cmdAnchor_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    m_CurrentAnchor = Index
    updateAnchorButtons
End Sub

'Perform a canvas resize operation
Private Sub CmdOK_Click()
    
    'Before resizing anything, check to make sure the textboxes have valid input
    If tudWidth.IsValid And tudHeight.IsValid Then
        
        Me.Visible = False
        Process "Canvas size", , buildParams(tudWidth, tudHeight, m_CurrentAnchor, colorPicker.Color)
        Unload Me
        
    End If
    
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

'Upon form activation, determine the ratio between the width and height of the image
Private Sub Form_Activate()
    
    'To prevent aspect ratio changes to one box resulting in recursion-type changes to the other, we only
    ' allow one box at a time to be updated.
    allowedToUpdateWidth = True
    allowedToUpdateHeight = True
    
    'Establish ratios
    wRatio = pdImages(CurrentImage).Width / pdImages(CurrentImage).Height
    hRatio = pdImages(CurrentImage).Height / pdImages(CurrentImage).Width
    
    'Automatically set the width and height text boxes to match the image's current dimensions
    tudWidth.Value = pdImages(CurrentImage).Width
    tudHeight.Value = pdImages(CurrentImage).Height
    
    'If the source image is 32bpp, hide the color selection box and change the text to match
    If pdImages(CurrentImage).mainLayer.getLayerColorDepth = 32 Then
        lblFill.Caption = g_Language.TranslateMessage("note: empty areas will be made transparent")
    Else
        lblFill.Caption = g_Language.TranslateMessage("fill empty areas with:")
    End If
    
    'Start with a default top-left position for the anchor
    m_CurrentAnchor = 0
    updateAnchorButtons
    
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
        
End Sub

'Certain actions are done at LOAD time instead of ACTIVATE time to minimize visible flickering
Private Sub Form_Load()

    'If the current image is 32bpp, we have no need to display the "background color" selection box, as any blank space
    ' will be filled with transparency.
    If pdImages(CurrentImage).mainLayer.getLayerColorDepth = 32 Then
    
        'Hide the background color selectors
        colorPicker.Visible = False
        
        Dim formHeightDifference As Long
        Me.ScaleMode = vbTwips
        formHeightDifference = Me.Height - Me.ScaleHeight
        Me.ScaleMode = vbPixels
    
        'Move the command bar into place
        lblBackground.Top = lblFill.Top + lblFill.Height + 16
        CmdOK.Top = lblBackground.Top + 10
        CmdCancel.Top = CmdOK.Top
        
        'Resize the form to match
        Me.Height = formHeightDifference + (CmdOK.Top + CmdOK.Height + 10) * Screen.TwipsPerPixelY
        
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Resize an image using any one of several resampling algorithms.  (Some algorithms are provided by FreeImage.)
Public Sub ResizeCanvas(ByVal iWidth As Long, ByVal iHeight As Long, ByVal anchorPosition As Long, Optional ByVal newBackColor As Long = vbWhite)

    Dim srcWidth As Long, srcHeight As Long
    srcWidth = pdImages(CurrentImage).Width
    srcHeight = pdImages(CurrentImage).Height
    
    'If the image contains an active selection, disable it before transforming the canvas
    If pdImages(CurrentImage).selectionActive Then
        pdImages(CurrentImage).selectionActive = False
        pdImages(CurrentImage).mainSelection.lockRelease
        metaToggle tSelection, False
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
    
    'Create a temporary layer to hold the new canvas
    Dim tmpLayer As pdLayer
    Set tmpLayer = New pdLayer
    tmpLayer.createBlank iWidth, iHeight, pdImages(CurrentImage).mainLayer.getLayerColorDepth, newBackColor

    'Bitblt the old image into its new position on the canvas
    BitBlt tmpLayer.getLayerDC, dstX, dstY, srcWidth, srcHeight, pdImages(CurrentImage).mainLayer.getLayerDC, 0, 0, vbSrcCopy
    
    'The temporary layer now holds the new canvas and image.  Copy it back into the main image.
    pdImages(CurrentImage).mainLayer.createFromExistingLayer tmpLayer
    Set tmpLayer = Nothing
    
    'Update the main image's size values
    pdImages(CurrentImage).updateSize
    DisplaySize pdImages(CurrentImage).Width, pdImages(CurrentImage).Height
    
    'Fit the new image on-screen and redraw its viewport
    PrepareViewport FormMain.ActiveForm, "Canvas resize"
    
    Message "Finished."
    
End Sub

'PhotoDemon now displays an approximate aspect ratio for the selected values.  This can be helpful when
' trying to select new width/height values for a specific application with a set aspect ratio (e.g. 16:9 screens).
Private Sub updateAspectRatio()

    Dim wholeNumber As Double, Numerator As Double, Denominator As Double
    
    If tudWidth.IsValid And tudHeight.IsValid Then
        convertToFraction tudWidth / tudHeight, wholeNumber, Numerator, Denominator, 4, 99.9
        
        'Aspect ratios are typically given in terms of base 10 if possible, so change values like 8:5 to 16:10
        If CLng(Denominator) = 5 Then
            Numerator = Numerator * 2
            Denominator = Denominator * 2
        End If
        
        lblAspectRatio.Caption = g_Language.TranslateMessage("new aspect ratio will be %1:%2", Numerator, Denominator)
    End If

End Sub

'If "Lock Image Aspect Ratio" is selected, these two routines keep all values in sync
Private Sub tudHeight_Change()
    If CBool(chkRatio) And allowedToUpdateWidth Then
        allowedToUpdateHeight = False
        tudWidth = Int((tudHeight * wRatio) + 0.5)
        allowedToUpdateHeight = True
    End If
    
    updateAspectRatio
    
End Sub

Private Sub tudWidth_Change()
    If CBool(chkRatio) And allowedToUpdateHeight Then
        allowedToUpdateWidth = False
        tudHeight = Int((tudWidth * hRatio) + 0.5)
        allowedToUpdateWidth = True
    End If
    
    updateAspectRatio
    
End Sub
