VERSION 5.00
Begin VB.Form FormResize 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Resize Image"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   10665
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
   ScaleHeight     =   354
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   711
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   4560
      Width           =   10665
      _ExtentX        =   18812
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
      AutoloadLastPreset=   -1  'True
   End
   Begin PhotoDemon.jcbutton cmdCategory 
      Height          =   780
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   1376
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      Caption         =   "Basic options "
      ForeColor       =   4210752
      ForeColorHover  =   4194304
      Mode            =   1
      Value           =   -1  'True
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_FormResize.frx":0000
      PictureAlign    =   0
      PictureEffectOnDown=   0
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipTitle    =   "Interface Options"
      ColorScheme     =   3
   End
   Begin PhotoDemon.jcbutton cmdCategory 
      Height          =   780
      Index           =   1
      Left            =   120
      TabIndex        =   16
      Top             =   960
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   1376
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      Caption         =   "Advanced options "
      ForeColor       =   4210752
      ForeColorHover  =   4194304
      Mode            =   1
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_FormResize.frx":1452
      PictureAlign    =   0
      PictureEffectOnDown=   0
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipTitle    =   "Load (Import) Options"
      ColorScheme     =   3
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4215
      Index           =   1
      Left            =   3360
      ScaleHeight     =   281
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   480
      TabIndex        =   6
      Top             =   120
      Width           =   7200
      Begin PhotoDemon.smartOptionButton optFit 
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   7
         Top             =   600
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   661
         Caption         =   "stretching"
         Value           =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin PhotoDemon.smartOptionButton optFit 
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   8
         Top             =   1320
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   661
         Caption         =   "fit inclusively"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin PhotoDemon.smartOptionButton optFit 
         Height          =   375
         Index           =   2
         Left            =   480
         TabIndex        =   9
         Top             =   2640
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   661
         Caption         =   "fit exclusively"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin PhotoDemon.colorSelector colorPicker 
         Height          =   495
         Left            =   960
         TabIndex        =   10
         Top             =   2040
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   873
      End
      Begin VB.Label lblFit 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "fit image to new size by:"
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
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   2610
      End
      Begin VB.Label lblSubtext 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "no distortion; empty borders will be filled with:"
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
         Height          =   240
         Index           =   1
         Left            =   960
         TabIndex        =   13
         Top             =   1710
         Width           =   4020
      End
      Begin VB.Label lblSubtext 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "no distortion; image edges will be cropped to fit"
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
         Height          =   240
         Index           =   2
         Left            =   960
         TabIndex        =   12
         Top             =   3030
         Width           =   4110
      End
      Begin VB.Label lblSubtext 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "maintain aspect ratio or the final image may look distorted"
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
         Height          =   240
         Index           =   0
         Left            =   960
         TabIndex        =   11
         Top             =   990
         Width           =   5010
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4215
      Index           =   0
      Left            =   3360
      ScaleHeight     =   281
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   120
      Width           =   7200
      Begin PhotoDemon.smartResize ucResize 
         Height          =   1695
         Left            =   480
         TabIndex        =   17
         Top             =   600
         Width           =   7695
         _ExtentX        =   11880
         _ExtentY        =   2990
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
      Begin VB.ComboBox cboResample 
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
         Left            =   480
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   3000
         Width           =   5895
      End
      Begin PhotoDemon.smartCheckBox chkNames 
         Height          =   480
         Left            =   480
         TabIndex        =   2
         Top             =   3480
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   847
         Caption         =   "show technical names"
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
      Begin VB.Label lblResample 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "resize quality:"
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
         Left            =   120
         TabIndex        =   5
         Top             =   2520
         Width           =   1470
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
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   990
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      X1              =   216
      X2              =   216
      Y1              =   296
      Y2              =   8
   End
End
Attribute VB_Name = "FormResize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Image Size Handler
'Copyright ©2001-2014 by Tanner Helland
'Created: 6/12/01
'Last updated: 31/January/14
'Last update: redesign dialog to account for new "SmartResize" user control.  Also, separate basic and
'              advanced options into separate panels.
'
'Handles all image-size related functions.  Currently supports nearest-neighbor and halftone resampling
' (via the API; not 100% accurate but faster than doing it manually), bilinear resampling via pure VB, and
' a number of more advanced resampling techniques via FreeImage.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'The list of available resampling algorithms changes based on the presence of FreeImage, and the use of
' "friendly" vs "technical" names.  As a result, we have to track them dynamically using a custom type.
Private Type resampleAlgorithm
    Name As String
    ProgramID As Long
End Type

Dim resampleTypes() As resampleAlgorithm
Dim numResamples As Long
Dim lastSelectedResample As Long

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

Private Sub addResample(ByVal rName As String, ByVal rID As Long)
    resampleTypes(numResamples).Name = rName
    resampleTypes(numResamples).ProgramID = rID
    numResamples = numResamples + 1
    ReDim Preserve resampleTypes(0 To numResamples) As resampleAlgorithm
End Sub

'Display all available resample algorithms in the combo box (contingent on the "show technical names" check box as well)
Private Sub refillResampleBox(Optional ByVal isFirstTime As Boolean = False)

    ReDim resampleTypes(0) As resampleAlgorithm
    numResamples = 0

    'Use friendly names
    If Not CBool(chkNames) Then
        
        lblResample.Caption = g_Language.TranslateMessage("resampling method:")
        
        'FreeImage is required for best output.  Without it, only a small number of resample algorithms are implemented.
        If g_ImageFormats.FreeImageEnabled Then
            addResample g_Language.TranslateMessage("best for photographs"), RESIZE_LANCZOS
            addResample g_Language.TranslateMessage("best for text and illustrations"), RESIZE_BICUBIC_MITCHELL
            addResample g_Language.TranslateMessage("fastest"), RESIZE_NORMAL
        Else
            addResample g_Language.TranslateMessage("best for photographs"), RESIZE_BILINEAR
            addResample g_Language.TranslateMessage("best for text and illustrations"), RESIZE_HALFTONE
            addResample g_Language.TranslateMessage("fastest"), RESIZE_NORMAL
        End If
    
    'Use technical names
    Else
        
        lblResample.Caption = g_Language.TranslateMessage("resampling algorithm:")
        
        'Prepare a list of available resample algorithms
        addResample g_Language.TranslateMessage("Nearest Neighbor"), RESIZE_NORMAL
        addResample g_Language.TranslateMessage("Halftone"), RESIZE_HALFTONE
        addResample g_Language.TranslateMessage("Bilinear"), RESIZE_BILINEAR
        
        'If the FreeImage library is available, add additional resize options to the combo box
        If g_ImageFormats.FreeImageEnabled Then
            addResample g_Language.TranslateMessage("B-Spline"), RESIZE_BSPLINE
            addResample g_Language.TranslateMessage("Bicubic (Mitchell and Netravali)"), RESIZE_BICUBIC_MITCHELL
            addResample g_Language.TranslateMessage("Bicubic (Catmull-Rom)"), RESIZE_BICUBIC_CATMULL
            addResample g_Language.TranslateMessage("Sinc (Lanczos 3-lobe)"), RESIZE_LANCZOS
        End If
                
    End If
    
    'Populate the combo box
    cboResample.Clear
    Dim i As Long
    For i = 0 To numResamples - 1
        cboResample.AddItem " " & resampleTypes(i).Name, i
    Next i
    
    'If this is the first time we are filling the combo box, provide an intelligent default setting
    If isFirstTime Then
    
        'Friendly names
        If Not CBool(chkNames) Then
            cboResample.ListIndex = 0
            
        'Technical names
        Else
        
            'FreeImage enabled
            If g_ImageFormats.FreeImageEnabled Then
                cboResample.ListIndex = 5
                
            'FreeImage not enabled
            Else
                cboResample.ListIndex = 2
            End If
        
        End If
    
    'If this is not the first time we are creating a list of resample methods, re-select whatever method the
    ' user had previously selected (if available; otherwise, redirect them to the best general-purpose algorithm)
    Else
    
        Dim targetResampleMethod As Long
        targetResampleMethod = lastSelectedResample
        
        'Some technical options are not available under friendly names, so redirect them to something similar
        If CBool(chkNames) And g_ImageFormats.FreeImageEnabled Then
            Select Case lastSelectedResample
                Case 1 To 3
                    targetResampleMethod = RESIZE_BICUBIC_CATMULL
            End Select
        End If
        
        'Find the matching resample method in the new combo box
        For i = 0 To cboResample.ListCount - 1
            If resampleTypes(i).ProgramID = targetResampleMethod Then
                cboResample.ListIndex = i
                Exit For
            End If
        Next i
    
    End If

End Sub

Private Sub cboResample_Click()
    lastSelectedResample = resampleTypes(cboResample.ListIndex).ProgramID
End Sub

'New to v6.0, PhotoDemon gives the user friendly resample names by default.  They can toggle these off at their liking.
Private Sub chkNames_Click()
    refillResampleBox
End Sub

'OK button
Private Sub cmdBar_OKClick()

    'Find the user's requested "how to fit" method; we pass this along to the master Resize function, even if it's
    ' not being used.
    Dim fitMethod As Long
    Dim i As Long
    For i = 0 To optFit.Count - 1
        If optFit(i).Value Then fitMethod = i
    Next i
        
    Process "Resize", , buildParams(ucResize.imgWidth, ucResize.imgHeight, resampleTypes(cboResample.ListIndex).ProgramID, fitMethod, colorPicker.Color)

End Sub

'I'm not sure that randomize serves any purpose on this dialog, but as I don't have a way to hide that button at
' present, simply randomize the width/height to +/- the current image's width/height divided by two.
Private Sub cmdBar_RandomizeClick()

    ucResize.lockAspectRatio = False
    ucResize.imgWidth = (pdImages(g_CurrentImage).Width / 2) + (Rnd * pdImages(g_CurrentImage).Width)
    ucResize.imgHeight = (pdImages(g_CurrentImage).Height / 2) + (Rnd * pdImages(g_CurrentImage).Height)

End Sub

Private Sub cmdBar_ResetClick()
    
    'Automatically set the width and height text boxes to match the image's current dimensions
    ucResize.setInitialDimensions pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height
    ucResize.lockAspectRatio = True
    
    'Make borders fill with black by default
    colorPicker.Color = RGB(0, 0, 0)
    
    'Stretch to new aspect ratio by default
    optFit(0).Value = True
    
    'Use friendly resample names by default
    chkNames.Value = vbUnchecked
    cboResample.ListIndex = 0
    
End Sub

Private Sub cmdCategory_Click(Index As Integer)
    
    Dim i As Long
    
    For i = 0 To cmdCategory.Count - 1
        If i = Index Then
            cmdCategory(i).Value = True
            picContainer(i).Visible = True
        Else
            cmdCategory(i).Value = False
            picContainer(i).Visible = False
        End If
    Next i

End Sub

Private Sub colorPicker_ColorChanged()
    optFit(1).Value = True
End Sub

'Upon form activation, determine the ratio between the width and height of the image
Private Sub Form_Activate()
    
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
    'By default, the basic options panel is always shown.
    Dim i As Long
    For i = 0 To cmdCategory.Count - 1
        If i = 0 Then
            cmdCategory(i).Value = True
            picContainer(i).Visible = True
        Else
            cmdCategory(i).Value = False
            picContainer(i).Visible = False
        End If
    Next i
    
    'Automatically set the width and height text boxes to match the image's current dimensions
    ucResize.setInitialDimensions pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height
    
    'TODO - TEMPORARY WARNING!!
    ' Let the user know that this form is being actively worked on
    'lblTemporaryWarning = "WARNING! This dialog is being completely reworked for PhotoDemon 6.2.  As long as this warning remains, resizing may not work as expected.  I hope to have all changes finalized by the end of February 2014, but until then, avoid using this developer release for serious resize tasks."
            
End Sub

'Certain actions are done at LOAD time instead of ACTIVATE time to minimize visible flickering
Private Sub Form_Load()

    'If the current image is 32bpp, we have no need to display the "background color" selection box, as any blank space
    ' will be filled with transparency.
    If pdImages(g_CurrentImage).mainDIB.getDIBColorDepth = 32 Then
    
        'Hide the background color selectors
        colorPicker.Visible = False
        
        'Move up the controls beneath it
        optFit(2).Top = optFit(1).Top + fixDPI(48)
        lblSubtext(2).Top = optFit(2).Top + fixDPI(26)
        
    End If
    
    'Populate the number of available resampling algorithms
    refillResampleBox True
    
    'If the source image is 32bpp, change the text of the "fit inclusive" subheading to match
    If pdImages(g_CurrentImage).mainDIB.getDIBColorDepth = 32 Then
        lblSubtext(1).Caption = g_Language.TranslateMessage("no distortion; empty borders will be transparent")
    Else
        lblSubtext(1).Caption = g_Language.TranslateMessage("no distortion; empty borders will be filled with:")
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Resize an image using the FreeImage library.  Very fast.
Private Sub FreeImageResize(ByRef dstDIB As pdDIB, ByRef srcDIB As pdDIB, ByVal iWidth As Long, iHeight As Long, ByVal interpolationMethod As Long)
    
    'Double-check that FreeImage exists
    If g_ImageFormats.FreeImageEnabled Then
        
        Message "Resampling image using the FreeImage plugin..."
        
        'If the original image is 32bpp, remove premultiplication now
        If srcDIB.getDIBColorDepth = 32 Then srcDIB.fixPremultipliedAlpha
        
        'Convert the current image to a FreeImage-type DIB
        Dim fi_DIB As Long
        fi_DIB = FreeImage_CreateFromDC(srcDIB.getDIBDC)
        
        'Use that handle to request an image resize
        If fi_DIB <> 0 Then
            
            Dim returnDIB As Long
            returnDIB = FreeImage_RescaleByPixel(fi_DIB, iWidth, iHeight, True, interpolationMethod)
            
            'Resize the destination DIB in preparation for the transfer
            dstDIB.createBlank iWidth, iHeight, srcDIB.getDIBColorDepth
            
            'Copy the bits from the FreeImage DIB to our DIB
            SetDIBitsToDevice dstDIB.getDIBDC, 0, 0, iWidth, iHeight, 0, 0, 0, iHeight, ByVal FreeImage_GetBits(returnDIB), ByVal FreeImage_GetInfo(returnDIB), 0&
     
            'With the transfer complete, release the FreeImage DIB and unload the library
            If returnDIB <> 0 Then FreeImage_UnloadEx returnDIB
            
        End If
        
        'If the original image is 32bpp, add back in premultiplication now
        If srcDIB.getDIBColorDepth = 32 Then dstDIB.fixPremultipliedAlpha True
        
    End If
    
End Sub

'Resize an image using any one of several resampling algorithms.  (Some algorithms are provided by FreeImage.)
Public Sub ResizeImage(ByVal iWidth As Long, ByVal iHeight As Long, ByVal resampleMethod As Byte, ByVal fitMethod As Long, Optional ByVal newBackColor As Long = vbWhite)

    'Depending on the requested fitting technique, we may have to resize the image to a slightly different size
    ' than the one requested.  Before doing anything else, calculate that new size.
    Dim fitWidth As Long, fitHeight As Long
    
    Dim srcWidth As Long, srcHeight As Long
    srcWidth = pdImages(g_CurrentImage).Width
    srcHeight = pdImages(g_CurrentImage).Height
    
    Select Case fitMethod
    
        'Stretch-to-fit.  Default behavior, and no size changes are required.
        Case 0
            fitWidth = iWidth
            fitHeight = iHeight
        
        'Fit inclusively.  Fit the image's largest dimension.  No cropping will occur, but blank space may be present.
        Case 1
            
            'We have an existing function for this purpose.  (It's used when rendering preview images, for example.)
            convertAspectRatio srcWidth, srcHeight, iWidth, iHeight, fitWidth, fitHeight
            
        'Fit exclusively.  Fit the image's smallest dimension.  Cropping will occur, but no blank space will be present.
        Case 2
        
            convertAspectRatio srcWidth, srcHeight, iWidth, iHeight, fitWidth, fitHeight, False
        
    End Select
    
    'If the image contains an active selection, automatically deactivate it
    If pdImages(g_CurrentImage).selectionActive Then
        pdImages(g_CurrentImage).selectionActive = False
        pdImages(g_CurrentImage).mainSelection.lockRelease
    End If

    'Because most resize methods require a temporary DIB, create one here
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB

    Select Case resampleMethod

        'Nearest neighbor...
        Case RESIZE_NORMAL
        
            Message "Resizing image..."
            
            'Copy the current DIB into this temporary DIB at the new size
            tmpDIB.createFromExistingDIB pdImages(g_CurrentImage).mainDIB, fitWidth, fitHeight, False
            
        'Halftone resampling... I'm not sure what to actually call it, but since it's based off the
        ' StretchBlt mode Microsoft calls "halftone," I'm sticking with that
        Case RESIZE_HALFTONE
            
            Message "Resizing image..."
            
            'Copy the current DIB into this temporary DIB at the new size
            tmpDIB.createFromExistingDIB pdImages(g_CurrentImage).mainDIB, fitWidth, fitHeight, True
        
        'True bilinear sampling
        Case RESIZE_BILINEAR
        
            'If FreeImage is enabled, use their bilinear filter.  Similar results, much faster.
            If g_ImageFormats.FreeImageEnabled Then
            
                FreeImageResize tmpDIB, pdImages(g_CurrentImage).mainDIB, fitWidth, fitHeight, FILTER_BILINEAR
            
            'If FreeImage is not enabled, we have to do the resample ourselves.
            Else
            
                Message "Resampling image..."
        
                'Create a local array and point it at the pixel data of the current image
                Dim srcImageData() As Byte
                Dim srcSA As SAFEARRAY2D
                prepImageData srcSA
                CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
                'Resize the temporary DIB to the target size, and point a second local array at it
                tmpDIB.createBlank fitWidth, fitHeight, pdImages(g_CurrentImage).mainDIB.getDIBColorDepth
                
                Dim dstImageData() As Byte
                Dim dstSA As SAFEARRAY2D
                
                prepSafeArray dstSA, tmpDIB
                CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
                
                'These values will help us access locations in the array more quickly.
                ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
                Dim qvDepth As Long
                qvDepth = tmpDIB.getDIBColorDepth \ 8
                
                'Create a filter support class, which will aid with edge handling and interpolation
                Dim fSupport As pdFilterSupport
                Set fSupport = New pdFilterSupport
                fSupport.setDistortParameters qvDepth, EDGE_CLAMP, True, curDIBValues.maxX, curDIBValues.MaxY
    
                'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
                ' based on the size of the area to be processed.
                Dim progBarCheck As Long
                SetProgBarMax iWidth
                progBarCheck = findBestProgBarValue()
            
                'Resampling requires many variables
                
                'Scaled ratios between the old x and y values and the new ones
                Dim xScale As Double, yScale As Double
                xScale = (srcWidth - 1) / fitWidth
                yScale = (srcHeight - 1) / fitHeight
                            
                'Coordinate variables for source and destination
                Dim x As Long, y As Long
                Dim srcX As Double, srcY As Double
                            
                For x = 0 To fitWidth - 1
                    
                    'Generate the x calculation variables
                    srcX = x * xScale
                    
                    'Draw each pixel in the new image
                    For y = 0 To fitHeight - 1
                        
                        'Generate the y calculation variables
                        srcY = y * yScale
                        
                        'The lovely .setPixels routine will handle edge detection and interpolation for us as necessary
                        fSupport.setPixels x, y, srcX, srcY, srcImageData, dstImageData
                                            
                    Next y
                
                    If (x And progBarCheck) = 0 Then SetProgBarVal x
                    
                Next x
                            
                'With our work complete, point both ImageData() arrays away from their DIBs and deallocate them
                CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
                Erase srcImageData
                
                CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
                Erase dstImageData
                
                SetProgBarVal 0
                
            End If
        
        Case RESIZE_BSPLINE
            FreeImageResize tmpDIB, pdImages(g_CurrentImage).mainDIB, fitWidth, fitHeight, FILTER_BSPLINE
            
        Case RESIZE_BICUBIC_MITCHELL
            FreeImageResize tmpDIB, pdImages(g_CurrentImage).mainDIB, fitWidth, fitHeight, FILTER_BICUBIC
            
        Case RESIZE_BICUBIC_CATMULL
            FreeImageResize tmpDIB, pdImages(g_CurrentImage).mainDIB, fitWidth, fitHeight, FILTER_CATMULLROM
        
        Case RESIZE_LANCZOS
            FreeImageResize tmpDIB, pdImages(g_CurrentImage).mainDIB, fitWidth, fitHeight, FILTER_LANCZOS3
            
    End Select
    
    'The temporary DIB now holds a copy of the resized image.
    
    'Calculate the aspect ratio of this DIB and the target picture box
    Dim srcAspect As Double, dstAspect As Double
    srcAspect = pdImages(g_CurrentImage).Width / pdImages(g_CurrentImage).Height
    dstAspect = iWidth / iHeight
    
    Dim dstX As Long, dstY As Long
    
    'We now want to copy the resized image into the current image using the technique requested by the user.
    Select Case fitMethod
    
        'Stretch-to-fit.  This is default resize behavior in all image editing software
        Case 0
    
            'Very simple - just copy the resized image back into the main DIB
            pdImages(g_CurrentImage).mainDIB.createFromExistingDIB tmpDIB
    
        'Fit inclusively.  This fits the image's largest dimension into the destination image, which can leave
        ' blank space - that space is filled by the background color parameter passed in.
        Case 1
        
            'Resize the main DIB (destructively!) to fit the new dimensions
            pdImages(g_CurrentImage).mainDIB.createBlank iWidth, iHeight, pdImages(g_CurrentImage).mainDIB.getDIBColorDepth, newBackColor
        
            'BitBlt the old image, centered, onto the new DIB
            If srcAspect > dstAspect Then
                dstY = CLng((iHeight - fitHeight) / 2)
                dstX = 0
            Else
                dstX = CLng((iWidth - fitWidth) / 2)
                dstY = 0
            End If
            
            BitBlt pdImages(g_CurrentImage).mainDIB.getDIBDC, dstX, dstY, fitWidth, fitHeight, tmpDIB.getDIBDC, 0, 0, vbSrcCopy
        
        'Fit exclusively.  This fits the image's smallest dimension into the destination image, which means no
        ' blank space - but parts of the image may get cropped out.
        Case 2
        
            'Resize the main DIB (destructively!) to fit the new dimensions
            pdImages(g_CurrentImage).mainDIB.createBlank iWidth, iHeight, pdImages(g_CurrentImage).mainDIB.getDIBColorDepth, newBackColor
        
            'BitBlt the old image, centered, onto the new DIB
            If srcAspect < dstAspect Then
                dstY = CLng((iHeight - fitHeight) / 2)
                dstX = 0
            Else
                dstX = CLng((iWidth - fitWidth) / 2)
                dstY = 0
            End If
            
            BitBlt pdImages(g_CurrentImage).mainDIB.getDIBDC, dstX, dstY, fitWidth, fitHeight, tmpDIB.getDIBDC, 0, 0, vbSrcCopy
        
    End Select
    
    'We are finished with the temporary DIB, so release it
    Set tmpDIB = Nothing
    
    'Update the main image's size values
    pdImages(g_CurrentImage).updateSize
    DisplaySize pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height
    
    'Fit the new image on-screen and redraw its viewport
    PrepareViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0), "Image resize"
    
    Message "Finished."
    
End Sub
