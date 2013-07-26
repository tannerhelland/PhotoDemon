VERSION 5.00
Begin VB.Form FormResize 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Resize Image"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   7020
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
   ScaleHeight     =   547
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   468
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picBackColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1080
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   359
      TabIndex        =   20
      Top             =   6090
      Width           =   5415
   End
   Begin PhotoDemon.smartOptionButton optFit 
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   16
      Top             =   4680
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
   Begin PhotoDemon.smartCheckBox chkNames 
      Height          =   480
      Left            =   600
      TabIndex        =   12
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
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   4080
      TabIndex        =   0
      Top             =   7590
      Width           =   1365
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   5520
      TabIndex        =   1
      Top             =   7590
      Width           =   1365
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
      Left            =   600
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3000
      Width           =   5895
   End
   Begin PhotoDemon.smartCheckBox chkRatio 
      Height          =   480
      Left            =   4200
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
      Min             =   1
      Max             =   32767
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
   Begin PhotoDemon.textUpDown tudHeight 
      Height          =   405
      Left            =   1440
      TabIndex        =   3
      Top             =   1335
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   714
      Min             =   1
      Max             =   32767
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
   Begin PhotoDemon.smartOptionButton optFit 
      Height          =   375
      Index           =   1
      Left            =   600
      TabIndex        =   17
      Top             =   5400
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
      Left            =   600
      TabIndex        =   18
      Top             =   6720
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
   Begin VB.Label lblSubtext 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "the final image may look distorted"
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
      Left            =   1080
      TabIndex        =   22
      Top             =   5070
      Width           =   2910
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
      Left            =   1080
      TabIndex        =   21
      Top             =   7110
      Width           =   4110
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
      Left            =   1080
      TabIndex        =   19
      Top             =   5790
      Width           =   4020
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
      Left            =   240
      TabIndex        =   15
      Top             =   4200
      Width           =   2610
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
      TabIndex        =   14
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
      TabIndex        =   13
      Top             =   1950
      Width           =   2490
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   272
      X2              =   272
      Y1              =   57
      Y2              =   106
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   272
      X2              =   232
      Y1              =   56
      Y2              =   56
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   232
      X2              =   272
      Y1              =   105
      Y2              =   105
   End
   Begin VB.Label lblBackground 
      Height          =   855
      Left            =   -2280
      TabIndex        =   11
      Top             =   7440
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
      TabIndex        =   10
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
      TabIndex        =   9
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
      TabIndex        =   8
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
      TabIndex        =   7
      Top             =   735
      Width           =   675
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
      Left            =   240
      TabIndex        =   6
      Top             =   2520
      Width           =   1470
   End
End
Attribute VB_Name = "FormResize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Image Size Handler
'Copyright ©2001-2013 by Tanner Helland
'Created: 6/12/01
'Last updated: 13/June/13
'Last update: redesigned the form to incorporate features originally meant for "Smart Resize"
'
'Handles all image-size related functions.  Currently supports nearest-neighbor and halftone resampling
' (via the API; not 100% accurate but faster than doing it manually), bilinear resampling via pure VB, and
' a number of more advanced resampling techniques via FreeImage.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
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

'Used for maintaining ratios when the check box is clicked
Private wRatio As Double, hRatio As Double

Dim allowedToUpdateWidth As Boolean, allowedToUpdateHeight As Boolean

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'If the user unselects the "lock aspect ratio" box, we need to recenter the dialog over the main window.
' Only do this once to avoid the dialog jumping all over the place.
Dim dialogNeedsCentering As Boolean

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

'New to v5.6, PhotoDemon gives the user friendly resample names by default.  They can toggle these off at their liking.
Private Sub chkNames_Click()
    refillResampleBox
End Sub

'If the ratio button is checked, update the height box to reflect the image's current aspect ratio
Private Sub ChkRatio_Click()
    If CBool(chkRatio.Value) Then tudHeight = Int((tudWidth * hRatio) + 0.5)
    updateFormLayout
    
    If dialogNeedsCentering Then
        dialogNeedsCentering = False
        Me.Top = FormMain.Top + ((FormMain.Height - Me.Height) \ 2)
    End If
End Sub

'Perform a resize operation
Private Sub CmdOK_Click()
    
    'Before resizing anything, check to make sure the textboxes have valid input
    If tudWidth.IsValid And tudHeight.IsValid Then
        
        Me.Visible = False
        
        Dim fitMethod As Long
        Dim i As Long
        For i = 0 To optFit.Count - 1
            If optFit(i).Value Then fitMethod = i
        Next i
            
        Process "Resize", , buildParams(tudWidth, tudHeight, resampleTypes(cboResample.ListIndex).ProgramID, fitMethod, picBackColor.backColor)
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

    'Populate the number of available resampling algorithms
    refillResampleBox True
    
    'Automatically set the width and height text boxes to match the image's current dimensions
    tudWidth.Value = pdImages(CurrentImage).Width
    tudHeight.Value = pdImages(CurrentImage).Height
    
    'If the source image is 32bpp, change the text of the "fit inclusive" subheading to match
    If pdImages(CurrentImage).mainLayer.getLayerColorDepth = 32 Then
        lblSubtext(1).Caption = g_Language.TranslateMessage("no distortion; empty borders will be transparent")
    Else
        lblSubtext(1).Caption = g_Language.TranslateMessage("no distortion; empty borders will be filled with:")
    End If
    
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    setHandCursor picBackColor
        
    'If the user unchecks the "lock aspect ratio" button, we will recenter the dialog once (as it's quite tall)
    dialogNeedsCentering = True
        
End Sub

'Certain actions are done at LOAD time instead of ACTIVATE time to minimize visible flickering
Private Sub Form_Load()

    'If the current image is 32bpp, we have no need to display the "background color" selection box, as any blank space
    ' will be filled with transparency.
    If pdImages(CurrentImage).mainLayer.getLayerColorDepth = 32 Then
    
        'Hide the background color selectors
        picBackColor.Visible = False
        
        'Move up the controls beneath it
        optFit(2).Top = optFit(1).Top + 48
        lblSubtext(2).Top = optFit(2).Top + 26
        
    End If
    
    updateFormLayout
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Resize an image using the FreeImage library.  Very fast.
Private Sub FreeImageResize(ByRef dstLayer As pdLayer, ByRef srcLayer As pdLayer, ByVal iWidth As Long, iHeight As Long, ByVal interpolationMethod As Long)
    
    'Double-check that FreeImage exists
    If g_ImageFormats.FreeImageEnabled Then
    
        'Load the FreeImage dll into memory
        Dim hLib As Long
        hLib = LoadLibrary(g_PluginPath & "FreeImage.dll")
        
        Message "Resampling image using the FreeImage plugin..."
        
        'Convert the current image to a FreeImage-type DIB
        Dim fi_DIB As Long
        fi_DIB = FreeImage_CreateFromDC(srcLayer.getLayerDC)
        
        'Use that handle to request an image resize
        If fi_DIB <> 0 Then
            
            Dim returnDIB As Long
            returnDIB = FreeImage_RescaleByPixel(fi_DIB, iWidth, iHeight, True, interpolationMethod)
            
            'Resize the destination layer in preparation for the transfer
            dstLayer.createBlank iWidth, iHeight, srcLayer.getLayerColorDepth
            
            'Copy the bits from the FreeImage DIB to our DIB
            SetDIBitsToDevice dstLayer.getLayerDC, 0, 0, iWidth, iHeight, 0, 0, 0, iHeight, ByVal FreeImage_GetBits(returnDIB), ByVal FreeImage_GetInfo(returnDIB), 0&
     
            'With the transfer complete, release the FreeImage DIB and unload the library
            If returnDIB <> 0 Then FreeImage_UnloadEx returnDIB
            FreeLibrary hLib
            
        Else
            FreeLibrary hLib
        End If
        
    End If
    
End Sub

'Resize an image using any one of several resampling algorithms.  (Some algorithms are provided by FreeImage.)
Public Sub ResizeImage(ByVal iWidth As Long, ByVal iHeight As Long, ByVal resampleMethod As Byte, ByVal fitMethod As Long, Optional ByVal newBackColor As Long = vbWhite)

    'Depending on the requested fitting technique, we may have to resize the image to a slightly different size
    ' than the one requested.  Before doing anything else, calculate that new size.
    Dim fitWidth As Long, fitHeight As Long
    
    Dim srcWidth As Long, srcHeight As Long
    srcWidth = pdImages(CurrentImage).Width
    srcHeight = pdImages(CurrentImage).Height
    
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
    If pdImages(CurrentImage).selectionActive Then
        pdImages(CurrentImage).selectionActive = False
        pdImages(CurrentImage).mainSelection.lockRelease
        tInit tSelection, False
    End If

    'Because most resize methods require a temporary layer, create one here
    Dim tmpLayer As pdLayer
    Set tmpLayer = New pdLayer

    Select Case resampleMethod

        'Nearest neighbor...
        Case RESIZE_NORMAL
        
            Message "Resizing image..."
            
            'Copy the current layer into this temporary layer at the new size
            tmpLayer.createFromExistingLayer pdImages(CurrentImage).mainLayer, fitWidth, fitHeight, False
            
        'Halftone resampling... I'm not sure what to actually call it, but since it's based off the
        ' StretchBlt mode Microsoft calls "halftone," I'm sticking with that
        Case RESIZE_HALFTONE
            
            Message "Resizing image..."
            
            'Copy the current layer into this temporary layer at the new size
            tmpLayer.createFromExistingLayer pdImages(CurrentImage).mainLayer, fitWidth, fitHeight, True
        
        'True bilinear sampling
        Case RESIZE_BILINEAR
        
            'If FreeImage is enabled, use their bilinear filter.  Similar results, much faster.
            If g_ImageFormats.FreeImageEnabled Then
            
                FreeImageResize tmpLayer, pdImages(CurrentImage).mainLayer, fitWidth, fitHeight, FILTER_BILINEAR
            
            'If FreeImage is not enabled, we have to do the resample ourselves.
            Else
            
                Message "Resampling image..."
        
                'Create a local array and point it at the pixel data of the current image
                Dim srcImageData() As Byte
                Dim srcSA As SAFEARRAY2D
                prepImageData srcSA
                CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
                'Resize the temporary layer to the target size, and point a second local array at it
                tmpLayer.createBlank fitWidth, fitHeight, pdImages(CurrentImage).mainLayer.getLayerColorDepth
                
                Dim dstImageData() As Byte
                Dim dstSA As SAFEARRAY2D
                
                prepSafeArray dstSA, tmpLayer
                CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
                
                'These values will help us access locations in the array more quickly.
                ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
                Dim qvDepth As Long
                qvDepth = tmpLayer.getLayerColorDepth \ 8
                
                'Create a filter support class, which will aid with edge handling and interpolation
                Dim fSupport As pdFilterSupport
                Set fSupport = New pdFilterSupport
                fSupport.setDistortParameters qvDepth, EDGE_CLAMP, True, curLayerValues.maxX, curLayerValues.MaxY
    
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
            FreeImageResize tmpLayer, pdImages(CurrentImage).mainLayer, fitWidth, fitHeight, FILTER_BSPLINE
            
        Case RESIZE_BICUBIC_MITCHELL
            FreeImageResize tmpLayer, pdImages(CurrentImage).mainLayer, fitWidth, fitHeight, FILTER_BICUBIC
            
        Case RESIZE_BICUBIC_CATMULL
            FreeImageResize tmpLayer, pdImages(CurrentImage).mainLayer, fitWidth, fitHeight, FILTER_CATMULLROM
        
        Case RESIZE_LANCZOS
            FreeImageResize tmpLayer, pdImages(CurrentImage).mainLayer, fitWidth, fitHeight, FILTER_LANCZOS3
            
    End Select
    
    'The temporary layer now holds a copy of the resized image.
    
    'Calculate the aspect ratio of this layer and the target picture box
    Dim srcAspect As Double, dstAspect As Double
    srcAspect = pdImages(CurrentImage).Width / pdImages(CurrentImage).Height
    dstAspect = iWidth / iHeight
    
    Dim dstX As Long, dstY As Long
    
    'We now want to copy the resized image into the current image using the technique requested by the user.
    Select Case fitMethod
    
        'Stretch-to-fit.  This is default resize behavior in all image editing software
        Case 0
    
            'Very simple - just copy the resized image back into the main layer
            pdImages(CurrentImage).mainLayer.createFromExistingLayer tmpLayer
    
        'Fit inclusively.  This fits the image's largest dimension into the destination image, which can leave
        ' blank space - that space is filled by the background color parameter passed in.
        Case 1
        
            'Resize the main layer (destructively!) to fit the new dimensions
            pdImages(CurrentImage).mainLayer.createBlank iWidth, iHeight, pdImages(CurrentImage).mainLayer.getLayerColorDepth, newBackColor
        
            'BitBlt the old image, centered, onto the new layer
            If srcAspect > dstAspect Then
                dstY = CLng((iHeight - fitHeight) / 2)
                dstX = 0
            Else
                dstX = CLng((iWidth - fitWidth) / 2)
                dstY = 0
            End If
            
            BitBlt pdImages(CurrentImage).mainLayer.getLayerDC, dstX, dstY, fitWidth, fitHeight, tmpLayer.getLayerDC, 0, 0, vbSrcCopy
        
        'Fit exclusively.  This fits the image's smallest dimension into the destination image, which means no
        ' blank space - but parts of the image may get cropped out.
        Case 2
        
            'Resize the main layer (destructively!) to fit the new dimensions
            pdImages(CurrentImage).mainLayer.createBlank iWidth, iHeight, pdImages(CurrentImage).mainLayer.getLayerColorDepth, newBackColor
        
            'BitBlt the old image, centered, onto the new layer
            If srcAspect < dstAspect Then
                dstY = CLng((iHeight - fitHeight) / 2)
                dstX = 0
            Else
                dstX = CLng((iWidth - fitWidth) / 2)
                dstY = 0
            End If
            
            BitBlt pdImages(CurrentImage).mainLayer.getLayerDC, dstX, dstY, fitWidth, fitHeight, tmpLayer.getLayerDC, 0, 0, vbSrcCopy
        
    End Select
    
    'We are finished with the temporary layer, so release it
    Set tmpLayer = Nothing
    
    'Update the main image's size values
    pdImages(CurrentImage).updateSize
    DisplaySize pdImages(CurrentImage).Width, pdImages(CurrentImage).Height
    
    'Fit the new image on-screen and redraw its viewport
    PrepareViewport FormMain.ActiveForm, "Image resize"
    
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

Private Sub picBackColor_Click()
    
    'Use the default color dialog to select a new color
    Dim newColor As Long
    If showColorDialog(newColor, Me, picBackColor.backColor) Then
        picBackColor.backColor = newColor
        'updatePreview
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

'When the image's aspect ratio is not being preserved, the user is provided with additional resize options.
Private Sub updateFormLayout()

    Dim i As Long
    Dim formHeightDifference As Long
    Me.ScaleMode = vbTwips
    formHeightDifference = Me.Height - Me.ScaleHeight
    Me.ScaleMode = vbPixels

    If CBool(chkRatio) Then
    
        'Hide all "fit image" controls
        lblFit.Visible = False
        
        For i = 0 To optFit.Count - 1
            optFit(i).Visible = False
        Next i
        
        For i = 0 To lblSubtext.Count - 1
            lblSubtext(i).Visible = False
        Next i
        
        picBackColor.Visible = False
        
        'Move the command bar into place
        lblBackground.Top = chkNames.Top + chkNames.Height + 16
        CmdOK.Top = lblBackground.Top + 10
        CmdCancel.Top = CmdOK.Top
        
        'Resize the form to match
        Me.Height = formHeightDifference + (CmdOK.Top + CmdOK.Height + 10) * Screen.TwipsPerPixelY
    
    Else
    
        'Show all "fit image" controls
        lblFit.Visible = True
        
        For i = 0 To optFit.Count - 1
            optFit(i).Visible = True
        Next i
        
        For i = 0 To lblSubtext.Count - 1
            lblSubtext(i).Visible = True
        Next i
        
        'Hide the background color selector only if the image is not 32bpp.  (If it is 32bpp, blank space will
        ' be filled by transparency, not color.)
        If pdImages(CurrentImage).mainLayer.getLayerColorDepth <> 32 Then
            picBackColor.Visible = True
        End If
        
        'Move the command bar into place
        lblBackground.Top = lblSubtext(2).Top + lblSubtext(2).Height + 16
        CmdOK.Top = lblBackground.Top + 10
        CmdCancel.Top = CmdOK.Top
        
        'Resize the form to match
        Me.Height = formHeightDifference + (CmdOK.Top + CmdOK.Height + 10) * Screen.TwipsPerPixelY
    
    End If

End Sub
