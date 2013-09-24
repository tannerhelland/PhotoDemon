VERSION 5.00
Begin VB.Form FormReduceColors 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Reduce Image Colors"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12315
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
   ScaleWidth      =   821
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5775
      Width           =   12315
      _ExtentX        =   21722
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
   Begin PhotoDemon.smartOptionButton optQuant 
      Height          =   375
      Index           =   0
      Left            =   6120
      TabIndex        =   4
      Top             =   2400
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   661
      Caption         =   "Xiaolin Wu"
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
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.smartOptionButton optQuant 
      Height          =   375
      Index           =   1
      Left            =   6120
      TabIndex        =   5
      Top             =   2880
      Width           =   3120
      _ExtentX        =   5503
      _ExtentY        =   661
      Caption         =   "NeuQuant neural network"
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
   Begin VB.Label lblWarning 
      BackStyle       =   0  'Transparent
      Caption         =   "Note: the FreeImage plugin is missing.  Please install it if you wish to use this tool."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1455
      Left            =   6000
      TabIndex        =   2
      Top             =   4200
      Visible         =   0   'False
      Width           =   6015
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblQuantMethod 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "quantization method:"
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
      Height          =   405
      Left            =   6000
      TabIndex        =   1
      Top             =   1920
      Width           =   2265
   End
End
Attribute VB_Name = "FormReduceColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Color Reduction Form
'Copyright ©2000-2013 by Tanner Helland
'Created: 4/October/00
'Last updated: 24/August/13
'Last update: move all manual reduction routines to the Posterize form, where they make more sense
'
'In the original incarnation of PhotoDemon, this was a central part of the project. I have since not used it much
' (since the project is now centered around 24/32bpp imaging), but as it costs nothing to tie into FreeImage's advanced
' color reduction routines, I figure it's worth keeping this dialog around.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'SetDIBitsToDevice is used to interact with the FreeImage DLL
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As Any, ByVal wUsage As Long) As Long

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'OK button
Private Sub cmdBar_OKClick()
    
    'Xiaolin Wu
    If optQuant(0).Value Then
        Process "Reduce colors", , buildParams(REDUCECOLORS_AUTO, FIQ_WUQUANT)
        
    'NeuQuant
    Else
        Process "Reduce colors", , buildParams(REDUCECOLORS_AUTO, FIQ_NNQUANT)
    End If
    
End Sub

Private Sub Form_Activate()
        
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
    'Render a preview
    cmdBar.markPreviewStatus True
    updatePreview
    
End Sub

Private Sub Form_Load()

    'Suspend previews until the dialog has been fully initialized
    cmdBar.markPreviewStatus False

    'Only allow AutoReduction stuff if the FreeImage dll was found.
    If Not g_ImageFormats.FreeImageEnabled Then
        optQuant(0).Enabled = False
        optQuant(1).Enabled = False
        lblWarning.Visible = True
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Enable/disable the manual settings depending on which option button has been selected
Private Sub OptQuant_Click(Index As Integer)
    updatePreview
End Sub

'Automatic 8-bit color reduction via the FreeImage DLL.
Public Sub ReduceImageColors_Auto(ByVal qMethod As Long, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)

    'If a selection is active, remove it.
    If pdImages(CurrentImage).selectionActive Then
        pdImages(CurrentImage).selectionActive = False
        pdImages(CurrentImage).mainSelection.lockRelease
        metaToggle tSelection, False
    End If

    'If this is a preview, we want to perform the color reduction on a temporary image
    If toPreview Then
        Dim tmpSA As SAFEARRAY2D
        prepImageData tmpSA, toPreview, dstPic
    End If

    'Make sure we found the FreeImage plug-in when the program was loaded
    If g_ImageFormats.FreeImageEnabled Then
    
        'Load the FreeImage dll into memory
        Dim hLib As Long
        hLib = LoadLibrary(g_PluginPath & "FreeImage.dll")
        
        If Not toPreview Then Message "Quantizing image using the FreeImage library..."
        
        'Convert our current layer to a FreeImage-type DIB
        Dim fi_DIB As Long
        
        If toPreview Then
            If workingLayer.getLayerColorDepth = 32 Then workingLayer.compositeBackgroundColor 255, 255, 255
            fi_DIB = FreeImage_CreateFromDC(workingLayer.getLayerDC)
        Else
            If pdImages(CurrentImage).mainLayer.getLayerColorDepth = 32 Then pdImages(CurrentImage).mainLayer.compositeBackgroundColor 255, 255, 255
            fi_DIB = FreeImage_CreateFromDC(pdImages(CurrentImage).mainLayer.getLayerDC)
        End If
        
        'Use that handle to save the image to GIF format, with required 8bpp (256 color) conversion
        If fi_DIB <> 0 Then
            
            Dim returnDIB As Long
            
            returnDIB = FreeImage_ColorQuantizeEx(fi_DIB, qMethod, True)
            
            'If this is a preview, render it to the temporary layer.  Otherwise, use the current main layer.
            If toPreview Then
                workingLayer.createBlank workingLayer.getLayerWidth, workingLayer.getLayerHeight, 24
                SetDIBitsToDevice workingLayer.getLayerDC, 0, 0, workingLayer.getLayerWidth, workingLayer.getLayerHeight, 0, 0, 0, workingLayer.getLayerHeight, ByVal FreeImage_GetBits(returnDIB), ByVal FreeImage_GetInfo(returnDIB), 0&
            Else
                pdImages(CurrentImage).mainLayer.createBlank pdImages(CurrentImage).Width, pdImages(CurrentImage).Height, 24
                SetDIBitsToDevice pdImages(CurrentImage).mainLayer.getLayerDC, 0, 0, pdImages(CurrentImage).Width, pdImages(CurrentImage).Height, 0, 0, 0, pdImages(CurrentImage).Height, ByVal FreeImage_GetBits(returnDIB), ByVal FreeImage_GetInfo(returnDIB), 0&
            End If
            
            'With the transfer complete, release the FreeImage DIB and unload the library
            If returnDIB <> 0 Then FreeImage_UnloadEx returnDIB
            FreeLibrary hLib
     
            'If this is a preview, draw the new image to the picture box and exit.  Otherwise, render the new main image layer.
            If toPreview Then
                finalizeImageData toPreview, dstPic
            Else
                ScrollViewport FormMain.ActiveForm
                Message "Image successfully quantized to %1 unique colors. ", 256
            End If
            
        End If
        
    Else
        pdMsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or disabled upon program initialization." & vbCrLf & vbCrLf & "To enable support for this feature, please copy the FreeImage.dll file into the plug-in directory and reload the program.", vbExclamation + vbOKOnly + vbApplicationModal, " FreeImage Interface Error"
        Exit Sub
    End If
    
End Sub

Private Sub sltBlue_Change()
    updatePreview
End Sub

Private Sub sltGreen_Change()
    updatePreview
End Sub

Private Sub sltRed_Change()
    updatePreview
End Sub

'Use this sub to update the on-screen preview
Private Sub updatePreview()
    
    If cmdBar.previewsAllowed Then
        If optQuant(0).Value Then
            ReduceImageColors_Auto FIQ_WUQUANT, True, fxPreview
        Else
            ReduceImageColors_Auto FIQ_NNQUANT, True, fxPreview
        End If
    End If
    
End Sub
