VERSION 5.00
Begin VB.Form FormReduceColors 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Indexed color"
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
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5775
      Width           =   12315
      _ExtentX        =   21722
      _ExtentY        =   1323
      BackColor       =   14802140
   End
   Begin PhotoDemon.smartOptionButton optQuant 
      Height          =   360
      Index           =   0
      Left            =   6120
      TabIndex        =   4
      Top             =   2040
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   635
      Caption         =   "Xiaolin Wu"
      FontSize        =   11
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
      Height          =   360
      Index           =   1
      Left            =   6120
      TabIndex        =   5
      Top             =   2520
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   635
      Caption         =   "NeuQuant neural network"
      FontSize        =   11
   End
   Begin VB.Label lblFlatten 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Note: this operation will flatten the image before converting it to indexed color mode."
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
      Height          =   1125
      Left            =   6000
      TabIndex        =   6
      Top             =   3360
      Visible         =   0   'False
      Width           =   6090
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblWarning 
      BackStyle       =   0  'Transparent
      Caption         =   "The FreeImage plugin is missing.  Please install it if you wish to use this tool."
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
      Height          =   975
      Left            =   6000
      TabIndex        =   2
      Top             =   4680
      Visible         =   0   'False
      Width           =   6015
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblQuantMethod 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "quantization method:="
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
      Top             =   1560
      Width           =   2445
   End
End
Attribute VB_Name = "FormReduceColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Color Reduction Form
'Copyright 2000-2015 by Tanner Helland
'Created: 4/October/00
'Last updated: 14/April/14
'Last update: rewrite function against layers; note that this will now flatten a layered image before proceeding
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

'OK button
Private Sub cmdBar_OKClick()
    
    'Xiaolin Wu
    If optQuant(0).Value Then
        Process "Reduce colors", , buildParams(REDUCECOLORS_AUTO, FIQ_WUQUANT), UNDO_IMAGE
        
    'NeuQuant
    Else
        Process "Reduce colors", , buildParams(REDUCECOLORS_AUTO, FIQ_NNQUANT), UNDO_IMAGE
    End If
    
End Sub

Private Sub Form_Activate()
        
    'Apply translations and visual themes
    MakeFormPretty Me
    
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
    
    'If the current image has more than one layer, warn the user that this action will flatten the image.
    If pdImages(g_CurrentImage).getNumOfLayers > 1 Then
        lblFlatten.Visible = True
    Else
        lblFlatten.Visible = False
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

    'If this is not a preview, and a selection is active on the main image, remove it.
    If (Not toPreview) And pdImages(g_CurrentImage).selectionActive Then
        pdImages(g_CurrentImage).selectionActive = False
        pdImages(g_CurrentImage).mainSelection.lockRelease
    End If
    
    'A temporary DIB is required to pass data back-and-forth with FreeImage
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    pdImages(g_CurrentImage).getCompositedImage tmpDIB, True
    
    'Color reduction only works on a flat copy of the image, so retrieve a composited version now.
    If toPreview Then
        Dim tmpSafeArray As SAFEARRAY2D
        previewNonStandardImage tmpSafeArray, tmpDIB, fxPreview
        
    'If this is not a preview, flatten the image before proceeding further
    Else
        
        SetProgBarMax 3
        SetProgBarVal 1
        Message "Flattening image..."
        Layer_Handler.flattenImage
        
    End If
    
    'At this point, we have two potential sources of our temporary DIB:
    ' 1) During a preview, the global workingDIB object contains a section of the image relevant to the
    '     preview window.
    ' 2) During the processing of a full image, pdImages(g_CurrentImage) has what we need (the flattened image).
    '
    'To simplify the code from here, we are going to conditionally copy the current flattened image into
    ' the global workingLayer DIB.  That way, we can use the same code path regardless of previews or
    ' actual processing.
    If Not toPreview Then
        Set workingDIB = New pdDIB
        workingDIB.createFromExistingDIB pdImages(g_CurrentImage).getLayerByIndex(0).layerDIB
    End If
    
    'FreeImage requires 24bpp images as color reduction targets.
    ' UPDATE MARCH 2015: the Wu quantizer supports a 32-bpp input, but it simply ignores alpha, resulting in a
    '                    nasty-looking image.  So we still forcibly downsample to 24bpp.
    If (workingDIB.getDIBColorDepth = 32) Then workingDIB.convertTo24bpp
    
    'Make sure we found the FreeImage plug-in when the program was loaded
    If g_ImageFormats.FreeImageEnabled Then
        
        'Convert our current DIB to a FreeImage-type DIB
        Dim fi_DIB As Long
        fi_DIB = FreeImage_CreateFromDC(workingDIB.getDIBDC)
        
        'Use that handle to request a color space transform from FreeImage
        If fi_DIB <> 0 Then
            
            If (Not toPreview) Then
                SetProgBarVal 2
                Message "Indexing colors..."
            End If
            
            Dim returnDIB As Long
            returnDIB = FreeImage_ColorQuantizeEx(fi_DIB, qMethod, True)
            
            Dim numOfQuantizedColors As Long
            
            'If this is a preview, copy the FreeImage data into the global workingDIB object.
            If toPreview Then
                Plugin_FreeImage_Interface.getPDDibFromFreeImageHandle returnDIB, workingDIB
                
            'This is not a preview.  Overwrite the current active layer with the quantized FreeImage data.
            Else
                
                SetProgBarVal 3
                Plugin_FreeImage_Interface.getPDDibFromFreeImageHandle returnDIB, pdImages(g_CurrentImage).getLayerByIndex(0).layerDIB
                pdImages(g_CurrentImage).getLayerByIndex(0).layerDIB.convertTo32bpp
                
                'Ask FreeImage for the size of the quantized image's palette
                numOfQuantizedColors = FreeImage_GetColorsUsed(returnDIB)
                
                'Notify the parent image of these changes
                pdImages(g_CurrentImage).notifyImageChanged UNDO_LAYER, 0
                pdImages(g_CurrentImage).notifyImageChanged UNDO_IMAGE
                
            End If
            
            'With the transfer complete, release the FreeImage DIB
            If returnDIB <> 0 Then FreeImage_UnloadEx returnDIB
            
            'If this is a preview, draw the new image to the picture box and exit.  Otherwise, render the new main image DIB.
            If toPreview Then
                finalizeNonstandardPreview dstPic
            Else
                Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
                SetProgBarVal 0
                releaseProgressBar
                Message "Image successfully quantized to %1 unique colors. ", numOfQuantizedColors
            End If
            
        End If
        
    Else
        PDMsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or disabled upon program initialization." & vbCrLf & vbCrLf & "To enable support for this feature, please copy the FreeImage.dll file into the plug-in directory and reload the program.", vbExclamation + vbOKOnly + vbApplicationModal, " FreeImage Interface Error"
        Exit Sub
    End If
    
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

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

