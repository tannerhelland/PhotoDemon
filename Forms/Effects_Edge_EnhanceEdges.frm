VERSION 5.00
Begin VB.Form FormEdgeEnhance 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Enhance image edges"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12195
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
   ScaleWidth      =   813
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5775
      Width           =   12195
      _ExtentX        =   21511
      _ExtentY        =   1323
      BackColor       =   14802140
   End
   Begin VB.ListBox LstEdgeOptions 
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
      Height          =   2220
      Left            =   6240
      TabIndex        =   1
      Top             =   480
      Width           =   5655
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
   Begin PhotoDemon.smartCheckBox chkDirection 
      Height          =   360
      Index           =   0
      Left            =   6240
      TabIndex        =   5
      Top             =   3360
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   582
      Caption         =   "horizontal"
   End
   Begin PhotoDemon.smartCheckBox chkDirection 
      Height          =   360
      Index           =   1
      Left            =   6240
      TabIndex        =   6
      Top             =   3840
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   582
      Caption         =   "vertical"
   End
   Begin PhotoDemon.sliderTextCombo sltStrength 
      Height          =   720
      Left            =   6000
      TabIndex        =   7
      Top             =   4560
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   1270
      Caption         =   "strength"
      Max             =   100
      NotchPosition   =   2
      NotchValueCustom=   50
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "detection direction(s)"
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
      Index           =   1
      Left            =   6000
      TabIndex        =   4
      Top             =   3000
      Width           =   2235
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "edge detection technique"
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
      Left            =   6000
      TabIndex        =   2
      Top             =   120
      Width           =   2640
   End
End
Attribute VB_Name = "FormEdgeEnhance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Edge Enhancement Interface
'Copyright 2002-2015 by Tanner Helland
'Created: sometimes 2002
'Last updated: 12/June/14
'Last update: give tool its own form, and open it up to all available edge detection techniques
'
'This edge enhancement function allows the user to selectively emphasize image edges using any available edge
' detection technique.  PD's compositor is then used to composite the results back onto the base image at some
' variable strength specified by the user.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'To prevent recursion when setting checkbox state, this value is used to notify the function that a state change
' is already underway
Private ignoreStateChanges As Boolean

'The direction checkboxes are somewhat odd; one or the other should always be selected, so we have to do some special
' checking to make sure that happens.
Private Sub chkDirection_Click(Index As Integer)

    If ignoreStateChanges Then Exit Sub
    
    ignoreStateChanges = True

    Dim otherIndex As Long
    If Index = 0 Then otherIndex = 1 Else otherIndex = 0

    If Not chkDirection(Index) Then
        If Not chkDirection(otherIndex) Then chkDirection(otherIndex).Value = vbChecked
    End If
    
    ignoreStateChanges = False
    
    updatePreview

End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Enhance edges", , buildParams(LstEdgeOptions.ListIndex, getDirectionality(), sltStrength.Value), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub Form_Activate()
        
    'Apply translations and visual themes
    MakeFormPretty Me
    
    'Update the descriptions (this will also draw a preview of the selected edge-detection algorithm)
    cmdBar.markPreviewStatus True
    updatePreview
    
End Sub

'Apply any supported edge detection filter to an image.  Directionality can be specified, but note that only some
' algorithms support the parameter.
Public Sub ApplyEdgeEnhancement(ByVal edgeDetectionType As PD_EDGE_DETECTION, ByVal edgeDirectionality As PD_EDGE_DETECTION_DIRECTION, ByVal enhanceStrength As Double, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)

    'Applying an edge detection filter generally happens via these steps:
    
    '1) Set up any parameters we know in advance, like generating a String name for the supplied filter, and converting
    '    the optional "blackBackground" parameter into PD's internal ParamString format
    '2) Retrieve a relevant convolution matrix for the requested filter
    '3) Supply the full ParamString, including convo matrix, to PD's central ApplyConvolutionFilter function
    '4) If necessary, repeat steps 2 and 3 to retrieve multiple directionality images
    
    Dim tmpParamString As String, convolutionMatrixString As String
    
    '1a) Generate a name for the requested filter
    tmpParamString = getNameOfEdgeDetector(edgeDetectionType) & "|"
    
    '1b) Add in the invert (black background) parameter
    tmpParamString = tmpParamString & "False" & "|"
    
    '2a) Retrieve the relevant convolution matrix for this filter
    convolutionMatrixString = getParamStringForEdgeDetector(edgeDetectionType, edgeDirectionality)
    
    '2b) Merge the retrieved convolution matrix string with our name and invert params
    tmpParamString = tmpParamString & convolutionMatrixString
    
    'Next, we need to obtain a DIB of the processed edge detection results for the image.  This requires two or
    ' three passes, contingent on the detection type.  In order to update the progress bar correctly, calculate the number
    ' of passes required in advance.
    Dim numPassesRequired As Long
    If (edgeDetectionType = PD_EDGE_ARTISTIC_CONTOUR) Then
        numPassesRequired = 2
    ElseIf isEdgeDetectionSinglePass(edgeDetectionType, edgeDirectionality) Then
        numPassesRequired = 2
    Else
        numPassesRequired = 3
    End If
    
    If Not toPreview Then Message "Applying pass %1 of %2 for %3 filter...", "1", numPassesRequired, getNameOfEdgeDetector(edgeDetectionType)
    
    'Use PD's central image handler to populate the public workingDIB object.
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    
    'Create a second DIB copy.  This will receive the edge-detection copy of the image.
    Dim edgeDIB As pdDIB
    Set edgeDIB = New pdDIB
    edgeDIB.createFromExistingDIB workingDIB
    
    '3a) If the function is single-pass compatible (e.g. it does not require us to traverse the image multiple times, then
    '     blend the edge detection results), supply the compiled param string to PD's central convolution function and exit
    If (edgeDetectionType = PD_EDGE_ARTISTIC_CONTOUR) Then
        CreateContourDIB True, workingDIB, edgeDIB, toPreview, workingDIB.getDIBWidth * numPassesRequired, 0
    
    Else
        ConvolveDIB tmpParamString, workingDIB, edgeDIB, toPreview, workingDIB.getDIBWidth * numPassesRequired, 0
    
    End If
    
    'A pdCompositor class is required to selectively blend the edge detection results back onto the main image
    Dim cComposite As pdCompositor
    Set cComposite = New pdCompositor
    
    '3b) If the requested edge function is not single-pass compatible, run a second pass in the opposite direction,
    '     the blend the results back onto edgeDIB.
    If Not isEdgeDetectionSinglePass(edgeDetectionType, edgeDirectionality) Then
    
        'Create a second DIB copy.  This will receive the edge-detection copy of the image.
        Dim tmpEdgeDIB As pdDIB
        Set tmpEdgeDIB = New pdDIB
        tmpEdgeDIB.createFromExistingDIB workingDIB
        
        'When two passes are required, the vertical direction is always applied first.  Thus we know we need to apply the
        ' horizontal direction next.  Generate a new param string for the horizontal direction.
        If Not toPreview Then Message "Applying pass %1 of %2 for %3 filter...", "2", numPassesRequired, getNameOfEdgeDetector(edgeDetectionType)
        
        tmpParamString = getNameOfEdgeDetector(edgeDetectionType) & "|"
        tmpParamString = tmpParamString & "False" & "|"
        convolutionMatrixString = getParamStringForEdgeDetector(edgeDetectionType, PD_EDGE_DIR_HORIZONTAL)
        tmpParamString = tmpParamString & convolutionMatrixString
                
        'Use the central ConvolveDIB function to apply the new convolution to workingDIB
        ConvolveDIB tmpParamString, workingDIB, tmpEdgeDIB, toPreview, workingDIB.getDIBWidth * numPassesRequired, workingDIB.getDIBWidth
        
        'The compositor requires premultiplied alpha, so convert both top and bottom layers now
        edgeDIB.setAlphaPremultiplication True
        tmpEdgeDIB.setAlphaPremultiplication True
        
        'Use the pdCompositor class to blend the results of the second edge detection pass with the first pass.
        cComposite.quickMergeTwoDibsOfEqualSize edgeDIB, tmpEdgeDIB, BL_SCREEN
        
        'Remove premultiplication
        edgeDIB.setAlphaPremultiplication False
        
        Set tmpEdgeDIB = Nothing
        
    End If
    
    
    'edgeDIB now contains a complete edge-detection copy of the image, using the supplied edge detector algorithm.
    ' We now need to selectively blend the results back onto the main image, at the strength requested.
    If Not toPreview Then Message "Applying pass %1 of %2 for %3 filter...", numPassesRequired, numPassesRequired, getNameOfEdgeDetector(edgeDetectionType)
    
    'Apply premultiplication prior to compositing
    edgeDIB.setAlphaPremultiplication True
    workingDIB.setAlphaPremultiplication True
    
    'To make the merge operation easier, we're going to place our DIBs inside temporary layers.  This allows us to use existing
    ' layer code to handle the merge.
    Dim tmpLayerTop As pdLayer, tmpLayerBottom As pdLayer
    Set tmpLayerTop = New pdLayer
    Set tmpLayerBottom = New pdLayer
    
    tmpLayerTop.InitializeNewLayer PDL_IMAGE, , edgeDIB
    tmpLayerBottom.InitializeNewLayer PDL_IMAGE, , workingDIB
    
    tmpLayerTop.setLayerBlendMode BL_SCREEN
    tmpLayerTop.setLayerOpacity enhanceStrength
    
    cComposite.mergeLayers tmpLayerTop, tmpLayerBottom, True
    
    'Copy the finished DIB from the bottom layer back into workingDIB
    workingDIB.createFromExistingDIB tmpLayerBottom.layerDIB
    
    Set tmpLayerTop = Nothing
    Set tmpLayerBottom = Nothing
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingDIB
    finalizeImageData toPreview, dstPic, True

End Sub

'Return the naem of an edge detection type as a human-readable string
Private Function getNameOfEdgeDetector(ByVal edgeDetectionType As PD_EDGE_DETECTION) As String

    Select Case edgeDetectionType
                
        Case PD_EDGE_HILITE
            getNameOfEdgeDetector = g_Language.TranslateMessage("Hilite edge detection")
            
        Case PD_EDGE_LAPLACIAN
            getNameOfEdgeDetector = g_Language.TranslateMessage("Laplacian edge detection")
        
        Case PD_EDGE_PHOTODEMON
            getNameOfEdgeDetector = g_Language.TranslateMessage("PhotoDemon edge detection")
            
        Case PD_EDGE_PREWITT
            getNameOfEdgeDetector = g_Language.TranslateMessage("Prewitt edge detection")
        
        Case PD_EDGE_ROBERTS
            getNameOfEdgeDetector = g_Language.TranslateMessage("Roberts cross edge detection")
            
        Case PD_EDGE_SOBEL
            getNameOfEdgeDetector = g_Language.TranslateMessage("Sobel edge detection")
            
    End Select

End Function

'Given an edge detection type and a direction, return TRUE if the requested edge detector can be applied in a single pass.
' Return FALSE if the function requires multiple image passes.
Private Function isEdgeDetectionSinglePass(ByVal edgeDetectionType As PD_EDGE_DETECTION, Optional ByVal edgeDirectionality As PD_EDGE_DETECTION_DIRECTION = PD_EDGE_DIR_ALL) As Boolean

    'Convolution matrix strings are assembled in two or three steps:
    ' 1) Add divisor and offset values
    ' 2 (optional) ) Check directionality and adjust behavior accordingly
    ' 3) Build actual convolution matrix
    Select Case edgeDetectionType
        
        Case PD_EDGE_ARTISTIC_CONTOUR
            isEdgeDetectionSinglePass = True
    
        'Hilite detection (doesn't support directionality)
        Case PD_EDGE_HILITE
            isEdgeDetectionSinglePass = True
        
        'Laplacian is unique because it supports a different operator for all directionalities, so even horizontal/vertical can
        ' be done in a single pass.
        Case PD_EDGE_LAPLACIAN
            isEdgeDetectionSinglePass = True
                
        'PhotoDemon edge detection (doesn't support directionality)
        Case PD_EDGE_PHOTODEMON
            isEdgeDetectionSinglePass = True
        
        'Prewitt edge detection is unidirectional
        Case PD_EDGE_PREWITT
            If (edgeDirectionality = PD_EDGE_DIR_HORIZONTAL) Or (edgeDirectionality = PD_EDGE_DIR_VERTICAL) Then
                isEdgeDetectionSinglePass = True
            Else
                isEdgeDetectionSinglePass = False
            End If
            
        'Roberts cross edge detection is unidirectional
        Case PD_EDGE_ROBERTS
            If (edgeDirectionality = PD_EDGE_DIR_HORIZONTAL) Or (edgeDirectionality = PD_EDGE_DIR_VERTICAL) Then
                isEdgeDetectionSinglePass = True
            Else
                isEdgeDetectionSinglePass = False
            End If
        
        'Sobel edge detection is unidirectional
        Case PD_EDGE_SOBEL
            If (edgeDirectionality = PD_EDGE_DIR_HORIZONTAL) Or (edgeDirectionality = PD_EDGE_DIR_VERTICAL) Then
                isEdgeDetectionSinglePass = True
            Else
                isEdgeDetectionSinglePass = False
            End If
        
    End Select

End Function

'Given an internal edge detection type (and optionally, a direction), calculate a matching convolution matrix and return it
Private Function getParamStringForEdgeDetector(ByVal edgeDetectionType As PD_EDGE_DETECTION, Optional ByVal edgeDirectionality As PD_EDGE_DETECTION_DIRECTION = PD_EDGE_DIR_ALL) As String

    Dim convoString As String
    convoString = ""
    
    'Convolution matrix strings are assembled in two or three steps:
    ' 1) Add divisor and offset values
    ' 2 (optional) Check directionality and adjust behavior accordingly
    ' 3) Build actual convolution matrix
    Select Case edgeDetectionType
    
        'Hilite detection (doesn't support directionality)
        Case PD_EDGE_HILITE
            
            'Divisor/offset
            convoString = convoString & "1|0|"
    
            'Actual convo matrix
            convoString = convoString & "0|0|0|0|0|"
            convoString = convoString & "0|-4|-2|-1|0|"
            convoString = convoString & "0|-2|10|0|0|"
            convoString = convoString & "0|-1|0|0|0|"
            convoString = convoString & "0|0|0|0|0"
        
        Case PD_EDGE_LAPLACIAN
            
            'Actual convo matrix varies according to direction
            If edgeDirectionality = PD_EDGE_DIR_HORIZONTAL Then
            
                'Divisor/offset
                convoString = convoString & "0.25|0|"
                
                convoString = convoString & "0|0|0|0|0|"
                convoString = convoString & "0|0|0|0|0|"
                convoString = convoString & "0|-1|2|-1|0|"
                convoString = convoString & "0|0|0|0|0|"
                convoString = convoString & "0|0|0|0|0"
                
            ElseIf edgeDirectionality = PD_EDGE_DIR_VERTICAL Then
            
                'Divisor/offset
                convoString = convoString & "0.25|0|"
                
                convoString = convoString & "0|0|0|0|0|"
                convoString = convoString & "0|0|-1|0|0|"
                convoString = convoString & "0|0|2|0|0|"
                convoString = convoString & "0|0|-1|0|0|"
                convoString = convoString & "0|0|0|0|0"
                
            Else
            
                'Divisor/offset
                convoString = convoString & "0.5|0|"
                
                convoString = convoString & "0|0|0|0|0|"
                convoString = convoString & "0|0|-1|0|0|"
                convoString = convoString & "0|-1|4|-1|0|"
                convoString = convoString & "0|0|-1|0|0|"
                convoString = convoString & "0|0|0|0|0"
            
            End If
        
        'PhotoDemon edge detection (doesn't support directionality)
        Case PD_EDGE_PHOTODEMON
        
            'Divisor/offset
            convoString = convoString & "1|0|"
            
            'Actual convo matrix
            convoString = convoString & "0|-1|0|0|0|"
            convoString = convoString & "0|0|0|0|-1|"
            convoString = convoString & "0|0|4|0|0|"
            convoString = convoString & "-1|0|0|0|0|"
            convoString = convoString & "0|0|0|-1|0"
        
        'Prewitt edge detection (directionality supported)
        Case PD_EDGE_PREWITT
        
            'Divisor/offset
            convoString = convoString & "1|0|"
            
            'Actual convo matrix varies according to direction
            If edgeDirectionality = PD_EDGE_DIR_HORIZONTAL Then
                convoString = convoString & "0|0|0|0|0|"
                convoString = convoString & "0|-1|0|1|0|"
                convoString = convoString & "0|-1|0|1|0|"
                convoString = convoString & "0|-1|0|1|0|"
                convoString = convoString & "0|0|0|0|0"
            Else
                convoString = convoString & "0|0|0|0|0|"
                convoString = convoString & "0|1|1|1|0|"
                convoString = convoString & "0|0|0|0|0|"
                convoString = convoString & "0|-1|-1|-1|0|"
                convoString = convoString & "0|0|0|0|0"
            End If
        
        'Roberts cross edge detection (directionality supported)
        Case PD_EDGE_ROBERTS
        
            'Divisor/offset
            convoString = convoString & "0.5|0|"
            
            'Actual convo matrix varies according to direction
            If edgeDirectionality = PD_EDGE_DIR_HORIZONTAL Then
                convoString = convoString & "0|0|0|0|0|"
                convoString = convoString & "0|-1|0|0|0|"
                convoString = convoString & "0|0|1|0|0|"
                convoString = convoString & "0|0|0|0|0|"
                convoString = convoString & "0|0|0|0|0"
            Else
                convoString = convoString & "0|0|0|0|0|"
                convoString = convoString & "0|0|0|-1|0|"
                convoString = convoString & "0|0|1|0|0|"
                convoString = convoString & "0|0|0|0|0|"
                convoString = convoString & "0|0|0|0|0"
            End If
        
        'Sobel edge detection (directionality supported)
        Case PD_EDGE_SOBEL
            
            'Divisor/offset
            convoString = convoString & "1|0|"
            
            'Actual convo matrix varies according to direction
            If edgeDirectionality = PD_EDGE_DIR_HORIZONTAL Then
                convoString = convoString & "0|0|0|0|0|"
                convoString = convoString & "0|-1|0|1|0|"
                convoString = convoString & "0|-2|0|2|0|"
                convoString = convoString & "0|-1|0|1|0|"
                convoString = convoString & "0|0|0|0|0"
            Else
                convoString = convoString & "0|0|0|0|0|"
                convoString = convoString & "0|1|2|1|0|"
                convoString = convoString & "0|0|0|0|0|"
                convoString = convoString & "0|-1|-2|-1|0|"
                convoString = convoString & "0|0|0|0|0"
            End If
    
    End Select
    
    getParamStringForEdgeDetector = convoString

End Function

'This code is a modified version of an algorithm originally developed by Manuel Augusto Santos.  A link to his original
' implementation is available from the "Help -> About PhotoDemon" menu option.
Public Sub FilterSmoothContour(Optional ByVal blackBackground As Boolean = False, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)

    If Not toPreview Then Message "Tracing image edges with virtual paintbrush..."
        
    'Create a local array and point it at the pixel data of the current image
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent blurred pixel values from spreading across the image as we go.)
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.createFromExistingDIB workingDIB
    
    CreateContourDIB blackBackground, srcDIB, workingDIB, toPreview
    
    srcDIB.eraseDIB
    Set srcDIB = Nothing
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingDIB
    finalizeImageData toPreview, dstPic
    
End Sub

Private Sub Form_Load()
    
    'Suspend previews until the list box has been populated
    cmdBar.markPreviewStatus False
    
    'Generate a list box with all the currently implemented edge detection algorithms
    LstEdgeOptions.AddItem "Artistic contour", 0
    LstEdgeOptions.AddItem "Hilite", 1
    LstEdgeOptions.AddItem "Laplacian", 2
    LstEdgeOptions.AddItem "PhotoDemon", 3
    LstEdgeOptions.AddItem "Prewitt", 4
    LstEdgeOptions.AddItem "Roberts cross", 5
    LstEdgeOptions.AddItem "Sobel", 6
    
    LstEdgeOptions.ListIndex = 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub LstEdgeOptions_Click()
    
    cmdBar.markPreviewStatus False
    
    'Directionality is only supported by some transforms, so de/activate the directionality check boxes to match the
    ' capabilities of the selected transform
    Select Case LstEdgeOptions.ListIndex
    
        Case PD_EDGE_ARTISTIC_CONTOUR
            changeCheckboxActivation False
        
        Case PD_EDGE_HILITE
            changeCheckboxActivation False
        
        Case PD_EDGE_LAPLACIAN
            changeCheckboxActivation True
        
        Case PD_EDGE_PHOTODEMON
            changeCheckboxActivation False
        
        Case PD_EDGE_PREWITT
            changeCheckboxActivation True
        
        Case PD_EDGE_ROBERTS
            changeCheckboxActivation True
            
        Case PD_EDGE_SOBEL
            changeCheckboxActivation True
    
    End Select
    
    cmdBar.markPreviewStatus True
    updatePreview
    
End Sub

'Dis/enable the directionality checkboxes to match the request; when checkboxes are disabled, their value is automatically
' forced to TRUE.
Private Sub changeCheckboxActivation(ByVal toEnable As Boolean)

    If toEnable Then
    
        chkDirection(0).Enabled = True
        chkDirection(1).Enabled = True
    
    'Activate both directions, then disable the checkboxes
    Else
    
        If Not chkDirection(0) Then chkDirection(0).Value = vbChecked
        If Not chkDirection(1) Then chkDirection(1).Value = vbChecked
        
        chkDirection(0).Enabled = False
        chkDirection(1).Enabled = False
    
    End If
    
End Sub

'Convert the directionality checkboxes to PD's internal edge detection definitions
Private Function getDirectionality() As PD_EDGE_DETECTION_DIRECTION

    If CBool(chkDirection(0)) And Not CBool(chkDirection(1)) Then
        getDirectionality = PD_EDGE_DIR_HORIZONTAL
    ElseIf CBool(chkDirection(1)) And Not CBool(chkDirection(0)) Then
        getDirectionality = PD_EDGE_DIR_VERTICAL
    Else
        getDirectionality = PD_EDGE_DIR_ALL
    End If

End Function

'Update the live preview of the selected edge detection options
Private Sub updatePreview()
    
    If cmdBar.previewsAllowed Then
        ApplyEdgeEnhancement LstEdgeOptions.ListIndex, getDirectionality(), sltStrength.Value, True, fxPreview
    End If
    
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

Private Sub sltStrength_Change()
    updatePreview
End Sub
