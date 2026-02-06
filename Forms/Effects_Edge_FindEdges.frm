VERSION 5.00
Begin VB.Form FormFindEdges 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Find edges"
   ClientHeight    =   6525
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12195
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
   ScaleHeight     =   435
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   813
   Visible         =   0   'False
   Begin PhotoDemon.pdListBox lstEdgeOptions 
      Height          =   2775
      Left            =   6000
      TabIndex        =   5
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   4895
      Caption         =   "edge detection technique"
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5775
      Width           =   12195
      _ExtentX        =   21511
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdCheckBox chkInvert 
      Height          =   330
      Left            =   6240
      TabIndex        =   3
      Top             =   5040
      Width           =   5610
      _ExtentX        =   9895
      _ExtentY        =   582
      Caption         =   "use black background"
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.pdCheckBox chkDirection 
      Height          =   360
      Index           =   0
      Left            =   6240
      TabIndex        =   1
      Top             =   3360
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   582
      Caption         =   "horizontal"
   End
   Begin PhotoDemon.pdCheckBox chkDirection 
      Height          =   360
      Index           =   1
      Left            =   6240
      TabIndex        =   4
      Top             =   3840
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   582
      Caption         =   "vertical"
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   2
      Left            =   6000
      Top             =   4560
      Width           =   5970
      _ExtentX        =   10530
      _ExtentY        =   503
      Caption         =   "other options"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   1
      Left            =   6000
      Top             =   3000
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   503
      Caption         =   "detection direction(s)"
      FontSize        =   12
      ForeColor       =   4210752
   End
End
Attribute VB_Name = "FormFindEdges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Edge Detection Interface
'Copyright 2000-2026 by Tanner Helland
'Created: 1/11/02
'Last updated: 30/July/17
'Last update: performance improvements, migrate to XML params
'
'All known edge-detection routines are handled from this form.  Most are simply convolution kernels that are passed off
' to the "ApplyConvolutionFilter" function, but at least one (Artistic Contour) resides here.
'
'As of June '14, directionality is now supported for all compatible filters, and the entire engine has been rewritten to make
' it easier to add additional operators in the future.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
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
    If (Index = 0) Then otherIndex = 1 Else otherIndex = 0

    If (Not chkDirection(Index)) Then
        If (Not chkDirection(otherIndex)) Then chkDirection(otherIndex).Value = True
    End If
    
    ignoreStateChanges = False
    
    UpdatePreview

End Sub

Private Sub chkInvert_Click()
    UpdatePreview
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Find edges", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

'Apply any supported edge detection filter to an image.  Directionality can be specified, but note that only some
' algorithms support the parameter.
Public Sub ApplyEdgeDetection(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)

    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim edgeDetectionType As PD_EdgeDetector, edgeDirectionality As PD_EdgeDirection, blackBackground As Boolean
    
    With cParams
        edgeDetectionType = .GetLong("method", lstEdgeOptions.ListIndex)
        edgeDirectionality = .GetLong("direction", pded_All)
        blackBackground = .GetBool("invert", True)
    End With
    
    'Applying an edge detection filter generally happens via these steps:
    
    '1) Set up any parameters we know in advance, like generating a String name for the supplied filter, and converting
    '    the optional "blackBackground" parameter into PD's internal ParamString format
    '2) Retrieve a relevant convolution matrix for the requested filter
    '3) Supply the full ParamString, including convo matrix, to PD's central ApplyConvolutionFilter function
    '4) If necessary, repeat steps 2 and 3 to retrieve multiple directionality images
    
    'Before doing anything else, check for the Artistic Contour filter type.  This is handled via its own dedicated
    ' function, separate from traditional convolution matrix processing
    If edgeDetectionType = pded_Contour Then
        FilterSmoothContour blackBackground, toPreview, dstPic
        
    Else
    
        'Because some of these parameters are handled separately, we now need to build a special parameter string
        ' for just the convolver.
        Dim cParamsOut As pdSerialize
        Set cParamsOut = New pdSerialize
        
        With cParamsOut
            .AddParam "name", GetNameOfEdgeDetector(edgeDetectionType)
            .AddParam "invert", Not blackBackground
            
            'We now need to calculate per-algorithm values using a separate helper function
            Dim fWeight As Double, fBias As Double, fMatrix As String
            GetParamStringForEdgeDetector edgeDetectionType, edgeDirectionality, fWeight, fBias, fMatrix
            .AddParam "weight", fWeight
            .AddParam "bias", fBias
            .AddParam "matrix", fMatrix
        End With
        
        '3a) If the function is single-pass compatible (e.g. it does not require us to traverse the image multiple times, then
        '     blend the edge detection results), supply the compiled param string to PD's central convolution function and exit
        If IsEdgeDetectionSinglePass(edgeDetectionType, edgeDirectionality) Then
            Filters_Area.ApplyConvolutionFilter_XML cParamsOut.GetParamString, toPreview, dstPic
            
        Else
        
            '3b) The requested edge operation cannot be applied in a single-pass.  We need to manually process the request by
            '     traversing the image twice, then blending the results.
                    
            'Note that the only purpose of the FilterType string is to display this message
            If (Not toPreview) Then Message "Applying pass %1 of %2 for %3 filter...", "1", "2", GetNameOfEdgeDetector(edgeDetectionType)
            
            'Create a local array and point it at the pixel data of the current image.  Note that the current layer is referred to as the
            ' DESTINATION image for the convolution; we will make a separate temp copy of the image to use as the SOURCE.
            Dim dstSA As SafeArray2D
            EffectPrep.PrepImageData dstSA, toPreview, dstPic
            
            'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
            ' (This is necessary to prevent processed pixel values from spreading across the image as we go.)
            Dim srcDIB As pdDIB
            Set srcDIB = New pdDIB
            srcDIB.CreateFromExistingDIB workingDIB
                
            'Use the central ConvolveDIB function to apply the first convolution to workingDIB
            ConvolveDIB_XML cParamsOut.GetParamString(), srcDIB, workingDIB, toPreview, srcDIB.GetDIBWidth * 2
            
            'Now we need a third copy of the image, which will receive the alternate direction transform
            Dim secondDstDIB As pdDIB
            Set secondDstDIB = New pdDIB
            secondDstDIB.CreateFromExistingDIB srcDIB
            
            'When two passes are required, the vertical direction is always applied first.  Thus we know we need to apply the
            ' horizontal direction next.  Generate a new param string for the horizontal direction.
            If (Not toPreview) Then Message "Applying pass %1 of %2 for %3 filter...", "2", "2", GetNameOfEdgeDetector(edgeDetectionType)
            
            cParamsOut.Reset
            
            With cParamsOut
                .AddParam "name", GetNameOfEdgeDetector(edgeDetectionType)
                .AddParam "invert", Not blackBackground
                GetParamStringForEdgeDetector edgeDetectionType, pded_Horizontal, fWeight, fBias, fMatrix
                .AddParam "weight", fWeight
                .AddParam "bias", fBias
                .AddParam "matrix", fMatrix
            End With
            
            'Use the central ConvolveDIB function to apply the new convolution to workingDIB
            ConvolveDIB_XML cParamsOut.GetParamString(), srcDIB, secondDstDIB, toPreview, srcDIB.GetDIBWidth * 2, srcDIB.GetDIBWidth
            
            'Free our temporary source DIB
            Set srcDIB = Nothing
            
            'The compositor requires premultiplied alpha, so convert both top and bottom layers now
            workingDIB.SetAlphaPremultiplication True
            secondDstDIB.SetAlphaPremultiplication True
            
            'Last step is to blend the two result arrays together.  Use the pdCompositor class to do this.
            Dim cComposite As pdCompositor
            Set cComposite = New pdCompositor
            
            If blackBackground Then
                cComposite.QuickMergeTwoDibsOfEqualSize workingDIB, secondDstDIB, BM_Screen
            Else
                cComposite.QuickMergeTwoDibsOfEqualSize workingDIB, secondDstDIB, BM_Multiply
            End If
            
            'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingDIB
            EffectPrep.FinalizeImageData toPreview, dstPic, True
            
        End If
        
    End If

End Sub

'Return the name of an edge detection type as a human-readable string
Private Function GetNameOfEdgeDetector(ByVal edgeDetectionType As PD_EdgeDetector) As String

    Select Case edgeDetectionType
                
        Case pded_Hilite
            GetNameOfEdgeDetector = g_Language.TranslateMessage("Hilite edge detection")
            
        Case pded_Laplacian
            GetNameOfEdgeDetector = g_Language.TranslateMessage("Laplacian edge detection")
        
        Case pded_PhotoDemon
            GetNameOfEdgeDetector = g_Language.TranslateMessage("PhotoDemon edge detection")
            
        Case pded_Prewitt
            GetNameOfEdgeDetector = g_Language.TranslateMessage("Prewitt edge detection")
        
        Case pded_Roberts
            GetNameOfEdgeDetector = g_Language.TranslateMessage("Roberts cross edge detection")
            
        Case pded_Sobel
            GetNameOfEdgeDetector = g_Language.TranslateMessage("Sobel edge detection")
            
    End Select

End Function

'Given an edge detection type and a direction, return TRUE if the requested edge detector can be applied in a single pass.
' Return FALSE if the function requires multiple image passes.
Private Function IsEdgeDetectionSinglePass(ByVal edgeDetectionType As PD_EdgeDetector, Optional ByVal edgeDirectionality As PD_EdgeDirection = pded_All) As Boolean

    'Convolution matrix strings are assembled in two or three steps:
    ' 1) Add divisor and offset values
    ' 2 (optional) ) Check directionality and adjust behavior accordingly
    ' 3) Build actual convolution matrix
    Select Case edgeDetectionType
        
        Case pded_Contour
            IsEdgeDetectionSinglePass = True
    
        'Hilite detection (doesn't support directionality)
        Case pded_Hilite
            IsEdgeDetectionSinglePass = True
        
        'Laplacian is unique because it supports a different operator for all directionalities, so even horizontal/vertical can
        ' be done in a single pass.
        Case pded_Laplacian
            IsEdgeDetectionSinglePass = True
                
        'PhotoDemon edge detection (doesn't support directionality)
        Case pded_PhotoDemon
            IsEdgeDetectionSinglePass = True
        
        'Prewitt edge detection is unidirectional
        Case pded_Prewitt
            If (edgeDirectionality = pded_Horizontal) Or (edgeDirectionality = pded_Vertical) Then
                IsEdgeDetectionSinglePass = True
            Else
                IsEdgeDetectionSinglePass = False
            End If
            
        'Roberts cross edge detection is unidirectional
        Case pded_Roberts
            If (edgeDirectionality = pded_Horizontal) Or (edgeDirectionality = pded_Vertical) Then
                IsEdgeDetectionSinglePass = True
            Else
                IsEdgeDetectionSinglePass = False
            End If
        
        'Sobel edge detection is unidirectional
        Case pded_Sobel
            If (edgeDirectionality = pded_Horizontal) Or (edgeDirectionality = pded_Vertical) Then
                IsEdgeDetectionSinglePass = True
            Else
                IsEdgeDetectionSinglePass = False
            End If
        
    End Select

End Function

'Given an internal edge detection type (and optionally, a direction), calculate a matching convolution matrix and return it
Private Function GetParamStringForEdgeDetector(ByVal edgeDetectionType As PD_EdgeDetector, ByVal edgeDirectionality As PD_EdgeDirection, ByRef fWeight As Double, ByRef fBias As Double, ByRef fMatrix As String) As String

    Dim convoString As String
    
    'Convolution matrix strings are assembled in two or three steps:
    ' 1) Add divisor and offset values
    ' 2 (optional) ) Check directionality and adjust behavior accordingly
    ' 3) Build actual convolution matrix
    Select Case edgeDetectionType
    
        'Hilite detection (doesn't support directionality)
        Case pded_Hilite
            
            'Divisor/offset
            fWeight = 1#: fBias = 0#
    
            'Actual convo matrix
            convoString = convoString & "0|0|0|0|0|"
            convoString = convoString & "0|-4|-2|-1|0|"
            convoString = convoString & "0|-2|10|0|0|"
            convoString = convoString & "0|-1|0|0|0|"
            convoString = convoString & "0|0|0|0|0"
        
        Case pded_Laplacian
            
            'Actual convo matrix varies according to direction
            If edgeDirectionality = pded_Horizontal Then
            
                'Divisor/offset
                fWeight = 0.25: fBias = 0#
                
                convoString = convoString & "0|0|0|0|0|"
                convoString = convoString & "0|0|0|0|0|"
                convoString = convoString & "0|-1|2|-1|0|"
                convoString = convoString & "0|0|0|0|0|"
                convoString = convoString & "0|0|0|0|0"
                
            ElseIf edgeDirectionality = pded_Vertical Then
            
                'Divisor/offset
                fWeight = 0.25: fBias = 0#
                
                convoString = convoString & "0|0|0|0|0|"
                convoString = convoString & "0|0|-1|0|0|"
                convoString = convoString & "0|0|2|0|0|"
                convoString = convoString & "0|0|-1|0|0|"
                convoString = convoString & "0|0|0|0|0"
                
            Else
            
                'Divisor/offset
                fWeight = 0.5: fBias = 0#
                
                convoString = convoString & "0|0|0|0|0|"
                convoString = convoString & "0|0|-1|0|0|"
                convoString = convoString & "0|-1|4|-1|0|"
                convoString = convoString & "0|0|-1|0|0|"
                convoString = convoString & "0|0|0|0|0"
            
            End If
        
        'PhotoDemon edge detection (doesn't support directionality)
        Case pded_PhotoDemon
        
            'Divisor/offset
            fWeight = 1#: fBias = 0#
            
            'Actual convo matrix
            convoString = convoString & "0|-1|0|0|0|"
            convoString = convoString & "0|0|0|0|-1|"
            convoString = convoString & "0|0|4|0|0|"
            convoString = convoString & "-1|0|0|0|0|"
            convoString = convoString & "0|0|0|-1|0"
        
        'Prewitt edge detection (directionality supported)
        Case pded_Prewitt
        
            'Divisor/offset
            fWeight = 1#: fBias = 0#
            
            'Actual convo matrix varies according to direction
            If edgeDirectionality = pded_Horizontal Then
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
        Case pded_Roberts
        
            'Divisor/offset
            fWeight = 0.5: fBias = 0#
            
            'Actual convo matrix varies according to direction
            If edgeDirectionality = pded_Horizontal Then
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
        Case pded_Sobel
            
            'Divisor/offset
            fWeight = 1#: fBias = 0#
            
            'Actual convo matrix varies according to direction
            If edgeDirectionality = pded_Horizontal Then
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
    
    fMatrix = convoString

End Function

'This code is a modified version of an algorithm originally developed by Manuel Augusto Santos.  A link to his original
' implementation is available from the "Help -> About PhotoDemon" menu option.
Private Sub FilterSmoothContour(Optional ByVal blackBackground As Boolean = False, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)

    If (Not toPreview) Then Message "Tracing image edges with virtual paintbrush..."
        
    'Create a local array and point it at the pixel data of the current image
    Dim dstSA As SafeArray2D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent blurred pixel values from spreading across the image as we go.)
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.CreateFromExistingDIB workingDIB
    
    CreateContourDIB blackBackground, srcDIB, workingDIB, toPreview
    
    Set srcDIB = Nothing
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingDIB
    EffectPrep.FinalizeImageData toPreview, dstPic
    
End Sub

Private Sub Form_Load()
    
    'Suspend previews until the list box has been populated
    cmdBar.SetPreviewStatus False
    
    'Generate a list box with all the currently implemented edge detection algorithms
    lstEdgeOptions.SetAutomaticRedraws False
    lstEdgeOptions.AddItem "Artistic contour", 0
    lstEdgeOptions.AddItem "Hilite", 1
    lstEdgeOptions.AddItem "Laplacian", 2
    lstEdgeOptions.AddItem "PhotoDemon", 3
    lstEdgeOptions.AddItem "Prewitt", 4
    lstEdgeOptions.AddItem "Roberts cross", 5
    lstEdgeOptions.AddItem "Sobel", 6
    lstEdgeOptions.ListIndex = 0
    lstEdgeOptions.SetAutomaticRedraws True, True
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub LstEdgeOptions_Click()
    
    cmdBar.SetPreviewStatus False
    
    'Directionality is only supported by some transforms, so de/activate the directionality check boxes to match the
    ' capabilities of the selected transform
    Select Case lstEdgeOptions.ListIndex
    
        Case pded_Contour
            ChangeCheckboxActivation False
        
        Case pded_Hilite
            ChangeCheckboxActivation False
        
        Case pded_Laplacian
            ChangeCheckboxActivation True
        
        Case pded_PhotoDemon
            ChangeCheckboxActivation False
        
        Case pded_Prewitt
            ChangeCheckboxActivation True
        
        Case pded_Roberts
            ChangeCheckboxActivation True
            
        Case pded_Sobel
            ChangeCheckboxActivation True
    
    End Select
    
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

'Dis/enable the directionality checkboxes to match the request; when checkboxes are disabled, their value is automatically
' forced to TRUE.
Private Sub ChangeCheckboxActivation(ByVal toEnable As Boolean)

    If toEnable Then
    
        chkDirection(0).Enabled = True
        chkDirection(1).Enabled = True
    
    'Activate both directions, then disable the checkboxes
    Else
    
        If (Not chkDirection(0)) Then chkDirection(0).Value = True
        If (Not chkDirection(1)) Then chkDirection(1).Value = True
        
        chkDirection(0).Enabled = False
        chkDirection(1).Enabled = False
    
    End If
    
End Sub

'Convert the directionality checkboxes to PD's internal edge detection definitions
Private Function GetDirectionality() As PD_EdgeDirection

    If chkDirection(0).Value And Not chkDirection(1).Value Then
        GetDirectionality = pded_Horizontal
    ElseIf chkDirection(1).Value And Not chkDirection(0).Value Then
        GetDirectionality = pded_Vertical
    Else
        GetDirectionality = pded_All
    End If

End Function

'Update the live preview of the selected edge detection options
Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then ApplyEdgeDetection GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "method", lstEdgeOptions.ListIndex
        .AddParam "direction", GetDirectionality()
        .AddParam "invert", chkInvert.Value
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
