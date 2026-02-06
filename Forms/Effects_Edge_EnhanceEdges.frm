VERSION 5.00
Begin VB.Form FormEdgeEnhance 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Enhance edges"
   ClientHeight    =   6525
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11775
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
   ScaleWidth      =   785
   Visible         =   0   'False
   Begin PhotoDemon.pdListBox lstEdgeOptions 
      Height          =   2775
      Left            =   6000
      TabIndex        =   5
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   4895
      Caption         =   "edge detection technique"
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5775
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   1323
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
      TabIndex        =   4
      Top             =   3360
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   635
      Caption         =   "horizontal"
   End
   Begin PhotoDemon.pdCheckBox chkDirection 
      Height          =   360
      Index           =   1
      Left            =   6240
      TabIndex        =   1
      Top             =   3840
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   635
      Caption         =   "vertical"
   End
   Begin PhotoDemon.pdSlider sltStrength 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   4560
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   1244
      Caption         =   "strength"
      Max             =   100
      Value           =   50
      NotchPosition   =   2
      NotchValueCustom=   50
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   1
      Left            =   6000
      Top             =   3000
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   503
      Caption         =   "detection direction(s)"
      FontSize        =   12
      ForeColor       =   4210752
   End
End
Attribute VB_Name = "FormEdgeEnhance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Edge Enhancement Interface
'Copyright 2002-2026 by Tanner Helland
'Created: sometimes 2002
'Last updated: 29/July/17
'Last update: performance improvements, migrate to XML params
'
'This edge enhancement function allows the user to selectively emphasize image edges using any available edge
' detection technique.  PD's compositor is then used to composite the results back onto the base image at some
' variable strength specified by the user.
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

'OK button
Private Sub cmdBar_OKClick()
    Process "Enhance edges", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

'Apply any supported edge detection filter to an image.  Directionality can be specified, but note that only some
' algorithms support the parameter.
Public Sub ApplyEdgeEnhancement(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)

    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim edgeDetectionType As PD_EdgeDetector, edgeDirectionality As PD_EdgeDirection
    Dim enhanceStrength As Double
    
    With cParams
        edgeDetectionType = .GetLong("method", lstEdgeOptions.ListIndex)
        edgeDirectionality = .GetLong("direction", pded_All)
        enhanceStrength = .GetDouble("strength", sltStrength.Value)
    End With
    
    'Applying an edge detection filter generally happens via these steps:
    
    '1) Set up any parameters we know in advance, like generating a String name for the supplied filter, and converting
    '    the optional "blackBackground" parameter into PD's internal ParamString format
    '2) Retrieve a relevant convolution matrix for the requested filter
    '3) Supply the full ParamString, including convo matrix, to PD's central ApplyConvolutionFilter function
    '4) If necessary, repeat steps 2 and 3 to retrieve multiple directionality images
    
    'Because some of these parameters are handled separately, we now need to build a special parameter string
    ' for just the convolver.
    Dim cParamsOut As pdSerialize
    Set cParamsOut = New pdSerialize
    
    With cParamsOut
        .AddParam "name", GetNameOfEdgeDetector(edgeDetectionType)
        .AddParam "invert", False
        
        'We now need to calculate per-algorithm values using a separate helper function
        Dim fWeight As Double, fBias As Double, fMatrix As String
        GetParamStringForEdgeDetector edgeDetectionType, edgeDirectionality, fWeight, fBias, fMatrix
        .AddParam "weight", fWeight
        .AddParam "bias", fBias
        .AddParam "matrix", fMatrix
    End With
    
    'Next, we need to obtain a DIB of the processed edge detection results for the image.  This requires two or
    ' three passes, contingent on the detection type.  In order to update the progress bar correctly, calculate the number
    ' of passes required in advance.
    Dim numPassesRequired As Long
    If (edgeDetectionType = pded_Contour) Then
        numPassesRequired = 2
    ElseIf IsEdgeDetectionSinglePass(edgeDetectionType, edgeDirectionality) Then
        numPassesRequired = 2
    Else
        numPassesRequired = 3
    End If
    
    If (Not toPreview) Then Message "Applying pass %1 of %2 for %3 filter...", "1", numPassesRequired, GetNameOfEdgeDetector(edgeDetectionType)
    
    'Use PD's central image handler to populate the public workingDIB object.
    Dim dstSA As SafeArray2D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic
    
    'Create a second DIB copy.  This will receive the edge-detection copy of the image.
    Dim edgeDIB As pdDIB
    Set edgeDIB = New pdDIB
    edgeDIB.CreateFromExistingDIB workingDIB
    
    '3a) If the function is single-pass compatible (e.g. it does not require us to traverse the image multiple times, then
    '     blend the edge detection results), supply the compiled param string to PD's central convolution function and exit
    If (edgeDetectionType = pded_Contour) Then
        Filters_Layers.CreateContourDIB True, workingDIB, edgeDIB, toPreview, workingDIB.GetDIBWidth * numPassesRequired, 0
    Else
        Filters_Area.ConvolveDIB_XML cParamsOut.GetParamString(), workingDIB, edgeDIB, toPreview, workingDIB.GetDIBWidth * numPassesRequired, 0
    End If
    
    'A pdCompositor class is required to selectively blend the edge detection results back onto the main image
    Dim cComposite As pdCompositor
    Set cComposite = New pdCompositor
    
    '3b) If the requested edge function is not single-pass compatible, run a second pass in the opposite direction,
    '     the blend the results back onto edgeDIB.
    If (Not IsEdgeDetectionSinglePass(edgeDetectionType, edgeDirectionality)) Then
    
        'Create a second DIB copy.  This will receive the edge-detection copy of the image.
        Dim tmpEdgeDIB As pdDIB
        Set tmpEdgeDIB = New pdDIB
        tmpEdgeDIB.CreateFromExistingDIB workingDIB
        
        'When two passes are required, the vertical direction is always applied first.  Thus we know we need to apply the
        ' horizontal direction next.  Generate a new param string for the horizontal direction.
        If (Not toPreview) Then Message "Applying pass %1 of %2 for %3 filter...", "2", numPassesRequired, GetNameOfEdgeDetector(edgeDetectionType)
        
        cParamsOut.Reset
        
        With cParamsOut
            .AddParam "name", GetNameOfEdgeDetector(edgeDetectionType)
            .AddParam "invert", False
            GetParamStringForEdgeDetector edgeDetectionType, pded_Horizontal, fWeight, fBias, fMatrix
            .AddParam "weight", fWeight
            .AddParam "bias", fBias
            .AddParam "matrix", fMatrix
        End With
        
        'Use the central ConvolveDIB function to apply the new convolution to workingDIB
        Filters_Area.ConvolveDIB_XML cParamsOut.GetParamString(), workingDIB, tmpEdgeDIB, toPreview, workingDIB.GetDIBWidth * numPassesRequired, workingDIB.GetDIBWidth
        
        'The compositor requires premultiplied alpha, so convert both top and bottom layers now
        edgeDIB.SetAlphaPremultiplication True
        tmpEdgeDIB.SetAlphaPremultiplication True
        
        'Use the pdCompositor class to blend the results of the second edge detection pass with the first pass.
        cComposite.QuickMergeTwoDibsOfEqualSize edgeDIB, tmpEdgeDIB, BM_Screen
        
        'Remove premultiplication
        edgeDIB.SetAlphaPremultiplication False
        
        Set tmpEdgeDIB = Nothing
        
    End If
    
    'edgeDIB now contains a complete edge-detection copy of the image, using the supplied edge detector algorithm.
    ' We now need to selectively blend the results back onto the main image, at the strength requested.
    If (Not toPreview) Then Message "Applying pass %1 of %2 for %3 filter...", numPassesRequired, numPassesRequired, GetNameOfEdgeDetector(edgeDetectionType)
    
    'Apply premultiplication prior to compositing
    edgeDIB.SetAlphaPremultiplication True
    workingDIB.SetAlphaPremultiplication True
    
    'Merge the two DIBs together
    cComposite.QuickMergeTwoDibsOfEqualSize workingDIB, edgeDIB, BM_Screen, enhanceStrength
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingDIB
    EffectPrep.FinalizeImageData toPreview, dstPic, True

End Sub

'Return the naem of an edge detection type as a human-readable string
Private Function GetNameOfEdgeDetector(ByVal edgeDetectionType As PD_EdgeDetector) As String

    Select Case edgeDetectionType
        
        Case pded_Contour
            GetNameOfEdgeDetector = g_Language.TranslateMessage("Artistic contour edge detection")
            
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
Private Sub GetParamStringForEdgeDetector(ByVal edgeDetectionType As PD_EdgeDetector, ByVal edgeDirectionality As PD_EdgeDirection, ByRef fWeight As Double, ByRef fBias As Double, ByRef fMatrix As String)

    Dim convoString As String
    
    'Convolution matrix strings are assembled in two or three steps:
    ' 1) Add divisor and offset values
    ' 2 (optional) Check directionality and adjust behavior accordingly
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
    
        If Not chkDirection(0).Value Then chkDirection(0).Value = True
        If Not chkDirection(1).Value Then chkDirection(1).Value = True
        
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
    If cmdBar.PreviewsAllowed Then ApplyEdgeEnhancement GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Sub sltStrength_Change()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "method", lstEdgeOptions.ListIndex
        .AddParam "direction", GetDirectionality()
        .AddParam "strength", sltStrength.Value
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
