VERSION 5.00
Begin VB.Form FormFindEdges 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Find Edges"
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
   Begin PhotoDemon.smartCheckBox chkInvert 
      Height          =   480
      Left            =   9120
      TabIndex        =   5
      Top             =   3360
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   847
      Caption         =   "use black background"
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
      Height          =   2460
      Left            =   6000
      TabIndex        =   1
      Top             =   1320
      Width           =   2895
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "description:"
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
      Height          =   375
      Left            =   9120
      TabIndex        =   3
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label LblDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "(no item selected)"
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
      Height          =   1575
      Left            =   9120
      TabIndex        =   2
      Top             =   1800
      Width           =   2895
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "FormFindEdges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Edge Detection Interface
'Copyright ©2000-2014 by Tanner Helland
'Created: 1/11/02
'Last updated: 22/August/13
'Last update: rewrote all ApplyConvolutionFilter calls with paramStrings
'
'All known edge-detection routines are handled from this form.  Most are simply convolution kernels that are passed off
' to the "ApplyConvolutionFilter" function, but at least one (Artistic Contour) resides here.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

Private Sub chkInvert_Click()
    updatePreview
End Sub

'OK button
Private Sub cmdBar_OKClick()

    Select Case LstEdgeOptions.ListIndex
        Case 0
            Process "Find edges (Prewitt horizontal)", , buildParams(CBool(chkInvert.Value)), UNDO_LAYER
        Case 1
            Process "Find edges (Prewitt vertical)", , buildParams(CBool(chkInvert.Value)), UNDO_LAYER
        Case 2
            Process "Find edges (Sobel horizontal)", , buildParams(CBool(chkInvert.Value)), UNDO_LAYER
        Case 3
            Process "Find edges (Sobel vertical)", , buildParams(CBool(chkInvert.Value)), UNDO_LAYER
        Case 4
            Process "Find edges (Laplacian)", , buildParams(CBool(chkInvert.Value)), UNDO_LAYER
        Case 5
            Process "Artistic contour", , buildParams(CBool(chkInvert.Value)), UNDO_LAYER
        Case 6
            Process "Find edges (Hilite)", , buildParams(CBool(chkInvert.Value)), UNDO_LAYER
        Case 7
            Process "Find edges (PhotoDemon linear)", , buildParams(CBool(chkInvert.Value)), UNDO_LAYER
        Case 8
            Process "Find edges (PhotoDemon cubic)", , buildParams(CBool(chkInvert.Value)), UNDO_LAYER
    End Select
    
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    LstEdgeOptions.ListIndex = 5
End Sub

Private Sub Form_Activate()
        
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
    'Update the descriptions (this will also draw a preview of the selected edge-detection algorithm)
    cmdBar.markPreviewStatus True
    updatePreview
    
End Sub

Public Sub FilterHilite(Optional ByVal blackBackground As Boolean = False, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    Dim tmpString As String
    
    'Start with a filter name
    tmpString = g_Language.TranslateMessage("Hilite edge detection") & "|"
    
    'Next comes an invert parameter
    tmpString = tmpString & Str(Not blackBackground) & "|"
    
    'Next is the divisor and offset
    tmpString = tmpString & "1|0|"
    
    'And finally, the convolution array itself
    tmpString = tmpString & "0|0|0|0|0|"
    tmpString = tmpString & "0|-4|-2|-1|0|"
    tmpString = tmpString & "0|-2|10|0|0|"
    tmpString = tmpString & "0|-1|0|0|0|"
    tmpString = tmpString & "0|0|0|0|0"
    
    'Pass our new parameter string to the main convolution filter function
    ApplyConvolutionFilter tmpString, toPreview, dstPic
    
End Sub

Public Sub PhotoDemonCubicEdgeDetection(Optional ByVal blackBackground As Boolean = False, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)

    Dim tmpString As String
    
    'Start with a filter name
    tmpString = g_Language.TranslateMessage("PhotoDemon cubic edge detection") & "|"
    
    'Next comes an invert parameter
    tmpString = tmpString & Str(Not blackBackground) & "|"
    
    'Next is the divisor and offset
    tmpString = tmpString & "1|0|"
    
    'And finally, the convolution array itself
    tmpString = tmpString & "0|1|0|0|0|"
    tmpString = tmpString & "0|0|0|0|1|"
    tmpString = tmpString & "0|0|-4|0|0|"
    tmpString = tmpString & "1|0|0|0|0|"
    tmpString = tmpString & "0|0|0|1|0"
    
    'Pass our new parameter string to the main convolution filter function
    ApplyConvolutionFilter tmpString, toPreview, dstPic
    
End Sub

Public Sub PhotoDemonLinearEdgeDetection(Optional ByVal blackBackground As Boolean = False, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    Dim tmpString As String
    
    'Start with a filter name
    tmpString = g_Language.TranslateMessage("PhotoDemon linear edge detection") & "|"
    
    'Next comes an invert parameter
    tmpString = tmpString & Str(Not blackBackground) & "|"
    
    'Next is the divisor and offset
    tmpString = tmpString & "1|0|"
    
    'And finally, the convolution array itself
    tmpString = tmpString & "0|0|0|0|0|"
    tmpString = tmpString & "0|-1|0|-1|0|"
    tmpString = tmpString & "0|0|4|0|0|"
    tmpString = tmpString & "0|-1|0|-1|0|"
    tmpString = tmpString & "0|0|0|0|0"
    
    'Pass our new parameter string to the main convolution filter function
    ApplyConvolutionFilter tmpString, toPreview, dstPic
    
End Sub

Public Sub FilterPrewittHorizontal(Optional ByVal blackBackground As Boolean = False, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)

    Dim tmpString As String
    
    'Start with a filter name
    tmpString = g_Language.TranslateMessage("Prewitt horizontal edge detection") & "|"
    
    'Next comes an invert parameter
    tmpString = tmpString & Str(Not blackBackground) & "|"
    
    'Next is the divisor and offset
    tmpString = tmpString & "1|0|"
    
    'And finally, the convolution array itself
    tmpString = tmpString & "0|0|0|0|0|"
    tmpString = tmpString & "0|-1|0|1|0|"
    tmpString = tmpString & "0|-1|0|1|0|"
    tmpString = tmpString & "0|-1|0|1|0|"
    tmpString = tmpString & "0|0|0|0|0"
    
    'Pass our new parameter string to the main convolution filter function
    ApplyConvolutionFilter tmpString, toPreview, dstPic
    
End Sub

Public Sub FilterPrewittVertical(Optional ByVal blackBackground As Boolean = False, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)

    Dim tmpString As String
    
    'Start with a filter name
    tmpString = g_Language.TranslateMessage("Prewitt vertical edge detection") & "|"
    
    'Next comes an invert parameter
    tmpString = tmpString & Str(Not blackBackground) & "|"
    
    'Next is the divisor and offset
    tmpString = tmpString & "1|0|"
    
    'And finally, the convolution array itself
    tmpString = tmpString & "0|0|0|0|0|"
    tmpString = tmpString & "0|1|1|1|0|"
    tmpString = tmpString & "0|0|0|0|0|"
    tmpString = tmpString & "0|-1|-1|-1|0|"
    tmpString = tmpString & "0|0|0|0|0"
    
    'Pass our new parameter string to the main convolution filter function
    ApplyConvolutionFilter tmpString, toPreview, dstPic
    
End Sub

Public Sub FilterSobelHorizontal(Optional ByVal blackBackground As Boolean = False, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)

    Dim tmpString As String
    
    'Start with a filter name
    tmpString = g_Language.TranslateMessage("Sobel horizontal edge detection") & "|"
    
    'Next comes an invert parameter
    tmpString = tmpString & Str(Not blackBackground) & "|"
    
    'Next is the divisor and offset
    tmpString = tmpString & "1|0|"
    
    'And finally, the convolution array itself
    tmpString = tmpString & "0|0|0|0|0|"
    tmpString = tmpString & "0|-1|0|1|0|"
    tmpString = tmpString & "0|-2|0|2|0|"
    tmpString = tmpString & "0|-1|0|1|0|"
    tmpString = tmpString & "0|0|0|0|0"
    
    'Pass our new parameter string to the main convolution filter function
    ApplyConvolutionFilter tmpString, toPreview, dstPic
    
End Sub

Public Sub FilterSobelVertical(Optional ByVal blackBackground As Boolean = False, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)

    Dim tmpString As String
    
    'Start with a filter name
    tmpString = g_Language.TranslateMessage("Sobel vertical edge detection") & "|"
    
    'Next comes an invert parameter
    tmpString = tmpString & Str(Not blackBackground) & "|"
    
    'Next is the divisor and offset
    tmpString = tmpString & "1|0|"
    
    'And finally, the convolution array itself
    tmpString = tmpString & "0|0|0|0|0|"
    tmpString = tmpString & "0|1|2|1|0|"
    tmpString = tmpString & "0|0|0|0|0|"
    tmpString = tmpString & "0|-1|-2|-1|0|"
    tmpString = tmpString & "0|0|0|0|0"
    
    'Pass our new parameter string to the main convolution filter function
    ApplyConvolutionFilter tmpString, toPreview, dstPic
    
End Sub

Public Sub FilterLaplacian(Optional ByVal blackBackground As Boolean = False, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    Dim tmpString As String
    
    'Start with a filter name
    tmpString = g_Language.TranslateMessage("Laplacian edge detection") & "|"
    
    'Next comes an invert parameter
    tmpString = tmpString & Str(Not blackBackground) & "|"
    
    'Next is the divisor and offset
    tmpString = tmpString & "1|0|"
    
    'And finally, the convolution array itself
    tmpString = tmpString & "0|0|0|0|0|"
    tmpString = tmpString & "0|0|-1|0|0|"
    tmpString = tmpString & "0|-1|4|-1|0|"
    tmpString = tmpString & "0|0|-1|0|0|"
    tmpString = tmpString & "0|0|0|0|0"
    
    'Pass our new parameter string to the main convolution filter function
    ApplyConvolutionFilter tmpString, toPreview, dstPic
    
End Sub

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
    LstEdgeOptions.AddItem "Prewitt Horizontal"
    LstEdgeOptions.AddItem "Prewitt Vertical"
    LstEdgeOptions.AddItem "Sobel Horizontal"
    LstEdgeOptions.AddItem "Sobel Vertical"
    LstEdgeOptions.AddItem "Laplacian"
    LstEdgeOptions.AddItem "Artistic Contour"
    LstEdgeOptions.AddItem "Hilite"
    LstEdgeOptions.AddItem "PhotoDemon Linear"
    LstEdgeOptions.AddItem "PhotoDemon Cubic"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub LstEdgeOptions_Click()
    updatePreview
End Sub

'Show the user a brief explanation of the algorithm in question.  Yes, the PhotoDemon routine descriptions are bullshit -
' I know that already.  :)  But the descriptions make them sound more impressive than they actually are.
' This sub also handles redrawing the edge detection preview.
Private Sub updatePreview()
    
    If cmdBar.previewsAllowed Then
    
        Select Case LstEdgeOptions.ListIndex
        
            Case 0
                LblDesc = g_Language.TranslateMessage("Simple matrix method:") & vbCrLf & vbCrLf & "-1 0 1" & vbCrLf & "-1 0 1" & vbCrLf & "-1 0 1"
                FilterPrewittHorizontal CBool(chkInvert.Value), True, fxPreview
            Case 1
                LblDesc = g_Language.TranslateMessage("Simple matrix method:") & vbCrLf & vbCrLf & "-1 -1 -1" & vbCrLf & " 0  0  0" & vbCrLf & " 1  1  1"
                FilterPrewittVertical CBool(chkInvert.Value), True, fxPreview
            Case 2
                LblDesc = g_Language.TranslateMessage("Simple matrix method:") & vbCrLf & vbCrLf & "-1 0 1" & vbCrLf & "-2 0 2" & vbCrLf & "-1 0 1"
                FilterSobelHorizontal CBool(chkInvert.Value), True, fxPreview
            Case 3
                LblDesc = g_Language.TranslateMessage("Simple matrix method:") & vbCrLf & vbCrLf & "-1 -2 -1" & vbCrLf & " 0  0  0" & vbCrLf & " 1  2  1"
                FilterSobelVertical CBool(chkInvert.Value), True, fxPreview
            Case 4
                LblDesc = g_Language.TranslateMessage("Simple matrix method:") & vbCrLf & vbCrLf & " 0 -1  0" & vbCrLf & "-1  4 -1" & vbCrLf & " 0 -1  0"
                FilterLaplacian CBool(chkInvert.Value), True, fxPreview
            Case 5
                LblDesc = g_Language.TranslateMessage("Algorithm designed to present a clean, artistic prediction of image edges.")
                FilterSmoothContour CBool(chkInvert.Value), True, fxPreview
            Case 6
                LblDesc = g_Language.TranslateMessage("Simple matrix method:") & vbCrLf & vbCrLf & "-4 -2 -1" & vbCrLf & "-2 10  0" & vbCrLf & "-1  0  0"
                FilterHilite CBool(chkInvert.Value), True, fxPreview
            Case 7
                LblDesc = g_Language.TranslateMessage("Simple mathematical routine based on linear relationships between diagonal pixels.")
                PhotoDemonLinearEdgeDetection CBool(chkInvert.Value), True, fxPreview
            Case 8
                LblDesc = g_Language.TranslateMessage("Advanced mathematical routine based on cubic relationships between diagonal pixels.")
                PhotoDemonCubicEdgeDetection CBool(chkInvert.Value), True, fxPreview
        
        End Select
        
    End If
    
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub


