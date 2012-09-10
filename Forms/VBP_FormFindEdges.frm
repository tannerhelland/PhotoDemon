VERSION 5.00
Begin VB.Form FormFindEdges 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Find Edges"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6270
   BeginProperty Font 
      Name            =   "Arial"
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
   ScaleHeight     =   426
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   418
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.PictureBox picPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2730
      Left            =   120
      ScaleHeight     =   180
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   191
      TabIndex        =   6
      Top             =   120
      Width           =   2895
   End
   Begin VB.PictureBox picEffect 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2730
      Left            =   3240
      ScaleHeight     =   180
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   191
      TabIndex        =   5
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3795
      TabIndex        =   1
      Top             =   5880
      Width           =   1125
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4995
      TabIndex        =   2
      Top             =   5880
      Width           =   1125
   End
   Begin VB.ListBox LstEdgeOptions 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1950
      Left            =   120
      TabIndex        =   0
      Top             =   3360
      Width           =   2895
   End
   Begin VB.Label lblBeforeandAfter 
      BackStyle       =   0  'Transparent
      Caption         =   "  Before                                                           After"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label LblDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "<No Item Selected>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   1575
      Left            =   3240
      TabIndex        =   3
      Top             =   3720
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
'Copyright ©2000-2012 by Tanner Helland
'Created: 1/11/02
'Last updated: 09/September/12
'Last update: added previewing!  Also, rewrote all functions against new layer code.
'
'All known edge-detection routines are handled from this form.  Most are simply convolution kernels that are passed off
' to the "DoFilter" function, but at least one (Artistic Contour) resides here.
'
'***************************************************************************

Option Explicit

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub CmdOK_Click()

    Me.Visible = False
    
    Select Case LstEdgeOptions.ListIndex
        Case 0
            Process PrewittHorizontal
        Case 1
            Process PrewittVertical
        Case 2
            Process SobelHorizontal
        Case 3
            Process SobelVertical
        Case 4
            Process Laplacian
        Case 5
            Process SmoothContour
        Case 6
            Process HiliteEdge
        Case 7
            Process PhotoDemonEdgeLinear
        Case 8
            Process PhotoDemonEdgeCubic
    End Select
    
    Unload Me

End Sub

'LOAD form
Private Sub Form_Load()
    
    'Draw a preview of the current image to the left box
    DrawPreviewImage picPreview
    
    'Generate a list box with all the various edge detection algorithms
    LstEdgeOptions.AddItem "Prewitt Horizontal"
    LstEdgeOptions.AddItem "Prewitt Vertical"
    LstEdgeOptions.AddItem "Sobel Horizontal"
    LstEdgeOptions.AddItem "Sobel Vertical"
    LstEdgeOptions.AddItem "Laplacian"
    LstEdgeOptions.AddItem "Artistic Contour"
    LstEdgeOptions.AddItem "Hilite"
    LstEdgeOptions.AddItem "PhotoDemon Linear"
    LstEdgeOptions.AddItem "PhotoDemon Cubic"
    
    LstEdgeOptions.ListIndex = 5
    
    'Update the descriptions (this will also draw a preview of the selected edge-detection algorithm)
    UpdateDescriptions
    
    'Assign the system hand cursor to all relevant objects
    setHandCursorForAll Me
    
End Sub

Public Sub FilterHilite(Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As PictureBox)
    FilterSize = 3
    ReDim FM(-1 To 1, -1 To 1) As Long
    FM(-1, -1) = -4
    FM(-1, 0) = -2
    FM(0, -1) = -2
    FM(1, -1) = -1
    FM(-1, 1) = -1
    FM(0, 0) = 10
    FilterWeight = 1
    FilterBias = 0
    DoFilter "Hilite edge detection", True, , toPreview, dstPic
End Sub

Public Sub PhotoDemonCubicEdgeDetection(Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As PictureBox)
    FilterSize = 5
    ReDim FM(-2 To 2, -2 To 2) As Long
    FM(-1, -2) = 1
    FM(-2, 1) = 1
    FM(1, 2) = 1
    FM(2, -1) = 1
    FM(0, 0) = -4
    FilterWeight = 1
    FilterBias = 0
    DoFilter "PhotoDemon cubic edge detection", True, , toPreview, dstPic
End Sub

Public Sub PhotoDemonLinearEdgeDetection(Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As PictureBox)
    FilterSize = 3
    ReDim FM(-1 To 1, -1 To 1) As Long
    FM(-1, -1) = -1
    FM(-1, 1) = -1
    FM(1, -1) = -1
    FM(1, 1) = -1
    FM(0, 0) = 4
    FilterWeight = 1
    FilterBias = 0
    DoFilter "PhotoDemon linear edge detection", True, , toPreview, dstPic
End Sub

Public Sub FilterPrewittHorizontal(Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As PictureBox)
    FilterSize = 3
    ReDim FM(-1 To 1, -1 To 1) As Long
    FM(-1, -1) = -1
    FM(-1, 0) = -1
    FM(-1, 1) = -1
    FM(1, -1) = 1
    FM(1, 0) = 1
    FM(1, 1) = 1
    FilterWeight = 1
    FilterBias = 0
    DoFilter "Prewitt horizontal edge detection", True, , toPreview, dstPic
End Sub

Public Sub FilterPrewittVertical(Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As PictureBox)
    FilterSize = 3
    ReDim FM(-1 To 1, -1 To 1) As Long
    FM(-1, -1) = 1
    FM(0, -1) = 1
    FM(1, -1) = 1
    FM(-1, 1) = -1
    FM(0, 1) = -1
    FM(1, 1) = -1
    FilterWeight = 1
    FilterBias = 0
    DoFilter "Prewitt vertical edge detection", True, , toPreview, dstPic
End Sub

Public Sub FilterSobelHorizontal(Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As PictureBox)
    FilterSize = 3
    ReDim FM(-1 To 1, -1 To 1) As Long
    FM(-1, -1) = -1
    FM(-1, 0) = -2
    FM(-1, 1) = -1
    FM(1, -1) = 1
    FM(1, 0) = 2
    FM(1, 1) = 1
    FilterWeight = 1
    FilterBias = 0
    DoFilter "Sobel horizontal edge detection", True, , toPreview, dstPic
End Sub

Public Sub FilterSobelVertical(Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As PictureBox)
    FilterSize = 3
    ReDim FM(-1 To 1, -1 To 1) As Long
    FM(-1, -1) = 1
    FM(0, -1) = 2
    FM(1, -1) = 1
    FM(-1, 1) = -1
    FM(0, 1) = -2
    FM(1, 1) = -1
    FilterWeight = 1
    FilterBias = 0
    DoFilter "Sobel vertical edge detection", True, , toPreview, dstPic
End Sub

Public Sub FilterLaplacian(Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As PictureBox)
    FilterSize = 3
    ReDim FM(-1 To 1, -1 To 1) As Long
    FM(-1, 0) = -1
    FM(0, -1) = -1
    FM(0, 1) = -1
    FM(1, 0) = -1
    FM(0, 0) = 4
    FilterWeight = 1
    FilterBias = 0
    DoFilter "Laplacian edge detection", True, , toPreview, dstPic
End Sub

'This code is a modified version of an algorithm originally developed by Manuel Augusto Santos.  A link to his original
' implementation is available from the "Help -> About PhotoDemon" menu option.
Public Sub FilterSmoothContour(Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As PictureBox)

    If toPreview = False Then Message "Tracing image edges with virtual paintbrush..."
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent already embossed pixels from screwing up our results for later pixels.)
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    
    Dim srcLayer As pdLayer
    Set srcLayer = New pdLayer
    srcLayer.createFromExistingLayer workingLayer
    
    prepSafeArray srcSA, srcLayer
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, z As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left + 1
    initY = curLayerValues.Top + 1
    finalX = curLayerValues.Right - 1
    finalY = curLayerValues.Bottom - 1
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, QuickValRight As Long, QuickValLeft As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Color variables
    Dim tmpColor As Long, tMin As Long
        
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
        QuickValRight = (x + 1) * qvDepth
        QuickValLeft = (x - 1) * qvDepth
    For y = initY To finalY
        For z = 0 To 2
    
            tMin = 255
            tmpColor = srcImageData(QuickValRight + z, y)
            If tmpColor < tMin Then tMin = tmpColor
            tmpColor = srcImageData(QuickValRight + z, y - 1)
            If tmpColor < tMin Then tMin = tmpColor
            tmpColor = srcImageData(QuickValRight + z, y + 1)
            If tmpColor < tMin Then tMin = tmpColor
            tmpColor = srcImageData(QuickValLeft + z, y)
            If tmpColor < tMin Then tMin = tmpColor
            tmpColor = srcImageData(QuickValLeft + z, y - 1)
            If tmpColor < tMin Then tMin = tmpColor
            tmpColor = srcImageData(QuickValLeft + z, y + 1)
            If tmpColor < tMin Then tMin = tmpColor
            tmpColor = srcImageData(QuickVal + z, y)
            If tmpColor < tMin Then tMin = tmpColor
            tmpColor = srcImageData(QuickVal + z, y - 1)
            If tmpColor < tMin Then tMin = tmpColor
            tmpColor = srcImageData(QuickVal + z, y + 1)
            If tmpColor < tMin Then tMin = tmpColor
            
            If tMin > 255 Then tMin = 255
            If tMin < 0 Then tMin = 0
            
            dstImageData(QuickVal + z, y) = 255 - (srcImageData(QuickVal + z, y) - tMin)
            
            'The edges of the image will always be missed, so manually check for and correct that
            If x = initX Then dstImageData(QuickValLeft + z, y) = dstImageData(QuickVal + z, y)
            If x = finalX Then dstImageData(QuickValRight + z, y) = dstImageData(QuickVal + z, y)
            If y = initY Then dstImageData(QuickVal + z, y - 1) = dstImageData(QuickVal + z, y)
            If y = finalY Then dstImageData(QuickVal + z, y + 1) = dstImageData(QuickVal + z, y)
        
        Next z
    Next y
        If toPreview = False Then
            If (x And progBarCheck) = 0 Then SetProgBarVal x
        End If
    Next x
    
    'With our work complete, point both ImageData() arrays away from their DIBs and deallocate them
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    Erase srcImageData
    
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    Erase dstImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
     
End Sub

Private Sub LstEdgeOptions_Click()
    UpdateDescriptions
End Sub

'Show the user a brief explanation of the algorithm in question.  Yes, the PhotoDemon routine descriptions are bullshit -
' I know that already.  :)  But the descriptions make them sound more impressive than they actually are.
' This sub also handles redrawing the edge detection preview.
Private Sub UpdateDescriptions()
    Dim l As String
    l = LstEdgeOptions.List(LstEdgeOptions.ListIndex)
    If l = "Prewitt Horizontal" Then
        lblDesc = "Simple matrix method:" & vbCrLf & vbCrLf & "-1 0 1" & vbCrLf & "-1 0 1" & vbCrLf & "-1 0 1"
        FilterPrewittHorizontal True, PicEffect
    ElseIf l = "Prewitt Vertical" Then
        lblDesc = "Simple matrix method:" & vbCrLf & vbCrLf & "-1 -1 -1" & vbCrLf & " 0  0  0" & vbCrLf & " 1  1  1"
        FilterPrewittVertical True, PicEffect
    ElseIf l = "Sobel Horizontal" Then
        lblDesc = "Simple matrix method:" & vbCrLf & vbCrLf & "-1 0 1" & vbCrLf & "-2 0 2" & vbCrLf & "-1 0 1"
        FilterSobelHorizontal True, PicEffect
    ElseIf l = "Sobel Vertical" Then
        lblDesc = "Simple matrix method:" & vbCrLf & vbCrLf & "-1 -2 -1" & vbCrLf & " 0  0  0" & vbCrLf & " 1  2  1"
        FilterSobelVertical True, PicEffect
    ElseIf l = "Laplacian" Then
        lblDesc = "Simple matrix method:" & vbCrLf & vbCrLf & " 0 -1  0" & vbCrLf & "-1  4 -1" & vbCrLf & " 0 -1  0"
        FilterLaplacian True, PicEffect
    ElseIf l = "Artistic Contour" Then
        lblDesc = "Algorithm designed to present a clean, artistic prediction of image edges."
        FilterSmoothContour True, PicEffect
    ElseIf l = "Hilite" Then
        lblDesc = "Simple matrix method:" & vbCrLf & vbCrLf & "-4 -2 -1" & vbCrLf & "-2 10  0" & vbCrLf & "-1  0  0"
        FilterHilite True, PicEffect
    ElseIf l = "PhotoDemon Linear" Then
        lblDesc = "Simple mathematical routine based on linear relationships between diagonal pixels."
        PhotoDemonLinearEdgeDetection True, PicEffect
    Else
        lblDesc = "Advanced mathematical routine based on cubic relationships between diagonal pixels."
        PhotoDemonCubicEdgeDetection True, PicEffect
    End If
End Sub

