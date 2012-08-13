VERSION 5.00
Begin VB.Form FormFindEdges 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Find Edges"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6585
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
   ScaleHeight     =   215
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   439
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
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
      Left            =   4080
      MouseIcon       =   "VBP_FormFindEdges.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2640
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
      Left            =   5280
      MouseIcon       =   "VBP_FormFindEdges.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2640
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
      MouseIcon       =   "VBP_FormFindEdges.frx":02A4
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   240
      Width           =   3135
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
      Left            =   3360
      TabIndex        =   4
      Top             =   240
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
      Height          =   1695
      Left            =   3360
      TabIndex        =   3
      Top             =   600
      Width           =   3015
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
'Last updated: 19/June/12
'Last update: rewritten descriptions, code clean-up.
'
'All known edge-detection routines are handled from this form.
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
    
    UpdateDescriptions
    
    'Assign the system hand cursor to all relevant objects
    setHandCursorForAll Me
    
End Sub

Public Sub FilterHilite()
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
    DoFilter "Hilite edge detection", True
End Sub

Public Sub PhotoDemonCubicEdgeDetection()
    FilterSize = 5
    ReDim FM(-2 To 2, -2 To 2) As Long
    FM(-1, -2) = 1
    FM(-2, 1) = 1
    FM(1, 2) = 1
    FM(2, -1) = 1
    FM(0, 0) = -4
    FilterWeight = 1
    FilterBias = 0
    DoFilter "PhotoDemon cubic edge detection", True
End Sub

Public Sub PhotoDemonLinearEdgeDetection()
    FilterSize = 3
    ReDim FM(-1 To 1, -1 To 1) As Long
    FM(-1, -1) = -1
    FM(-1, 1) = -1
    FM(1, -1) = -1
    FM(1, 1) = -1
    FM(0, 0) = 4
    FilterWeight = 1
    FilterBias = 0
    DoFilter "PhotoDemon linear edge detection", True
End Sub

Public Sub FilterPrewittHorizontal()
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
    DoFilter "Prewitt horizontal edge detection", True
End Sub

Public Sub FilterPrewittVertical()
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
    DoFilter "Prewitt vertical edge detection", True
End Sub

Public Sub FilterSobelHorizontal()
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
    DoFilter "Sobel horizontal edge detection", True
End Sub

Public Sub FilterSobelVertical()
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
    DoFilter "Sobel vertical edge detection", True
End Sub

Public Sub FilterLaplacian()
    FilterSize = 3
    ReDim FM(-1 To 1, -1 To 1) As Long
    FM(-1, 0) = -1
    FM(0, -1) = -1
    FM(0, 1) = -1
    FM(1, 0) = -1
    FM(0, 0) = 4
    FilterWeight = 1
    FilterBias = 0
    DoFilter "Laplacian edge detection", True
End Sub

'This code is a modified version of an algorithm originally developed by Manuel Augusto Santos.  A link to his original implementation
' is available from the "Help -> About PhotoDemon" menu option
Public Sub FilterSmoothContour()
    
    Message "Analyzing image edges..."
    
    SetProgBarMax PicHeightL
    
    Dim TC As Long, tMin As Long
    ReDim tData(0 To PicWidthL * 3 + 3, 0 To PicHeightL + 1)
    
    Dim QuickX As Long, QuickXLeft As Long, QuickXRight As Long
    
    For y = 1 To PicHeightL - 1
    For x = 1 To PicWidthL - 1
        QuickX = x * 3
        QuickXLeft = (x - 1) * 3
        QuickXRight = (x + 1) * 3
    For z = 0 To 2
        tMin = 255
        TC = ImageData(QuickXRight + z, y)
        If TC < tMin Then tMin = TC
        TC = ImageData(QuickXRight + z, y - 1)
        If TC < tMin Then tMin = TC
        TC = ImageData(QuickXRight + z, y + 1)
        If TC < tMin Then tMin = TC
        TC = ImageData(QuickXLeft + z, y)
        If TC < tMin Then tMin = TC
        TC = ImageData(QuickXLeft + z, y - 1)
        If TC < tMin Then tMin = TC
        TC = ImageData(QuickXLeft + z, y + 1)
        If TC < tMin Then tMin = TC
        TC = ImageData(QuickX + z, y)
        If TC < tMin Then tMin = TC
        TC = ImageData(QuickX + z, y - 1)
        If TC < tMin Then tMin = TC
        TC = ImageData(QuickX + z, y + 1)
        If TC < tMin Then tMin = TC
        
        If tMin > 255 Then tMin = 255
        If tMin < 0 Then tMin = 0
        
        tData(QuickX + z, y) = 255 - (ImageData(QuickX + z, y) - tMin)
        
    Next z
    Next x
        If y Mod 20 = 0 Then SetProgBarVal y
    Next y
    
    TransferImageData
    SetImageData
    
End Sub

Private Sub TransferImageData()
    Message "Transferring data..."
    Dim QuickX As Long
    For x = 0 To PicWidthL
        QuickX = x * 3
    For y = 0 To PicHeightL
        For z = 0 To 2
            ImageData(QuickX + z, y) = tData(QuickX + z, y)
        Next z
    Next y
    Next x
End Sub

Private Sub LstEdgeOptions_Click()
    UpdateDescriptions
End Sub

'Show the user a brief explanation of the algorithm in question.  Yes, the PhotoDemon routines are bullshit - I know that already.  :)
Private Sub UpdateDescriptions()
    Dim l As String
    l = LstEdgeOptions.List(LstEdgeOptions.ListIndex)
    If l = "Prewitt Horizontal" Then
        lblDesc = "Simple matrix method:" & vbCrLf & vbCrLf & "-1 0 1" & vbCrLf & "-1 0 1" & vbCrLf & "-1 0 1"
    ElseIf l = "Prewitt Vertical" Then
        lblDesc = "Simple matrix method:" & vbCrLf & vbCrLf & "-1 -1 -1" & vbCrLf & " 0  0  0" & vbCrLf & " 1  1  1"
    ElseIf l = "Sobel Horizontal" Then
        lblDesc = "Simple matrix method:" & vbCrLf & vbCrLf & "-1 0 1" & vbCrLf & "-2 0 2" & vbCrLf & "-1 0 1"
    ElseIf l = "Sobel Vertical" Then
        lblDesc = "Simple matrix method:" & vbCrLf & vbCrLf & "-1 -2 -1" & vbCrLf & " 0  0  0" & vbCrLf & " 1  2  1"
    ElseIf l = "Laplacian" Then
        lblDesc = "Simple matrix method:" & vbCrLf & vbCrLf & " 0 -1  0" & vbCrLf & "-1  4 -1" & vbCrLf & " 0 -1  0"
    ElseIf l = "Artistic Contour" Then
        lblDesc = "Algorithm designed to present a clean, artistic prediction of image edges."
    ElseIf l = "Hilite" Then
        lblDesc = "Simple matrix method:" & vbCrLf & vbCrLf & "-4 -2 -1" & vbCrLf & "-2 10  0" & vbCrLf & "-1  0  0"
    ElseIf l = "PhotoDemon Linear" Then
        lblDesc = "Simple mathematical routine based on linear relationships between diagonal pixels."
    Else
        lblDesc = "Advanced mathematical routine based on cubic relationships between diagonal pixels."
    End If
End Sub
