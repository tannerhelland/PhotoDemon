VERSION 5.00
Begin VB.Form FormGrayscale 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " "
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10350
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
   ScaleHeight     =   458
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   690
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
      Left            =   4080
      ScaleHeight     =   180
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   191
      TabIndex        =   13
      TabStop         =   0   'False
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
      Left            =   7200
      ScaleHeight     =   180
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   191
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   120
      Width           =   2895
   End
   Begin VB.HScrollBar hsShades 
      Height          =   255
      Left            =   4320
      Max             =   254
      Min             =   2
      TabIndex        =   4
      Top             =   5280
      Value           =   3
      Width           =   4785
   End
   Begin VB.TextBox txtShades 
      Alignment       =   2  'Center
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
      Height          =   315
      Left            =   9240
      MaxLength       =   3
      TabIndex        =   3
      Text            =   "3"
      Top             =   5250
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox cboMethod 
      Appearance      =   0  'Flat
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
      Height          =   330
      Left            =   4320
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   4080
      Width           =   3975
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   7560
      TabIndex        =   0
      Top             =   6270
      Width           =   1245
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   8880
      TabIndex        =   1
      Top             =   6270
      Width           =   1245
   End
   Begin VB.PictureBox picChannel 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4320
      ScaleHeight     =   375
      ScaleWidth      =   3855
      TabIndex        =   15
      Top             =   5280
      Width           =   3855
      Begin VB.OptionButton optChannel 
         Appearance      =   0  'Flat
         Caption         =   "blue"
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
         Height          =   255
         Index           =   2
         Left            =   2520
         TabIndex        =   9
         Top             =   0
         Width           =   1095
      End
      Begin VB.OptionButton optChannel 
         Appearance      =   0  'Flat
         Caption         =   "red"
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
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   0
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optChannel 
         Appearance      =   0  'Flat
         Caption         =   "green"
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
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   8
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox picDecompose 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4320
      ScaleHeight     =   375
      ScaleWidth      =   3975
      TabIndex        =   14
      Top             =   5280
      Width           =   3975
      Begin VB.OptionButton optDecompose 
         Appearance      =   0  'Flat
         Caption         =   "maximum"
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
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   6
         Top             =   0
         Width           =   1455
      End
      Begin VB.OptionButton optDecompose 
         Appearance      =   0  'Flat
         Caption         =   "minimum"
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
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   0
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Label lblLink 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Grayscale Conversion: An In-Depth Look"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFCBA1&
      Height          =   255
      Left            =   120
      MouseIcon       =   "VBP_FormGrayscale.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   21
      Top             =   6375
      Width           =   3495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      X1              =   264
      X2              =   680
      Y1              =   406
      Y2              =   406
   End
   Begin VB.Label lblExplanation 
      BackStyle       =   0  'Transparent
      Caption         =   "(Description appears here)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   5415
      Left            =   360
      TabIndex        =   20
      Top             =   930
      Width           =   3135
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "grayscale tool"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   0
      TabIndex        =   19
      Top             =   240
      Width           =   3825
   End
   Begin VB.Label lblBackColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   6855
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   3825
   End
   Begin VB.Label lblAfter 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "after"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   7320
      TabIndex        =   17
      Top             =   2880
      Width           =   360
   End
   Begin VB.Label lblBefore 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "before"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   4200
      TabIndex        =   16
      Top             =   2880
      Width           =   480
   End
   Begin VB.Label lblAdditional 
      AutoSize        =   -1  'True
      Caption         =   "additional options for this method:"
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
      Left            =   4200
      TabIndex        =   11
      Top             =   4800
      Width           =   3690
   End
   Begin VB.Label lblAlgorithm 
      AutoSize        =   -1  'True
      Caption         =   "grayscale method:"
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
      Left            =   4200
      TabIndex        =   10
      Top             =   3645
      Width           =   1950
   End
End
Attribute VB_Name = "FormGrayscale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Grayscale Conversion Handler
'Copyright ©2000-2012 by Tanner Helland
'Created: 1/12/02
'Last updated: 09/September/12
'Last update: improved accuracy of dithered conversion by using more Single-type variables.
'
'Updated version of the grayscale handler; utilizes five different methods
'(average, ISU, desaturate, X # of shades, X # of shades dithered).
'
'***************************************************************************

Option Explicit

'This routine is used to call the appropriate grayscale routine with the preview flag set
Private Sub drawGrayscalePreview()

    'Error checking
    If EntryValid(txtShades, hsShades.Min, hsShades.Max, False, False) Then
        
        Select Case cboMethod.ListIndex
            Case 0
                MenuGrayscaleAverage True, picEffect
            Case 1
                MenuGrayscale True, picEffect
            Case 2
                MenuDesaturate True, picEffect
            Case 3
                If optDecompose(0).Value = True Then
                    MenuDecompose 0, True, picEffect
                Else
                    MenuDecompose 1, True, picEffect
                End If
            Case 4
                If optChannel(0).Value = True Then
                    MenuGrayscaleSingleChannel 0, True, picEffect
                ElseIf optChannel(1).Value = True Then
                    MenuGrayscaleSingleChannel 1, True, picEffect
                Else
                    MenuGrayscaleSingleChannel 2, True, picEffect
                End If
            Case 5
                fGrayscaleCustom hsShades.Value, True, picEffect
            Case 6
                fGrayscaleCustomDither hsShades.Value, True, picEffect
        End Select
    
    End If

End Sub

Private Sub cboMethod_Click()
    UpdateVisibleControls
    drawGrayscalePreview
End Sub

Private Sub cboMethod_KeyDown(KeyCode As Integer, Shift As Integer)
    UpdateVisibleControls
    drawGrayscalePreview
End Sub

'Certain algorithms require additional user input.  This routine enables/disables the controls associated with a given algorithm.
Private Sub UpdateVisibleControls()
    
    Select Case cboMethod.ListIndex
        Case 3
            txtShades.Visible = False
            hsShades.Visible = False
            lblAdditional.Caption = "decompose using these values:"
            lblAdditional.Visible = True
            picDecompose.Visible = True
            picChannel.Visible = False
        Case 4
            txtShades.Visible = False
            hsShades.Visible = False
            lblAdditional.Caption = "use this channel:"
            lblAdditional.Visible = True
            picDecompose.Visible = False
            picChannel.Visible = True
        Case 5
            txtShades.Visible = True
            hsShades.Visible = True
            lblAdditional.Caption = "use this many shades of gray:"
            lblAdditional.Visible = True
            picDecompose.Visible = False
            picChannel.Visible = False
        Case 6
            txtShades.Visible = True
            hsShades.Visible = True
            lblAdditional.Caption = "use this many shades of gray:"
            lblAdditional.Visible = True
            picDecompose.Visible = False
            picChannel.Visible = False
        Case Else
            txtShades.Visible = False
            hsShades.Visible = False
            lblAdditional.Visible = False
            picDecompose.Visible = False
            picChannel.Visible = False
    End Select
    
End Sub

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub CmdOK_Click()
    
    'Error checking
    If EntryValid(txtShades, hsShades.Min, hsShades.Max) Then
        
        Me.Visible = False
        
        Select Case cboMethod.ListIndex
            Case 0
                Process GrayscaleAverage
            Case 1
                Process GrayScale
            Case 2
                Process Desaturate
            Case 3
                If optDecompose(0).Value = True Then
                    Process GrayscaleDecompose, 0
                Else
                    Process GrayscaleDecompose, 1
                End If
            Case 4
                If optChannel(0).Value = True Then
                    Process GrayscaleSingleChannel, 0
                ElseIf optChannel(1).Value = True Then
                    Process GrayscaleSingleChannel, 1
                Else
                    Process GrayscaleSingleChannel, 2
                End If
            Case 5
                Process GrayscaleCustom, hsShades.Value
            Case 6
                Process GrayscaleCustomDither, hsShades.Value
        End Select
        
        Unload Me
        
    Else
        AutoSelectText txtShades
    End If

End Sub

Private Sub Form_Activate()
        
    'Set up the grayscale options combo box
    cboMethod.AddItem "Fastest Calculation (average value)", 0
    cboMethod.AddItem "Highest Quality (ITU Standard)", 1
    cboMethod.AddItem "Desaturate", 2
    cboMethod.AddItem "Decompose", 3
    cboMethod.AddItem "Single color channel", 4
    cboMethod.AddItem "Specific # of shades", 5
    cboMethod.AddItem "Specific # of shades (dithered)", 6
    cboMethod.ListIndex = 1
    
    UpdateVisibleControls
    
    'Populate the explanation label
    lblExplanation = "This tool removes color data from an image.  The new image contains only shades of gray." & vbCrLf & vbCrLf & "Sometimes this tool is called a ""black and white"" tool, but that name is misleading because the processed image contains many more shades than just ""black"" and ""white"".  A separate ""Black and White"" tool (found in the ""Color"" menu) exists if you want an image with just black and just white." & vbCrLf & vbCrLf & "While there are many ways to remove color from an image, most users stick with the ""Highest Quality (ITU Standard)"" method, which produces the best grayscale image.  Other options are provided for artistic effect." & vbCrLf & vbCrLf & "To learn more about the various grayscale conversion options, please visit this link:"
    
    'Render the preview images
    DrawPreviewImage picPreview
    
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
    
    
    'Draw the initial preview
    drawGrayscalePreview
    
End Sub

'Reduce to X # gray shades
Public Sub fGrayscaleCustom(ByVal numOfShades As Long, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As PictureBox)

    If toPreview = False Then Message "Converting image to " & numOfShades & " shades of gray..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    
    prepImageData tmpSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Color variables
    Dim r As Long, g As Long, b As Long, grayVal As Long
    
    'This conversion factor is the value we need to turn grayscale values in the [0,255] range into a specific subset of values
    Dim conversionFactor As Single
    conversionFactor = (255 / (numOfShades - 1))
    
    'Build a look-up table for our custom grayscale conversion results
    Dim LookUp(0 To 255) As Byte
    
    For x = 0 To 255
        grayVal = Int((CDbl(x) / conversionFactor) + 0.5) * conversionFactor
        If grayVal > 255 Then grayVal = 255
        LookUp(x) = CByte(grayVal)
    Next x
    
    'Build another look-up table for our initial grayscale index calculation
    Dim grayLookUp(0 To 765) As Byte
    For x = 0 To 765
        grayLookUp(x) = x \ 3
    Next x
        
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        grayVal = grayLookUp(r + g + b)
        
        'Assign all color channels the new gray value
        ImageData(QuickVal + 2, y) = LookUp(grayVal)
        ImageData(QuickVal + 1, y) = LookUp(grayVal)
        ImageData(QuickVal, y) = LookUp(grayVal)
        
    Next y
        If toPreview = False Then
            If (x And progBarCheck) = 0 Then SetProgBarVal x
        End If
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
    
End Sub

'Reduce to X # gray shades (dithered)
Public Sub fGrayscaleCustomDither(ByVal numOfShades As Long, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As PictureBox)

    If toPreview = False Then Message "Converting to " & numOfShades & " shades of gray, with dithering..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    
    prepImageData tmpSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    SetProgBarMax finalY
    progBarCheck = findBestProgBarValue()
    
    'Color variables
    Dim r As Long, g As Long, b As Long, grayVal As Long
    
    'This conversion factor is the value we need to turn grayscale values in the [0,255] range into a specific subset of values
    Dim conversionFactor As Single
    conversionFactor = (255 / (numOfShades - 1))
    
    'Build another look-up table for our initial grayscale index calculation
    Dim grayLookUp(0 To 765) As Byte
    For x = 0 To 765
        grayLookUp(x) = x \ 3
    Next x
    
    'Unfortunately, this algorithm (unlike its non-dithering counterpart) is not well-suited to using a look-up table,
    ' so all calculations have been moved into the loop
    Dim grayTempCalc As Single
    
    'This value tracks the drifting error of our conversions, which allows us to dither
    Dim errorValue As Single
    errorValue = 0
        
    'Loop through each pixel in the image, converting values as we go
    For y = initY To finalY
    For x = initX To finalX
    
        QuickVal = x * qvDepth
    
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Look up our initial grayscale value in the table
        grayVal = grayLookUp(r + g + b)
        
        'Add the error value (a cumulative value of the difference between actual gray values and gray values we've selected) to the current gray value
        grayTempCalc = grayVal + errorValue
        
        'Rebuild our temporary calculation variable using the shade reduction formula
        grayTempCalc = Int((CSng(grayTempCalc) / conversionFactor) + 0.5) * conversionFactor
        
        'Adjust our error value to include this latest calculation
        errorValue = CLng(grayVal) + errorValue - grayTempCalc
        
        If grayTempCalc < 0 Then grayTempCalc = 0
        If grayTempCalc > 255 Then grayTempCalc = 255
        
        grayVal = CByte(grayTempCalc)
        
        'Assign all color channels the new gray value
        ImageData(QuickVal + 2, y) = grayVal
        ImageData(QuickVal + 1, y) = grayVal
        ImageData(QuickVal, y) = grayVal
        
    Next x
        
        'Reset the error value at the end of each line
        errorValue = 0
        
        If toPreview = False Then
            If (y And progBarCheck) = 0 Then SetProgBarVal y
        End If
        
    Next y
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
    
End Sub

'Reduce to gray via (r+g+b)/3
Public Sub MenuGrayscaleAverage(Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As PictureBox)
    
    If toPreview = False Then Message "Converting image to grayscale..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    
    prepImageData tmpSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Color and grayscale variables
    Dim r As Long, g As Long, b As Long
    Dim grayVal As Byte
    
    'Build a look-up table of grayscale values (faster than calculating it manually for each pixel)
    Dim grayLookUp(0 To 765) As Byte
    For x = 0 To 765
        grayLookUp(x) = x \ 3
    Next x
    
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Calculate the gray value using the look-up table
        grayVal = grayLookUp(r + g + b)
        
        'Assign that gray value to each color channel
        ImageData(QuickVal, y) = grayVal
        ImageData(QuickVal + 1, y) = grayVal
        ImageData(QuickVal + 2, y) = grayVal
        
    Next y
        If toPreview = False Then
            If (x And progBarCheck) = 0 Then SetProgBarVal x
        End If
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
    
End Sub

'Reduce to gray in a more human-eye friendly manner
Public Sub MenuGrayscale(Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As PictureBox)
    
    If toPreview = False Then Message "Generating ITU-R compatible grayscale image..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    
    prepImageData tmpSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Color and grayscale variables
    Dim r As Long, g As Long, b As Long
    Dim grayVal As Long
        
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Calculate a grayscale value using the original ITU-R recommended formula (BT.709, specifically)
        grayVal = (213 * r + 715 * g + 72 * b) \ 1000
        If grayVal > 255 Then grayVal = 255
        
        'Assign that gray value to each color channel
        ImageData(QuickVal, y) = grayVal
        ImageData(QuickVal + 1, y) = grayVal
        ImageData(QuickVal + 2, y) = grayVal
        
    Next y
        If toPreview = False Then
            If (x And progBarCheck) = 0 Then SetProgBarVal x
        End If
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
    
End Sub

'Reduce to gray via HSL -> convert S to 0
Public Sub MenuDesaturate(Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As PictureBox)
        
    If toPreview = False Then Message "Desaturating image..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    
    prepImageData tmpSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Color and grayscale variables
    Dim r As Long, g As Long, b As Long
    Dim grayVal As Byte
       
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Calculate a grayscale value by using a short-hand RGB <-> HSL conversion
        grayVal = CByte(getLuminance(r, g, b))
        
        'Assign that gray value to each color channel
        ImageData(QuickVal, y) = grayVal
        ImageData(QuickVal + 1, y) = grayVal
        ImageData(QuickVal + 2, y) = grayVal
        
    Next y
        If toPreview = False Then
            If (x And progBarCheck) = 0 Then SetProgBarVal x
        End If
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
    
End Sub

'Reduce to gray by selecting the minimum (maxOrMin = 0) or maximum (maxOrMin = 1) color in each pixel
Public Sub MenuDecompose(ByVal maxOrMin As Long, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As PictureBox)

    If toPreview = False Then Message "Decomposing image..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    
    prepImageData tmpSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Color and grayscale variables
    Dim r As Long, g As Long, b As Long
    Dim grayVal As Byte
        
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Find the highest or lowest of the RGB values
        If maxOrMin = 0 Then grayVal = CByte(MinimumInt(r, g, b)) Else grayVal = CByte(MaximumInt(r, g, b))
        
        'Assign that gray value to each color channel
        ImageData(QuickVal, y) = grayVal
        ImageData(QuickVal + 1, y) = grayVal
        ImageData(QuickVal + 2, y) = grayVal
        
    Next y
        If toPreview = False Then
            If (x And progBarCheck) = 0 Then SetProgBarVal x
        End If
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
    
End Sub

'Reduce to gray by selecting a single color channel (represeted by cChannel: 0 = Red, 1 = Green, 2 = Blue)
Public Sub MenuGrayscaleSingleChannel(ByVal cChannel As Long, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As PictureBox)

    Dim cString As String
     
    Select Case cChannel
        Case 0
            cString = "red"
        Case 1
            cString = "green"
        Case 2
            cString = "blue"
    End Select

    If toPreview = False Then Message "Converting image to grayscale using " & cString & " values..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    
    prepImageData tmpSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Color and grayscale variables
    Dim r As Long, g As Long, b As Long
    Dim grayVal As Byte
        
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Assign the gray value to a single color channel based on the value of cChannel
        Select Case cChannel
            Case 0
                grayVal = r
            Case 1
                grayVal = g
            Case 2
                grayVal = b
        End Select
        
        'Assign that gray value to each color channel
        ImageData(QuickVal, y) = grayVal
        ImageData(QuickVal + 1, y) = grayVal
        ImageData(QuickVal + 2, y) = grayVal
        
    Next y
        If toPreview = False Then
            If (x And progBarCheck) = 0 Then SetProgBarVal x
        End If
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
        
End Sub

'When the "# of shades" horizontal scroll bar is changed, update the text box to match
Private Sub hsShades_Change()
    copyToTextBoxI txtShades, hsShades.Value
    drawGrayscalePreview
End Sub

Private Sub hsShades_Scroll()
    copyToTextBoxI txtShades, hsShades.Value
    drawGrayscalePreview
End Sub

Private Sub lblLink_Click()
    OpenURL "http://www.tannerhelland.com/3643/grayscale-image-algorithm-vb6/"
End Sub

'When option buttons are used, update the preview accordingly
Private Sub optChannel_Click(Index As Integer)
    drawGrayscalePreview
End Sub

Private Sub optDecompose_Click(Index As Integer)
    drawGrayscalePreview
End Sub

'As a convenience to the user, when they click the "# of shades" text box, automatically select the text for them
Private Sub txtShades_GotFocus()
    AutoSelectText txtShades
End Sub

'When the "# of shades" text box is changed, check the value for errors and redraw the preview
Private Sub txtShades_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtShades
    If EntryValid(txtShades, hsShades.Min, hsShades.Max, False, False) Then hsShades.Value = Val(txtShades)
End Sub
