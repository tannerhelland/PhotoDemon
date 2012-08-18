VERSION 5.00
Begin VB.Form FormGrayscale 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Color to Grayscale Conversion"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6495
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
   ScaleHeight     =   399
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   433
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
      Left            =   240
      ScaleHeight     =   180
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   191
      TabIndex        =   8
      Top             =   240
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
      Left            =   3360
      ScaleHeight     =   180
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   191
      TabIndex        =   7
      Top             =   240
      Width           =   2895
   End
   Begin VB.HScrollBar hsShades 
      Height          =   255
      Left            =   2760
      Max             =   254
      Min             =   2
      TabIndex        =   2
      Top             =   4200
      Value           =   3
      Width           =   3345
   End
   Begin VB.TextBox txtShades 
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
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      Text            =   "3"
      Top             =   4170
      Visible         =   0   'False
      Width           =   495
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
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   3600
      Width           =   3975
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   5400
      Width           =   1125
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Top             =   5400
      Width           =   1125
   End
   Begin VB.PictureBox picChannel 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2160
      ScaleHeight     =   375
      ScaleWidth      =   3975
      TabIndex        =   13
      Top             =   4200
      Width           =   3975
      Begin VB.OptionButton optChannel 
         Appearance      =   0  'Flat
         Caption         =   "Blue"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   2400
         TabIndex        =   16
         Top             =   0
         Width           =   1095
      End
      Begin VB.OptionButton optChannel 
         Appearance      =   0  'Flat
         Caption         =   "Red"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   0
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optChannel 
         Appearance      =   0  'Flat
         Caption         =   "Green"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   14
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox picDecompose 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2160
      ScaleHeight     =   375
      ScaleWidth      =   3975
      TabIndex        =   10
      Top             =   4200
      Width           =   3975
      Begin VB.OptionButton optDecompose 
         Appearance      =   0  'Flat
         Caption         =   "Maximum"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   12
         Top             =   0
         Width           =   1455
      End
      Begin VB.OptionButton optDecompose 
         Appearance      =   0  'Flat
         Caption         =   "Minimum"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   0
         Value           =   -1  'True
         Width           =   1455
      End
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
      Left            =   240
      TabIndex        =   9
      Top             =   3000
      Width           =   3975
   End
   Begin VB.Label lblAdditional 
      AutoSize        =   -1  'True
      Caption         =   "Additional options:"
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
      Height          =   210
      Left            =   360
      TabIndex        =   6
      Top             =   4200
      Width           =   1515
   End
   Begin VB.Label lblAlgorithm 
      AutoSize        =   -1  'True
      Caption         =   "Grayscale algorithm:"
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
      Height          =   210
      Left            =   360
      TabIndex        =   5
      Top             =   3645
      Width           =   1620
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
'Last updated: 18/August/09
'Last update: homebrew methods now use the simpler (R+G+B)\3 method
'
'NOTE: this code still needs to be optimized and cleaned up - look to the
' grayscale project on THDC for specifics.
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
                MenuGrayscaleAverage True, PicPreview, PicEffect
            Case 1
                MenuGrayscale True, PicPreview, PicEffect
            Case 2
                MenuDesaturate True, PicPreview, PicEffect
            Case 3
                If optDecompose(0).Value = True Then
                    MenuDecompose 0, True, PicPreview, PicEffect
                Else
                    MenuDecompose 1, True, PicPreview, PicEffect
                End If
            Case 4
                If optChannel(0).Value = True Then
                    MenuGrayscaleSingleChannel 0, True, PicPreview, PicEffect
                ElseIf optChannel(1).Value = True Then
                    MenuGrayscaleSingleChannel 1, True, PicPreview, PicEffect
                Else
                    MenuGrayscaleSingleChannel 2, True, PicPreview, PicEffect
                End If
            Case 5
                fGrayscaleCustom hsShades.Value, True, PicPreview, PicEffect
            Case 6
                fGrayscaleCustomDither hsShades.Value, True, PicPreview, PicEffect
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
            lblAdditional.Visible = True
            picDecompose.Visible = True
            picChannel.Visible = False
        Case 4
            txtShades.Visible = False
            hsShades.Visible = False
            lblAdditional.Visible = True
            picDecompose.Visible = False
            picChannel.Visible = True
        Case 5
            txtShades.Visible = True
            hsShades.Visible = True
            lblAdditional.Visible = True
            picDecompose.Visible = False
            picChannel.Visible = False
        Case 6
            txtShades.Visible = True
            hsShades.Visible = True
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

'Initialize the combo box
Private Sub Form_Load()
        
    'Set up the grayscale options combo box
    cboMethod.AddItem "Average value [(R+G+B) / 3]", 0
    cboMethod.AddItem "Human eye equivalent [ITU Standard]", 1
    cboMethod.AddItem "Desaturate", 2
    cboMethod.AddItem "Decompose", 3
    cboMethod.AddItem "Single color channel", 4
    cboMethod.AddItem "Specific # of shades", 5
    cboMethod.AddItem "Specific # of shades (dithered)", 6
    cboMethod.ListIndex = 1
    
    UpdateVisibleControls
    
    'Render the preview images
    DrawPreviewImage PicPreview
    DrawPreviewImage PicEffect
    
    'Assign the system hand cursor to all relevant objects
    setHandCursorForAll Me
    
    'Draw the initial preview
    drawGrayscalePreview
    
End Sub

'Reduce to X # gray shades
Public Sub fGrayscaleCustom(ByVal numOfShades As Long, Optional ByVal toPreview As Boolean = False, Optional ByRef srcPic As PictureBox, Optional ByRef dstPic As PictureBox)
    
    'Get the appropriate set of image data contingent on whether this is a preview or not
    If toPreview = True Then
        GetPreviewData srcPic
    Else
        Message "Converting to " & numOfShades & " shades of gray..."
        GetImageData
        SetProgBarMax PicHeightL
    End If
    
    'This conversion factor is the value we need to turn grayscale values in the [0,255] range into a specific subset of values
    Dim conversionFactor As Single
    conversionFactor = (255 / (numOfShades - 1))
    
    'Build a look-up table for our custom grayscale conversion results
    Dim LookUp(0 To 255) As Byte
    Dim grayTempCalc As Long
    
    For x = 0 To 255
        grayTempCalc = Int((CDbl(x) / conversionFactor) + 0.5) * conversionFactor
        ByteMeL grayTempCalc
        LookUp(x) = CByte(grayTempCalc)
    Next x
    
    'Build another look-up table for our initial grayscale index calculation
    Dim grayLookUp(0 To 765) As Byte
    For x = 0 To 765
        grayLookUp(x) = x \ 3
    Next x
    
    Dim r As Long, g As Long, b As Long, gray As Long
    
    'Calculate loop values based on previewing/not-previewing
    Dim initX As Long, initY As Long, finX As Long, finY As Long
    If toPreview = True Then
        initX = PreviewX
        finX = PreviewX + PreviewWidth
        initY = PreviewY
        finY = PreviewY + PreviewHeight
    Else
        initX = 0
        finX = PicWidthL
        initY = 0
        finY = PicHeightL
    End If
    
    'Loop through each pixel in the image, converting values as we go
    Dim QuickVal As Long
    For y = initY To finY
    For x = initX To finX
        
        QuickVal = x * 3
        
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        gray = grayLookUp(r + g + b)
        
        'Assign all color channels the new gray value
        ImageData(QuickVal + 2, y) = LookUp(gray)
        ImageData(QuickVal + 1, y) = LookUp(gray)
        ImageData(QuickVal, y) = LookUp(gray)
        
    Next x
        If toPreview = False Then
            If (y Mod 20 = 0) Then SetProgBarVal y
        End If
    Next y
    
    'Render the finished output to the appropriate image container
    If toPreview = True Then
        SetPreviewData dstPic
    Else
        SetImageData
        Message "Finished."
    End If
    
End Sub

'Reduce to X # gray shades (dithered)
Public Sub fGrayscaleCustomDither(ByVal numOfShades As Long, Optional ByVal toPreview As Boolean = False, Optional ByRef srcPic As PictureBox, Optional ByRef dstPic As PictureBox)
    
    'Get the appropriate set of image data contingent on whether this is a preview or not
    If toPreview = True Then
        GetPreviewData srcPic
    Else
        Message "Converting to " & numOfShades & " shades of gray, with dithering..."
        GetImageData
        SetProgBarMax PicHeightL
    End If
    
    'This conversion factor is the value we need to turn grayscale values in the [0,255] range into a specific subset of values
    Dim conversionFactor As Single
    conversionFactor = (255 / (numOfShades - 1))
    
    'Build another look-up table for our initial grayscale index calculation
    Dim grayLookUp(0 To 765) As Byte
    For x = 0 To 765
        grayLookUp(x) = x \ 3
    Next x
    
    'Unfortunately, this algorithm (unlike its non-dithering counterpart) is not well-suited to using a look-up table, so all calculations have been moved into the loop
    Dim grayTempCalc As Long
    
    'This value tracks the drifting error of our conversions, which allows us to dither
    Dim errorValue As Long
    errorValue = 0
    
    'Color and grayscale variables
    Dim r As Long, g As Long, b As Long
    Dim gray As Byte

    'Calculate loop values based on previewing/not-previewing
    Dim initX As Long, initY As Long, finX As Long, finY As Long
    If toPreview = True Then
        initX = PreviewX
        finX = PreviewX + PreviewWidth
        initY = PreviewY
        finY = PreviewY + PreviewHeight
    Else
        initX = 0
        finX = PicWidthL
        initY = 0
        finY = PicHeightL
    End If
    
    'Loop through each pixel in the image, converting values as we go
    Dim QuickVal As Long
    For y = initY To finY
    For x = initX To finX
    
        QuickVal = x * 3
        
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Look up our initial grayscale value in the table
        gray = grayLookUp(r + g + b)
        grayTempCalc = gray
        
        'Add the error value (a cumulative value of the difference between actual gray values and gray values we've selected) to the current gray value
        grayTempCalc = grayTempCalc + errorValue
        
        'Rebuild our temporary calculation variable using the shade reduction formula
        grayTempCalc = Int((CDbl(grayTempCalc) / conversionFactor) + 0.5) * conversionFactor
        
        'Adjust our error value to include this latest calculation
        errorValue = CLng(gray) + errorValue - grayTempCalc
        
        ByteMeL grayTempCalc
        gray = CByte(grayTempCalc)
        
        'Assign all color channels the new gray value
        ImageData(QuickVal + 2, y) = gray
        ImageData(QuickVal + 1, y) = gray
        ImageData(QuickVal, y) = gray
        
    Next x
    
        'Reset the error value at the end of each line
        errorValue = 0
        
        'If we aren't previewing, update the progress bar
        If toPreview = False Then
            If (y Mod 20 = 0) Then SetProgBarVal y
        End If
        
    Next y
    
    'Render the finished output to the appropriate image container
    If toPreview = True Then
        SetPreviewData dstPic
    Else
        SetImageData
        Message "Finished."
    End If
    
End Sub

'Reduce to gray via (r+g+b)/3
Public Sub MenuGrayscaleAverage(Optional ByVal toPreview As Boolean = False, Optional ByRef srcPic As PictureBox, Optional ByRef dstPic As PictureBox)
    
    'Get the appropriate set of image data contingent on whether this is a preview or not
    If toPreview = True Then
        GetPreviewData srcPic
    Else
        Message "Converting image to grayscale..."
        GetImageData
        SetProgBarMax PicWidthL
    End If
    
    'Color and grayscale variables
    Dim r As Long, g As Long, b As Long
    Dim gray As Byte
    
    'Build a look-up table of grayscale values (faster than calculating it manually for each pixel)
    Dim grayLookUp(0 To 765) As Byte
    For x = 0 To 765
        grayLookUp(x) = x \ 3
    Next x
    
    'Calculate loop values based on previewing/not-previewing
    Dim initX As Long, initY As Long, finX As Long, finY As Long
    If toPreview = True Then
        initX = PreviewX
        finX = PreviewX + PreviewWidth
        initY = PreviewY
        finY = PreviewY + PreviewHeight
    Else
        initX = 0
        finX = PicWidthL
        initY = 0
        finY = PicHeightL
    End If
    
    'Loop through each pixel in the image, converting values as we go
    Dim QuickVal As Long
    For x = initX To finX
        QuickVal = x * 3
    For y = initY To finY
    
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Calculate the gray value using the look-up table
        gray = grayLookUp(r + g + b)
        
        'Assign that gray value to each color channel
        ImageData(QuickVal, y) = gray
        ImageData(QuickVal + 1, y) = gray
        ImageData(QuickVal + 2, y) = gray
        
    Next y
        If toPreview = False Then
            If (x Mod 20 = 0) Then SetProgBarVal x
        End If
    Next x
    
    'Render the finished output to the appropriate image container
    If toPreview = True Then
        SetPreviewData dstPic
    Else
        SetImageData
        Message "Finished."
    End If
    
End Sub

'Reduce to gray in a more human-eye friendly manner
Public Sub MenuGrayscale(Optional ByVal toPreview As Boolean = False, Optional ByRef srcPic As PictureBox, Optional ByRef dstPic As PictureBox)

    'Get the appropriate set of image data contingent on whether this is a preview or not
    If toPreview = True Then
        GetPreviewData srcPic
    Else
        Message "Generating ITU-R compatible grayscale image..."
        GetImageData
        SetProgBarMax PicWidthL
    End If

    'Color and gray variables
    Dim r As Long, g As Long, b As Long
    Dim gray As Long
    
    'Calculate loop values based on previewing/not-previewing
    Dim initX As Long, initY As Long, finX As Long, finY As Long
    If toPreview = True Then
        initX = PreviewX
        finX = PreviewX + PreviewWidth
        initY = PreviewY
        finY = PreviewY + PreviewHeight
    Else
        initX = 0
        finX = PicWidthL
        initY = 0
        finY = PicHeightL
    End If
    
    'Loop through each pixel in the image, converting values as we go
    Dim QuickVal As Long
    For x = initX To finX
        QuickVal = x * 3
    For y = initY To finY
    
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Calculate a grayscale value using the original ITU-R recommended formula (BT.709, specifically)
        gray = (213 * r + 715 * g + 72 * b) \ 1000
        ByteMeL gray
        
        'Assign all color channels the new gray value
        ImageData(QuickVal, y) = gray
        ImageData(QuickVal + 1, y) = gray
        ImageData(QuickVal + 2, y) = gray
        
    Next y
        If toPreview = False Then
            If (x Mod 20 = 0) Then SetProgBarVal x
        End If
    Next x
    
    'Render the finished output to the appropriate image container
    If toPreview = True Then
        SetPreviewData dstPic
    Else
        SetImageData
        Message "Finished."
    End If
    
End Sub

'Reduce to gray via HSL -> convert S to 0
Public Sub MenuDesaturate(Optional ByVal toPreview As Boolean = False, Optional ByRef srcPic As PictureBox, Optional ByRef dstPic As PictureBox)
    
    'Get the appropriate set of image data contingent on whether this is a preview or not
    If toPreview = True Then
        GetPreviewData srcPic
    Else
        Message "Desaturating image..."
        GetImageData
        SetProgBarMax PicWidthL
    End If
    
    'These variables will hold temporary pixel color values (red, green, blue)
    Dim r As Long, g As Long, b As Long
    
    'This value will hold the grayscale value of each pixel
    Dim gray As Long
    
    'These variables will hold the maximum and minimum channel values for each pixel
    Dim cMax As Long, cMin As Long
    
    'Calculate initial and terminal loop values based on previewing/not-previewing
    Dim initX As Long, initY As Long, finX As Long, finY As Long
    If toPreview = True Then
        initX = PreviewX
        finX = PreviewX + PreviewWidth
        initY = PreviewY
        finY = PreviewY + PreviewHeight
    Else
        initX = 0
        finX = PicWidthL
        initY = 0
        finY = PicHeightL
    End If
    
    'Loop through each pixel in the image, converting values as we go
    Dim QuickVal As Long
    For x = initX To finX
        QuickVal = x * 3
    For y = initY To finY
    
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Find the highest and lowest of the RGB values
        cMax = Maximum(r, g, b)
        cMin = Minimum(r, g, b)
        
        'Calculate luminance and make sure it falls between 255 and 0 (it always should, but it doesn't hurt to check)
        gray = (cMax + cMin) \ 2
        ByteMeL gray
        
        'Assign all color channels to the new gray value
        ImageData(QuickVal + 2, y) = CByte(gray)
        ImageData(QuickVal + 1, y) = CByte(gray)
        ImageData(QuickVal, y) = CByte(gray)
    Next y
        If toPreview = False Then
            If (x Mod 20 = 0) Then SetProgBarVal x
        End If
    Next x
    
    'Render the finished output to the appropriate image container
    If toPreview = True Then
        SetPreviewData dstPic
    Else
        SetImageData
        Message "Finished."
    End If
    
End Sub

'Reduce to gray by selecting the minimum (maxOrMin = 0) or maximum (maxOrMin = 1) color in each pixel
Public Sub MenuDecompose(ByVal maxOrMin As Long, Optional ByVal toPreview As Boolean = False, Optional ByRef srcPic As PictureBox, Optional ByRef dstPic As PictureBox)
    
    'Get the appropriate set of image data contingent on whether this is a preview or not
    If toPreview = True Then
        GetPreviewData srcPic
    Else
        Message "Decomposing image..."
        GetImageData
        SetProgBarMax PicWidthL
    End If
    
    'These variables will hold temporary pixel color values (red, green, blue)
    Dim r As Long, g As Long, b As Long
    
    'This value will hold the grayscale value of each pixel
    Dim gray As Long
    
    'Calculate initial and terminal loop values based on previewing/not-previewing
    Dim initX As Long, initY As Long, finX As Long, finY As Long
    If toPreview = True Then
        initX = PreviewX
        finX = PreviewX + PreviewWidth
        initY = PreviewY
        finY = PreviewY + PreviewHeight
    Else
        initX = 0
        finX = PicWidthL
        initY = 0
        finY = PicHeightL
    End If
    
    'Loop through each pixel in the image, converting values as we go
    Dim QuickVal As Long
    For x = initX To finX
        QuickVal = x * 3
    For y = initY To finY
    
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Find the highest or lowest of the RGB values
        If maxOrMin = 0 Then gray = Minimum(r, g, b) Else gray = Maximum(r, g, b)
        
        'Assign all color channels to the new gray value
        ImageData(QuickVal + 2, y) = CByte(gray)
        ImageData(QuickVal + 1, y) = CByte(gray)
        ImageData(QuickVal, y) = CByte(gray)
    Next y
        If toPreview = False Then
            If (x Mod 20 = 0) Then SetProgBarVal x
        End If
    Next x
    
    'Render the finished output to the appropriate image container
    If toPreview = True Then
        SetPreviewData dstPic
    Else
        SetImageData
        Message "Finished."
    End If
    
End Sub

'Reduce to gray by selecting a single color channel (represeted by cChannel: 0 = Red, 1 = Green, 2 = Blue)
Public Sub MenuGrayscaleSingleChannel(ByVal cChannel As Long, Optional ByVal toPreview As Boolean = False, Optional ByRef srcPic As PictureBox, Optional ByRef dstPic As PictureBox)
    
    'Get the appropriate set of image data contingent on whether this is a preview or not
    If toPreview = True Then
        GetPreviewData srcPic
    Else
        Message "Converting to grayscale by isolating single color channel..."
        GetImageData
        SetProgBarMax PicWidthL
    End If
    
    'These variables will hold temporary pixel color values (red, green, blue)
    Dim r As Long, g As Long, b As Long
    
    'This value will hold the grayscale value of each pixel
    Dim gray As Byte
    
    'Calculate initial and terminal loop values based on previewing/not-previewing
    Dim initX As Long, initY As Long, finX As Long, finY As Long
    If toPreview = True Then
        initX = PreviewX
        finX = PreviewX + PreviewWidth
        initY = PreviewY
        finY = PreviewY + PreviewHeight
    Else
        initX = 0
        finX = PicWidthL
        initY = 0
        finY = PicHeightL
    End If
    
    'Loop through each pixel in the image, converting values as we go
    Dim QuickVal As Long
    For x = initX To finX
        QuickVal = x * 3
    For y = initY To finY
    
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Assign the gray value to a single color channel based on the value of cChannel
        Select Case cChannel
            Case 0
                gray = r
            Case 1
                gray = g
            Case 2
                gray = b
        End Select
        
        'Assign all color channels to the new gray value
        ImageData(QuickVal + 2, y) = gray
        ImageData(QuickVal + 1, y) = gray
        ImageData(QuickVal, y) = gray
    Next y
        If toPreview = False Then
            If (x Mod 20 = 0) Then SetProgBarVal x
        End If
    Next x
    
    'Render the finished output to the appropriate image container
    If toPreview = True Then
        SetPreviewData dstPic
    Else
        SetImageData
        Message "Finished."
    End If
    
End Sub

'When the "# of shades" horizontal scroll bar is changed, update the text box to match
Private Sub hsShades_Change()
    txtShades.Text = hsShades.Value
End Sub

Private Sub hsShades_Scroll()
    txtShades.Text = hsShades.Value
End Sub

'When option buttons are used, update the preview accordingly
Private Sub optChannel_Click(Index As Integer)
    drawGrayscalePreview
End Sub

Private Sub optDecompose_Click(Index As Integer)
    drawGrayscalePreview
End Sub

'When the "# of shades" text box is changed, check the value for errors and redraw the preview
Private Sub txtShades_Change()
    If EntryValid(txtShades, hsShades.Min, hsShades.Max, False, False) Then
        hsShades.Value = val(txtShades)
        drawGrayscalePreview
    End If
End Sub

'As a convenience to the user, when they click the "# of shades" text box, automatically select the text for them
Private Sub txtShades_GotFocus()
    AutoSelectText txtShades
End Sub

'Return the maximum of three Long-type variables
Private Function Maximum(rR As Long, rG As Long, rB As Long) As Long
   If (rR > rG) Then
      If (rR > rB) Then
         Maximum = rR
      Else
         Maximum = rB
      End If
   Else
      If (rB > rG) Then
         Maximum = rB
      Else
         Maximum = rG
      End If
   End If
End Function

'Return the minimum of three Long-type variables
Private Function Minimum(rR As Long, rG As Long, rB As Long) As Long
   If (rR < rG) Then
      If (rR < rB) Then
         Minimum = rR
      Else
         Minimum = rB
      End If
   Else
      If (rB < rG) Then
         Minimum = rB
      Else
         Minimum = rG
      End If
   End If
End Function
