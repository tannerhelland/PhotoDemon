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
      Left            =   3000
      Max             =   254
      Min             =   3
      TabIndex        =   2
      Top             =   4200
      Value           =   3
      Width           =   2865
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
      Left            =   2400
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
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   3600
      Width           =   3495
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
      Left            =   600
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
      Left            =   600
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
                MenuGrayscaleAverage True
            Case 1
                MenuGrayscale True
            Case 2
                MenuDesaturate True
            Case 3
                fGrayscaleCustom hsShades.Value, True
            Case 4
                fGrayscaleCustomDither hsShades.Value, True
        End Select
        
    End If

End Sub

'*********************************************************************
'The next three methods exist to activate/deactivate the textbox and scrollbar
Private Sub cboMethod_Click()
    If cboMethod.ListIndex = 3 Or cboMethod.ListIndex = 4 Then
        ShowControls True
    Else
        ShowControls False
    End If
    drawGrayscalePreview
End Sub

Private Sub cboMethod_KeyDown(KeyCode As Integer, Shift As Integer)
    If cboMethod.ListIndex = 3 Or cboMethod.ListIndex = 4 Then
        ShowControls True
    Else
        ShowControls False
    End If
    drawGrayscalePreview
End Sub

Private Sub ShowControls(ByVal toShow As Boolean)
    txtShades.Visible = toShow
    hsShades.Visible = toShow
    lblAdditional.Visible = toShow
End Sub
'*********************************************************************

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
                Process GrayscaleCustom, hsShades.Value
            Case 4
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
    cboMethod.AddItem "Average", 0
    cboMethod.AddItem "ITU Standard", 1
    cboMethod.AddItem "Desaturate", 2
    cboMethod.AddItem "X # of shades", 3
    cboMethod.AddItem "X # of shades (dithered)", 4
    cboMethod.ListIndex = 1
    
    'Render the preview images
    DrawPreviewImage PicPreview
    DrawPreviewImage PicEffect
    
    'Assign the system hand cursor to all relevant objects
    setHandCursorForAll Me
    
    'Draw the initial preview
    drawGrayscalePreview
    
End Sub

'Reduce to X # gray shades
Public Sub fGrayscaleCustom(ByVal NumToConvertTo As Long, Optional ByVal toPreview As Boolean = False)
    
    'Get the appropriate set of image data contingent on whether this is a preview or not
    If toPreview = True Then
        GetPreviewData PicPreview
    Else
        Message "Converting to " & NumToConvertTo & " shades of gray..."
        GetImageData
        SetProgBarMax PicHeightL
    End If
    
    'Build a look-up table for our custom grayscale conversion results
    Dim MagicNum As Single
    MagicNum = (255 / (NumToConvertTo - 1))
    Dim LookUp(0 To 255) As Long
    For x = 0 To 255
        LookUp(x) = Int((CDbl(x) / MagicNum) + 0.5) * MagicNum
        If LookUp(x) > 255 Then LookUp(x) = 255
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
    
    Dim QuickVal As Long
    For y = initY To finY
    For x = initX To finX
        
        QuickVal = x * 3
        
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        gray = grayLookUp(r + g + b)
        
        ImageData(QuickVal + 2, y) = LookUp(gray)
        ImageData(QuickVal + 1, y) = LookUp(gray)
        ImageData(QuickVal, y) = LookUp(gray)
        
    Next x
        If toPreview = False Then
            If (y Mod 20 = 0) Then SetProgBarVal y
        End If
    Next y
    
    If toPreview = True Then
        SetPreviewData PicEffect
    Else
        SetImageData
        Message "Finished."
    End If
    
End Sub

'Reduce to X # gray shades (dithered)
Public Sub fGrayscaleCustomDither(ByVal NumToConvertTo As Long, Optional ByVal toPreview As Boolean = False)
    
    'Get the appropriate set of image data contingent on whether this is a preview or not
    If toPreview = True Then
        GetPreviewData PicPreview
    Else
        Message "Converting to " & NumToConvertTo & " shades of gray, with dithering..."
        GetImageData
        SetProgBarMax PicHeightL
    End If
    
    Dim MagicNum As Single
    MagicNum = (255 / (NumToConvertTo - 1))
    
    Dim CurrentColor As Long
    
    Dim EV As Long
    Dim CC As Long
    Dim r1 As Long, g1 As Long, b1 As Long

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
    
    Dim QuickVal As Long
    For y = initY To finY
    For x = initX To finX
        QuickVal = x * 3
        r1 = ImageData(QuickVal + 2, y)
        g1 = ImageData(QuickVal + 1, y)
        b1 = ImageData(QuickVal, y)
        CurrentColor = (r1 + g1 + b1) \ 3
        CurrentColor = CurrentColor + EV
        CC = Int((CDbl(CurrentColor) / MagicNum) + 0.5) * MagicNum
        EV = CurrentColor - CC
        If CC > 255 Then CC = 255
        If CC < 0 Then CC = 0
        ImageData(QuickVal + 2, y) = CC
        ImageData(QuickVal + 1, y) = CC
        ImageData(QuickVal, y) = CC
    Next x
        EV = 0
        If toPreview = False Then
            If (y Mod 20 = 0) Then SetProgBarVal y
        End If
    Next y
    
    If toPreview = True Then
        SetPreviewData PicEffect
    Else
        SetImageData
        Message "Finished."
    End If
    
End Sub

'Reduce to gray via (r+g+b)/3
Public Sub MenuGrayscaleAverage(Optional ByVal toPreview As Boolean = False)
    
    'Get the appropriate set of image data contingent on whether this is a preview or not
    If toPreview = True Then
        GetPreviewData PicPreview
    Else
        Message "Converting image to grayscale..."
        GetImageData
        SetProgBarMax PicWidthL
    End If
    
    Dim CurrentColor As Byte
    Dim r As Long, g As Long, b As Long
    
    Dim LookUp(0 To 765) As Byte
    For x = 0 To 765
        LookUp(x) = x \ 3
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
    
    Dim QuickVal As Long
    For x = initX To finX
        QuickVal = x * 3
    For y = initY To finY
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        CurrentColor = LookUp(r + g + b)
        ImageData(QuickVal, y) = CurrentColor
        ImageData(QuickVal + 1, y) = CurrentColor
        ImageData(QuickVal + 2, y) = CurrentColor
    Next y
        If toPreview = False Then
            If (x Mod 20 = 0) Then SetProgBarVal x
        End If
    Next x
    
    If toPreview = True Then
        SetPreviewData PicEffect
    Else
        SetImageData
        Message "Finished."
    End If
    
End Sub

'Reduce to gray in a more human-eye friendly manner
Public Sub MenuGrayscale(Optional ByVal toPreview As Boolean = False)

    'Get the appropriate set of image data contingent on whether this is a preview or not
    If toPreview = True Then
        GetPreviewData PicPreview
    Else
        Message "Generating ITU standard grayscale image..."
        GetImageData
        SetProgBarMax PicWidthL
    End If

    Dim CurrentColor As Long
    Dim r As Long, g As Long, b As Long
    
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
    
    Dim QuickVal As Long
    For x = initX To finX
        QuickVal = x * 3
    For y = initY To finY
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        CurrentColor = (222 * r + 707 * g + 71 * b) \ 1000
        ByteMeL CurrentColor
        ImageData(QuickVal, y) = CurrentColor
        ImageData(QuickVal + 1, y) = CurrentColor
        ImageData(QuickVal + 2, y) = CurrentColor
    Next y
        If toPreview = False Then
            If (x Mod 20 = 0) Then SetProgBarVal x
        End If
    Next x
    
    If toPreview = True Then
        SetPreviewData PicEffect
    Else
        SetImageData
        Message "Finished."
    End If
    
End Sub

'Reduce to gray via HSL -> convert S to 0
Public Sub MenuDesaturate(Optional ByVal toPreview As Boolean = False)
    
    'Get the appropriate set of image data contingent on whether this is a preview or not
    If toPreview = True Then
        GetPreviewData PicPreview
    Else
        Message "Desaturating image..."
        GetImageData
        SetProgBarMax PicWidthL
    End If
    
    Dim r As Long, g As Long, b As Long
    Dim HH As Single, SS As Single, LL As Single
    
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
    
    Dim QuickVal As Long
    For x = initX To finX
        QuickVal = x * 3
    For y = initY To finY
    
        'Get the temporary values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Use the RGB values to calculate corresponding hue, saturation, and luminance values
        tRGBToHSL r, g, b, HH, SS, LL
        
        'Set saturation to zero, then convert HSL back into RGB values
        tHSLToRGB HH, 0, LL, r, g, b
        
        'Assign those values into the array
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
    Next y
        If toPreview = False Then
            If (x Mod 20 = 0) Then SetProgBarVal x
        End If
    Next x
    
    If toPreview = True Then
        SetPreviewData PicEffect
    Else
        SetImageData
        Message "Finished."
    End If
    
End Sub

Private Sub hsShades_Change()
    txtShades.Text = hsShades.Value
End Sub

Private Sub hsShades_Scroll()
    txtShades.Text = hsShades.Value
End Sub

Private Sub txtShades_Change()
    If EntryValid(txtShades, hsShades.Min, hsShades.Max, False, False) Then
        hsShades.Value = val(txtShades)
        drawGrayscalePreview
    End If
End Sub

Private Sub txtShades_GotFocus()
    AutoSelectText txtShades
End Sub
