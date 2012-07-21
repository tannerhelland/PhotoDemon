VERSION 5.00
Begin VB.Form FormReduceColors 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Reduce Image Colors"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6480
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
   ScaleHeight     =   332
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   432
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.HScrollBar hsBlue 
      Height          =   255
      Left            =   3360
      Max             =   256
      Min             =   1
      MouseIcon       =   "VBP_FormReduceColors.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   2895
      Value           =   6
      Width           =   2895
   End
   Begin VB.HScrollBar hsGreen 
      Height          =   255
      Left            =   3360
      Max             =   256
      Min             =   1
      MouseIcon       =   "VBP_FormReduceColors.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   2535
      Value           =   7
      Width           =   2895
   End
   Begin VB.HScrollBar hsRed 
      Height          =   255
      Left            =   3360
      Max             =   256
      Min             =   1
      MouseIcon       =   "VBP_FormReduceColors.frx":02A4
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2175
      Value           =   6
      Width           =   2895
   End
   Begin VB.OptionButton OptQuant 
      Appearance      =   0  'Flat
      Caption         =   "PhotoDemon Advanced (manual)"
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   2
      Left            =   840
      MouseIcon       =   "VBP_FormReduceColors.frx":03F6
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1320
      Width           =   2775
   End
   Begin VB.TextBox TxtB 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2640
      TabIndex        =   7
      Text            =   "6"
      Top             =   2880
      Width           =   615
   End
   Begin VB.TextBox TxtG 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2640
      TabIndex        =   5
      Text            =   "7"
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox TxtR 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2640
      TabIndex        =   3
      Text            =   "6"
      Top             =   2160
      Width           =   615
   End
   Begin VB.CheckBox ChkColorDither 
      Appearance      =   0  'Flat
      Caption         =   "Use error diffusion dithering"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      MouseIcon       =   "VBP_FormReduceColors.frx":0548
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   3840
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.CheckBox chkSmartColors 
      Appearance      =   0  'Flat
      Caption         =   "Use intelligent coloring"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3480
      MouseIcon       =   "VBP_FormReduceColors.frx":069A
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   3840
      Value           =   1  'Checked
      Width           =   2535
   End
   Begin VB.OptionButton OptQuant 
      Appearance      =   0  'Flat
      Caption         =   "Xiaolin Wu (automatic)"
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   0
      Left            =   840
      MouseIcon       =   "VBP_FormReduceColors.frx":07EC
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   600
      Value           =   -1  'True
      Width           =   2055
   End
   Begin VB.OptionButton OptQuant 
      Appearance      =   0  'Flat
      Caption         =   "NeuQuant by Anthony Dekker (automatic)"
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   1
      Left            =   840
      MouseIcon       =   "VBP_FormReduceColors.frx":093E
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   960
      Width           =   3375
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      MouseIcon       =   "VBP_FormReduceColors.frx":0A90
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   4440
      Width           =   1125
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5160
      MouseIcon       =   "VBP_FormReduceColors.frx":0BE2
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   4440
      Width           =   1125
   End
   Begin VB.Label lblWarning 
      Caption         =   "Note: some options on this page have been disabled because the FreeImage plugin could not be found."
      ForeColor       =   &H000000C0&
      Height          =   855
      Left            =   4320
      TabIndex        =   19
      Top             =   600
      Visible         =   0   'False
      Width           =   1935
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblOptions 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PhotoDemon Advanced Quantization options:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   240
      TabIndex        =   18
      Top             =   1800
      Width           =   3825
   End
   Begin VB.Label lblBlue 
      BackStyle       =   0  'Transparent
      Caption         =   "Possible blue values (1-256):"
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   2925
      Width           =   2175
   End
   Begin VB.Label lblGreen 
      BackStyle       =   0  'Transparent
      Caption         =   "Possible green values (1-256):"
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   2565
      Width           =   2535
   End
   Begin VB.Label lblRed 
      BackStyle       =   0  'Transparent
      Caption         =   "Possible red values (1-256):"
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   2205
      Width           =   2175
   End
   Begin VB.Label lblMaxColors 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "These parameters allow for a maximum of 252 colors in the quantized image."
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   360
      TabIndex        =   14
      Top             =   3360
      Width           =   5505
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Quantization method:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   240
      TabIndex        =   13
      Top             =   240
      Width           =   1830
   End
End
Attribute VB_Name = "FormReduceColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Color Reduction Form
'Copyright ©2000-2012 by Tanner Helland
'Created: 4/October/00
'Last updated: 20/June/12
'Last update: Rewrote much of the code to align with the new UI.
'
'In the original incarnation of PhotoDemon, this was a central part of the
'project. I have since not used it as much (since the project has become
'almost entirely cenetered around 24-bit imaging), but the code is solid
'and the feature set is large.
'
'***************************************************************************

Option Explicit

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub CmdOK_Click()

    'Check to see which method the user has requested
    
    'Xiaolin Wu
    If OptQuant(0).Value = True Then
        FormReduceColors.Visible = False
        Process ReduceColors, REDUCECOLORS_AUTO, FIQ_WUQUANT
        Unload Me
    
    'NeuQuant
    ElseIf OptQuant(1).Value = True Then
        FormReduceColors.Visible = False
        Process ReduceColors, REDUCECOLORS_AUTO, FIQ_NNQUANT
        Unload Me
    
    'Manual
    Else
    
        'Before reducing anything, check to make sure the textboxes have valid input
        If Not EntryValid(TxtR, hsRed.Min, hsRed.Max, True, True) Then
            AutoSelectText TxtR
            Exit Sub
        End If
        If Not EntryValid(TxtG, hsGreen.Min, hsGreen.Max, True, True) Then
            AutoSelectText TxtG
            Exit Sub
        End If
        If Not EntryValid(TxtB, hsBlue.Min, hsBlue.Max, True, True) Then
            AutoSelectText TxtB
            Exit Sub
        End If
        
        Me.Visible = False
        
        'Do the appropriate method of color reduction
        If ChkColorDither.Value = vbUnchecked Then
            If chkSmartColors.Value = vbUnchecked Then
                Process ReduceColors, REDUCECOLORS_MANUAL, TxtR, TxtG, TxtB, False
            Else
                Process ReduceColors, REDUCECOLORS_MANUAL, TxtR, TxtG, TxtB, True
            End If
        Else
            If chkSmartColors.Value = vbUnchecked Then
                Process ReduceColors, REDUCECOLORS_MANUAL_ERRORDIFFUSION, TxtR, TxtG, TxtB, False
            Else
                Process ReduceColors, REDUCECOLORS_MANUAL_ERRORDIFFUSION, TxtR, TxtG, TxtB, True
            End If
        End If
        
        Unload Me
        
    End If
    
End Sub

Private Sub Form_Load()
    'Only allow AutoReduction stuff if the FreeImage dll was found.
    If FreeImageEnabled = False Then
        OptQuant(0).Enabled = False
        OptQuant(1).Enabled = False
        OptQuant(2).Value = True
        DisplayManualOptions True
        lblWarning.Visible = True
    Else
        OptQuant(0).Value = True
        DisplayManualOptions False
    End If
End Sub

'Enable/disable the manual settings depending on which option button has been selected
Private Sub OptQuant_Click(Index As Integer)
    If OptQuant(2).Value = True Then DisplayManualOptions True Else DisplayManualOptions False
End Sub

Private Sub OptQuant_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If OptQuant(2).Value = True Then DisplayManualOptions True Else DisplayManualOptions False
End Sub

'The large chunk of subs that follow serve to keep the text box and scroll bar values in lock-step
Private Sub hsRed_Change()
    TxtR = hsRed.Value
End Sub

Private Sub hsGreen_Change()
    TxtG = hsGreen.Value
End Sub

Private Sub hsBlue_Change()
    TxtB = hsBlue.Value
End Sub

Private Sub hsRed_Scroll()
    TxtR = hsRed.Value
End Sub

Private Sub hsGreen_Scroll()
    TxtG = hsGreen.Value
End Sub

Private Sub hsBlue_Scroll()
    TxtB = hsBlue.Value
End Sub

Private Sub TxtB_Change()
    If EntryValid(TxtB, hsBlue.Min, hsBlue.Max, False, False) Then hsBlue.Value = val(TxtB)
    updateColorLabel
End Sub

Private Sub TxtB_GotFocus()
    AutoSelectText TxtB
End Sub

Private Sub TxtG_Change()
    If EntryValid(TxtG, hsGreen.Min, hsGreen.Max, False, False) Then hsGreen.Value = val(TxtG)
    updateColorLabel
End Sub

Private Sub TxtG_GotFocus()
    AutoSelectText TxtG
End Sub

Private Sub TxtR_Change()
    If EntryValid(TxtR, hsRed.Min, hsRed.Max, False, False) Then hsRed.Value = val(TxtR)
    updateColorLabel
End Sub

Private Sub TxtR_GotFocus()
    AutoSelectText TxtR
End Sub

'This lets the user know the max number of colors that the current set of quantization parameters will allow for
Private Sub updateColorLabel()
    If EntryValid(TxtR, hsRed.Min, hsRed.Max, False, False) And EntryValid(TxtG, hsGreen.Min, hsGreen.Max, False, False) And EntryValid(TxtB, hsBlue.Min, hsBlue.Max, False, False) Then
        lblMaxColors = "These parameters allow for a maximum of " & val(TxtR) * val(TxtG) * val(TxtB) & " colors in the quantized image."
    Else
        lblMaxColors = "Color count could not be calculated due to invalid text box values."
    End If
End Sub

'Enable/disable the manual options depending on which quantization method has been selected
Private Sub DisplayManualOptions(ByVal toDisplay As Boolean)
    If toDisplay = False Then
        lblOptions.ForeColor = RGB(160, 160, 160)
        lblRed.ForeColor = RGB(160, 160, 160)
        lblGreen.ForeColor = RGB(160, 160, 160)
        lblBlue.ForeColor = RGB(160, 160, 160)
        lblMaxColors.ForeColor = RGB(160, 160, 160)
        TxtR.Enabled = False
        TxtG.Enabled = False
        TxtB.Enabled = False
        hsRed.Enabled = False
        hsGreen.Enabled = False
        hsBlue.Enabled = False
        ChkColorDither.Enabled = False
        chkSmartColors.Enabled = False
    Else
        lblOptions.ForeColor = &H400000
        lblRed.ForeColor = &H400000
        lblGreen.ForeColor = &H400000
        lblBlue.ForeColor = &H400000
        lblMaxColors.ForeColor = &H400000
        TxtR.Enabled = True
        TxtG.Enabled = True
        TxtB.Enabled = True
        hsRed.Enabled = True
        hsGreen.Enabled = True
        hsBlue.Enabled = True
        ChkColorDither.Enabled = True
        chkSmartColors.Enabled = True
    End If
End Sub

'Automatic 8-bit color reduction via the FreeImage DLL.
Public Sub ReduceImageColors_Auto(ByVal qMethod As Long)

    'Make sure we found the plug-in when we loaded the program
    If FreeImageEnabled = False Then
        MsgBox "The FreeImage interface plug-ins (FreeImage.dll) was marked as missing or corrupted upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy this file into the plug-in directory and reload " & PROGRAMNAME & ".", vbCritical + vbOKOnly + vbApplicationModal, PROGRAMNAME & " FreeImage Interface Error"
        Unload FormMain.ActiveForm
        Exit Sub
    End If
    
    'Load the FreeImage dll into memory
    Dim hLib As Long
    hLib = LoadLibrary(PluginPath & "FreeImage.dll")
    
    Message "Preparing image..."
    
    'Dump the image to a temporary file (a temporary work-around so we can capture
    'the handle required by FreeImage)
    Dim temporaryImg As String
    temporaryImg = TempPath & "PDQuantize.bmp"
    SavePicture FormMain.ActiveForm.BackBuffer.Picture, temporaryImg
    
    'These two variables will hold pointers to the bitmaps created by FreeImage calls
    Dim initDib As Long
    Dim quantizeDib As Long

    'Load the temp image into FreeImage
    initDib = FreeImage_LoadEx(temporaryImg)

    'Quantize the image
    Message "Quantizing image..."
    quantizeDib = FreeImage_ColorQuantizeEx(initDib, qMethod)
    
    'Paint the quantized image to the current picture box
    Message "Rendering..."
    Dim PaintReturn As Long
    PaintReturn = FreeImage_PaintDC(FormMain.ActiveForm.BackBuffer.hdc, quantizeDib)
    FormMain.ActiveForm.BackBuffer.Picture = FormMain.ActiveForm.BackBuffer.Image
    FormMain.ActiveForm.BackBuffer.Refresh
    ScrollViewport FormMain.ActiveForm
    
    'Clear out the images generated by FreeImage
    If initDib <> 0 Then FreeImage_UnloadEx initDib
    If quantizeDib <> 0 Then FreeImage_UnloadEx quantizeDib
    
    'Release the library
    FreeLibrary hLib
    
    'Delete the temp file
    If FileExist(temporaryImg) Then Kill temporaryImg
    
    SetProgBarVal 0
    Message "Colors reduced successfully"
End Sub

'Bit RGB color reduction (no error diffusion)
Public Sub ReduceImageColors_BitRGB(ByVal rValue As Byte, ByVal gValue As Byte, ByVal bValue As Byte, Optional ByVal smartColors As Boolean = False)

    Dim r As Long, g As Long, b As Long
    Dim mR As Single, mG As Single, mB As Single

    Message "Converting picture by manual RGB reduction..."
    
    Dim cR As Integer, cG As Integer, cB As Integer
    
    GetImageData
    SetProgBarMax PicWidthL
    
    'New code for so-called "Intelligent Coloring"
    Dim rLookup() As Long
    Dim gLookup() As Long
    Dim bLookup() As Long
    Dim countLookup() As Long
    
    ReDim rLookup(0 To rValue, 0 To gValue, 0 To bValue) As Long
    ReDim gLookup(0 To rValue, 0 To gValue, 0 To bValue) As Long
    ReDim bLookup(0 To rValue, 0 To gValue, 0 To bValue) As Long
    ReDim countLookup(0 To rValue, 0 To gValue, 0 To bValue) As Long
    
    'For safety, initialize our lookup tables to zero (because we'll possibly access them without first assigning values)
    If smartColors = True Then
        Message "Preparing look-up tables for Intelligent Coloring..."
        For r = 0 To rValue
        For g = 0 To gValue
        For b = 0 To bValue
            rLookup(r, g, b) = 0
            gLookup(r, g, b) = 0
            bLookup(r, g, b) = 0
            countLookup(r, g, b) = 0
        Next b
        Next g
        Next r
    End If
    
    'Prepare inputted variables for the requisite maths
    rValue = rValue - 1
    gValue = gValue - 1
    bValue = bValue - 1
    mR = (256 / rValue)
    mG = (256 / gValue)
    mB = (256 / bValue)
    
    'Faster loop indexing
    Dim QuickX As Long
    
    Message "Analyzing and converting image to specified RGB parameters..."
    
    For x = 0 To PicWidthL
        QuickX = x * 3
    For y = 0 To PicHeightL

        r = ImageData(QuickX + 2, y)
        g = ImageData(QuickX + 1, y)
        b = ImageData(QuickX, y)
        
        'Truncate R, G, and B values (posterize-style) into discreet increments.  0.5 is added for rounding purposes.
        cR = Int((r / mR) + 0.5)
        cG = Int((g / mG) + 0.5)
        cB = Int((b / mB) + 0.5)
        
        'If we're doing Intelligent Coloring, place color values into a look-up table
        If smartColors = True Then
            rLookup(cR, cG, cB) = rLookup(cR, cG, cB) + r
            gLookup(cR, cG, cB) = gLookup(cR, cG, cB) + g
            bLookup(cR, cG, cB) = bLookup(cR, cG, cB) + b
            'Also, keep track of how many colors fall into this bucket (so we can later determine an average color)
            countLookup(cR, cG, cB) = countLookup(cR, cG, cB) + 1
        End If
        
        'Multiply the now-discretely divided R, G, and B values to (0-255) equivalents
        cR = cR * mR
        cG = cG * mG
        cB = cB * mB
        ByteMe cR
        ByteMe cG
        ByteMe cB
        
        'If we are not doing Intelligent Coloring, assign the colors now (to avoid having to do another loop at the end)
        If smartColors = False Then
            ImageData(QuickX + 2, y) = cR
            ImageData(QuickX + 1, y) = cG
            ImageData(QuickX, y) = cB
        End If
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    
    'If we are not doing Intelligent Coloring, draw the pixels and quit
    If smartColors = False Then
        SetImageData
    
        Message "Color reduction complete."
        SetProgBarVal 0
        Exit Sub
    End If
    
    'If we're still here at this point, asssme we are doing Intelligent Coloring
    If smartColors = True Then
    
        Message "Assigning new values from calculated look-up tables..."
        
        'Find average colors based on color counts
        For r = 0 To rValue
        For g = 0 To gValue
        For b = 0 To bValue
            If countLookup(r, g, b) <> 0 Then
                rLookup(r, g, b) = Int(Int(rLookup(r, g, b)) / Int(countLookup(r, g, b)))
                gLookup(r, g, b) = Int(Int(gLookup(r, g, b)) / Int(countLookup(r, g, b)))
                bLookup(r, g, b) = Int(Int(bLookup(r, g, b)) / Int(countLookup(r, g, b)))
                If rLookup(r, g, b) > 255 Then rLookup(r, g, b) = 255
                If gLookup(r, g, b) > 255 Then gLookup(r, g, b) = 255
                If bLookup(r, g, b) > 255 Then bLookup(r, g, b) = 255
            End If
        Next b
        Next g
        Next r
        
        'Assign average colors back into the picture
        For x = 0 To PicWidthL
            QuickX = x * 3
        For y = 0 To PicHeightL
        
            r = ImageData(QuickX + 2, y)
            g = ImageData(QuickX + 1, y)
            b = ImageData(QuickX, y)
            cR = Int((r / mR) + 0.5)
            cG = Int((g / mG) + 0.5)
            cB = Int((b / mB) + 0.5)
            ImageData(QuickX + 2, y) = rLookup(cR, cG, cB)
            ImageData(QuickX + 1, y) = gLookup(cR, cG, cB)
            ImageData(QuickX, y) = bLookup(cR, cG, cB)
        Next y
        Next x
        
        SetImageData
    
        Message "Color reduction complete."
        SetProgBarVal 0
        Exit Sub
    
    End If
    
End Sub

'Error Diffusion dithering to x# shades of color per component
Public Sub ReduceImageColors_BitRGB_ErrorDif(ByVal rValue As Byte, ByVal gValue As Byte, ByVal bValue As Byte, Optional ByVal smartColors As Boolean = False)
    
    Dim r As Long, g As Long, b As Long
    Dim mR As Single, mG As Single, mB As Single
    
    Message "Converting picture by dithered RGB reduction..."
    
    Dim cR As Long, cG As Long, cB As Long
    Dim eR As Long, eG As Long, eB As Long
    Dim Offset As Long
    
    GetImageData
    SetProgBarMax PicHeightL
    
    'New code for so-called "Intelligent Coloring"
    Dim rLookup() As Long
    Dim gLookup() As Long
    Dim bLookup() As Long
    Dim countLookup() As Long
    
    ReDim rLookup(0 To rValue, 0 To gValue, 0 To bValue) As Long
    ReDim gLookup(0 To rValue, 0 To gValue, 0 To bValue) As Long
    ReDim bLookup(0 To rValue, 0 To gValue, 0 To bValue) As Long
    ReDim countLookup(0 To rValue, 0 To gValue, 0 To bValue) As Long
    
    'For safety, initialize our lookup tables to zero (because we'll possibly access them without first assigning values)
    If smartColors = True Then
        Message "Preparing look-up tables for Intelligent Coloring..."
        For r = 0 To rValue
        For g = 0 To gValue
        For b = 0 To bValue
            rLookup(r, g, b) = 0
            gLookup(r, g, b) = 0
            bLookup(r, g, b) = 0
            countLookup(r, g, b) = 0
        Next b
        Next g
        Next r
    End If
    
    'Prepare inputted variables for the requisite maths
    rValue = rValue - 1
    gValue = gValue - 1
    bValue = bValue - 1
    mR = (256 / rValue)
    mG = (256 / gValue)
    mB = (256 / bValue)
    
    Dim QuickX As Long
    
    Message "Analyzing and converting image to specified RGB parameters..."
    
    For y = 0 To PicHeightL
        Offset = y Mod 2
    For x = Offset To PicWidthL
    
        QuickX = x * 3
        
        r = ImageData(QuickX + 2, y)
        g = ImageData(QuickX + 1, y)
        b = ImageData(QuickX, y)
        r = r + eR
        g = g + eG
        b = b + eB
        cR = Int((r / mR) + 0.5)
        cG = Int((g / mG) + 0.5)
        cB = Int((b / mB) + 0.5)
        If cR < 0 Then cR = 0
        If cG < 0 Then cG = 0
        If cB < 0 Then cB = 0
        If cR > 255 Then cR = 255
        If cG > 255 Then cG = 255
        If cB > 255 Then cB = 255
        
        'If we're doing Intelligent Coloring, place color values into a look-up table
        If smartColors = True Then
            rLookup(cR, cG, cB) = rLookup(cR, cG, cB) + r
            gLookup(cR, cG, cB) = gLookup(cR, cG, cB) + g
            bLookup(cR, cG, cB) = bLookup(cR, cG, cB) + b
            'Also, keep track of how many colors fall into this bucket (so we can later determine an average color)
            countLookup(cR, cG, cB) = countLookup(cR, cG, cB) + 1
        End If
        
        'Multiply the now-discretely divided R, G, and B values to (0-255) equivalents
        cR = cR * mR
        cG = cG * mG
        cB = cB * mB
        
        eR = r - cR
        eG = g - cG
        eB = b - cB
        
        ByteMeL cR
        ByteMeL cG
        ByteMeL cB
        
        'If ER < 0 Then ER = 0
        'If EG < 0 Then EG = 0
        'If EB < 0 Then EB = 0
        
        'If we are not doing Intelligent Coloring, assign the colors now (to avoid having to do another loop at the end)
        If smartColors = False Then
            ImageData(QuickX + 2, y) = cR
            ImageData(QuickX + 1, y) = cG
            ImageData(QuickX, y) = cB
        End If
        
    Next x
        eR = 0
        eG = 0
        eB = 0
        If y Mod 20 = 0 Then SetProgBarVal y
    Next y
    
    'If we are not doing Intelligent Coloring, draw the pixels and quit
    If smartColors = False Then
        SetImageData
    
        Message "Color reduction complete."
        SetProgBarVal 0
        Exit Sub
    End If
    
    'If we're still here at this point, assume we are doing Intelligent Coloring
    If smartColors = True Then
    
        Message "Assigning new values from calculated look-up tables..."
        
        'Find average colors based on color counts
        For r = 0 To rValue
        For g = 0 To gValue
        For b = 0 To bValue
            If countLookup(r, g, b) <> 0 Then
                rLookup(r, g, b) = Int(Int(rLookup(r, g, b)) / Int(countLookup(r, g, b)))
                gLookup(r, g, b) = Int(Int(gLookup(r, g, b)) / Int(countLookup(r, g, b)))
                bLookup(r, g, b) = Int(Int(bLookup(r, g, b)) / Int(countLookup(r, g, b)))
                If rLookup(r, g, b) > 255 Then rLookup(r, g, b) = 255
                If gLookup(r, g, b) > 255 Then gLookup(r, g, b) = 255
                If bLookup(r, g, b) > 255 Then bLookup(r, g, b) = 255
                If rLookup(r, g, b) < 0 Then rLookup(r, g, b) = 0
                If gLookup(r, g, b) < 0 Then gLookup(r, g, b) = 0
                If bLookup(r, g, b) < 0 Then bLookup(r, g, b) = 0
            End If
        Next b
        Next g
        Next r
        
        'Assign average colors back into the picture
        For x = 0 To PicWidthL
            QuickX = x * 3
            eR = 0
            eG = 0
            eB = 0
        For y = 0 To PicHeightL
        
            r = ImageData(QuickX + 2, y)
            g = ImageData(QuickX + 1, y)
            b = ImageData(QuickX, y)
            r = r + eR
            g = g + eG
            b = b + eB
            cR = Int((r / mR) + 0.5)
            cG = Int((g / mG) + 0.5)
            cB = Int((b / mB) + 0.5)
            ImageData(QuickX + 2, y) = rLookup(cR, cG, cB)
            ImageData(QuickX + 1, y) = gLookup(cR, cG, cB)
            ImageData(QuickX, y) = bLookup(cR, cG, cB)
            
            cR = cR * mR
            cG = cG * mG
            cB = cB * mB
        
            eR = r - cR
            eG = g - cG
            eB = b - cB
            
        Next y
        Next x
        
        SetImageData
    
        Message "Color reduction complete."
        SetProgBarVal 0
        Exit Sub
    
    End If
    
End Sub
