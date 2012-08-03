VERSION 5.00
Begin VB.Form FormHistogram 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   " Histogram"
   ClientHeight    =   7485
   ClientLeft      =   120
   ClientTop       =   360
   ClientWidth     =   11910
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
   ScaleHeight     =   499
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   794
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkLog 
      Appearance      =   0  'Flat
      Caption         =   "Use logarithmic values"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4920
      MouseIcon       =   "VBP_FormHistogram.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   6900
      Width           =   2055
   End
   Begin VB.CheckBox chkSmooth 
      Appearance      =   0  'Flat
      Caption         =   "Use smooth lines"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3000
      MouseIcon       =   "VBP_FormHistogram.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   6900
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CommandButton cmdExportHistogram 
      Caption         =   "Export Histogram to File..."
      Height          =   495
      Left            =   240
      MouseIcon       =   "VBP_FormHistogram.frx":02A4
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   6840
      Width           =   2535
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   10440
      MouseIcon       =   "VBP_FormHistogram.frx":03F6
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   6840
      Width           =   1245
   End
   Begin VB.PictureBox picH 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   120
      MousePointer    =   2  'Cross
      ScaleHeight     =   319
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   775
      TabIndex        =   1
      Top             =   120
      Width           =   11655
   End
   Begin VB.PictureBox picGradient 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   775
      TabIndex        =   0
      Top             =   4950
      Width           =   11655
   End
   Begin VB.Label lblLevel 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "lblLevel"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   720
      TabIndex        =   16
      Top             =   5880
      Width           =   525
   End
   Begin VB.Label lblMaxCount 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "lblMaxCount"
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   1440
      TabIndex        =   15
      Top             =   6360
      Width           =   2415
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Maximum count:"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   240
      TabIndex        =   14
      Top             =   6360
      Width           =   1170
   End
   Begin VB.Label lblCountRed 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "lblRed"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   600
      TabIndex        =   13
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Red:"
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   6120
      Width           =   375
   End
   Begin VB.Line lineStats 
      BorderColor     =   &H80000002&
      X1              =   16
      X2              =   784
      Y1              =   387
      Y2              =   387
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Level:"
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label lblTotalPixels 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total pixels:"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   5520
      Width           =   870
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Green:"
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   1560
      TabIndex        =   9
      Top             =   6120
      Width           =   495
   End
   Begin VB.Label lblCountGreen 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "lblGreen"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   2160
      TabIndex        =   8
      Top             =   6120
      Width           =   855
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Blue:"
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label lblCountBlue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "lblBlue"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Luminance:"
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   4920
      TabIndex        =   5
      Top             =   6120
      Width           =   855
   End
   Begin VB.Label lblCountLuminance 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "lblLuminance"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5760
      TabIndex        =   4
      Top             =   6120
      Width           =   1215
   End
End
Attribute VB_Name = "FormHistogram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Histogram Handler
'Copyright ©2001-2012 by Tanner Helland
'Created: 6/12/01
'Last updated: 01/August/12
'Last update: added "smooth lines" option, which uses a cubic spline function and antialiasing
'             to smooth the histogram output.  Note that antialiasing can be enabled for the non-cubic
'             drawing routine; check the comments for details.
'
'This form runs the basic code for calculating and displaying an image's histogram. Throughout the code, the
' following array locations refer to a type of histogram:
' 0 - Red
' 1 - Green
' 2 - Blue
' 3 - Luminance
' This applies especially to the hData() and hMax() arrays.
'
'Also, I owe great thanks to the original author of the cubic spline routine I've used (Jason Bullen).
' His original cubic spline code can be downloaded from:
' http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=11488&lngWId=1
   '
   'ORIGINAL COMMENTS FOR JASON'S CUBIC SPLINE CODE:
   'Here is an absolute minimum Cubic Spline routine.
   'It's a VB rewrite from a Java applet I found by by Anthony Alto 4/25/99
   'Computes coefficients based on equations mathematically derived from the curve
   'constraints.   i.e. :
   '    curves meet at knots (predefined points)  - These must be sorted by X
   '    first derivatives must be equal at knots
   '    second derivatives must be equal at knots
'
'Additional thanks go out to Ron van Tilburg, for his native-VB implementation of Xiaolin Wu's line antialiasing algorithm.
' (See the "Outside_mGfxWu" module for details.)
'
'***************************************************************************

Option Explicit

'Old functions use these:
Dim rData(0 To 255) As Long, gData(0 To 255) As Long, bData(0 To 255) As Long

'Histogram data for each particular type (r/g/b/luminance)
Dim hData(0 To 3, 0 To 255) As Single
Dim hDataLog(0 To 3, 0 To 255) As Single

'Maximum histogram values (r/g/b/luminance)
'NOTE: As of 2012, a single max value is calculated for red, green, blue, and luminance (because all lines are drawn simultaneously).  No longer needed: Dim HMax(0 To 3) As Single
Dim hMax As Single, hMaxLog As Single

'Loop and position variables
Dim x As Long, y As Long

'Modified cubic spline variables:
Dim nPoints As Integer
Private iX() As Single
Private iY() As Single
Private p() As Single
Private u() As Single
Private results() As Long   'Stores the y-values for each x-value in the final spline

Private Sub chkLog_Click()
    DrawHistogram
End Sub

'When the smoothing option is changed, redraw the histogram
Private Sub chkSmooth_Click()
    DrawHistogram
End Sub

'Export the histogram image to file
Private Sub cmdExportHistogram_Click()
    Dim CC As cCommonDialog
    Set CC = New cCommonDialog
    
    'Get the last "save image" path from the INI file
    Dim tempPathString As String
    tempPathString = GetFromIni("Program Paths", "MainSave")
    
    Dim cdfStr As String
    
    cdfStr = "BMP - Windows Bitmap|*.bmp"
    
    'FreeImage allows us to save more filetypes
    If FreeImageEnabled = True Then
        cdfStr = cdfStr & "|GIF - Graphics Interchange Format|*.gif"
        cdfStr = cdfStr & "|PNG - Portable Network Graphic|*.png"
    End If
    
    Dim sFile As String
    sFile = "Histogram for " & pdImages(CurrentImage).OriginalFileName
    
    'If FreeImage is enabled, suggest PNG as the default format; otherwise, bitmaps is all they get
    Dim defFormat As Long
    Dim defExtension As String
    If FreeImageEnabled = False Then defFormat = 1 Else defFormat = 3
    If FreeImageEnabled = False Then defExtension = ".bmp" Else defExtension = ".png"
    
    'Display the save dialog
    If CC.VBGetSaveFileName(sFile, , True, cdfStr, defFormat, tempPathString, "Save histogram to file", defExtension, FormHistogram.HWnd, 0) Then
        
        'Save the new directory as the default path for future usage
        tempPathString = sFile
        StripDirectory tempPathString
        WriteToIni "Program Paths", "MainSave", tempPathString
        
        Message "Saving histogram to file..."
        
        'Now we do a bit of hackery to fool PhotoDemon into saving the histogram image using the stock SaveImage routine.
        ' First, create an object to hold a copy of the active form's image
        Dim tmppic As StdPicture
        Set tmppic = FormMain.ActiveForm.BackBuffer.Picture
        
        Dim oWidth As Long, oHeight As Long
        oWidth = FormMain.ActiveForm.BackBuffer.Width
        oHeight = FormMain.ActiveForm.BackBuffer.Height
        
        'Now copy the histogram image over that image.  Hackish, isn't it?  Don't say I didn't warn you. ;)
        FormMain.ActiveForm.BackBuffer.Width = FormHistogram.picH.Width
        FormMain.ActiveForm.BackBuffer.Height = FormHistogram.picH.Height
        FormMain.ActiveForm.BackBuffer.Picture = FormHistogram.picH.Image
        
        'With our hackery complete, use the core PhotoDemon save function to save the histogram image to file
        PhotoDemon_SaveImage CurrentImage, sFile, False, &H8
        
        'Replace the main image and exit out
        FormMain.ActiveForm.BackBuffer.Width = oWidth
        FormMain.ActiveForm.BackBuffer.Height = oHeight
        FormMain.ActiveForm.BackBuffer.Picture = tmppic
        
        Message "Save complete."
    End If
End Sub

'OK button
Private Sub CmdOK_Click()
    Unload Me
End Sub

'Generate the histogram data upon form load (we only need to do it once per image)
Private Sub Form_Load()
    
    TallyHistogramValues

    DrawHistogram
    
    'Commented code below is from a previous build where the user could specify which channel to draw.
    ' I haven't implemented per-channel checkboxes yet.  It's coming.
    
    'DrawHistogram lastHistSource, lastHistMethod
    
    'Clear out the combo box that displays histogram sources and fill it with
    'appropriate choices
    'cmbHistSource.Clear
    'cmbHistSource.AddItem "Red"
    'cmbHistSource.AddItem "Green"
    'cmbHistSource.AddItem "Blue"
    'cmbHistSource.AddItem "Luminance"
    'Set the current combo box option to be whatever we used last
    'cmbHistSource.ListIndex = lastHistSource
    
    'Clear out the combo box that displays histogram drawing methods and fill it
    'with appropriate choices
    'cmbHistMethod.Clear
    'cmbHistMethod.AddItem "Connecting lines"
    'cmbHistMethod.AddItem "Solid bars"
    'Set the current combo box option to be whatever we used last
    'cmbHistMethod.ListIndex = lastHistMethod
    
End Sub

'Subroutine to draw a histogram.  hType tells us what histogram to draw:
'0 - Red
'1 - Green
'2 - Blue
'3 - Luminance
'drawMethod tells us what kind of histogram to draw:
'0 - Connected lines (like a line graph)
'1 - Solid bars (like a bar graph) - CURRENTLY UNUSED
'2 - Smooth lines (using cubic spline code adopted from the Curves function)
Public Sub DrawHistogram()

    Dim drawMethod As Long
    If chkSmooth.Value = vbUnchecked Then drawMethod = 0 Else drawMethod = 2
    
    'Clear out whatever was there before
    'picH.Cls
    picH.Picture = LoadPicture("")
    
    'tHeight is used to determine the height of the maximum value in the
    'histogram.  We want it to be slightly shorter than the height of the
    'picture box; this way the tallest histogram value fills the entire box
    Dim tHeight As Long
    tHeight = picH.ScaleHeight - 2
    
    'LastX and LastY are used to draw a connecting line between histogram points
    Dim LastX As Long, LastY As Long
    
    'Now draw a little gradient below the histogram window, to help orient the user
    DrawHistogramGradient picGradient, RGB(0, 0, 0), RGB(255, 255, 255)
    
    'We'll need to draw four lines - one each for red, green, blue, and luminance
    Dim hType As Long
    
    For hType = 0 To 3
    
        'The type of histogram we're drawing will determine the color of the histogram
        'line - we'll make it match what we're drawing (red/green/blue/black)
        Select Case hType
            'Red
            Case 0
                picH.ForeColor = RGB(255, 0, 0)
            'Green
            Case 1
                picH.ForeColor = RGB(0, 255, 0)
            'Blue
            Case 2
                picH.ForeColor = RGB(0, 0, 255)
            'Luminance
            Case 3
                picH.ForeColor = RGB(0, 0, 0)
        End Select
    

    
        'Now we'll draw the histogram.  The drawing code will change based on the drawMethod specified by the user.
        'Remember: 0 - Connected lines, 1 - Solid bars, 2 - Smooth lines
        Select Case drawMethod
        
            Case 0
        
                'For the first point there is no last 'x' or 'y', so we'll just make it the
                'same as the first value in the histogram. (We care about this only if we're
                'drawing a "connected lines" type of histogram.)
                LastX = 0
                If chkLog.Value = vbChecked Then
                    LastY = tHeight - (hDataLog(hType, 0) / hMaxLog) * tHeight
                Else
                    LastY = tHeight - (hData(hType, 0) / hMax) * tHeight
                End If
                    
                Dim xCalc As Long
                
                'Run a loop through every histogram value...
                For x = 0 To picH.ScaleWidth
            
                    'The y-value of the histogram is drawn as a percentage (RData(x) / MaxVal) * tHeight) with tHeight being
                    ' the tallest possible value (when RData(x) = MaxVal).  We then subtract that value from tHeight because
                    ' y values INCREASE as we move DOWN a picture box - remember that (0,0) is in the top left.
                    xCalc = Int((x / picH.ScaleWidth) * 256)
                    If xCalc > 255 Then xCalc = 255
                    
                    If chkLog.Value = vbChecked Then
                        y = tHeight - (hDataLog(hType, xCalc) / hMaxLog) * tHeight
                    Else
                        y = tHeight - (hData(hType, xCalc) / hMax) * tHeight
                    End If
                    
                    'For connecting lines...
                    If drawMethod = 0 Then
                        'Then draw a line from the last (x,y) to the current (x,y)
                        picH.Line (LastX, LastY + 2)-(x, y + 2)
                        'The line below can be used for antialiased drawing, FYI
                        'DrawLineWuAA picH.hDC, LastX, LastY + 2, x, y + 2, picH.ForeColor
                        LastX = x
                        LastY = y
                        
                    'For a bar graph...
                    ElseIf drawMethod = 1 Then
                        'Draw a line from the bottom of the picture box to the calculated y-value
                        picH.Line (x, tHeight + 2)-(x, y + 2)
                    End If
                Next x
                
            Case 2
        
                'Drawing a cubic spline line is complex enough to warrant its own subroutine.  Check there for details.
                drawCubicSplineHistogram hType, tHeight
                
        End Select
                
    Next hType
    
    picH.Picture = picH.Image
    picH.Refresh
    
    'Last but not least, generate the statistics at the bottom of the form
    
    'Total number of pixels
    GetImageData
    lblTotalPixels.Caption = "Total pixels: " & (PicWidthL * PicHeightL)
    
    'Maximum value
    lblMaxCount.Caption = hMax

End Sub

'If the form is resized, adjust all the controls to match
Private Sub Form_Resize()

    picH.Width = Me.ScaleWidth - picH.Left - 8
    picGradient.Width = Me.ScaleWidth - picGradient.Left - 8
    
    CmdOK.Left = Me.ScaleWidth - CmdOK.Width - 8
    lineStats.x2 = Me.ScaleWidth - lineStats.x1
    
    DrawHistogram
    
End Sub

Private Sub Form_Terminate()
    Unload Me
End Sub

'When the mouse moves over the histogram, display the level and count for the histogram
'entry at the x-value over which the mouse passes
Private Sub picH_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim xCalc As Long
    xCalc = Int((x / picH.ScaleWidth) * 256)
    If xCalc > 255 Then xCalc = 255
    lblLevel.Caption = xCalc
    lblCountRed.Caption = hData(0, xCalc)
    lblCountGreen.Caption = hData(1, xCalc)
    lblCountBlue.Caption = hData(2, xCalc)
    lblCountLuminance.Caption = hData(3, xCalc)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Message "Finished."
End Sub

'We'll use this routine only to draw the gradient below the histogram window
'(like Photoshop does).  This code is old, but it works ;)
Public Sub DrawHistogramGradient(ByRef DstObject As PictureBox, ByVal Color1 As Long, ByVal Color2 As Long)
    'RGB() variables for each color
    Dim r As Long, g As Long, b As Long
    Dim r2 As Long, g2 As Long, b2 As Long
    
    'Extract the r,g,b values from the colors passed by the user
    r = Color1 Mod 256
    g = (Color1 \ 256) And 255
    b = (Color1 \ 65536) And 255
    r2 = Color2 Mod 256
    g2 = (Color2 \ 256) And 255
    b2 = (Color2 \ 65536) And 255
    
    'Calculation variables for the gradiency
    Dim VR As Single, VG As Single, VB As Single
    
    'Size of the picture box we'll be drawing to
    Dim iWidth As Long, iHeight As Long
    iWidth = DstObject.ScaleWidth
    iHeight = DstObject.ScaleHeight
    
    'Here, create a calculation variable for determining the step between
    'each level of the gradient
    VR = Abs(r - r2) / iWidth
    VG = Abs(g - g2) / iWidth
    VB = Abs(b - b2) / iWidth
    'If the second value is lower then the first value, make the step negative
    If r2 < r Then VR = -VR
    If g2 < g Then VG = -VG
    If b2 < b Then VB = -VB
    'Last, run a loop through the width of the picture box, incrementing the color as
    'we go (thus creating a gradient effect)
    Dim x As Long
    For x = 0 To iWidth
        r2 = r + VR * x
        g2 = g + VG * x
        b2 = b + VB * x
        DstObject.Line (x, 0)-(x, iHeight), RGB(r2, g2, b2)
    Next x
End Sub

'Stretch the histogram to reach from 0 to 255 (white balance correction is a better method, FYI)
Public Sub StretchHistogram()
    
    Dim RMax As Byte, GMax As Byte, BMax As Byte
    
    Dim RMin As Byte, GMin As Byte, BMin As Byte
    RMin = 255
    GMin = 255
    BMin = 255

    Dim r As Long, g As Long, b As Long

    Message "Gathering histogram data..."
    GetImageData
    Dim QuickVal As Long
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        If r < RMin Then RMin = r
        If r > RMax Then RMax = r
        If g < GMin Then GMin = g
        If g > GMax Then GMax = g
        If b < BMin Then BMin = b
        If b > BMax Then BMax = b
    Next y
    Next x
    
    Message "Stretching histogram..."
    SetProgBarMax PicWidthL
    
    Dim Rdif As Integer, Gdif As Integer, Bdif As Integer
    
    Rdif = RMax - RMin
    Gdif = GMax - GMin
    Bdif = BMax - BMin
    
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        If Rdif <> 0 Then r = 255 * ((r - RMin) / Rdif)
        If Gdif <> 0 Then g = 255 * ((g - GMin) / Gdif)
        If Bdif <> 0 Then b = 255 * ((b - BMin) / Bdif)
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    
    SetImageData
    
    Message "Finished."
End Sub

'Equalize the red, green, and/or blue channels of an image
Public Sub EqualizeHistogram(ByVal HandleR As Boolean, ByVal HandleG As Boolean, ByVal HandleB As Boolean)
    GetImageData
    Dim r As Integer, g As Integer, b As Integer
    Message "Gathering histogram data..."
    SetProgBarMax (PicWidthL * 2)
    SetProgBarVal 0
    
    Dim QuickVal As Long
    
    'First, tally the amount of each color (i.e. build the histogram)
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        rData(r) = rData(r) + 1
        gData(g) = gData(g) + 1
        bData(b) = bData(b) + 1
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    'Compute our scaling factor
    Dim scalefactor As Double
    scalefactor = 255 / (PicWidthL * PicHeightL)
    'Handle red as necessary
    If HandleR = True Then
        rData(0) = rData(0) * scalefactor
        For x = 1 To 255
            rData(x) = rData(x - 1) + (scalefactor * rData(x))
        Next x
    End If
    'Handle green as necessary
    If HandleG = True Then
        gData(0) = gData(0) * scalefactor
        For x = 1 To 255
            gData(x) = gData(x - 1) + (scalefactor * gData(x))
        Next x
    End If
    'Handle blue as necessary
    If HandleB = True Then
        bData(0) = bData(0) * scalefactor
        For x = 1 To 255
            bData(x) = bData(x - 1) + (scalefactor * bData(x))
        Next x
    End If
    'Integerize all the look-up values
    For x = 0 To 255
        rData(x) = Int(rData(x))
        If rData(x) > 255 Then rData(x) = 255
        gData(x) = Int(gData(x))
        If gData(x) > 255 Then gData(x) = 255
        bData(x) = Int(bData(x))
        If bData(x) > 255 Then bData(x) = 255
    Next x
    
    'Apply the equalized values
    Message "Equalizing image..."

    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        If HandleR = True Then ImageData(QuickVal + 2, y) = rData(ImageData(QuickVal + 2, y))
        If HandleG = True Then ImageData(QuickVal + 1, y) = gData(ImageData(QuickVal + 1, y))
        If HandleB = True Then ImageData(QuickVal, y) = bData(ImageData(QuickVal, y))
    Next y
        If x Mod 20 = 0 Then SetProgBarVal PicWidthL + x
    Next x
    
    SetImageData
    Message "Finished."
End Sub

'Equalize an image using only luminance values
Public Sub EqualizeLuminance()
    GetImageData
    Dim Lum(0 To 255) As Single
    Dim r As Long, g As Long, b As Long
    Dim HH As Single, SS As Single, LL As Single
    Message "Gathering histogram data..."
    SetProgBarMax (PicWidthL * 2)
    SetProgBarVal 0
    
    Dim QuickVal As Long
    
    'First, tally the luminance amounts (i.e. build the histogram)
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        tRGBToHSL r, g, b, HH, SS, LL
        Lum(LL) = Lum(LL) + 1
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    'Compute our scaling factor
    Dim scalefactor As Double
    scalefactor = 255 / (PicWidthL * PicHeightL)
    'Equalize the luminance
    Lum(0) = Lum(0) * scalefactor
    For x = 1 To 255
        Lum(x) = Lum(x - 1) + (scalefactor * Lum(x))
    Next x
   'Integerize all the look-up values
    For x = 0 To 255
        Lum(x) = Int(Lum(x))
        If Lum(x) > 255 Then Lum(x) = 255
    Next x
    
    'Apply the equalized values
    Message "Equalizing image..."

    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        'Get the temporary values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        'Get the hue and saturation
        tRGBToHSL r, g, b, HH, SS, LL
        'Convert back to RGB using our artificial luminance values
        tHSLToRGB HH, SS, Lum(LL) / 255, r, g, b
        'Assign those values into the array
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
    Next y
        If x Mod 20 = 0 Then SetProgBarVal PicWidthL + x
    Next x
    
    SetImageData
    Message "Finished."
End Sub

'The next four functions are required for converting between the HSL and RGB color spaces
Private Sub tRGBToHSL(r As Long, g As Long, b As Long, h As Single, s As Single, l As Single)
Dim Max As Single
Dim Min As Single
Dim delta As Single
Dim rR As Single, rG As Single, rB As Single
   rR = r / 255: rG = g / 255: rB = b / 255

'{Given: rgb each in [0,1].
' Desired: h in [0,360] and s in [0,1], except if s=0, then h=UNDEFINED.}
        Max = Maximum(rR, rG, rB)
        Min = Minimum(rR, rG, rB)
        l = (Max + Min) / 2    '{This is the lightness}
        '{Next calculate saturation}
        If Max = Min Then
            'begin {Achromatic case}
            s = 0
            h = 0
           'end {Achromatic case}
        Else
           'begin {Chromatic case}
                '{First calculate the saturation.}
           If l <= 0.5 Then
               s = (Max - Min) / (Max + Min)
           Else
               s = (Max - Min) / (2 - Max - Min)
            End If
            '{Next calculate the hue.}
            delta = Max - Min
           If rR = Max Then
                h = (rG - rB) / delta    '{Resulting color is between yellow and magenta}
           ElseIf rG = Max Then
                h = 2 + (rB - rR) / delta '{Resulting color is between cyan and yellow}
           ElseIf rB = Max Then
                h = 4 + (rR - rG) / delta '{Resulting color is between magenta and cyan}
           End If
            'Debug.Print h
            'h = h * 60
           'If h < 0# Then
           '     h = h + 360            '{Make degrees be nonnegative}
           'End If
        'end {Chromatic Case}
      End If

'Tanner's hack: transfer the values into ones I can use; this yields
' hue on [0,240], saturation on [0,255], and luminance on [0,255]
    'H = Int(H * 40 + 40)
    'S = Int(S * 255)
    l = Int(l * 255)
End Sub

Public Sub tHSLToRGB(h As Single, s As Single, l As Single, r As Long, g As Long, b As Long)
Dim rR As Single, rG As Single, rB As Single
Dim Min As Single, Max As Single
'This one requires the stupid values; such is life

   If s = 0 Then
      ' Achromatic case:
      rR = l: rG = l: rB = l
   Else
      ' Chromatic case:
      ' delta = Max-Min
      If l <= 0.5 Then
         's = (Max - Min) / (Max + Min)
         ' Get Min value:
         Min = l * (1 - s)
      Else
         's = (Max - Min) / (2 - Max - Min)
         ' Get Min value:
         Min = l - s * (1 - l)
      End If
      ' Get the Max value:
      Max = 2 * l - Min
      
      ' Now depending on sector we can evaluate the h,l,s:
      If (h < 1) Then
         rR = Max
         If (h < 0) Then
            rG = Min
            rB = rG - h * (Max - Min)
         Else
            rB = Min
            rG = h * (Max - Min) + rB
         End If
      ElseIf (h < 3) Then
         rG = Max
         If (h < 2) Then
            rB = Min
            rR = rB - (h - 2) * (Max - Min)
         Else
            rR = Min
            rB = (h - 2) * (Max - Min) + rR
         End If
      Else
         rB = Max
         If (h < 4) Then
            rR = Min
            rG = rR - (h - 4) * (Max - Min)
         Else
            rG = Min
            rR = (h - 4) * (Max - Min) + rG
         End If
         
      End If
            
   End If
   r = rR * 255: g = rG * 255: b = rB * 255
End Sub

Private Function Maximum(rR As Single, rG As Single, rB As Single) As Single
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

Private Function Minimum(rR As Single, rG As Single, rB As Single) As Single
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

'This routine draws the histogram using cubic splines to smooth the output
Private Function drawCubicSplineHistogram(ByVal histogramChannel As Long, ByVal tHeight As Long)
    
    'Create an array consisting of 256 points, where each point corresponds to a histogram value
    nPoints = 256
    ReDim iX(nPoints) As Single
    ReDim iY(nPoints) As Single
    ReDim p(nPoints) As Single
    ReDim u(nPoints) As Single
    
    'Now, populate the iX and iY arrays with the histogram values for the specified channel (0-3, corresponds to hType above)
    Dim i As Long
    For i = 1 To nPoints
        iX(i) = (i - 1) * (picH.ScaleWidth / 255)
        If chkLog.Value = vbChecked Then
            iY(i) = tHeight - (hDataLog(histogramChannel, i - 1) / hMaxLog) * tHeight
        Else
            iY(i) = tHeight - (hData(histogramChannel, i - 1) / hMax) * tHeight
        End If
    Next i
    
    'results() will hold the actual pixel (x,y) values for each line to be drawn to the picture box
    ReDim results(0 To picH.ScaleWidth) As Long
    
    'Now run a loop through the knots, calculating spline values as we go
    Call SetPandU
    Dim Xpos As Long, Ypos As Single
    For i = 1 To nPoints - 1
        For Xpos = iX(i) To iX(i + 1)
            Ypos = getCurvePoint(i, Xpos)
            'If yPos > 255 Then yPos = 254       'Force values to be in the 1-254 range (0-255 also
            'If yPos < 0 Then yPos = 1           ' works, but is harder to see on the picture box)
            results(Xpos) = Ypos
        Next Xpos
    Next i
    
    'Draw the finished spline
    For i = 1 To picH.ScaleWidth
        'picH.Line (i, results(i) + 2)-(i - 1, results(i - 1) + 2)
        DrawLineWuAA picH.hDC, i, results(i) + 2, i - 1, results(i - 1) + 2, picH.ForeColor
    Next i
    
    picH.Picture = picH.Image
    
End Function


'Original required spline function:
Private Function getCurvePoint(ByVal i As Long, ByVal v As Single) As Single
    Dim t As Single
    'derived curve equation (which uses p's and u's for coefficients)
    t = (v - iX(i)) / u(i)
    getCurvePoint = t * iY(i + 1) + (1 - t) * iY(i) + u(i) * u(i) * (f(t) * p(i + 1) + f(1 - t) * p(i)) / 6#
End Function

'Original required spline function:
Private Function f(x As Single) As Single
        f = x * x * x - x
End Function

'Original required spline function:
Private Sub SetPandU()
    Dim i As Integer
    Dim d() As Single
    Dim w() As Single
    ReDim d(nPoints) As Single
    ReDim w(nPoints) As Single
    'Routine to compute the parameters of our cubic spline.  Based on equations derived from some basic facts...
    'Each segment must be a cubic polynomial.  Curve segments must have equal first and second derivatives
    'at knots they share.  General algorithm taken from a book which has long since been lost.
    
    'The math that derived this stuff is pretty messy...  expressions are isolated and put into
    'arrays.  we're essentially trying to find the values of the second derivative of each polynomial
    'at each knot within the curve.  That's why theres only N-2 p's (where N is # points).
    'later, we use the p's and u's to calculate curve points...

    For i = 2 To nPoints - 1
        d(i) = 2 * (iX(i + 1) - iX(i - 1))
    Next
    For i = 1 To nPoints - 1
        u(i) = iX(i + 1) - iX(i)
    Next
    For i = 2 To nPoints - 1
        w(i) = 6# * ((iY(i + 1) - iY(i)) / u(i) - (iY(i) - iY(i - 1)) / u(i - 1))
    Next
    For i = 2 To nPoints - 2
        w(i + 1) = w(i + 1) - w(i) * u(i) / d(i)
        d(i + 1) = d(i + 1) - u(i) * u(i) / d(i)
    Next
    p(1) = 0#
    For i = nPoints - 1 To 2 Step -1
        p(i) = (w(i) - u(i) * p(i + 1)) / d(i)
    Next
    p(nPoints) = 0#
End Sub

Public Sub TallyHistogramValues()
    
    Message "Gathering histogram data..."
    
    'Blank the red, green, blue, and luminance count text boxes
    lblCountRed = ""
    lblCountGreen = ""
    lblCountBlue = ""
    lblCountLuminance = ""
    
    'Grab image information
    GetImageData
    
    'These variables will hold temporary histogram values
    Dim r As Long, g As Long, b As Long, l As Long
    
    'If the histogram has already been used, we need to clear out all the
    'maximum values and histogram values
    hMax = 0
    For x = 0 To 3
        For y = 0 To 255
            hData(x, y) = 0
        Next y
    Next x
    
    'Run a quick loop through the image, gathering what we need to
    'calculate our histogram
    Dim QuickVal As Long
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        'We have to gather the red, green, and blue in order to calculate luminance
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        'Rather than generate authentic luminance (which requires an HSL
        ' conversion routine), we'll use an average value.  It's accurate
        ' enough for a project like this.
        l = (r + g + b) \ 3
        'Increment each value in the array, depending on its present value;
        'this will let us see how many of each color value (and luminance
        'value) there is in the image
        'Red
        hData(0, r) = hData(0, r) + 1
        'Green
        hData(1, g) = hData(1, g) + 1
        'Blue
        hData(2, b) = hData(2, b) + 1
        'Luminance
        hData(3, l) = hData(3, l) + 1
    Next y
    Next x
    
    'Run a quick loop through the completed array to find maximum values
    For x = 0 To 255
        'Red
        If hData(0, x) > hMax Then hMax = hData(0, x)
        'Green
        If hData(1, x) > hMax Then hMax = hData(1, x)
        'Blue
        If hData(2, x) > hMax Then hMax = hData(2, x)
        'Luminance
        If hData(3, x) > hMax Then hMax = hData(3, x)
    Next x
    
    'Now calculate the logarithmic version of the histogram
    hMaxLog = Log(hMax)
    For x = 0 To 3
        For y = 0 To 255
            hDataLog(x, y) = Log(hData(x, y))
        Next y
    Next x
    
    Message "Finished."

End Sub
