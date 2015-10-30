VERSION 5.00
Begin VB.Form FormHistogram 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   " Histogram"
   ClientHeight    =   9045
   ClientLeft      =   120
   ClientTop       =   360
   ClientWidth     =   10590
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
   ScaleHeight     =   603
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   706
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdButton cmdOK 
      Height          =   510
      Left            =   7320
      TabIndex        =   25
      Top             =   8400
      Width           =   3135
      _ExtentX        =   7011
      _ExtentY        =   873
      Caption         =   "Exit histogram"
   End
   Begin PhotoDemon.smartCheckBox chkLog 
      Height          =   330
      Left            =   7320
      TabIndex        =   22
      Top             =   6000
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   582
      Caption         =   "use logarithmic values"
   End
   Begin PhotoDemon.smartCheckBox chkSmooth 
      Height          =   330
      Left            =   7320
      TabIndex        =   21
      Top             =   5040
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   582
      Caption         =   "use smooth lines"
   End
   Begin PhotoDemon.smartCheckBox chkChannel 
      Height          =   330
      Index           =   0
      Left            =   4680
      TabIndex        =   17
      Top             =   5040
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   582
      Caption         =   "red"
   End
   Begin VB.PictureBox picH 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4035
      Left            =   120
      MousePointer    =   2  'Cross
      ScaleHeight     =   267
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   687
      TabIndex        =   1
      Top             =   120
      Width           =   10335
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
      ScaleWidth      =   687
      TabIndex        =   0
      Top             =   4200
      Width           =   10335
   End
   Begin PhotoDemon.smartCheckBox chkChannel 
      Height          =   330
      Index           =   1
      Left            =   4680
      TabIndex        =   18
      Top             =   5520
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   582
      Caption         =   "green"
   End
   Begin PhotoDemon.smartCheckBox chkChannel 
      Height          =   330
      Index           =   2
      Left            =   4680
      TabIndex        =   19
      Top             =   6000
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   582
      Caption         =   "blue"
   End
   Begin PhotoDemon.smartCheckBox chkChannel 
      Height          =   330
      Index           =   3
      Left            =   4680
      TabIndex        =   20
      Top             =   6480
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   582
      Caption         =   "luminance"
   End
   Begin PhotoDemon.smartCheckBox chkFillCurve 
      Height          =   330
      Left            =   7320
      TabIndex        =   23
      Top             =   5520
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   582
      Caption         =   "fill histogram curves"
   End
   Begin VB.Label lblVisibleChannels 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "visible channels"
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
      Left            =   4320
      TabIndex        =   24
      Top             =   4680
      Width           =   1650
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "statistics"
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
      Index           =   0
      Left            =   240
      TabIndex        =   16
      Top             =   4680
      Width           =   885
   End
   Begin VB.Label lblMouseInstructions 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "(Note: move the mouse over the histogram to calculate these values)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   210
      Left            =   480
      TabIndex        =   15
      Top             =   7800
      Width           =   5805
   End
   Begin VB.Label lblDrawOptions 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "rendering options"
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
      Left            =   6960
      TabIndex        =   14
      Top             =   4680
      Width           =   1875
   End
   Begin VB.Label lblValue 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "level"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   240
      Index           =   0
      Left            =   1080
      TabIndex        =   13
      Top             =   5880
      Width           =   390
   End
   Begin VB.Label lblMaxCount 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "maximum count:"
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
      Height          =   240
      Left            =   480
      TabIndex        =   12
      Top             =   5520
      Width           =   1440
   End
   Begin VB.Label lblValue 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   1
      Left            =   1080
      TabIndex        =   11
      Top             =   6240
      Width           =   285
   End
   Begin VB.Label lblValueTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "red:"
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
      Height          =   240
      Index           =   1
      Left            =   480
      TabIndex        =   10
      Top             =   6240
      Width           =   360
   End
   Begin VB.Label lblValueTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "level:"
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
      Height          =   240
      Index           =   0
      Left            =   480
      TabIndex        =   9
      Top             =   5880
      Width           =   465
   End
   Begin VB.Label lblTotalPixels 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "total pixels:"
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
      Height          =   240
      Left            =   480
      TabIndex        =   8
      Top             =   5160
      Width           =   990
   End
   Begin VB.Label lblValueTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "green:"
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
      Height          =   240
      Index           =   2
      Left            =   480
      TabIndex        =   7
      Top             =   6600
      Width           =   570
   End
   Begin VB.Label lblValue 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00008000&
      Height          =   240
      Index           =   2
      Left            =   1200
      TabIndex        =   6
      Top             =   6600
      Width           =   495
   End
   Begin VB.Label lblValueTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "blue:"
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
      Height          =   240
      Index           =   3
      Left            =   480
      TabIndex        =   5
      Top             =   6960
      Width           =   435
   End
   Begin VB.Label lblValue 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   3
      Left            =   1200
      TabIndex        =   4
      Top             =   6960
      Width           =   360
   End
   Begin VB.Label lblValueTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "luminance:"
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
      Height          =   240
      Index           =   4
      Left            =   480
      TabIndex        =   3
      Top             =   7320
      Width           =   945
   End
   Begin VB.Label lblValue 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "luminance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   4
      Left            =   1680
      TabIndex        =   2
      Top             =   7320
      Width           =   870
   End
End
Attribute VB_Name = "FormHistogram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Histogram Handler
'Copyright 2001-2015 by Tanner Helland
'Created: 6/12/01
'Last updated: 30/September/13
'Last update: when drawing cubic spline histograms, cache various GDI+ handles to improve performance.
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
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Have we generated a histogram yet?
Private histogramGenerated As Boolean

'Histogram data for each particular type (r/g/b/luminance)
Private hData() As Double
Private hDataLog() As Double

'Maximum histogram values (r/g/b/luminance)
'NOTE: As of 2012, a single max value is calculated for red, green, blue, and luminance (because all lines are drawn simultaneously).  No longer needed: Private HMax(0 To 3) As Double
Private hMax As Double, hMaxLog As Double
Private channelMax() As Double
Private channelMaxLog() As Double
Private channelMaxPosition() As Byte
Private maxChannel As Byte          'This identifies the channel with the highest value (red, green, or blue)

'Modified cubic spline variables:
Private nPoints As Integer
Private iX() As Double
Private iy() As Double
Private p() As Double
Private u() As Double
Private results() As Long   'Stores the y-values for each x-value in the final spline

'To improve histogram render performance, we cache a number of translated strings; this saves us having to re-translate them
' every time the histogram is redrawn.
Private strTotalPixels As String
Private strMaxCount As String
Private strRed As String, strGreen As String, strBlue As String, strLuminance As String
Private strLevel As String

'When channels are enabled or disabled, redraw the histogram
Private Sub chkChannel_Click(Index As Integer)
    DrawHistogram
End Sub

Private Sub chkFillCurve_Click()
    DrawHistogram
End Sub

Private Sub chkLog_Click()
    DrawHistogram
End Sub

'When the smoothing option is changed, redraw the histogram
Private Sub chkSmooth_Click()
    DrawHistogram
End Sub

'OK button
Private Sub CmdOK_Click()
    Unload Me
End Sub

'Just to be safe, regenerate the histogram whenever the form receives focus
Private Sub Form_Activate()
    
    'Apply visual themes and translations
    MakeFormPretty Me
    
    'Cache the translation for several dynamic strings; this is more efficient than retranslating them over and over
    strTotalPixels = g_Language.TranslateMessage("total pixels") & ": "
    strMaxCount = g_Language.TranslateMessage("max count") & ": "
    strRed = g_Language.TranslateMessage("red")
    strGreen = g_Language.TranslateMessage("green")
    strBlue = g_Language.TranslateMessage("blue")
    strLuminance = g_Language.TranslateMessage("luminance")
    strLevel = g_Language.TranslateMessage("level")
    
    'Blank out the specific level labels populated by moving the mouse across the form
    ' Also, align the value labels with their (potentially translated) corresponding title labels
    Dim i As Long
    For i = 0 To lblValue.Count - 1
        lblValue(i).Left = lblValueTitle(i).Left + lblValueTitle(i).Width + FixDPI(8)
        lblValue(i) = ""
    Next i
    
    If Not histogramGenerated Then TallyHistogramValues
    DrawHistogram
    
End Sub

'Subroutine to draw a histogram.  Note that a variable called "hType" is used frequently in the sub; it tells us which histogram to draw:
'0 - Red
'1 - Green
'2 - Blue
'3 - Luminance
Public Sub DrawHistogram()
    
    'If histogram data hasn't been generated, exit
    If Not histogramGenerated Then Exit Sub
    
    'Clear out whatever was there before
    picH.Picture = LoadPicture("")
    
    'tHeight is used to determine the height of the maximum value in the histogram.  We want it to be slightly
    ' shorter than the height of the picture box; this way the tallest histogram value fills the entire box
    Dim tHeight As Long
    tHeight = picH.ScaleHeight - 2
    
    'LastX and LastY are used to draw a connecting line between histogram points
    Dim LastX As Long, LastY As Long
    
    'We now need to calculate a max histogram value based on which RGB channels are enabled
    hMax = 0:    hMaxLog = 0:   maxChannel = 4  'Set maxChannel to an arbitrary value higher than 2
    
    Dim i As Long
    For i = 0 To 2
        If CBool(chkChannel(i)) Then
            If channelMax(i) > hMax Then
                hMax = channelMax(i)
                hMaxLog = channelMaxLog(i)
                maxChannel = i
            End If
        End If
    Next i
    
    'We'll need to draw up to four lines - one each for red, green, blue, and luminance,
    ' depending on what channels the user has enabled.
    Dim hType As Long
    
    For hType = 0 To 3
    
        'Only draw this histogram channel if the user has requested it
        If CBool(chkChannel(hType)) Then
        
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
            
            'The luminance channel is a special case - it uses its own max values, so check for that here
            If hType = 3 Then
                hMax = channelMax(hType)
                hMaxLog = channelMaxLog(hType)
            End If
    
            'Now we'll draw the histogram.  The drawing code will change based on the "smooth lines" setting on the histogram dialog.
            If CBool(chkSmooth) Then
            
                'Drawing a cubic spline line is complex enough to warrant its own subroutine.  Check there for details.
                drawCubicSplineHistogram hType, tHeight, CBool(chkFillCurve)
                
            Else
                    
                'For the first point there is no last 'x' or 'y', so we'll just make it the same as the first value in the histogram.
                LastX = 0
                If CBool(chkLog) Then
                    LastY = tHeight - (hDataLog(hType, 0) / hMaxLog) * tHeight
                Else
                    LastY = tHeight - (hData(hType, 0) / hMax) * tHeight
                End If
                    
                Dim xCalc As Long
                
                'Run a loop through every histogram value...
                Dim x As Long, y As Long
                For x = 0 To picH.ScaleWidth
            
                    'The y-value of the histogram is drawn as a percentage (RData(x) / MaxVal) * tHeight) with tHeight being
                    ' the tallest possible value (when RData(x) = MaxVal).  We then subtract that value from tHeight because
                    ' y values INCREASE as we move DOWN a picture box - remember that (0,0) is in the top left.
                    xCalc = Int((x / picH.ScaleWidth) * 256)
                    If xCalc > 255 Then xCalc = 255
                    
                    'Use logarithmic values if requested by the user
                    If CBool(chkLog) Then
                        y = tHeight - (hDataLog(hType, xCalc) / hMaxLog) * tHeight
                    Else
                        y = tHeight - (hData(hType, xCalc) / hMax) * tHeight
                    End If
                    
                    'Draw a line from the last (x,y) to the current (x,y)
                    picH.Line (LastX, LastY + 2)-(x, y + 2)
                        
                    'If "fill curve" is selected, fill the area beneath this point.  (Note that luminance curve is never filled!)
                    If hType < 3 And CBool(chkFillCurve) Then GDIPlusDrawLineToDC picH.hDC, x, y + 2, x, picH.ScaleHeight, picH.ForeColor, 64, 1, False
                        
                    'Update the LastX/Y values
                    LastX = x
                    LastY = y
                    
                Next x
            
            End If
                
        End If
                
    Next hType
    
    picH.Picture = picH.Image
    picH.Refresh
    
    'Last but not least, generate the statistics at the bottom of the form
    
    'Total number of pixels
    lblTotalPixels.Caption = strTotalPixels & (pdImages(g_CurrentImage).Width * pdImages(g_CurrentImage).Height)
    
    'Maximum value; if a color channel is enabled, use that
    If CBool(chkChannel(0)) Or CBool(chkChannel(1)) Or CBool(chkChannel(2)) Then
        
        'Reset hMax, which may have been changed if the luminance histogram was rendered
        hMax = channelMax(maxChannel)
        lblMaxCount.Caption = strMaxCount & hMax
        
        'Also display the channel with that max value, if applicable
        Select Case maxChannel
            Case 0
                lblMaxCount.Caption = lblMaxCount.Caption & " (" & strRed
            Case 1
                lblMaxCount.Caption = lblMaxCount.Caption & " (" & strGreen
            Case 2
                lblMaxCount.Caption = lblMaxCount.Caption & " (" & strBlue
        End Select
        
        lblMaxCount.Caption = lblMaxCount.Caption & ", " & strLevel & " " & channelMaxPosition(maxChannel) & ")"
    
    'Otherwise, default to luminance
    Else
        lblMaxCount.Caption = channelMax(3) & " (" & strLuminance
        lblMaxCount.Caption = lblMaxCount.Caption & ", " & strLevel & " " & channelMaxPosition(3) & ")"
    End If
        
End Sub

Private Sub Form_Deactivate()
    histogramGenerated = False
End Sub

Private Sub Form_Load()
    histogramGenerated = False
    
    'On XP, GDI+'s line function is hideously slow.  Disable filled curves by default.
    If Not g_IsVistaOrLater Then chkFillCurve.Value = vbUnchecked Else chkFillCurve.Value = vbChecked
    
End Sub

'If the form is resized, adjust all the controls to match
Private Sub Form_Resize()

    picH.Width = Me.ScaleWidth - picH.Left - FixDPI(8)
    picGradient.Width = Me.ScaleWidth - picGradient.Left - FixDPI(8)
    cmdOK.Left = Me.ScaleWidth - cmdOK.Width - FixDPI(8)
    
    'Now draw a little gradient below the histogram window, to help orient the user
    DrawHistogramGradient picGradient, RGB(0, 0, 0), RGB(255, 255, 255)
    
    'Only draw the histogram if the histogram data has been initialized
    ' (This is necessary because VB triggers the Resize event before the Activate event)
    If histogramGenerated Then DrawHistogram
    
End Sub

'When the mouse moves over the histogram, display the level and count for the histogram
'entry at the x-value over which the mouse passes
Private Sub picH_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim xCalc As Long
    xCalc = Int((x / picH.ScaleWidth) * 256)
    If xCalc > 255 Then xCalc = 255
    lblValue(0).Caption = xCalc
    lblValue(1).Caption = hData(0, xCalc)
    lblValue(2).Caption = hData(1, xCalc)
    lblValue(3).Caption = hData(2, xCalc)
    lblValue(4).Caption = hData(3, xCalc)
End Sub

'UNLOAD form
Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
    Message "Finished."
End Sub

'We'll use this routine only to draw the gradient below the histogram window.  This code is old, but it works ;)
Private Sub DrawHistogramGradient(ByRef dstObject As PictureBox, ByVal Color1 As Long, ByVal Color2 As Long)
    
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
    Dim vR As Double, vG As Double, vB As Double
    
    'Size of the picture box we'll be drawing to
    Dim iWidth As Long, iHeight As Long
    iWidth = dstObject.ScaleWidth
    iHeight = dstObject.ScaleHeight
    
    'Here, create a calculation variable for determining the step between
    'each level of the gradient
    vR = Abs(r - r2) / iWidth
    vG = Abs(g - g2) / iWidth
    vB = Abs(b - b2) / iWidth
    
    'If the second value is lower then the first value, make the step negative
    If r2 < r Then vR = -vR
    If g2 < g Then vG = -vG
    If b2 < b Then vB = -vB
    
    'Last, run a loop through the width of the picture box, incrementing the color as
    'we go (thus creating a gradient effect)
    Dim x As Long
    For x = 0 To iWidth
        r2 = r + vR * x
        g2 = g + vG * x
        b2 = b + vB * x
        dstObject.Line (x, 0)-(x, iHeight), RGB(r2, g2, b2)
    Next x
    
End Sub

'This routine draws the histogram using cubic splines to smooth the output
Private Sub drawCubicSplineHistogram(ByVal histogramChannel As Long, ByVal tHeight As Long, ByVal fillCurve As Boolean)
    
    'Initialize a few variables that are simply copies of image properties; this is faster than repeatedly accessing the properties themselves.
    Dim histWidth As Long, histHeight As Long
    histWidth = picH.ScaleWidth
    histHeight = picH.ScaleHeight
    
    Dim curHistColor As Long
    curHistColor = picH.ForeColor
    
    'Create an array consisting of 256 points, where each point corresponds to a histogram value
    nPoints = 256
    ReDim iX(nPoints) As Double
    ReDim iy(nPoints) As Double
    ReDim p(nPoints) As Double
    ReDim u(nPoints) As Double
    
    'Now, populate the iX and iY arrays with the histogram values for the specified channel (0-3, corresponds to hType above)
    Dim logMode As Boolean
    logMode = CBool(chkLog)
    
    Dim i As Long
    For i = 1 To nPoints
        iX(i) = (i - 1) * (histWidth / 255)
        
        If logMode Then
            iy(i) = tHeight - (hDataLog(histogramChannel, i - 1) / hMaxLog) * tHeight
        Else
            iy(i) = tHeight - (hData(histogramChannel, i - 1) / hMax) * tHeight
        End If
        
    Next i
    
    'results() will hold the actual pixel (x,y) values for each line to be drawn to the picture box
    ReDim results(0 To histWidth) As Long
    
    'Now run a loop through the knots, calculating spline values as we go
    Call SetPandU
    Dim xPos As Long, yPos As Double
    For i = 1 To nPoints - 1
        For xPos = iX(i) To iX(i + 1)
            yPos = getCurvePoint(i, xPos)
            
            'Add two to the final point, to shift the histogram slightly downward
            results(xPos) = yPos + 2
        Next xPos
    Next i
    
    'The area under the curve is filled if: 1) the "fill curve" checkbox is selected, and 2) the curve is something other than luminance
    Dim needToFill As Boolean
    If (histogramChannel < 3) And fillCurve Then needToFill = True Else needToFill = False
    
    'For performance reasons, cache the handle to the GDI+ image container and GDI+ pens we will be using.  This is faster than recreating
    ' them for every line, especially if the histogram window has been resized to something large.
    Dim gdiHistogram As Long
    gdiHistogram = getGDIPlusGraphicsFromDC(picH.hDC)
    
    Dim gdiPenSolid As Long, gdiPenTranslucent As Long
    gdiPenSolid = getGDIPlusPenHandle(curHistColor)
    gdiPenTranslucent = getGDIPlusPenHandle(curHistColor, 64)
        
    'If "Fill curve" is selected, we need to manually draw the left-most column (as the draw loop starts at 1)
    ' Note that the luminance curve is never filled.
    If needToFill Then GDIPlusDrawLine_Fast gdiHistogram, gdiPenTranslucent, 0, results(0), 0, histHeight
    
    'Draw the finished spline, using GDI+ for antialiasing
    For i = 1 To histWidth
        GDIPlusDrawLine_Fast gdiHistogram, gdiPenSolid, i, results(i), i - 1, results(i - 1)
        If needToFill Then GDIPlusDrawLine_Fast gdiHistogram, gdiPenTranslucent, i, results(i), i, histHeight
    Next i
    
    'Free the GDI+ handles
    releaseGDIPlusGraphics gdiHistogram
    releaseGDIPlusPen gdiPenSolid
    releaseGDIPlusPen gdiPenTranslucent
    
    'Refresh the picture
    picH.Picture = picH.Image
    
End Sub

'Original required spline function:
Private Function getCurvePoint(ByRef i As Long, ByVal v As Double) As Double

    Dim t As Double
    'derived curve equation (which uses p's and u's for coefficients)
    t = (v - iX(i)) / u(i)
    getCurvePoint = t * iy(i + 1) + (1 - t) * iy(i) + u(i) * u(i) * (f(t) * p(i + 1) + f(1 - t) * p(i)) / 6#
    
End Function

'Original required spline function:
Private Function f(ByRef x As Double) As Double
        f = x * x * x - x
End Function

'Original required spline function:
Private Sub SetPandU()

    Dim i As Long
    Dim d() As Double
    Dim w() As Double
    ReDim d(nPoints) As Double
    ReDim w(nPoints) As Double
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
        w(i) = 6# * ((iy(i + 1) - iy(i)) / u(i) - (iy(i) - iy(i - 1)) / u(i - 1))
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

'Build the histogram tables.  This only needs to be called once, when the image is changed. It will generate all histogram
' data for all channels (including luminance, and all log variants).
Public Sub TallyHistogramValues()

    Debug.Print "Tallying histogram values..."

    'Notify the user that the histogram is being generated
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    
    'If a histogram has already been drawn, render the "please wait" text over the top of it.  Otherwise, render it to a blank white image.
    If (picH.Picture.Width = 0) Then
        tmpDIB.createBlank picH.ScaleWidth, picH.ScaleHeight
    Else
        tmpDIB.CreateFromPicture picH.Picture
    End If
    
    Dim notifyFont As pdFont
    Set notifyFont = New pdFont
    notifyFont.SetFontFace g_InterfaceFont
    notifyFont.SetFontSize 14
    notifyFont.SetFontColor 0
    notifyFont.SetFontBold True
    notifyFont.SetTextAlignment vbCenter
    notifyFont.CreateFontObject
    notifyFont.AttachToDC tmpDIB.getDIBDC
    
    notifyFont.FastRenderText picH.ScaleWidth / 2, picH.ScaleHeight / 2, g_Language.TranslateMessage("Please wait while the histogram is updated...")
    tmpDIB.renderToPictureBox picH
    
    notifyFont.ReleaseFromDC
    Set tmpDIB = Nothing

    Message "Updating histogram..."
    
    'Blank the red, green, blue, and luminance count text boxes
    Dim i As Long
    For i = 0 To lblValue.Count - 1
        lblValue(i) = ""
    Next i
    
    'Use our new external function to fill the important histogram arrays
    fillHistogramArrays hData, hDataLog, channelMax, channelMaxLog, channelMaxPosition
    
    'If the histogram has already been used, we need to clear out two additional maximum values
    hMax = 0
    hMaxLog = 0
    
    histogramGenerated = True
    
    Message "Finished."

End Sub

'Stretch the histogram to reach from 0 to 255 (white balance correction is a far better method, FYI)
Public Sub StretchHistogram()
   
    Message "Analyzing image histogram for maximum and minimum values..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    
    prepImageData tmpSA
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Color variables
    Dim r As Long, g As Long, b As Long
    
    'Max and min values
    Dim RMax As Long, gMax As Long, bMax As Long
    Dim RMin As Long, gMin As Long, bMin As Long
    RMin = 255
    gMin = 255
    bMin = 255
        
    'Loop through each pixel in the image, checking max/min values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        If r < RMin Then RMin = r
        If r > RMax Then RMax = r
        If g < gMin Then gMin = g
        If g > gMax Then gMax = g
        If b < bMin Then bMin = b
        If b > bMax Then bMax = b
        
    Next y
    Next x
    
    Message "Stretching histogram..."
    Dim rDif As Long, gDif As Long, bDif As Long
    
    rDif = RMax - RMin
    gDif = gMax - gMin
    bDif = bMax - bMin
    
    'Lookup tables make the stretching go faster
    Dim rLookup(0 To 255) As Byte, gLookUp(0 To 255) As Byte, bLookup(0 To 255) As Byte
    
    For x = 0 To 255
        If rDif <> 0 Then
            r = 255 * ((x - RMin) / rDif)
            If r < 0 Then r = 0
            If r > 255 Then r = 255
            rLookup(x) = r
        Else
            rLookup(x) = x
        End If
        If gDif <> 0 Then
            g = 255 * ((x - gMin) / gDif)
            If g < 0 Then g = 0
            If g > 255 Then g = 255
            gLookUp(x) = g
        Else
            gLookUp(x) = x
        End If
        If bDif <> 0 Then
            b = 255 * ((x - bMin) / bDif)
            If b < 0 Then b = 0
            If b > 255 Then b = 255
            bLookup(x) = b
        Else
            bLookup(x) = x
        End If
    Next x
    
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
                
        ImageData(QuickVal + 2, y) = rLookup(r)
        ImageData(QuickVal + 1, y) = gLookUp(g)
        ImageData(QuickVal, y) = bLookup(b)
        
    Next y
        If (x And progBarCheck) = 0 Then SetProgBarVal x
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData
        
End Sub
