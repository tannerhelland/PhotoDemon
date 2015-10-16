VERSION 5.00
Begin VB.Form dialog_ExportColorDepth 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Please Choose Exported Color Depth"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6435
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
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   429
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.commandBarMini cmdBarMini 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   6
      Top             =   3750
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   1323
      BackColor       =   14802140
      dontAutoUnloadParent=   -1  'True
   End
   Begin PhotoDemon.smartOptionButton optColorDepth 
      CausesValidation=   0   'False
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   1
      Top             =   960
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   582
      Caption         =   "1 bpp (monochrome)"
   End
   Begin PhotoDemon.smartOptionButton optColorDepth 
      CausesValidation=   0   'False
      Height          =   375
      Index           =   1
      Left            =   600
      TabIndex        =   2
      Top             =   1440
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   582
      Caption         =   "4 bpp (16 shades of gray)"
   End
   Begin PhotoDemon.smartOptionButton optColorDepth 
      CausesValidation=   0   'False
      Height          =   375
      Index           =   2
      Left            =   600
      TabIndex        =   3
      Top             =   1920
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   582
      Caption         =   "8 bpp (256 colors or full grayscale)"
   End
   Begin PhotoDemon.smartOptionButton optColorDepth 
      CausesValidation=   0   'False
      Height          =   375
      Index           =   3
      Left            =   600
      TabIndex        =   4
      Top             =   2400
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   582
      Caption         =   "24 bpp (16 million colors)"
   End
   Begin PhotoDemon.smartOptionButton optColorDepth 
      CausesValidation=   0   'False
      Height          =   375
      Index           =   4
      Left            =   600
      TabIndex        =   5
      Top             =   2880
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   582
      Caption         =   "32 bpp (16 million colors + full transparency)"
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "please choose an output color depth"
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
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3840
   End
End
Attribute VB_Name = "dialog_ExportColorDepth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Exported Color Depth Dialog
'Copyright 2011-2015 by Tanner Helland
'Created: 11/December/12
'Last updated: 11/December/12
'Last update: initial build
'
'Dialog for presenting the user a choice of exported color depths.  I prefer this to be
' handled automatically by the software, but in certain rare cases it may be desirable
' for a user to manually export a certain color depth
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'The user input from the dialog
Private userAnswer As VbMsgBoxResult

'The desired output format (used to activate available color depths)
Private outputFormat As Long

'We want to temporarily suspend an hourglass cursor if necessary
Private restoreCursor As Boolean

'The user's answer is returned via this property
Public Property Get DialogResult() As VbMsgBoxResult
    DialogResult = userAnswer
End Property

'This form can be notified of the image being exported.  This may be used in the future to provide a preview.
Public Property Let imageFormat(ByVal imageFormat As Long)
    outputFormat = imageFormat
End Property

'The ShowDialog routine presents the user with this form.
Public Sub showDialog()

    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    userAnswer = vbCancel
    
    'Make sure that the proper cursor is set
    If Screen.MousePointer = vbHourglass Then
        restoreCursor = True
        Screen.MousePointer = vbNormal
    Else
        restoreCursor = False
    End If
        
    'Based on the supplied image format, disable invalid color depths
    If g_ImageFormats.isColorDepthSupported(outputFormat, 1) Then optColorDepth(0).Enabled = True Else optColorDepth(0).Enabled = False
    If g_ImageFormats.isColorDepthSupported(outputFormat, 4) Then optColorDepth(1).Enabled = True Else optColorDepth(1).Enabled = False
    If g_ImageFormats.isColorDepthSupported(outputFormat, 8) Then optColorDepth(2).Enabled = True Else optColorDepth(2).Enabled = False
    If g_ImageFormats.isColorDepthSupported(outputFormat, 24) Then optColorDepth(3).Enabled = True Else optColorDepth(3).Enabled = False
    If g_ImageFormats.isColorDepthSupported(outputFormat, 32) Then optColorDepth(4).Enabled = True Else optColorDepth(4).Enabled = False
        
    'Out of politeness, set the default color depth to the current image's color depth
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    
    'Retrieve a composite copy of the image no larger than 512x512 pixels
    If pdImages(g_CurrentImage).Width > 512 Or pdImages(g_CurrentImage).Height > 512 Then
        
        Dim imgWidth As Long, imgHeight As Long
        convertAspectRatio pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height, 512, 512, imgWidth, imgHeight
        pdImages(g_CurrentImage).getCompositedRect tmpDIB, 0, 0, imgWidth, imgHeight, 0, 0, pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height, InterpolationModeHighQualityBicubic, , CLC_Generic
        
    Else
        pdImages(g_CurrentImage).getCompositedImage tmpDIB
    End If
    
    'If the image has a useless alpha channel, blank it out now.
    If Not DIB_Handler.verifyDIBAlphaChannel(tmpDIB) Then tmpDIB.convertTo24bpp
    
    'Select 24bpp or 32bpp by default, depending on the alpha channel status of the composited DIB.
    If (tmpDIB.getDIBColorDepth = 24) And (optColorDepth(3).Enabled) Then
        optColorDepth(3).Value = True
    ElseIf (tmpDIB.getDIBColorDepth = 32) And (optColorDepth(4).Enabled) Then
        optColorDepth(4).Value = True
    Else
        'If both 24 and 32bpp are disabled (not possible at present, but whatever), select the highest possible
        ' color depth by default
        Dim i As Long
        For i = optColorDepth.Count - 1 To 0 Step -1
            If optColorDepth(i).Enabled Then
                optColorDepth(i).Value = True
                Exit For
            End If
        Next i
    End If
        
    Message "Waiting for user to specify color depth... "
        
    'Apply translations and visual themes
    MakeFormPretty Me
    
    'Display the dialog
    ShowPDDialog vbModal, Me, True

End Sub

Private Sub cmdBarMini_CancelClick()
    
    If restoreCursor Then Screen.MousePointer = vbHourglass
    
    userAnswer = vbCancel
    Me.Hide
    
End Sub

Private Sub cmdBarMini_OKClick()

    'Restore a busy cursor if necessary
    If restoreCursor Then Screen.MousePointer = vbHourglass
        
    'Save the selected color depth to the corresponding global variable (so other functions can access it
    ' after this form is unloaded)
    If optColorDepth(0).Value = True Then g_ColorDepth = 1
    If optColorDepth(1).Value = True Then g_ColorDepth = 4
    If optColorDepth(2).Value = True Then g_ColorDepth = 8
    If optColorDepth(3).Value = True Then g_ColorDepth = 24
    If optColorDepth(4).Value = True Then g_ColorDepth = 32
     
    userAnswer = vbOK
    Me.Hide

End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub
