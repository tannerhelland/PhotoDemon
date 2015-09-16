VERSION 5.00
Begin VB.Form FormNewImage 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Create new image"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9630
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
   ScaleHeight     =   469
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   642
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   6285
      Width           =   9630
      _ExtentX        =   16986
      _ExtentY        =   1323
      BackColor       =   14802140
   End
   Begin PhotoDemon.smartOptionButton optBackground 
      CausesValidation=   0   'False
      Height          =   330
      Index           =   0
      Left            =   480
      TabIndex        =   2
      Top             =   3840
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   582
      Caption         =   "transparent"
      Value           =   -1  'True
   End
   Begin PhotoDemon.smartOptionButton optBackground 
      CausesValidation=   0   'False
      Height          =   330
      Index           =   1
      Left            =   480
      TabIndex        =   3
      Top             =   4200
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   582
      Caption         =   "black"
   End
   Begin PhotoDemon.smartOptionButton optBackground 
      CausesValidation=   0   'False
      Height          =   330
      Index           =   2
      Left            =   480
      TabIndex        =   4
      Top             =   4560
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   582
      Caption         =   "white"
   End
   Begin PhotoDemon.smartOptionButton optBackground 
      CausesValidation=   0   'False
      Height          =   330
      Index           =   3
      Left            =   480
      TabIndex        =   5
      Top             =   4920
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   582
      Caption         =   "custom color"
   End
   Begin PhotoDemon.colorSelector colorPicker 
      Height          =   615
      Left            =   840
      TabIndex        =   6
      Top             =   5280
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   1085
      curColor        =   16749332
   End
   Begin PhotoDemon.smartResize ucResize 
      Height          =   2850
      Left            =   360
      TabIndex        =   7
      Top             =   360
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   5027
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DisablePercentOption=   -1  'True
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "background"
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
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   3480
      Width           =   1245
   End
End
Attribute VB_Name = "FormNewImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'New Image Dialog
'Copyright 2014-2015 by Tanner Helland
'Created: 29/December/14
'Last updated: 31/December/14
'Last update: wrap up initial build
'
'Basic "create new image" dialog.  Image size and background can be specified directly from the dialog,
' and the command bar allows for saving/loading presets just like every other tool.  In the future, it might be nice
' to provide some kind of "template" dropdown for convenience.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private Sub cmdBar_OKClick()

    'Retrieve the layer type from the active command button
    Dim backgroundType As Long
    
    Dim i As Long
    For i = 0 To optBackground.Count - 1
        If optBackground(i).Value Then
            backgroundType = i
            Exit For
        End If
    Next i
        
    Process "New image", False, buildParams(ucResize.imgWidthInPixels, ucResize.imgHeightInPixels, ucResize.imgDPIAsPPI, backgroundType, colorPicker.Color), UNDO_NOTHING
    
End Sub

Private Sub cmdBar_RandomizeClick()
    calculateDefaultSize
    optBackground(3).Value = True
End Sub

Private Sub cmdBar_ResetClick()
    calculateDefaultSize
    colorPicker.Color = RGB(60, 160, 255)
End Sub

Private Sub Form_Activate()
    calculateDefaultSize
    
    makeFormPretty Me
    
End Sub

Private Sub Form_Load()
    
    'Set a default size for the image.  This is calculated according to the following formula:
    ' If another image is loaded and active, default to that image size.  GIMP does this and it's extremely helpful.
    ' If no images have been loaded, default to desktop wallpaper size
    
    'Obviously, the user can use the save/load preset functionality to save favorite image sizes.
    
    'Fill in the boxes with the default size
    calculateDefaultSize
    
End Sub

Private Sub calculateDefaultSize()

    'Default to pixels
    ucResize.unitOfMeasurement = MU_PIXELS
    
    'Is another image loaded?
    If g_OpenImageCount > 0 Then
        
        'Default to the dimensions of the currently active image
        ucResize.setInitialDimensions pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height, pdImages(g_CurrentImage).getDPI
        
    Else
    
        'Default to primary monitor size
        Dim pDisplay As pdDisplay
        Set pDisplay = g_Displays.PrimaryDisplay
                
        Dim pDisplayRect As RECTL
        If Not pDisplay Is Nothing Then
            pDisplay.getRect pDisplayRect
        Else
            With pDisplayRect
                .Left = 0
                .Top = 0
                .Right = 1920
                .Bottom = 1080
            End With
        End If
        
        ucResize.setInitialDimensions pDisplayRect.Right - pDisplayRect.Left, pDisplayRect.Bottom - pDisplayRect.Top, 96
        
    End If

End Sub
