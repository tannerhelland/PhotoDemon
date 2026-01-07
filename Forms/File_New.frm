VERSION 5.00
Begin VB.Form FormNewImage 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " New image"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9630
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   469
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   642
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   6285
      Width           =   9630
      _ExtentX        =   16986
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdRadioButton optBackground 
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
   Begin PhotoDemon.pdRadioButton optBackground 
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
   Begin PhotoDemon.pdRadioButton optBackground 
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
   Begin PhotoDemon.pdRadioButton optBackground 
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
   Begin PhotoDemon.pdColorSelector csBackground 
      Height          =   615
      Left            =   840
      TabIndex        =   6
      Top             =   5280
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   1085
      curColor        =   16749332
   End
   Begin PhotoDemon.pdResize ucResize 
      Height          =   2850
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   5027
      DisablePercentOption=   -1  'True
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   1
      Left            =   240
      Top             =   3480
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   503
      Caption         =   "background"
      FontSize        =   12
      ForeColor       =   4210752
   End
End
Attribute VB_Name = "FormNewImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'New Image Dialog
'Copyright 2014-2026 by Tanner Helland
'Created: 29/December/14
'Last updated: 31/December/14
'Last update: wrap up initial build
'
'Basic "create new image" dialog.  Image size and background can be specified directly from the dialog,
' and the command bar allows for saving/loading presets just like every other tool.  In the future, it might be nice
' to provide some kind of "template" dropdown for convenience.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
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
    
    'As of v7.0, all parameter strings must be XML-based
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "WidthInPixels", ucResize.ResizeWidthInPixels
        .AddParam "HeightInPixels", ucResize.ResizeHeightInPixels
        .AddParam "DPI", ucResize.ResizeDPIAsPPI
        .AddParam "BackgroundType", backgroundType
        .AddParam "OptionalBackcolor", csBackground.Color
    End With
    
    Processor.Process "New image", False, cParams.GetParamString(), UNDO_Nothing
    
End Sub

Private Sub cmdBar_RandomizeClick()
    CalculateDefaultSize
    optBackground(3).Value = True
End Sub

Private Sub cmdBar_ResetClick()
    CalculateDefaultSize
    csBackground.Color = RGB(60, 160, 255)
End Sub

Private Sub Form_Load()
    
    'Set a default size for the image.  This is calculated according to the following formula:
    ' - If another image is loaded and active, default to that image size.
    '  (GIMP does this and it's extremely helpful.)
    ' - If no images have been loaded, default to desktop wallpaper size
    
    'Obviously, the user can use the save/load preset functionality to save favorite image sizes.
    
    'Fill in the boxes with the default size
    CalculateDefaultSize
    Interface.ApplyThemeAndTranslations Me
    
End Sub

Private Sub CalculateDefaultSize()

    'Default to pixels
    ucResize.UnitOfMeasurement = mu_Pixels
    
    'Is another image loaded?
    If PDImages.IsImageActive() Then
        
        'Default to the dimensions of the currently active image
        ucResize.SetInitialDimensions PDImages.GetActiveImage.Width, PDImages.GetActiveImage.Height, PDImages.GetActiveImage.GetDPI
        
    Else
    
        'Default to primary monitor size
        Dim pDisplay As pdDisplay
        Set pDisplay = g_Displays.PrimaryDisplay
                
        Dim pDisplayRect As RectL
        If (Not pDisplay Is Nothing) Then
            pDisplay.GetRect pDisplayRect
        Else
            With pDisplayRect
                .Left = 0
                .Top = 0
                .Right = 1920
                .Bottom = 1080
            End With
        End If
        
        ucResize.SetInitialDimensions pDisplayRect.Right - pDisplayRect.Left, pDisplayRect.Bottom - pDisplayRect.Top, 96
        
    End If

End Sub
