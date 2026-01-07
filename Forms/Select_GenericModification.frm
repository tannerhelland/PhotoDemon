VERSION 5.00
Begin VB.Form FormSelectionDialogs 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Selection options"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6660
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
   ScaleHeight     =   176
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   444
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdCommandBarMini cmdBarMini 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   1
      Top             =   1890
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   1323
      DontAutoUnloadParent=   -1  'True
   End
   Begin PhotoDemon.pdSlider sltSelValue 
      Height          =   405
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   5895
      _ExtentX        =   10186
      _ExtentY        =   873
      Min             =   1
      Max             =   500
      Value           =   1
      DefaultValue    =   1
   End
End
Attribute VB_Name = "FormSelectionDialogs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Multi-purpose Selection Dialog
'Copyright 2013-2026 by Tanner Helland
'Created: 11/July/13
'Last updated: 11/July/13
'Last update: initial build
'
'Custom dialog box for asking the user for a selection-related parameters.  Because all selection-related options
' (grow, shrink, border, feather, etc) don't provide previews, it is easy to handle their dialogs using a single
' form - saving on resources in the process.
'
'This form is not designed to be displayed on its own.  Use the displaySelectionDialog function in the
' Selection_Handler module to properly initialize it (and properly capture all return values).
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'The user input from the dialog
Private userAnswer As VbMsgBoxResult

'The slider value, if the dialog is closed via OK
Private userValue As Double

Public Property Get DialogResult() As VbMsgBoxResult
    DialogResult = userAnswer
End Property

Public Property Get paramValue() As Double
    paramValue = userValue
End Property

'The ShowDialog routine presents the user with the form.
Public Sub ShowDialog(ByVal typeOfDialog As PD_SelectionDialog)
    
    'Based on the type of dialog requested, rebuild the dialog's text
    Dim titleText As String, sliderText As String
    
    Select Case typeOfDialog
    
        Case pdsd_Grow
            titleText = g_Language.TranslateMessage("Grow selection")
            sliderText = g_Language.TranslateMessage("radius")
        
        Case pdsd_Shrink
            titleText = g_Language.TranslateMessage("Shrink selection")
            sliderText = g_Language.TranslateMessage("radius")
        
        Case pdsd_Border
            titleText = g_Language.TranslateMessage("Border selection")
            sliderText = g_Language.TranslateMessage("radius")
        
        Case pdsd_Feather
            titleText = g_Language.TranslateMessage("Feather selection")
            sliderText = g_Language.TranslateMessage("radius")
        
        Case pdsd_Sharpen
            titleText = g_Language.TranslateMessage("Sharpen selection")
            sliderText = g_Language.TranslateMessage("radius")
    
    End Select
    
    'Use Unicode-aware form captions
    If (Not g_WindowManager Is Nothing) Then g_WindowManager.SetWindowCaptionW Me.hWnd, titleText
    sltSelValue.Caption = sliderText
    
    'Provide a default answer of "first image only" (in the event that the user clicks the "x" button in the top-right)
    userAnswer = vbNo

    'Apply translations and visual themes
    ApplyThemeAndTranslations Me

    'Display the form
    ShowPDDialog vbModal, Me

End Sub

Private Sub cmdBarMini_CancelClick()
    userAnswer = vbCancel
    userValue = 0
    Me.Hide
End Sub

Private Sub cmdBarMini_OKClick()
    If sltSelValue.IsValid Then
        userAnswer = vbOK
        userValue = sltSelValue.Value
        Me.Hide
    Else
        cmdBarMini.DoNotUnloadForm
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub
