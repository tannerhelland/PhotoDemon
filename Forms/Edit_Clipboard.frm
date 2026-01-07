VERSION 5.00
Begin VB.Form FormClipboard 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " "
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6360
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
   ScaleHeight     =   283
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   424
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdCheckBox chkMerged 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   661
      Caption         =   "use merged image"
   End
   Begin PhotoDemon.pdListBox lstFormats 
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   7435
      Caption         =   "available formats"
   End
   Begin PhotoDemon.pdCommandBarMini cmdBar 
      Align           =   2  'Align Bottom
      CausesValidation=   0   'False
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   3510
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   1296
      DontAutoUnloadParent=   -1  'True
   End
End
Attribute VB_Name = "FormClipboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Median Filter Tool
'Copyright 2013-2026 by Tanner Helland
'Created: 08/Feb/13
'Last updated: 23/August/13
'Last update: added a mode-tracking variable to help with the new command bar addition
'
'This is a heavily optimized median filter function.  An "accumulation" technique is used instead of the standard sliding
' window mechanism.  (See http://web.archive.org/web/20060718054020/http://www.acm.uiuc.edu/siggraph/workshops/wjarosz_convolution_2001.pdf)
' This allows the algorithm to perform extremely well, despite being written in pure VB.
'
'That said, it is still unfortunately slow in the IDE.  I STRONGLY recommend compiling the project before applying any
' median filter of a large radius (> 20).
'
'An optional percentile option is available.  At minimum value, this performs identically to an erode (minimum) filter.
' Similarly, at max value it performs identically to a dilate (maximum) filter.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Because this tool can be used for multiple actions (median, dilate, erode), we need to track which mode is currently active.
' When the reset or randomize buttons are pressed, we will automatically adjust our behavior to match.
Public Enum PD_ClipboardOp
    co_Cut = 0
    co_Copy = 1
    co_Paste = 2
End Enum

#If False Then
    Private Const co_Cut = 0, co_Copy = 1, co_Paste = 2
#End If

Private m_OpMode As PD_ClipboardOp

'Remember last-used settings
' (Normally this is declared WithEvents, but this dialog doesn't require custom settings behavior.)
Private m_Settings As pdLastUsedSettings
Attribute m_Settings.VB_VarHelpID = -1

'The user input from the dialog.  If the user cancels this dialog, default settings will be used.
Private m_CmdBarAnswer As VbMsgBoxResult

Public Property Get DialogResult() As VbMsgBoxResult
    DialogResult = m_CmdBarAnswer
End Property

Public Function GetParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    Select Case m_OpMode
        Case co_Cut, co_Copy
            cParams.AddParam "clipboard-format", GetClipFormatFromListIndex_CutCopy(), True
            cParams.AddParam "merged", CBool(chkMerged.Value)
        Case co_Paste
    End Select
    
    GetParamString = cParams.GetParamString()
    
End Function

Private Sub cmdBar_CancelClick()
    m_CmdBarAnswer = vbCancel
    Me.Hide
End Sub

Private Sub cmdBar_OKClick()
    m_CmdBarAnswer = vbOK
    Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'The median dialog is reused for several tools: minimum, median, maximum.
Public Sub ShowClipboardDialog(ByVal opMode As PD_ClipboardOp)
    
    'Provide a default answer (in case the user closes the dialog via some means other than the command bar)
    m_CmdBarAnswer = vbCancel
    
    'Cache the current mode for future reference
    m_OpMode = opMode
    
    If (opMode = co_Cut) Then
        If (Not g_WindowManager Is Nothing) Then g_WindowManager.SetWindowCaptionW Me.hWnd, g_Language.TranslateMessage("Cut special")
        ListSupportedFormats
        lstFormats.ListIndex = 0
        
    ElseIf (opMode = co_Copy) Then
        If (Not g_WindowManager Is Nothing) Then g_WindowManager.SetWindowCaptionW Me.hWnd, g_Language.TranslateMessage("Copy special")
        ListSupportedFormats
        lstFormats.ListIndex = 0
        
    Else
        If (Not g_WindowManager Is Nothing) Then g_WindowManager.SetWindowCaptionW Me.hWnd, g_Language.TranslateMessage("Paste special")
        
    End If
    
    'Load any last-used settings for this form
    Set m_Settings = New pdLastUsedSettings
    m_Settings.SetParentForm Me
    m_Settings.LoadAllControlValues
    
    ApplyThemeAndTranslations Me
    
    ShowPDDialog vbModal, Me, True

End Sub

Private Sub ListSupportedFormats()
    lstFormats.SetAutomaticRedraws False
    lstFormats.Clear
    lstFormats.AddItem "PhotoDemon Image", 0
    lstFormats.AddItem "Bitmap", 1
    lstFormats.AddItem "DIB", 2
    lstFormats.AddItem "DIB v5", 3
    lstFormats.AddItem "PNG", 4
    lstFormats.SetAutomaticRedraws True, True
End Sub

Private Function GetClipFormatFromListIndex_CutCopy() As PD_ClipboardFormats

    Select Case lstFormats.ListIndex
        Case 0
            GetClipFormatFromListIndex_CutCopy = pdcf_InternalPD
        Case 1
            GetClipFormatFromListIndex_CutCopy = pdcf_Bitmap
        Case 2
            GetClipFormatFromListIndex_CutCopy = pdcf_Dib
        Case 3
            GetClipFormatFromListIndex_CutCopy = pdcf_DibV5
        Case 4
            GetClipFormatFromListIndex_CutCopy = pdcf_PNG
        Case Else
            GetClipFormatFromListIndex_CutCopy = pdcf_Dib
    End Select

End Function
