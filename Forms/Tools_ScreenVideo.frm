VERSION 5.00
Begin VB.Form FormScreenCapPNG 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Capture Screen Animation"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8550
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
   ScaleHeight     =   451
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   570
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdTextBox txtCoords 
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Text            =   "0"
   End
   Begin PhotoDemon.pdSlider sldFrameRate 
      Height          =   735
      Left            =   240
      TabIndex        =   4
      Top             =   2520
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1296
      Caption         =   "max frame rate (fps)"
      FontSizeCaption =   10
      Min             =   1
      Max             =   30
      Value           =   10
      NotchPosition   =   2
      NotchValueCustom=   10
   End
   Begin PhotoDemon.pdButton cmdExit 
      Height          =   495
      Left            =   7200
      TabIndex        =   3
      Top             =   6120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "Cancel"
   End
   Begin PhotoDemon.pdButton cmdStart 
      Height          =   495
      Left            =   4560
      TabIndex        =   2
      Top             =   6120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      Caption         =   "Start"
   End
   Begin PhotoDemon.pdButton cmdExport 
      Height          =   375
      Left            =   7680
      TabIndex        =   1
      Top             =   1440
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      Caption         =   "..."
   End
   Begin PhotoDemon.pdTextBox txtExport 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   661
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   1
      Left            =   120
      Top             =   120
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   503
      Caption         =   "capture area"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   0
      Left            =   120
      Top             =   1080
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   503
      Caption         =   "destination file"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   2
      Left            =   120
      Top             =   2040
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   503
      Caption         =   "capture settings"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdTextBox txtCoords 
      Height          =   375
      Index           =   1
      Left            =   1800
      TabIndex        =   6
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Text            =   "0"
   End
   Begin PhotoDemon.pdTextBox txtCoords 
      Height          =   375
      Index           =   2
      Left            =   3960
      TabIndex        =   7
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Text            =   "800"
   End
   Begin PhotoDemon.pdTextBox txtCoords 
      Height          =   375
      Index           =   3
      Left            =   5520
      TabIndex        =   8
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Text            =   "600"
   End
End
Attribute VB_Name = "FormScreenCapPNG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Animated screen capture dialog
'Copyright 2020-2020 by Tanner Helland
'Created: 01/July/20
'Last updated: 01/July/20
'Last update: this is just an experiment at present!
'
'PD can write animated PNGs.  APNGs seem like a great fit for animated screen captures.
' Let's see if we can merge the two, eh?
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'A timer triggers capture events
Private WithEvents m_Timer As pdTimer
Attribute m_Timer.VB_VarHelpID = -1

'Frame count is tracked manually; certain things need to be handled differently on
' e.g. the first frame vs subsequent frames.
Private m_FrameCount As Long

'A pdPNG instance handles the actual PNG writing
Private m_PNG As pdPNG

'Destination file, if one is selected (check for null before using)
Private m_DstFilename As String

'Capture rect; once populated (at the start of the capture), it *cannot* be changed
Private m_CaptureRect As RectL

'If a capture event is ACTIVE, this will be set to TRUE
Private m_CaptureActive As Boolean

'Capture DIB.  Reused on successive frames for perf reasons.
Private m_captureDIB As pdDIB

Private Sub cmdExit_Click()
    StopTimer_Forcibly
    Unload Me
End Sub

Private Sub cmdExport_Click()
    
    Dim cSave As pdOpenSaveDialog
    Set cSave = New pdOpenSaveDialog
    
    'Get the last "save image" path from the preferences file
    Dim tempPathString As String
    tempPathString = UserPrefs.GetPref_String("Paths", "Save Image", vbNullString)
    
    'Provide a string to the common dialog; it will fill this with the user's chosen path + filename
    Dim tmpFilename As String
    tmpFilename = g_Language.TranslateMessage("capture")
    
    Dim sFile As String
    sFile = tempPathString & IncrementFilename(tempPathString, tmpFilename, "png")
    
    'Present a common dialog to the user
    If cSave.GetSaveFileName(sFile, , True, "Animated PNG (.png)|*.png", 1, tempPathString, g_Language.TranslateMessage("Export screen capture animation"), "png", Me.hWnd) Then
        txtExport.Text = sFile
    End If
    
End Sub

Private Sub cmdStart_Click()
    
    'If a capture is already active, we STOP the timer now
    If m_CaptureActive Then
        
        'Immediately stop the capture timer
        m_CaptureActive = False
        If (Not m_Timer Is Nothing) Then m_Timer.StopTimer
        
        'Notify the PNG encoder
        If (Not m_PNG Is Nothing) Then m_PNG.SaveAPNG_Streaming_Stop 0, 3000
        
        'Reset this button's caption
        cmdStart.Caption = g_Language.TranslateMessage("Start")
        
    'If a capture is NOT active, we START the timer now
    Else
        
        'Validate inputs
        m_DstFilename = Trim$(txtExport.Text)
        
        Dim dstFileOK As Boolean
        dstFileOK = (LenB(m_DstFilename) <> 0)
        
        'If the destination file looks okay, see if a file exists at that locale.
        ' Provide a failsafe overwrite warning.
        If dstFileOK Then
            
            If Files.FileExists(m_DstFilename) Then
                Dim msgReturn As VbMsgBoxResult
                msgReturn = PDMsgBox("%1 already exists.  Do you want to overwrite it?", vbYesNo Or vbExclamation Or vbApplicationModal, "Overwrite warning", m_DstFilename)
                If (msgReturn = vbNo) Then Exit Sub
            End If
            
        End If
        
        'If all preliminary checks passed, activate the capture timer.  Note that this
        ' *will* forcibly overwrite the file at the destination location, if one exists.
        If dstFileOK Then
            
            'Initialize the timer
            Set m_Timer = New pdTimer
            m_Timer.Interval = Int(1000# / sldFrameRate.Value + 0.5)
            
            Debug.Print m_Timer.Interval
            
            'Determine the capture rect
            With m_CaptureRect
                .Left = txtCoords(0).Text
                .Top = txtCoords(1).Text
                .Right = .Left + txtCoords(2).Text
                .Bottom = .Top + txtCoords(3).Text
            End With
            
            'Prep the capture DIB
            Set m_captureDIB = New pdDIB
            m_captureDIB.CreateBlank m_CaptureRect.Right - m_CaptureRect.Left, m_CaptureRect.Bottom - m_CaptureRect.Top, 24, vbBlack
            
            'Initialize the PNG writer
            m_FrameCount = 0
            Set m_PNG = New pdPNG
            If (m_PNG.SaveAPNG_Streaming_Start(m_DstFilename, m_captureDIB.GetDIBWidth, m_captureDIB.GetDIBHeight) < png_Failure) Then
                
                'Change the start button to a STOP button
                cmdStart.Caption = g_Language.TranslateMessage("Stop")
                
                'Start the capture timer
                m_CaptureActive = True
                m_Timer.StartTimer
            
            Else
                PDDebug.LogAction "WARNING!  APNG screen capture failed for unknown reason.  Consult debug log."
            End If
            
        Else
            PDMsgBox "You must enter a valid filename.", vbCritical Or vbOKOnly Or vbApplicationModal, "Invalid filename"
        End If
        
    End If
        
End Sub

Private Sub Form_Load()
    Interface.ApplyThemeAndTranslations Me
End Sub

'On timer events, capture the current frame by calling this function
Private Sub CaptureFrameNow()
    
    'Ensure we have an active APNG instance
    If (Not m_PNG Is Nothing) Then
        
        'Note the current time
        Dim capTime As Currency
        capTime = VBHacks.GetHighResTimeEx()
        
        'Capture the frame in question into a pdDIB object
        ScreenCapture.GetPartialDesktopAsDIB m_captureDIB, m_CaptureRect
        
        'Pass the frame off to the PNG encoder
        m_PNG.SaveAPNG_Streaming_Frame m_captureDIB, capTime
        
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'Forcibly stop the current capture (if any)
    StopTimer_Forcibly
    
End Sub

Private Sub m_Timer_Timer()
    Debug.Print Timer
    CaptureFrameNow
End Sub

'Want to emergency stop capture for whatever reason?  Call this function.
Private Sub StopTimer_Forcibly()
    
    If m_CaptureActive Then
        
        'Cancel any future timer events
        If (Not m_Timer Is Nothing) Then m_Timer.StopTimer
        Set m_Timer = Nothing
        
        'Cancel the PNG writer
        If (Not m_PNG Is Nothing) Then m_PNG.SaveAPNG_Streaming_Cancel
        Set m_PNG = Nothing
        
        'Erase the destination file (if one exists)
        If Files.FileExists(m_DstFilename) Then Files.FileDelete m_DstFilename
        m_DstFilename = vbNullString
        
        'Deactivate capture mode
        m_CaptureActive = False
        
    End If
    
End Sub
