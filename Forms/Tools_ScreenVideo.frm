VERSION 5.00
Begin VB.Form FormScreenCapPNG 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Animated screen capture"
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
   Begin PhotoDemon.pdLabel lblProgress 
      Height          =   735
      Left            =   240
      Top             =   5280
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   1296
      Caption         =   ""
   End
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

'A pdPNG instance handles the actual PNG writing
Private m_PNG As pdPNG

'Destination file, if one is selected (check for null before using)
Private m_DstFilename As String

'Capture rect; once populated (at the start of the capture), it *cannot* be changed
Private m_CaptureRect As RectL

'If a capture event is ACTIVE, this will be set to TRUE
Private m_CaptureActive As Boolean

'Capture DIBs.  Reused on successive frames for perf reasons.  Separate 24- and 32-bpp DIBs
' are used because we perform the actual capture using a 24-bpp DIB (there is a measurable
' perf benefit to this on Win 10), but the PNG export engine requires 32-bpp DIBs because
' it needs to perform heuristics to see if there's benefits to writing 32-bpp frames if
' transparency provides meaningful benefits during frame optimizations.  (Don't sweat the
' details!)  Anyway, the point is that all captured frames are captured and cached as 24-bpp,
' but they are upsampled to 32-bpp before getting passed to the APNG encoder.
Private m_captureDIB24 As pdDIB, m_captureDIB32 As pdDIB

'Captured frames are stored as a collection of lz4-compressed arrays.  I don't currently have
' access to a DEFLATE library that can compress screen-capture-sized frames (e.g. 1024x768) in
' real-time on an XP-era PC.  lz4 is a better solution here, although it requires us to
' "play back" the frames when the capture ends.
Private Type FrameCapture
    fcTimeStamp As Currency
    frameSizeOrig As Long
    frameSizeCompressed As Long
    frameData() As Byte
End Type

'Frame count is tracked manually; certain things need to be handled differently on
' e.g. the first frame vs subsequent frames.
Private Const INIT_FRAME_BUFFER As Long = 64
Private m_FrameCount As Long
Private m_Frames() As FrameCapture

'For perf reasons, a persistent compression buffer is used; it is auto-enlarged to
' a "worst-case" size before capture begins.
Private m_CompressionBuffer() As Byte

Private Sub cmdExit_Click()
    StopTimer_Forcibly
    Unload Me
End Sub

Private Sub cmdExport_Click()
    
    Dim cSave As pdOpenSaveDialog
    Set cSave = New pdOpenSaveDialog
    
    'Provide a string to the common dialog; it will fill this with the user's chosen path + filename
    Dim tmpPath As String, tmpFilename As String, sFile As String
    If (LenB(txtExport.Text) <> 0) Then
        sFile = txtExport.Text
        tmpPath = Files.FileGetPath(sFile)
    Else
        tmpPath = UserPrefs.GetPref_String("Paths", "ScreenCapture", vbNullString)
        tmpFilename = g_Language.TranslateMessage("capture")
        sFile = tmpPath & IncrementFilename(tmpPath, tmpFilename, "png")
    End If
    
    'Present a common dialog to the user
    If cSave.GetSaveFileName(sFile, , True, "Animated PNG (.png)|*.png", 1, tmpPath, g_Language.TranslateMessage("Export screen capture animation"), "png", Me.hWnd) Then
        txtExport.Text = sFile
    End If
    
End Sub

Private Sub cmdStart_Click()
    
    'If a capture is already active, we STOP the timer now
    If m_CaptureActive Then
        
        'Immediately stop the capture timer
        m_CaptureActive = False
        If (Not m_Timer Is Nothing) Then m_Timer.StopTimer
        
        'Write all frames?
        Dim i As Long
        For i = 0 To m_FrameCount - 1
            
            lblProgress.Caption = g_Language.TranslateMessage("Processing frame %1 of %2", i + 1, m_FrameCount)
            lblProgress.RequestRefresh
            VBHacks.DoEvents_PaintOnly False
            
            'Extract this frame into the capture DIB
            Compression.DecompressPtrToPtr m_captureDIB24.GetDIBPointer, m_Frames(i).frameSizeOrig, VarPtr(m_Frames(i).frameData(0)), m_Frames(i).frameSizeCompressed, cf_Lz4
            
            'Convert the 24-bpp DIB to 32-bpp before handing it off to the APNG encoder
            If (m_captureDIB32 Is Nothing) Then
                Set m_captureDIB32 = New pdDIB
                m_captureDIB32.CreateBlank m_captureDIB24.GetDIBWidth, m_captureDIB24.GetDIBHeight, 32, 0, 255
            End If
            
            GDI.BitBltWrapper m_captureDIB32.GetDIBDC, 0, 0, m_captureDIB32.GetDIBWidth, m_captureDIB32.GetDIBHeight, m_captureDIB24.GetDIBDC, 0, 0, vbSrcCopy
            m_captureDIB32.ForceNewAlpha 255
            
            'Pass the frame off to the PNG encoder
            'If (i = m_FrameCount - 1) Then m_Frames(i).fcTimeStamp = m_Frames(i).fcTimeStamp + 3000
            m_PNG.SaveAPNG_Streaming_Frame m_captureDIB32, m_Frames(i).fcTimeStamp
            
        Next i
        
        'Notify the PNG encoder that the stream has ended
        lblProgress.Caption = g_Language.TranslateMessage("Capture complete!")
        If (Not m_PNG Is Nothing) Then m_PNG.SaveAPNG_Streaming_Stop 0, 0
        
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
            
            'Save the current export path to the user's preference file
            UserPrefs.SetPref_String "Paths", "ScreenCapture", m_DstFilename
            
            'Initialize the timer
            Set m_Timer = New pdTimer
            m_Timer.Interval = Int(1000# / sldFrameRate.Value + 0.5)
            
            'Determine the capture rect
            With m_CaptureRect
                .Left = txtCoords(0).Text
                .Top = txtCoords(1).Text
                .Right = .Left + txtCoords(2).Text
                .Bottom = .Top + txtCoords(3).Text
            End With
            
            'Prep the capture DIB
            Set m_captureDIB24 = New pdDIB
            m_captureDIB24.CreateBlank m_CaptureRect.Right - m_CaptureRect.Left, m_CaptureRect.Bottom - m_CaptureRect.Top, 24, vbBlack, 255
            
            'Prepare a persistent compression buffer
            Dim cmpBufferSize As Long
            cmpBufferSize = Compression.GetWorstCaseSize(m_captureDIB24.GetDIBStride * m_captureDIB24.GetDIBHeight, cf_Lz4)
            ReDim m_CompressionBuffer(0 To cmpBufferSize - 1) As Byte
            
            'Initialize the frame collection
            m_FrameCount = 0
            ReDim m_Frames(0 To INIT_FRAME_BUFFER - 1) As FrameCapture
            
            'Initialize the PNG writer
            Set m_PNG = New pdPNG
            If (m_PNG.SaveAPNG_Streaming_Start(m_DstFilename, m_captureDIB24.GetDIBWidth, m_captureDIB24.GetDIBHeight) < png_Failure) Then
                
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
    
    'Get the last "apng capture" path from the preferences file, and if it doesn't exist,
    ' default to the user's current "save image" path
    Dim tempPathString As String
    tempPathString = UserPrefs.GetPref_String("Paths", "ScreenCapture", vbNullString)
    If (LenB(tempPathString) = 0) Then
        tempPathString = UserPrefs.GetPref_String("Paths", "Save Image", vbNullString)
        If (LenB(tempPathString) <> 0) Then tempPathString = tempPathString & g_Language.TranslateMessage("capture") & ".png"
    End If
    
    'Place that path in the export box with "capture" appended, as the default export path
    txtExport.Text = tempPathString
    
    Interface.ApplyThemeAndTranslations Me
    
End Sub

'On timer events, capture the current frame by calling this function
Private Sub CaptureFrameNow()
    
    'Ensure we have an active APNG instance
    If (Not m_PNG Is Nothing) Then
        
        'Make sure we have room to store this frame
        If (m_FrameCount > UBound(m_Frames)) Then ReDim Preserve m_Frames(0 To m_FrameCount * 2 - 1) As FrameCapture
        
        '*Immediately* before capture, note the current time
        Dim capTime As Currency, testTime As Currency
        capTime = VBHacks.GetHighResTimeEx()
        
        'Capture the frame in question into a pdDIB object
        ScreenCapture.GetPartialDesktopAsDIB m_captureDIB24, m_CaptureRect
        
        'Compress the captured frame into the temporary frame buffer
        Dim cmpSize As Long
        cmpSize = UBound(m_CompressionBuffer) + 1
        Compression.CompressPtrToPtr VarPtr(m_CompressionBuffer(0)), cmpSize, m_captureDIB24.GetDIBPointer, m_captureDIB24.GetDIBStride * m_captureDIB24.GetDIBHeight, cf_Lz4
        
        'Allocate memory for this frame and store all data
        With m_Frames(m_FrameCount)
            .fcTimeStamp = capTime
            .frameSizeOrig = m_captureDIB24.GetDIBStride * m_captureDIB24.GetDIBHeight
            .frameSizeCompressed = cmpSize
            ReDim .frameData(0 To cmpSize - 1) As Byte
            CopyMemoryStrict VarPtr(.frameData(0)), VarPtr(m_CompressionBuffer(0)), cmpSize
        End With
        
        PDDebug.LogAction "Total frame time: " & VBHacks.GetTimeDiffNowAsString(capTime) & ", compression reduced size by " & CStr(100# * (1# - (cmpSize / (m_captureDIB24.GetDIBStride * m_captureDIB24.GetDIBHeight)))) & "%"
        
        'Increment frame count
        m_FrameCount = m_FrameCount + 1
        
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'Forcibly stop the current capture (if any)
    StopTimer_Forcibly
    
End Sub

Private Sub m_Timer_Timer()
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
