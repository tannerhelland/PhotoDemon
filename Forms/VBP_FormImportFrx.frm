VERSION 5.00
Begin VB.Form FormImportFrx 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Import From VB Binary File"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   7710
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
   ScaleHeight     =   363
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   514
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picDemo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   2940
      Left            =   3960
      ScaleHeight     =   194
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   238
      TabIndex        =   8
      Top             =   1440
      Width           =   3600
      Begin VB.Label LblData 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "No Image Selected"
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
         Height          =   2535
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   3315
      End
   End
   Begin VB.PictureBox PicLoadImage 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
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
      Left            =   240
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   7
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6240
      TabIndex        =   1
      Top             =   4800
      Width           =   1245
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Left            =   4920
      TabIndex        =   0
      Top             =   4800
      Width           =   1245
   End
   Begin VB.ListBox LstInfo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2940
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   3600
   End
   Begin VB.Image imgTemp 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   240
      Top             =   4680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblPreview 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "image preview:"
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
      Height          =   405
      Left            =   3960
      TabIndex        =   6
      Top             =   1080
      Width           =   1635
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "data type | offset | size (bytes)"
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
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   3495
   End
   Begin VB.Label LblCurFile 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "current file:"
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
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   1230
   End
End
Attribute VB_Name = "FormImportFrx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'VB Binary File Import Tool
'Copyright ©2000-2012 by Tanner Helland
' (Some segments adopted from the original version, which is ©1997-1999 by Brad Martinez, http://www.mvps.org - see Outside_ModFrx.bas for more details)
'Created: 2/14/03
'Last updated: 26/September/12
'Last update: vastly improved preview rendering.  Small images are drawn at actual size.  Large images are resized to fit inside the
'             demonstration picture box (with aspect ratio preserved).
'
'Module for importing images from VB binary files.  Allows the user to browse through all data
' within the resource file, and optionally load any images (ico, jpeg, whatever) contained therein.
'
'***************************************************************************

Option Explicit

Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long

Dim m_cff As New cFrxFile
Dim requestUnload As Boolean

'CANCEL button
Private Sub CmdCancel_Click()
    
    Message "VB binary file import canceled"
    Unload Me
    
End Sub

'OK button
Private Sub CmdOK_Click()
    Me.Visible = False
    Message "Importing binary image data..."
    PicLoadImage.Picture = m_cff(LstInfo.ListIndex + 1).Picture
    
    'Save this picture to a temporary file
    Dim tmpFRXFile As String
    tmpFRXFile = userPreferences.getTempPath & "PDFRXInterface.tmp"
    SavePicture PicLoadImage.Picture, tmpFRXFile
    
    'Because PreLoadImage requires a string array, create one to send it
    Dim sFile(0) As String
    sFile(0) = tmpFRXFile
        
    PreLoadImage sFile, False, "VB Binary Image (Imported)", "VB Binary Image (" & Day(Now) & " " & MonthName(Month(Now)) & " " & Year(Now) & ")"
        
    'Be polite and remove the temporary file
    Kill tmpFRXFile
        
    Message "Image imported successfully "
        
    If FormMain.Enabled Then FormMain.SetFocus
    
    Unload Me
End Sub

'If the user cancels the initial common dialog box, the requestUnload flag will be set to TRUE
Private Sub Form_Activate()
    If requestUnload = True Then Unload Me
End Sub

'LOAD form
Private Sub Form_Load()
    
    Dim cfi As cFrxItem
    Dim CC As cCommonDialog
    
    requestUnload = False
    
    'Get the last "open FRX" path from the INI file
    Dim tempPathString As String
    tempPathString = userPreferences.GetPreference_String("Program Paths", "ImportFRX", "")
    
    'File returned from the CommonDialog
    Dim sFile As String
    
    Message "Please select a VB binary file to scan for images..."
    
    'Call up the CommonDialog routine
TryBinaryImportAgain:
    
    Set CC = New cCommonDialog
    
    If CC.VBGetOpenFileName(sFile, , , , , True, "All VB Binary Files (*.frx,*.ctx,*.dsx,*.dox,*.pgx)|*.frx;*.ctx;*.dsx;*.dox;*.pgx|All files (*.*)|*.*", , tempPathString, "Select a VB Binary File", , FormMain.HWnd, 0) Then
    
        Message "Scanning binary file..."
    
        'Save the new path for future usage
        sFile = Left$(sFile, lstrlen(sFile))
        tempPathString = sFile
        StripDirectory tempPathString
        userPreferences.SetPreference_String "Program Paths", "ImportFRX", tempPathString
   
        'Assign the file (note: this may take some time if the file is invalid)
        m_cff.Path = sFile
    
        If m_cff.Count Then
            LblCurFile.Caption = sFile
            LstInfo.Clear
            DoEvents

            For Each cfi In m_cff
                LstInfo.AddItem cfi.FileTypeName & vbTab & cfi.FileOffset & vbTab & cfi.ImageSize
            Next
            
            'LstInfo.ListIndex = 0
      
            Message "Scan complete.  Please select an image to import."
      
        Else
            MsgBox "Unfortunately, no images were found in " & sFile & ".  Please select a new file.", vbCritical + vbApplicationModal + vbOKOnly, "No Images Found"
            GoTo TryBinaryImportAgain
        End If
        
        'Assign the system hand cursor to all relevant objects
        makeFormPretty Me
        
    'If the commondialog box is canceled...
    Else
        requestUnload = True
    End If
    
End Sub

Private Sub LstInfo_Click()
  
    Dim tmpImportLayer As pdLayer
    Set tmpImportLayer = New pdLayer
  
    With m_cff(LstInfo.ListIndex + 1)
        If .PictureType Then
        
            'Convert the stupid HiMetric size of this StdPicture to pixels
            Dim ImgWidth As Long, ImgHeight As Long
            ImgWidth = CInt(picDemo.ScaleX(.Picture.Width, vbHiMetric, vbPixels))
            ImgHeight = CInt(picDemo.ScaleY(.Picture.Height, vbHiMetric, vbPixels))
        
            'Icons and small images can be drawn at scale.  Large images must be scaled to the size of the sample picture box
            If (.PictureType <> ptICO) And ((ImgWidth > picDemo.ScaleWidth) Or (ImgHeight > picDemo.ScaleHeight)) Then
        
                'Use a temporary layer to render the image to the sample picture box
                tmpImportLayer.CreateFromPicture m_cff(LstInfo.ListIndex + 1).Picture
                If tmpImportLayer.getLayerWidth <> 0 And tmpImportLayer.getLayerHeight <> 0 Then tmpImportLayer.renderToPictureBox picDemo
                
                CmdOK.Enabled = True
                LblData.Visible = False
                DoEvents
            
            Else
                
                picDemo.Picture = LoadPicture("")
                PicLoadImage.Picture = .Picture
                
                'Center the image in the sample area
                BitBlt picDemo.hDC, (picDemo.ScaleWidth \ 2) - (ImgWidth \ 2), (picDemo.ScaleHeight \ 2) - (ImgHeight \ 2), ImgWidth, ImgHeight, PicLoadImage.hDC, 0, 0, vbSrcCopy
                picDemo.Picture = picDemo.Image
                picDemo.Refresh
                
                CmdOK.Enabled = True
                LblData.Visible = False
                DoEvents
            
            End If
            
        Else
            picDemo.Picture = Nothing
            LblData.Visible = True
            CmdOK.Enabled = False
            If .ImageSize And (.ImageSize < 2 ^ 15) Then
                LblData.Caption = "Binary data: " & StrConv(.Bits, vbUnicode)
            Else
                LblData.Caption = PROGRAMNAME & " is unable to display this data.  It may be from an incompatible version of Visual Basic, the source file may be corrupted, or the data may exceed 32k in size."
            End If
        End If
    End With
    
    tmpImportLayer.eraseLayer
    Set tmpImportLayer = Nothing
    
End Sub
