VERSION 5.00
Begin VB.Form FormImportFrx 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Import From VB Binary File"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   7005
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "VBP_FormImportFrx.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   467
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Left            =   3360
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5640
      MouseIcon       =   "VBP_FormImportFrx.frx":0BC2
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   3960
      Width           =   1125
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   4440
      MouseIcon       =   "VBP_FormImportFrx.frx":0D14
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   3960
      Width           =   1125
   End
   Begin VB.ListBox LstInfo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   3000
      Left            =   120
      MouseIcon       =   "VBP_FormImportFrx.frx":0E66
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   720
      Width           =   3495
   End
   Begin VB.Label LblData 
      BackStyle       =   0  'Transparent
      Caption         =   "No Image Selected"
      ForeColor       =   &H00400000&
      Height          =   2535
      Left            =   3960
      TabIndex        =   7
      Top             =   840
      Width           =   2835
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2985
      Left            =   3840
      Stretch         =   -1  'True
      Top             =   720
      Width           =   2970
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Image Preview:"
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   3840
      TabIndex        =   6
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Type  -  Offset  -  Size (in bytes)"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label LblCurFile 
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current file:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   870
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
'(Some segments adopted from the original version, which is ©1997-1999 by Brad Martinez, http://www.mvps.org - see Outside_ModFrx.bas for more details)
'Created: 2/14/03
'Last updated: 5/June/12
'Last update: if binary import fails, allow the user to try another file
'
'Module for importing images from VB binary files.  Allows the user to browse through
'all data within the resource file, and load any images (ico, jpeg, whatever) contained
'therein.
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
    tmpFRXFile = TempPath & "PDFRXInterface.tmp"
    SavePicture PicLoadImage.Picture, tmpFRXFile
    
    'Because PreLoadImage requires a string array, create one to send it
    Dim sFile(0) As String
    sFile(0) = tmpFRXFile
        
    PreLoadImage sFile, False, "VB Binary Image (Imported)", "VB Binary Image (" & Day(Now) & " " & MonthName(Month(Now)) & " " & Year(Now) & ")"
        
    'Be polite and remove the temporary file
    Kill tmpFRXFile
        
    Message "Image imported successfully "
        
    FormMain.SetFocus
    
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
    tempPathString = GetFromIni("Program Paths", "ImportFRX")
    
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
        WriteToIni "Program Paths", "ImportFRX", tempPathString
   
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
        setHandCursorForAll Me
        
    'If the commondialog box is canceled...
    Else
        requestUnload = True
    End If
    
End Sub

Private Sub LstInfo_Click()
  
    With m_cff(LstInfo.ListIndex + 1)
        If .PictureType Then
            Image1.Picture = .Picture
            CmdOK.Enabled = True
            LblData.Visible = False
            DoEvents
        Else
            Image1.Picture = Nothing
            LblData.Visible = True
            CmdOK.Enabled = False
            If .ImageSize And (.ImageSize < 2 ^ 15) Then
                LblData.Caption = "Binary data: " & StrConv(.Bits, vbUnicode)
            Else
                LblData.Caption = PROGRAMNAME & " is unable to display this data.  It may be from an incompatible version of Visual Basic, the source file may be corrupted, or the data may exceed 32k in size."
            End If
        End If
  End With
  
End Sub
