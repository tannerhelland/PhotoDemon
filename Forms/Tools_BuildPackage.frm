VERSION 5.00
Begin VB.Form FormPackage 
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create standalone pdPackage"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11295
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
   ScaleHeight     =   498
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   753
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin PhotoDemon.pdCheckBox chkOptions 
      Height          =   375
      Index           =   0
      Left            =   7560
      TabIndex        =   4
      Top             =   5760
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      Caption         =   "compress individual files"
   End
   Begin PhotoDemon.pdButton cmdSave 
      Height          =   855
      Left            =   3840
      TabIndex        =   3
      Top             =   5760
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1508
      Caption         =   "save the final package..."
   End
   Begin PhotoDemon.pdButton cmdAdd 
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   5760
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1508
      Caption         =   "add file(s) to the package..."
   End
   Begin PhotoDemon.pdListBox lstFiles 
      Height          =   5535
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   9763
      Caption         =   "files in this package:"
   End
   Begin PhotoDemon.pdCommandBarMini cmdBar 
      Align           =   2  'Align Bottom
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   6735
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   1296
   End
   Begin PhotoDemon.pdCheckBox chkOptions 
      Height          =   375
      Index           =   1
      Left            =   7560
      TabIndex        =   5
      Top             =   6240
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      Caption         =   "also compress directory"
   End
End
Attribute VB_Name = "FormPackage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()

    Dim srcFiles As String, finalList As pdStringStack
    Set finalList = New pdStringStack
    
    Dim cCommonDialog As pdOpenSaveDialog: Set cCommonDialog = New pdOpenSaveDialog
    If cCommonDialog.GetOpenFileName(srcFiles, , True, True, , , UserPrefs.GetAppPath, "Select one or more files", , Me.hWnd) Then
        
        'Take the return string (a null-delimited list of filenames) and split it out into a string array
        Dim listOfFiles() As String
        listOfFiles = Split(srcFiles, vbNullChar)
        
        Dim i As Long
        
        'Due to the buffering required by the API call, UBound(listOfFiles) should ALWAYS > 0 but
        ' let's check it anyway (just to be safe)
        If (UBound(listOfFiles) > 0) Then
        
            'Remove all empty strings from the array (which are a byproduct of the aforementioned buffering)
            For i = UBound(listOfFiles) To 0 Step -1
                If (LenB(listOfFiles(i)) <> 0) Then Exit For
            Next
            
            'With all the empty strings removed, all that's left is legitimate file paths
            ReDim Preserve listOfFiles(0 To i) As String
            
        End If
        
        'If multiple files were selected, we need to do some additional processing to the array
        If (UBound(listOfFiles) > 0) Then
        
            'The common dialog function returns a unique array. Index (0) contains the folder path (without a
            ' trailing backslash), so first things first - add a trailing backslash
            Dim basePath As String
            basePath = Files.PathAddBackslash(listOfFiles(0))
            
            'The remaining indices contain a filename within that folder.  To get the full filename, we must
            ' append the path from (0) to the start of each filename.  This will relieve the burden on
            ' whatever function called us - it can simply loop through the full paths, loading files as it goes
            For i = 1 To UBound(listOfFiles)
                finalList.AddString basePath & listOfFiles(i)
            Next i
            
        'If there is only one file in the array (e.g. the user only opened one image), we don't need to do all
        ' that extra processing - just retrieve the only file path as-is
        Else
            finalList.AddString listOfFiles(0)
        End If
        
        'Sort the list of files, then add it to the primary list box
        finalList.SortAlphabetically
        
        For i = 0 To finalList.GetNumOfStrings - 1
            lstFiles.AddItem finalList.GetString(i)
        Next i
        
    End If
    
End Sub

Private Sub cmdSave_Click()

    'Start a new pdPackage object; it does all the heavy lifting for us.
    Dim cPackage As pdPackager
    Set cPackage = New pdPackager
    cPackage.PrepareNewPackage
    
    'Set compression options for individual files (settable by the user)
    Dim cmpType As PD_CompressionFormat, cmpLevel As Long
    If chkOptions(0).Value Then
        cmpType = cf_Zstd
        cmpLevel = Compression.GetMaxCompressionLevel(cf_Zstd)
    Else
        cmpType = cf_None
        cmpLevel = Compression.GetDefaultCompressionLevel(cf_None)
    End If
    
    'Load all file data into the package
    Dim i As Long, idxNode As Long, tmpBytes() As Byte
    For i = 0 To lstFiles.ListCount - 1
        idxNode = cPackage.AddNode(vbNullString, i)
        If Files.FileLoadAsByteArray(lstFiles.List(i), tmpBytes) Then
            cPackage.AddNodeDataFromString idxNode, True, Files.FileGetName(lstFiles.List(i)), cmpType, cmpLevel
            cPackage.AddNodeDataFromByteArray idxNode, False, tmpBytes, cmpType, cmpLevel
        Else
            Debug.Print "WARNING!  Couldn't load file: " & lstFiles.List(i)
        End If
    Next i
    
    'Save the package using any additional settings supplied by the user
    Dim cDialog As pdOpenSaveDialog
    Set cDialog = New pdOpenSaveDialog
    
    Dim dstFile As String
    If cDialog.GetSaveFileName(dstFile, , True, , , UserPrefs.GetAppPath, "Save pdPackage", "pdp", Me.hWnd) Then
    
        If chkOptions(1).Value Then
            cmpType = cf_Zstd
            cmpLevel = Compression.GetMaxCompressionLevel(cf_Zstd)
        Else
            cmpType = cf_None
            cmpLevel = Compression.GetDefaultCompressionLevel(cf_None)
        End If
        
        cPackage.WritePackageToFile dstFile, cmpType, False, cmpLevel
    
    End If

End Sub

Private Sub Form_Load()

    Interface.ApplyThemeAndTranslations Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub
