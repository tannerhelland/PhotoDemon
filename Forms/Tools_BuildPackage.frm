VERSION 5.00
Begin VB.Form FormPackage 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create standalone pdPackage"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11295
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
End
Attribute VB_Name = "FormPackage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Developer package construction dialog
'Copyright 2019-2026 by Tanner Helland
'Created: 04/March/19
'Last updated: 14/October/19
'Last update: use the new pdPackageChunky format for resource collections
'
'As of v8.0, PD ships with some resource collections (e.g. prebuilt gradients).  These are stored
' in PD's resource segment, and it is helpful to package such collections into their own sub-archives.
' This dialog helps you construct sub-archives like this.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Sub cmdAdd_Click()

    Dim finalList As pdStringStack
    
    Dim cCommonDialog As pdOpenSaveDialog: Set cCommonDialog = New pdOpenSaveDialog
    If cCommonDialog.GetOpenFileNames_AsStringStack(finalList, vbNullString, vbNullString, True, , , UserPrefs.GetAppPath, "Select one or more files", , Me.hWnd) Then
        
        'Sort the list of files, then add it to the primary list box
        finalList.SortAlphabetically
        
        Dim i As Long
        For i = 0 To finalList.GetNumOfStrings - 1
            lstFiles.AddItem finalList.GetString(i)
        Next i
        
    End If
    
End Sub

Private Sub cmdSave_Click()

    'Set compression options for individual files (settable by the user)
    Dim cmpType As PD_CompressionFormat, cmpLevel As Long
    If chkOptions(0).Value Then
        cmpType = cf_Zstd
        cmpLevel = Compression.GetMaxCompressionLevel(cf_Zstd)
    Else
        cmpType = cf_None
        cmpLevel = Compression.GetDefaultCompressionLevel(cf_None)
    End If
    
    'Save the package using any additional settings supplied by the user
    Dim cDialog As pdOpenSaveDialog
    Set cDialog = New pdOpenSaveDialog
    
    Dim dstFile As String
    If cDialog.GetSaveFileName(dstFile, , True, , , UserPrefs.GetAppPath, "Save pdPackage", "pdp", Me.hWnd) Then
    
        'Start a new pdPackage; it does all the heavy lifting for us
        Dim cPackage As pdPackageChunky
        Set cPackage = New pdPackageChunky
        If cPackage.StartNewPackage_File(dstFile) Then
            
            'Load each file as raw binary data, then add it to the running package collection
            Dim i As Long, tmpBytes() As Byte
            For i = 0 To lstFiles.ListCount - 1
                If Files.FileLoadAsByteArray(lstFiles.List(i), tmpBytes) Then
                    If (Not cPackage.AddChunk_NameValuePair("NAME", Files.FileGetName(lstFiles.List(i)), "DATA", VarPtr(tmpBytes(0)), UBound(tmpBytes) + 1, cmpType, cmpLevel)) Then Debug.Print "WARNING!  Couldn't add chunk: " & lstFiles.List(i)
                Else
                    Debug.Print "WARNING!  Couldn't load file: " & lstFiles.List(i)
                End If
            Next i
        
        End If
        
    End If

End Sub

Private Sub Form_Load()
    Interface.ApplyThemeAndTranslations Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub
