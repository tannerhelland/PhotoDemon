VERSION 5.00
Begin VB.Form frmUpdate 
   BackColor       =   &H80000005&
   Caption         =   "PhotoDemon Update Generator"
   ClientHeight    =   6915
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12630
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   461
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   842
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAction 
      Caption         =   "Assemble nightly build files"
      Height          =   615
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   12135
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Step 1: copy all relevant nightly build files into a dedicated /nightly folder"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00DC7032&
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8040
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Program Auto-Update Generator
'Copyright ©2015-2015 by Tanner Helland
'Created: 28/January/15
'Last updated: 10/February/15
'Last update: continued work on initial build
'
'This project was built to help assemble automatic update information for PhotoDemon.  It is run by the nightly build batch
' script, and it assembles a few different things:
' - New pdPackage archives for the latest stable, beta (if relevant), and nightly build entries.
' - A master update file with version numbers and checksums for each of the pdPackage files
'
'NOTE: this project is intended only as a support tool for PhotoDemon.  It is not designed or tested for general-purpose use.
'       I do not have any intention of supporting this tool outside its intended use, so please do not submit bug reports
'       regarding this project unless they directly relate to its intended purpose (generating PhotoDemon update files).
'
'       Also, given this project's purpose, the code is pretty ugly.  Organization is minimal.  Read at your own risk.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

'PD update patch identifier.  IMPORTANT NOTE: this constant is shared with the main PhotoDemon project.  DO NOT CHANGE IT!
Private Const PD_PATCH_IDENTIFIER As Long = &H50554450   'PD update patch data (ASCII characters "PDUP", as hex, little-endian)

'A module-level pdFSO object is provided, for convenience.
Private m_File As pdFSO

'If silent mode has been activated via command line, this will be set to TRUE.
Dim m_SilentMode As Boolean

Private Sub cmdAction_Click(Index As Integer)
    
    Select Case Index
    
        'Assemble nightly build files
        Case 0
            AssembleNightlyBuild
            
    
    End Select
    
End Sub

'Copy the relevant nightly build files from their default VB project location, to a dedicated /Nightly folder.
' This greatly simplifies the pdPackage generation step, as we can handle the dedicated /Nightly folder the same way
' we handle the /Stable and /Beta folders.
Private Sub AssembleNightlyBuild()

    'This list of relevant files is hardcoded to match the nightly build script's instructions for 7zip.
    Dim nightlyList As pdStringStack
    Set nightlyList = New pdStringStack
    
    nightlyList.AddString "C:\PhotoDemon v4\PhotoDemon\PhotoDemon.exe"
    nightlyList.AddString "C:\PhotoDemon v4\PhotoDemon\README.txt"
    nightlyList.AddString "C:\PhotoDemon v4\PhotoDemon\Donate to PhotoDemon.url"
    
    'For the /App subfolder, we forcibly restrict which extensions are allowed, to avoid copying any backup files
    ' or other unwanted entries.
    m_File.retrieveAllFiles "C:\PhotoDemon v4\PhotoDemon\App\", nightlyList, True, False, "exe|txt|TXT|dll"
    
    'Assemble the corresponding pdPackage
    Dim nightlyPackage As pdPackager
    Set nightlyPackage = New pdPackager
    nightlyPackage.init_ZLib App.Path & "\zlibwapi.dll"
    nightlyPackage.prepareNewPackage 4, PD_PATCH_IDENTIFIER
    
    nightlyPackage.autoAddNodesFromStringStack nightlyList, "C:\PhotoDemon v4\PhotoDemon\", 0, True, True
    
    'Write the completed package out to the updates folder
    nightlyPackage.writePackageToFile "C:\PhotoDemon v4\PhotoDemon\no_sync\PD_Updates\nightly.pdz", True, True
    
    'TEMPORARY TEST ONLY!  Extract the files to a temp folder, to make sure they unpack correctly.
    nightlyPackage.readPackageFromFile "C:\PhotoDemon v4\PhotoDemon\no_sync\PD_Updates\nightly.pdz", PD_PATCH_IDENTIFIER
    nightlyPackage.autoExtractAllFiles "C:\PhotoDemon v4\PhotoDemon\no_sync\PD_Updates\nightly\"
    

End Sub

Private Sub Form_Load()
    
    Set m_File = New pdFSO
    
    'Check the command line.  This project can be run in silent mode as part of my nightly build batch script.
    Dim chkCommandLine As String
    chkCommandLine = Command$
    
    If Len(Trim$(chkCommandLine)) <> 0 Then
        If InStr(1, chkCommandLine, "-s", vbTextCompare) Then m_SilentMode = True Else m_SilentMode = False
    End If
    
    'If silent mode is activated, automatically "click" the relevant button
    If m_SilentMode Then
    
        'Assemble the nightly build .pdz
        Call cmdAction_Click(0)
        
    End If
    
End Sub
