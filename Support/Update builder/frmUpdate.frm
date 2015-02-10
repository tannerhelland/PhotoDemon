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
      Caption         =   "Assemble stable and beta build packages (dedicated folders)"
      Height          =   615
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   12135
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Assemble nightly build package (direct from current development folder)"
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
      Caption         =   "Step 2: assemble the stable and beta update packages"
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
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   5925
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Step 1: assemble the nightly build update package"
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
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5505
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
            
        'Assemble stable + beta build files
        Case 1
            AssembleStableAndBetaBuilds
            
    
    End Select
    
End Sub

'The nightly build is unique, because we generate it directly from the current PD development folder.  As such, it uses a
' different series of assembly steps (compared to the stable and beta builds).
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
    'nightlyPackage.readPackageFromFile "C:\PhotoDemon v4\PhotoDemon\no_sync\PD_Updates\nightly.pdz", PD_PATCH_IDENTIFIER
    'nightlyPackage.autoExtractAllFiles "C:\PhotoDemon v4\PhotoDemon\no_sync\PD_Updates\nightly\"
    

End Sub

'Stable and Beta update channels use custom, dedicated folders.  The contents of these folders are updated manually,
' only when necessary (as opposed to the nightly channel which is built directly from the current PD codebase).
Private Sub AssembleStableAndBetaBuilds()

    'Stable and beta builds can be constructed directly from their folders, no special work required.
    
    'Assemble a basic pdPackage instance
    Dim cPackage As pdPackager
    Set cPackage = New pdPackager
    cPackage.init_ZLib App.Path & "\zlibwapi.dll"
    
    'Build the stable update file directly from its folder.  Unlike the nightly build, all files are allowed, including
    ' XML language files.
    cPackage.prepareNewPackage 4, PD_PATCH_IDENTIFIER
    cPackage.autoAddNodesFromFolder "C:\PhotoDemon v4\PhotoDemon\no_sync\PD_Updates\stable\"
    cPackage.writePackageToFile "C:\PhotoDemon v4\PhotoDemon\no_sync\PD_Updates\stable.pdz", True, True
    
    'Repeat the above steps for the beta update folder
    cPackage.prepareNewPackage 4, PD_PATCH_IDENTIFIER
    cPackage.autoAddNodesFromFolder "C:\PhotoDemon v4\PhotoDemon\no_sync\PD_Updates\beta\"
    cPackage.writePackageToFile "C:\PhotoDemon v4\PhotoDemon\no_sync\PD_Updates\beta.pdz", True, True
    
    'TEMPORARY TEST ONLY!  Extract an update package to a temp folder, to make sure everything's in order.
    'cPackage.readPackageFromFile "C:\PhotoDemon v4\PhotoDemon\no_sync\PD_Updates\stable.pdz", PD_PATCH_IDENTIFIER
    'cPackage.autoExtractAllFiles "C:\PhotoDemon v4\PhotoDemon\no_sync\PD_Updates\testing_only\"
    

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
    
        'Assemble the nightly build update package
        Call cmdAction_Click(0)
        
        'assemble the stable and beta build update packages
        Call cmdAction_Click(1)
        
    End If
    
End Sub
