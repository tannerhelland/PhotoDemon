VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "PhotoDemon Compression Support Utility by Tanner Helland"
   ClientHeight    =   1530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5175
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   102
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   345
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSave 
      Caption         =   "Decompress a File"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton CmdCompress 
      Caption         =   "Compress a File"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "This utility requires the zLib WAPI-variant DLL."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   4695
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Plugin Compression Tool (uses zLib)
'Copyright ©2002-2013 by Tanner Helland
'Created: 3/02/02
'Last updated: 18/June/2013
'Last update: moved project into main PD Git repository
'
'Module to handle file compression and decompression to a custom file format via the zLib compression library.
'
'NOTE: this project is intended only as a support tool for PhotoDemon.  It is not designed or tested for general-purpose use.
'       I do not have any intention of supporting this tool outside its intended use, so please do not submit bug reports
'       regarding this project unless they directly relate to its intended purpose (compressing PhotoDemon plugins).
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Private Sub CmdCompress_Click()
        
    'String returned from the common dialog wrapper
    Dim sFile As String
    
    'Basic filter string
    Dim cdfStr As String
    cdfStr = "All Files|*.*"
    
    'Common dialog interface
    Dim CC As cCommonDialog
    Set CC = New cCommonDialog
    
    'If cancel isn't selected, load a picture from the user-specified file
    If CC.VBGetOpenFileName(sFile, , , , , True, cdfStr, , , "Open a file for compression", , Me.hWnd, 0) Then
    
        CompressFile sFile, True, True
        
    End If
    
End Sub

Private Sub CmdSave_Click()
    'String returned from the common dialog wrapper
    Dim sFile As String
    
    'Basic filter string
    Dim cdfStr As String
    cdfStr = "All Files|*.*"
    
    'Common dialog interface
    Dim CC As cCommonDialog
    Set CC = New cCommonDialog
    
    'If cancel isn't selected, load a picture from the user-specified file
    If CC.VBGetOpenFileName(sFile, , , , , True, cdfStr, , , "Open a file for compression", , Me.hWnd, 0) Then
    
        DecompressFile sFile, True, True
        
    End If
End Sub

Private Sub Form_Load()
    initializeZLib
End Sub
