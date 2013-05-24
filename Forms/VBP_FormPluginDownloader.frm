VERSION 5.00
Begin VB.Form FormPluginDownloader 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " PhotoDemon Plugin Downloader"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12405
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
   ScaleHeight     =   7320
   ScaleWidth      =   12405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picInitial 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7335
      Left            =   0
      ScaleHeight     =   489
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   825
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   12375
      Begin PhotoDemon.jcbutton cmdChoice 
         Default         =   -1  'True
         Height          =   1605
         Index           =   0
         Left            =   180
         TabIndex        =   0
         Top             =   5520
         Width           =   5970
         _ExtentX        =   10530
         _ExtentY        =   2831
         ButtonStyle     =   13
         ShowFocusRect   =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Yes. Download these files to the plugins folder."
         HandPointer     =   -1  'True
         PictureNormal   =   "VBP_FormPluginDownloader.frx":0000
         PictureAlign    =   0
         DisabledPictureMode=   1
         CaptionEffects  =   0
         TooltipTitle    =   "Download All Plugins"
      End
      Begin PhotoDemon.jcbutton cmdChoice 
         Cancel          =   -1  'True
         Height          =   765
         Index           =   1
         Left            =   6240
         TabIndex        =   1
         Top             =   5520
         Width           =   5970
         _ExtentX        =   10530
         _ExtentY        =   1349
         ButtonStyle     =   13
         ShowFocusRect   =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Not right now, but please remind me later."
         HandPointer     =   -1  'True
         PictureNormal   =   "VBP_FormPluginDownloader.frx":1052
         PictureAlign    =   0
         DisabledPictureMode=   1
         CaptionEffects  =   0
         TooltipTitle    =   "Postpone Plugin Download"
      End
      Begin PhotoDemon.jcbutton cmdChoice 
         Height          =   765
         Index           =   2
         Left            =   6240
         TabIndex        =   2
         Top             =   6360
         Width           =   5970
         _ExtentX        =   10530
         _ExtentY        =   1349
         ButtonStyle     =   13
         ShowFocusRect   =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " Not now, not ever. Do not prompt me again."
         HandPointer     =   -1  'True
         PictureNormal   =   "VBP_FormPluginDownloader.frx":20A4
         PictureAlign    =   0
         DisabledPictureMode=   1
         CaptionEffects  =   0
         TooltipTitle    =   "Never Download Plugins"
      End
      Begin VB.Label lblInterfaceTitle 
         AutoSize        =   -1  'True
         Caption         =   "ExifTool 9.29"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C07031&
         Height          =   285
         Index           =   4
         Left            =   4440
         MouseIcon       =   "VBP_FormPluginDownloader.frx":30F6
         MousePointer    =   99  'Custom
         TabIndex        =   21
         Top             =   3240
         Width           =   1425
      End
      Begin VB.Label lblDescription 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "free, open-source image metadata library (camera settings, GPS coordinates, etc)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   735
         Index           =   4
         Left            =   4680
         TabIndex        =   20
         Top             =   3600
         Width           =   3600
      End
      Begin VB.Label lblExplanation 
         BackStyle       =   0  'Transparent
         Caption         =   "Explanation appears here at run-time..."
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
         Height          =   1335
         Left            =   480
         TabIndex        =   19
         Top             =   720
         Width           =   9735
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblDescription 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "free, open-source library for optimizing portable network graphics (PNG files)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   735
         Index           =   3
         Left            =   600
         TabIndex        =   18
         Top             =   3600
         Width           =   3600
      End
      Begin VB.Label lblDescription 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "free, open-source interface for importing images from scanners and digital cameras"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   735
         Index           =   2
         Left            =   8520
         TabIndex        =   17
         Top             =   2520
         Width           =   3600
      End
      Begin VB.Label lblDescription 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "free, open-source compression library; required to decompress all other plugins"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   735
         Index           =   1
         Left            =   4440
         TabIndex        =   16
         Top             =   2520
         Width           =   3600
      End
      Begin VB.Label lblDescription 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "free, open-source library for importing and exporting a variety of image formats"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   735
         Index           =   0
         Left            =   600
         TabIndex        =   15
         Top             =   2520
         Width           =   3600
      End
      Begin VB.Label lblInterfaceTitle 
         AutoSize        =   -1  'True
         Caption         =   "pngnq-s9 2.0.1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C07031&
         Height          =   285
         Index           =   3
         Left            =   360
         MouseIcon       =   "VBP_FormPluginDownloader.frx":3248
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   3240
         Width           =   1635
      End
      Begin VB.Label lblInterfaceTitle 
         AutoSize        =   -1  'True
         Caption         =   "zLib 1.2.5"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C07031&
         Height          =   285
         Index           =   1
         Left            =   4440
         MouseIcon       =   "VBP_FormPluginDownloader.frx":339A
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   2160
         Width           =   1050
      End
      Begin VB.Label lblInterfaceTitle 
         AutoSize        =   -1  'True
         Caption         =   "EZTwain 1.18"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C07031&
         Height          =   285
         Index           =   2
         Left            =   8280
         MouseIcon       =   "VBP_FormPluginDownloader.frx":34EC
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   2160
         Width           =   1470
      End
      Begin VB.Label lblInterfaceTitle 
         AutoSize        =   -1  'True
         Caption         =   "FreeImage 3.15.4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C07031&
         Height          =   285
         Index           =   0
         Left            =   360
         MouseIcon       =   "VBP_FormPluginDownloader.frx":363E
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   2160
         Width           =   1890
      End
      Begin VB.Label lblPermission 
         AutoSize        =   -1  'True
         Caption         =   "Would you like PhotoDemon to download these plugins for you?"
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
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   4920
         Width           =   6855
      End
      Begin VB.Label lblDownloadSize 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "total download size of all plugins: "
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
         Height          =   240
         Left            =   480
         TabIndex        =   9
         Top             =   4380
         Width           =   2925
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Core Plugins Missing - Download Recommended"
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
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   180
         Width           =   5145
      End
   End
   Begin VB.PictureBox picYes 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   0
      ScaleHeight     =   145
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   833
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   7920
      Visible         =   0   'False
      Width           =   12495
      Begin VB.PictureBox picProgBar 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2640
         ScaleHeight     =   375
         ScaleWidth      =   9495
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1200
         Width           =   9495
      End
      Begin VB.Label lblDownloadInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Downloading file %1 of %2 (%3 of %4 bytes received)..."
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
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   12015
      End
      Begin VB.Label lblDownload 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "download progress:"
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
         Left            =   360
         TabIndex        =   4
         Top             =   1200
         Width           =   2115
      End
   End
End
Attribute VB_Name = "FormPluginDownloader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Automatic Plugin Downloader (for downloading core plugins that were not found at program start)
'Copyright ©2012-2013 by Tanner Helland
'Created: 10/June/12
'Last updated: 24/May/13
'Last update: interface overhaul due to inclusion of ExifTool as official plugin #5
'
'Interface for downloading plugins marked as missing upon program load.  The download code itself is a heavily
' modified version of publicly available code by Alberto Falossi (http://www.devx.com/vb2themax/Tip/19203).
'
'A number of features have been added to the original version of this code.  The routine checks plugin download
' size, and updates the user (via progress bar) on the download progress.  Many checks are in place to protect
' against Internet and download errors.  Full compression support is implemented, so if zLib is not found, it will be
' downloaded first then used to decompress the other plugins.  This cuts total download size from 8.5 to 5mb.
'
'Note that compression of the original plugin files must be performed using a custom PhotoDemon-based tool.  These are
' NOT generic .zip files (they are actually smaller than generic .zip files, owing to their simpler headers).
'
'Additionally, this form interfaces with the .INI file so the user can opt to not download the plugins and never be
' reminded again. (FYI - this option can be enabled/disabled from the 'Edit' -> 'Program Preferences' menu.)
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

'Whether or not the Internet is currently connected
Dim isInternetConnected As Boolean

'Download sizes of the five core plugins
Dim zLibSize As Double
Dim freeImageSize As Double
Dim ezTW32Size As Double
Dim pngnqSize As Double
Dim exifToolSize As Double

'Download size estimates if the user is not connected to the Internet
Private Const estZLibSize As Long = 139000
Private Const estFreeImageSize As Long = 1007000
Private Const estEzTW32Size As Long = 27000
Private Const estPngnqSize As Long = 298000
Private Const estExifToolSize As Long = 3510000

'Total expected download size, amount download thus far
Dim totalDownloadSize As Double, curDownloadSize As Double

'Number of files to download
Dim numOfFiles As Long, curNumOfFiles As Long

'We'll use a single internet session handle for this operation
Dim hInternetSession As Long

'The progress bar class we'll use to update the user on download progress
Dim dProgBar As cProgressBar

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

Private Sub cmdChoice_Click(Index As Integer)

    Select Case Index
    
        'Yes
        Case 0
        
            Dim pluginSuccess As Boolean
        
            pluginSuccess = downloadAllPlugins()
        
            'downloadAllPlugins() provides all its own error-checking
            If Not pluginSuccess Then Message "Plugins could not be downloaded at this time.  Carry on!"
        
            Unload Me
        
        'Not now
        Case 1
        
            'Store this preference
            g_UserPreferences.SetPreference_Boolean "General Preferences", "PromptForPluginDownload", True
    
            'Close our Internet connection, if any
            If hInternetSession Then InternetCloseHandle hInternetSession
            Message "Automatic plugin download canceled.  Plugin-related features disabled for this session."
            
            Unload Me
            
        'Not ever
        Case 2
            
            'Store this preference
            g_UserPreferences.SetPreference_Boolean "General Preferences", "PromptForPluginDownload", False
    
            'Close our Internet connection, if any
            If hInternetSession Then InternetCloseHandle hInternetSession
            Message "Automatic plugin download canceled.  Plugin-related features permanently disabled."
            
            Unload Me
    
    End Select

End Sub

'LOAD form
Private Sub Form_Load()

    'First things first - if the user isn't connected to the Internet, the wording of this page must be adjusted

    'So attempt to open an Internet session and assign it a handle
    Message "Checking for Internet connection..."
    hInternetSession = InternetOpen(App.EXEName, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    
    'If the user is NOT connected, adjust the text accordingly
    If hInternetSession = 0 Then
        isInternetConnected = False
        lblExplanation.Caption = g_Language.TranslateMessage("Thank you for using PhotoDemon.  Unfortunately, one or more required plugins could not be located.  PhotoDemon will still work without these plugins, but a number of features will be deactivated." & vbCrLf & vbCrLf & "To improve your user experience, please connect to the Internet, then allow the program to automatically download the following free, open-source plugin(s):")
    Else
        isInternetConnected = True
        lblExplanation.Caption = g_Language.TranslateMessage("Thank you for using PhotoDemon.  Unfortunately, one or more required plugins could not be located.  PhotoDemon will still work without these plugins, but a number of features will be deactivated." & vbCrLf & vbCrLf & "To improve your user experience, please allow the program to automatically download the following free, open-source plugin(s):")
    End If
    
    'This string will be used to hold the locations of the files to be downloaded
    Dim URL As String
    
    Message "Missing plugins detected.  Generating download information (this feature can be disabled from the Tools -> Options menu)..."
    
    totalDownloadSize = 0
    numOfFiles = 0
    
    'Manually apply some long tooltips at run-time.  This allows the translation engine to find the text and translate it.
    cmdChoice(0).ToolTip = g_Language.TranslateMessage("This is the recommended selection.  PhotoDemon will automatically download the selected plugins for you.  These plugins will be saved in the program's /Data/Plugins subdirectory.")
    cmdChoice(1).ToolTip = g_Language.TranslateMessage("This option is not recommended.  PhotoDemon will still run without plugins, but many features and tools will be temporarily disabled until these plugins are available.")
    cmdChoice(2).ToolTip = g_Language.TranslateMessage("This option is not recommended.  PhotoDemon will still run without plugins, but many features and tools will be permanently disabled.")
    
    'Upon program load, populate the list of files to be downloaded based on which could not be found.
    
    'zLib
    If isInternetConnected Then
        URL = "http://www.tannerhelland.com/photodemon_files/zlibwapi.pdc"
        zLibSize = getPluginSize(hInternetSession, URL)
        
        'If getPluginSize fails, it will return -1.  Set an estimated size and allow the software to continue
        If zLibSize = -1 Then zLibSize = estZLibSize
            
    Else
        zLibSize = estZLibSize
    End If
    
    'EZTwain
    If isInternetConnected Then
        URL = "http://www.tannerhelland.com/photodemon_files/eztw32.pdc"
        ezTW32Size = getPluginSize(hInternetSession, URL)
        
        'If getPluginSize fails, it will return -1.  Set an estimated size and allow the software to continue
        If ezTW32Size = -1 Then ezTW32Size = estEzTW32Size
            
    Else
        ezTW32Size = estEzTW32Size
    End If
    
    'FreeImage
    If isInternetConnected Then
        URL = "http://www.tannerhelland.com/photodemon_files/freeimage.pdc"
        freeImageSize = getPluginSize(hInternetSession, URL)
        
        'If getPluginSize fails, it will return -1.  Set an estimated size and allow the software to continue
        If freeImageSize = -1 Then freeImageSize = estFreeImageSize
        
    Else
        freeImageSize = estFreeImageSize
    End If
    
    'pngnq-s9
    If isInternetConnected Then
        URL = "http://www.tannerhelland.com/photodemon_files/pngnq-s9.pdc"
        pngnqSize = getPluginSize(hInternetSession, URL)
        
        'If getPluginSize fails, it will return -1.  Set an estimated size and allow the software to continue
        If pngnqSize = -1 Then pngnqSize = estPngnqSize
        
    Else
        pngnqSize = estPngnqSize
    End If
    
    'ExifTool
    If isInternetConnected Then
        URL = "http://www.tannerhelland.com/photodemon_files/exiftool.pdc"
        exifToolSize = getPluginSize(hInternetSession, URL)
        
        'If getPluginSize fails, it will return -1.  Set an estimated size and allow the software to continue
        If exifToolSize = -1 Then exifToolSize = estExifToolSize
        
    Else
        exifToolSize = estExifToolSize
    End If
    
    updateDownloadSize
    
    'Clear the download progress placeholder text
    lblDownloadInfo.Caption = ""
    
    Message "Ready to download required plugins. Awaiting user permission..."
    
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me, m_ToolTip

End Sub

'Simple routine to check the file size of a provided file URL
Private Function getPluginSize(ByVal hInternet As Long, ByVal pluginURL As String) As Long
    
    'Check the size of the file to be downloaded...
    Dim tmpStrBuffer As String
    tmpStrBuffer = String$(1024, 0)
    Dim hUrl As Long
    hUrl = InternetOpenUrl(hInternet, pluginURL, vbNullString, 0, INTERNET_FLAG_RELOAD, 0)
    Call HttpQueryInfo(ByVal hUrl, HTTP_QUERY_CONTENT_LENGTH, ByVal tmpStrBuffer, Len(tmpStrBuffer), 0)
    If hUrl <> 0 Then
        getPluginSize = CLng(tmpStrBuffer)
    Else
        getPluginSize = -1
    End If
    InternetCloseHandle hUrl
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Allow the user to visit any plugin homepage by clicking the plugin's name
Private Sub lblInterfaceTitle_Click(Index As Integer)
    
    Select Case Index
    
        'FreeImage
        Case 0
            OpenURL "http://freeimage.sourceforge.net/download.html"
            
        'zLib
        Case 1
            OpenURL "http://www.winimage.com/zLibDll/index.html"
        
        'EZTwain
        Case 2
            OpenURL "http://eztwain.com/eztwain1.htm"
            
        'pngnq-s9
        Case 3
            OpenURL "http://sourceforge.net/projects/pngnqs9/"
            
        'ExifTool
        Case 4
            OpenURL "http://www.sno.phy.queensu.ca/~phil/exiftool/"
            
    End Select
    
End Sub

Private Function downloadAllPlugins() As Boolean

    If isInternetConnected = False Then
    
        'Hopefully the user established an internet connection before clicking this button.  If not, prompt them to do so.
        Message "Checking again for Internet connection..."
        hInternetSession = InternetOpen(App.EXEName, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
        
        If hInternetSession = 0 Then
            Message "No Internet connection found."
            pdMsgBox "Unfortunately, %1 could not connect to the Internet.  Please connect to the Internet and try again.", vbApplicationModal + vbOKOnly + vbExclamation, "No Internet Connection", PROGRAMNAME
            downloadAllPlugins = False
            Exit Function
        End If
        
    End If
    
    'If we've made it here, assume the user has a valid Internet connection
    
    'Bring the picture box with the download info to the foreground
    picInitial.Visible = False
    picYes.Left = 0
    picYes.Top = 0
        
    'Set up a progress bar control
    Set dProgBar = New cProgressBar
    dProgBar.DrawObject = FormPluginDownloader.picProgBar
    dProgBar.BarColor = RGB(48, 117, 255)
    dProgBar.Min = 0
    dProgBar.Max = 100
    dProgBar.xpStyle = True
    dProgBar.ShowText = False
    dProgBar.Draw
    
    dProgBar.Max = totalDownloadSize
    dProgBar.Value = 0
    FormPluginDownloader.Height = 2475
    picYes.Visible = True
    
    'This DoEvents is necessary to refresh the picture boxes we have hidden and displayed
    DoEvents
    
    'Begin by creating a plugin subdirectory if it doesn't exist
    Message "Checking for plugin directory..."
    If DirectoryExist(g_PluginPath) = False Then
        Message "Creating plugin directory..."
        MkDir g_PluginPath
    End If
    
    Message "Downloading core plugin files..."
    
    Dim downloadSuccessful As Boolean
    curDownloadSize = 0
    curNumOfFiles = 1
    
    'Time to get the files.  Start with zLib.
    downloadSuccessful = downloadPlugin("http://www.tannerhelland.com/photodemon_files/zlibwapi.pdc", curNumOfFiles, numOfFiles, zLibSize, False)
    If downloadSuccessful = False Then
        pdMsgBox "Due to an unforeseen error, %1 is postponing plugin downloading for the moment.  Next time you run this application, it will try the download again.  (Apologies for the inconvenience.)", vbOKOnly + vbInformation + vbApplicationModal, "Unspecified Download Error", PROGRAMNAME
        downloadAllPlugins = False
        Exit Function
    Else
        g_ZLibEnabled = True
    End If
        
    curNumOfFiles = curNumOfFiles + 1
    
    'Next comes EZTW32
    downloadSuccessful = downloadPlugin("http://www.tannerhelland.com/photodemon_files/eztw32.pdc", curNumOfFiles, numOfFiles, ezTW32Size, True)
    If downloadSuccessful = False Then
        pdMsgBox "Due to an unforeseen error, %1 is postponing plugin downloading for the moment.  Next time you run this application, it will try the download again.  (Apologies for the inconvenience.)", vbOKOnly + vbInformation + vbApplicationModal, "Unspecified Download Error", PROGRAMNAME
        downloadAllPlugins = False
        Exit Function
    Else
        g_ScanEnabled = True
    End If
    
    curNumOfFiles = curNumOfFiles + 1
            
    'Next is FreeImage
    downloadSuccessful = downloadPlugin("http://www.tannerhelland.com/photodemon_files/freeimage.pdc", curNumOfFiles, numOfFiles, freeImageSize, True)
    If downloadSuccessful = False Then
        pdMsgBox "Due to an unforeseen error, %1 is postponing plugin downloading for the moment.  Next time you run this application, it will try the download again.  (Apologies for the inconvenience.)", vbOKOnly + vbInformation + vbApplicationModal, "Unspecified Download Error", PROGRAMNAME
        downloadAllPlugins = False
        Exit Function
    Else
        g_ImageFormats.FreeImageEnabled = True
    End If
    
    curNumOfFiles = curNumOfFiles + 1
    
    'Next is pngnq-s9
    downloadSuccessful = downloadPlugin("http://www.tannerhelland.com/photodemon_files/pngnq-s9.pdc", curNumOfFiles, numOfFiles, pngnqSize, True)
    If downloadSuccessful = False Then
        pdMsgBox "Due to an unforeseen error, %1 is postponing plugin downloading for the moment.  Next time you run this application, it will try the download again.  (Apologies for the inconvenience.)", vbOKOnly + vbInformation + vbApplicationModal, "Unspecified Download Error", PROGRAMNAME
        downloadAllPlugins = False
        Exit Function
    Else
        g_ImageFormats.pngnqEnabled = True
    End If
    
    curNumOfFiles = curNumOfFiles + 1
    
    'Last is ExifTool
    downloadSuccessful = downloadPlugin("http://www.tannerhelland.com/photodemon_files/exiftool.pdc", curNumOfFiles, numOfFiles, exifToolSize, True)
    If downloadSuccessful = False Then
        pdMsgBox "Due to an unforeseen error, %1 is postponing plugin downloading for the moment.  Next time you run this application, it will try the download again.  (Apologies for the inconvenience.)", vbOKOnly + vbInformation + vbApplicationModal, "Unspecified Download Error", PROGRAMNAME
        downloadAllPlugins = False
        Exit Function
    Else
        g_ExifToolEnabled = True
    End If
    
    dProgBar.Value = dProgBar.Max
    
    If hInternetSession Then InternetCloseHandle hInternetSession
    
    lblDownloadInfo.Caption = g_Language.TranslateMessage("All downloads successful.  This screen will automatically close in three seconds.")
    
    Dim OT As Double
    OT = Timer
    Do While Timer - OT < 3#
        DoEvents
    Loop
    
    Message "Plugins downloaded successfully.  To complete plugin setup, please restart the program."
    
    Unload Me

End Function

Private Function downloadPlugin(ByVal pluginURL As String, ByVal curNumFile As Long, ByVal maxNumFile As Long, ByVal downloadSize As Long, ByVal toDecompress As Boolean)

    'First, attempt to find the plugin URL; if found, assign it a handle
    lblDownloadInfo.Caption = "Verifying plugin URL..."
    
    Dim hUrl As Long
    hUrl = InternetOpenUrl(hInternetSession, pluginURL, vbNullString, 0, INTERNET_FLAG_RELOAD, 0)

    If hUrl = 0 Then
        pdMsgBox "%1 could not locate the plugin server.  Please double-check your Internet connection.  If the problem persists, please try again at another time.", vbExclamation + vbApplicationModal + vbOKOnly, "Plugin Server Not Responding", PROGRAMNAME
        If hInternetSession Then InternetCloseHandle hInternetSession
        downloadPlugin = False
        Message "Plugin download postponed."
        Exit Function
    End If
    
    'We need a temporary file to house the image; generate it automatically, using the extension of the original image
    lblDownloadInfo.Caption = "Creating temporary file..."
    Dim tmpFilename As String
    tmpFilename = pluginURL
    StripFilename tmpFilename
    
    Dim tmpFile As String
    If toDecompress = False Then
        StripOffExtension tmpFilename
        tmpFile = g_PluginPath & tmpFilename & ".dll"
    Else
        tmpFile = g_PluginPath & tmpFilename
    End If
    
    'Open the temporary file and begin downloading the image to it
    lblDownloadInfo.Caption = g_Language.TranslateMessage("Downloading file %1 of %2...", curNumFile, maxNumFile)
    Dim fileNum As Integer
    fileNum = FreeFile
    
    If FileExist(tmpFile) Then Kill tmpFile
    
    Open tmpFile For Binary As fileNum
    
        'Prepare a receiving buffer (this will be used to hold chunks of the image)
        Dim Buffer As String
        Buffer = Space(4096)
   
        'We will need to verify each chunk as its downloaded
        Dim chunkOK As Boolean
   
        'This will track the size of each chunk
        Dim numOfBytesRead As Long
   
        'This will track of how many bytes we've downloaded so far
        Dim totalBytesRead As Long
        totalBytesRead = 0
   
        Do
   
            'Read the next chunk of the image
            chunkOK = InternetReadFile(hUrl, Buffer, Len(Buffer), numOfBytesRead)
   
            'If something went wrong, terminate
            If chunkOK = False Then
                pdMsgBox "%1 lost access to the Internet. Please double-check your Internet connection.  If the problem persists, please try the download again at a later time.", vbExclamation + vbApplicationModal + vbOKOnly, "Internet Connection Error", PROGRAMNAME
                If FileExist(tmpFile) Then
                    Close #fileNum
                    Kill tmpFile
                End If
                If hUrl Then InternetCloseHandle hUrl
                If hInternetSession Then InternetCloseHandle hInternetSession
                downloadPlugin = False
                Exit Function
            End If
   
            'If the file is done, exit this loop
            If numOfBytesRead = 0 Then
                Exit Do
            End If
   
            'If we've made it this far, assume we've received legitimate data.  Place that data into the file.
            Put #fileNum, , Left$(Buffer, numOfBytesRead)
   
            totalBytesRead = totalBytesRead + numOfBytesRead
            
            curDownloadSize = curDownloadSize + numOfBytesRead
            
            If downloadSize <> 0 Then
                If curDownloadSize < dProgBar.Max Then dProgBar.Value = curDownloadSize
                lblDownloadInfo.Caption = g_Language.TranslateMessage("Downloading file %1 of %2 (%3 of %4 bytes received)...", curNumFile, maxNumFile, totalBytesRead, downloadSize)
                lblDownloadInfo.Refresh
            End If
            
        'Carry on
        Loop
        
    'Close the temporary file
    Close #fileNum
    
    'Close this URL and Internet session
    If hUrl Then InternetCloseHandle hUrl
    
    lblDownloadInfo.Caption = "Download complete. Verifying file integrity..."
    
    'If requested, decompress the file
    If toDecompress = False Then
        downloadPlugin = True
    Else
        Dim verifyDecompression As Boolean
        verifyDecompression = DecompressFile(tmpFile, False)
        
        If verifyDecompression = False Then
            downloadPlugin = False
        Else
            downloadPlugin = True
        End If
        
    End If

End Function

'Add up the sizes of the selected plugins to give the user a "total download size" estimate.
Private Sub updateDownloadSize()

    totalDownloadSize = 0
    numOfFiles = 0
    
    totalDownloadSize = totalDownloadSize + zLibSize
    numOfFiles = numOfFiles + 1
    
    totalDownloadSize = totalDownloadSize + freeImageSize
    numOfFiles = numOfFiles + 1
        
    totalDownloadSize = totalDownloadSize + ezTW32Size
    numOfFiles = numOfFiles + 1
    
    totalDownloadSize = totalDownloadSize + pngnqSize
    numOfFiles = numOfFiles + 1
        
    totalDownloadSize = totalDownloadSize + exifToolSize
    numOfFiles = numOfFiles + 1
    
    lblDownloadSize.Caption = g_Language.TranslateMessage("total download size:") & " " & Format(CStr(totalDownloadSize / 1000000), "0.00") & " MB"

End Sub
