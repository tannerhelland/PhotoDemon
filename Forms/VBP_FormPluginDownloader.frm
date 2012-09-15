VERSION 5.00
Begin VB.Form FormPluginDownloader 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " PhotoDemon Plugin Downloader"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8295
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
   ScaleHeight     =   401
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   553
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOKNo 
      Caption         =   "OK"
      Height          =   495
      Left            =   6720
      TabIndex        =   18
      Top             =   5280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox picYes 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   8280
      ScaleHeight     =   145
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   553
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   8295
      Begin VB.PictureBox picProgBar 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2160
         ScaleHeight     =   255
         ScaleWidth      =   6015
         TabIndex        =   17
         Top             =   1080
         Width           =   6015
      End
      Begin VB.Label lblDownloadInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Downloading file 1 of 3 (XXX of YYY bytes received)..."
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
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Width           =   7815
      End
      Begin VB.Label lblDownload 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Download progress:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   360
         TabIndex        =   15
         Top             =   1080
         Width           =   1725
      End
   End
   Begin VB.PictureBox picNo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5295
      Left            =   0
      MousePointer    =   1  'Arrow
      ScaleHeight     =   353
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   545
      TabIndex        =   6
      Top             =   6480
      Visible         =   0   'False
      Width           =   8175
      Begin VB.TextBox txtNoExplanation 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00400000&
         Height          =   1335
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   7
         Text            =   "VBP_FormPluginDownloader.frx":0000
         Top             =   1440
         Width           =   7935
      End
      Begin VB.Label lblPluginTitle3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FreeImage:  "
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1920
         TabIndex        =   13
         Top             =   3840
         Width           =   930
      End
      Begin VB.Label lblPluginTitle2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EZTW32 (""EZTwain Classic""): "
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1920
         TabIndex        =   12
         Top             =   3360
         Width           =   2115
      End
      Begin VB.Label lblPluginTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "zLib (WAPI variant): "
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1800
         TabIndex        =   11
         Top             =   2880
         Width           =   1500
      End
      Begin VB.Label lblFreeImage 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://freeimage.sourceforge.net/download.html"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   2880
         MouseIcon       =   "VBP_FormPluginDownloader.frx":01BC
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   3840
         Width           =   3540
      End
      Begin VB.Label lblEZTW32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://eztwain.com/eztwain1.htm"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   4080
         MouseIcon       =   "VBP_FormPluginDownloader.frx":030E
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   3360
         Width           =   2415
      End
      Begin VB.Label lblzLib 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.winimage.com/zLibDll/index.html"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   3360
         MouseIcon       =   "VBP_FormPluginDownloader.frx":0460
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   2880
         Width           =   3210
      End
   End
   Begin VB.TextBox txtExplanation 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H00400000&
      Height          =   1095
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "VBP_FormPluginDownloader.frx":05B2
      Top             =   240
      Width           =   7815
   End
   Begin VB.TextBox txtPlugins 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H00800000&
      Height          =   1815
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "VBP_FormPluginDownloader.frx":071A
      Top             =   1440
      Width           =   7815
   End
   Begin VB.CommandButton cmdYesDownload 
      Caption         =   "Yes.  Please download these files to the plugins directory."
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   4080
      Width           =   7815
   End
   Begin VB.CommandButton cmdNoDownload 
      Caption         =   "Not right now, but please remind me again in the future."
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   4680
      Width           =   7815
   End
   Begin VB.CommandButton cmdNoDownloadNoReminder 
      Caption         =   "Not now, not ever.  Do not download these files, and do not prompt me again."
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   5280
      Width           =   7815
   End
   Begin VB.Label lblPermission 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Would you like PhotoDemon to download these plugins for you?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   3480
      Width           =   7815
   End
End
Attribute VB_Name = "FormPluginDownloader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Automatic Plugin Downloader (for downloading core plugins that were not found at program start)
'Copyright ©2000-2012 by Tanner Helland
'Created: 10/June/12
'Last updated: 13/June/12
'Last update: added compression support.  The total download size for all three plugins is now under 1.0M.  Sweet!
'
'Interface for downloading plugins marked as missing upon program load.  This code is a heavily modified version
' of publicly available code by Alberto Falossi (http://www.devx.com/vb2themax/Tip/19203).
'
'A number of features have been added to the original version of this code.  The routine checks the file download
' size, and updates the user (via progress bar) on the download progress.  Many checks are in place to protect
' against Internet and download errors.  Full compression support is implemented, so if zLib is not found, it will be
' downloaded first then used to decompress the other plugins.  This cut total download size from 2.8M to just under 1.0M.
'
'Additionally, this form interfaces with the .INI file so the user can opt to not download the plugins and never be
' reminded again. (FYI - this option can be enabled/disabled from the 'Edit' -> 'Program Preferences' menu.)
'
'***************************************************************************

Option Explicit

'Whether or not the Internet is currently connected
Dim isInternetConnected As Boolean

'Download sizes of the three core plugins
Dim zLibSize As Single
Dim freeImageSize As Single
Dim ezTW32Size As Single

'Total expected download size, amount download thus far
Dim totalDownloadSize As Single, curDownloadSize As Single

'Number of files to download
Dim numOfFiles As Long, curNumOfFiles As Long

'We'll use a single internet session handle for this operation
Dim hInternetSession As Long

'The progress bar class we'll use to update the user on download progress
Dim dProgBar As cProgressBar

'Do not download the plugins, but prompt the user again in the future
Private Sub cmdNoDownload_Click()
    WriteToIni "General Preferences", "PromptForPluginDownload", 1
    If hInternetSession Then InternetCloseHandle hInternetSession
    Message "Automatic update canceled."
    
    cmdYesDownload.Visible = False
    cmdNoDownload.Visible = False
    cmdNoDownloadNoReminder.Visible = False
    picNo.Left = 0
    picNo.Top = 0
    DoEvents
    txtNoExplanation.Text = "The next time you launch " & PROGRAMNAME & ", it will repeat this check for missing plugins." & vbCrLf & vbCrLf & "Note: if you're the adventurous type, you can manually download these plugin files from their respective sites.  " & PROGRAMNAME & " will look for the DLL versions of these libraries in the 'plugins' subdirectory of wherever the " & PROGRAMNAME & " executable file is located."
    picNo.Visible = True
    cmdOKNo.Visible = True
    cmdOKNo.SetFocus
    
End Sub

'Do not download the plugins, and do not prompt the user again
Private Sub cmdNoDownloadNoReminder_Click()
    WriteToIni "General Preferences", "PromptForPluginDownload", 0
    If hInternetSession Then InternetCloseHandle hInternetSession
    Message "Automatic update canceled."
    
    cmdYesDownload.Visible = False
    cmdNoDownload.Visible = False
    cmdNoDownloadNoReminder.Visible = False
    picNo.Left = 0
    picNo.Top = 0
    txtNoExplanation.Text = PROGRAMNAME & " will no longer prompt you about missing plugins.  (If you change your mind in the future, this setting can be reversed from the 'Edit' -> 'Program Preferences' menu.)" & vbCrLf & vbCrLf & "Note: if you're the adventurous type, you can manually download these plugin files from their respective sites.  " & PROGRAMNAME & " will look for the DLL versions of these libraries in the 'plugins' subdirectory of wherever the " & PROGRAMNAME & " executable file is located."
    DoEvents
    picNo.Visible = True
    cmdOKNo.Visible = True
    cmdOKNo.SetFocus
    
End Sub

'This OK button only appears on the picture box that contains additional information when either of the two "No" buttons are selected
Private Sub cmdOKNo_Click()
    Unload Me
End Sub

'Yes, the user wants us to download the plugins.  Go for it!
Private Sub cmdYesDownload_Click()

    Dim pluginSuccess As Boolean
    
    pluginSuccess = downloadAllPlugins()
    
    'downloadAllPlugins() provides all its own error-checking
    If pluginSuccess = False Then Message "Plugins could not be downloaded at this time.  Carry on!"
    
    Unload Me
    
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
        txtExplanation.Text = "Thank you for using PhotoDemon." & vbCrLf & vbCrLf & "Unfortunately, one or more core plugins could not be located.  " & PROGRAMNAME & " will work without these files but certain features will be disabled.  To improve your user experience, please connect to the Internet and restart this program. Then, when prompted, please allow it to download the following free, open-source plugin(s):"
    Else
        isInternetConnected = True
        txtExplanation.Text = "Thank you for using PhotoDemon." & vbCrLf & vbCrLf & "Unfortunately, one or more core plugins could not be located.  " & PROGRAMNAME & " will work without these files but certain features will be disabled.  To improve your user experience, please allow the program to automatically download the following free, open-source plugin(s):"
    End If
    
    'This string will be used to hold the locations of the files to be downloaded
    Dim URL As String
    
    Message "Missing plugins detected.  Generating automatic update information (this feature can be disabled from the Edit -> Preferences menu)..."
    
    txtPlugins.Text = ""
    totalDownloadSize = 0
    numOfFiles = 0
    
    'Upon program load, populate the list of files to be downloaded based on which turned up missing
    If zLibEnabled = False Then
        If isInternetConnected = True Then
            URL = "http://www.tannerhelland.com/photodemon_files/zlibwapi.pdc"
            zLibSize = getPluginSize(hInternetSession, URL)
            
            'If getPluginSize fails, it will return -1.  Set an estimated size and allow the software to continue
            If zLibSize = -1 Then zLibSize = 139000
            totalDownloadSize = zLibSize
            
            txtPlugins.Text = ">> zLib: a compression library used to save PhotoDemon Image (PDI) files, and decompress the other plugins after they've been downloaded.  Size: " & Int(zLibSize \ 1000) & " kB" & vbCrLf
        Else
            totalDownloadSize = 139000
            txtPlugins.Text = ">> zLib: a compression library used to save PhotoDemon Image (PDI) files, and decompress the other plugins after they've been downloaded.  Size: ~138 kB" & vbCrLf
        End If
        
        numOfFiles = numOfFiles + 1
        
    End If
    
    If ScanEnabled = False Then
        If isInternetConnected = True Then
            URL = "http://www.tannerhelland.com/photodemon_files/eztw32.pdc"
            ezTW32Size = getPluginSize(hInternetSession, URL)
            
            'If getPluginSize fails, it will return -1.  Set an estimated size and allow the software to continue
            If ezTW32Size = -1 Then ezTW32Size = 28000
            totalDownloadSize = totalDownloadSize + ezTW32Size
            
            txtPlugins.Text = txtPlugins.Text & vbCrLf & ">> EZTW32: enables scanner and digital camera access via the TWAIN32 protocol.  Size: " & Int(ezTW32Size \ 1000) & " kB" & vbCrLf
        Else
            totalDownloadSize = totalDownloadSize + 28000
            txtPlugins.Text = txtPlugins.Text & vbCrLf & ">> EZTW32: enables scanner and digital camera access via the TWAIN32 protocol.  Size: ~27 kB" & vbCrLf
        End If
        
        numOfFiles = numOfFiles + 1
        
    End If
    
    If FreeImageEnabled = False Then
        If isInternetConnected = True Then
            URL = "http://www.tannerhelland.com/photodemon_files/freeimage.pdc"
            freeImageSize = getPluginSize(hInternetSession, URL)
            
            'If getPluginSize fails, it will return -1.  Set an estimated size and allow the software to continue
            If freeImageSize = -1 Then freeImageSize = 977000
            totalDownloadSize = totalDownloadSize + freeImageSize
            
            txtPlugins.Text = txtPlugins.Text & vbCrLf & ">> FreeImage: advanced file format support, including PSD, PICT, TGA, HDR, and many more.  Also used for advanced image resize filters (Mitchell and Netravali, Catmull-Rom, Lanczos).  Size: " & Int(freeImageSize \ 1000) & " kB"
        Else
            totalDownloadSize = totalDownloadSize + 977000
            txtPlugins.Text = txtPlugins.Text & vbCrLf & ">> FreeImage: advanced file format support, including PSD, PICT, TGA, HDR, and many more.  Also used for advanced image resize filters (Mitchell and Netravali, Catmull-Rom, Lanczos).  Size: ~976 kB"
        End If
        
        numOfFiles = numOfFiles + 1
        
    End If

    txtPlugins.Text = txtPlugins.Text & vbCrLf & vbCrLf & "Total download size: " & Int(totalDownloadSize \ 1000) & " kB"

    Message "Ready to update. Awaiting user permission..."
    
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me

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

'Launch the website for downloading the EZTW32 DLL
Private Sub lblEZTW32_Click()
    ShellExecute FormMain.HWnd, "Open", "http://eztwain.com/eztwain1.htm", "", 0, SW_SHOWNORMAL
End Sub

'Launch the website for downloading the FreeImage DLL
Private Sub lblFreeImage_Click()
    ShellExecute FormMain.HWnd, "Open", "http://freeimage.sourceforge.net/download.html", "", 0, SW_SHOWNORMAL
End Sub

'Launch the website for downloading the zLibwapi DLL
Private Sub lblzLib_Click()
    ShellExecute FormMain.HWnd, "Open", "http://www.winimage.com/zLibDll/index.html", "", 0, SW_SHOWNORMAL
End Sub

Private Function downloadAllPlugins() As Boolean

    If isInternetConnected = False Then
        'Hopefully the user established an internet connection before clicking this button.  If not, prompt them to do so.
        Message "Checking again for Internet connection..."
        hInternetSession = InternetOpen(App.EXEName, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
        
        If hInternetSession = 0 Then
            Message "No Internet connection found."
            MsgBox "Unfortunately, " & PROGRAMNAME & " could not connect to the Internet.  Please connect to the Internet and try again.", vbApplicationModal + vbOKOnly + vbCritical, "No Internet Connection"
            downloadAllPlugins = False
            Exit Function
        End If
        
    End If
    
    'If we've made it here, assume the user has a valid Internet connection
    
    'Bring the picture box with the download info to the foreground
    picYes.Left = 0
    picYes.Top = 0
        
    'Set up a progress bar control
    Set dProgBar = New cProgressBar
    dProgBar.DrawObject = FormPluginDownloader.picProgBar
    dProgBar.Min = 0
    dProgBar.Max = 100
    dProgBar.XpStyle = True
    dProgBar.ShowText = False
    dProgBar.Draw
    
    dProgBar.Max = totalDownloadSize
    dProgBar.Value = 0
    FormPluginDownloader.Height = 2475
    picYes.Visible = True
    DoEvents
    
    'Begin by creating a plugin subdirectory if it doesn't exist
    Message "Checking for plugin directory..."
    If DirectoryExist(PluginPath) = False Then
        Message "Creating plugin directory..."
        MkDir PluginPath
    End If
    
    Dim downloadSuccessful As Boolean
    curDownloadSize = 0
    curNumOfFiles = 1
    
    'Time to get the files.  Start with zLib.
    If zLibEnabled = False Then
        downloadSuccessful = downloadPlugin("http://www.tannerhelland.com/photodemon_files/zlibwapi.pdc", curNumOfFiles, numOfFiles, zLibSize, False)
        If downloadSuccessful = False Then
            MsgBox "Due to an unforeseen error, " & PROGRAMNAME & " is postponing plugin downloading for the moment.  Next time you run this application, it will try the download again.  (Apologies for the inconvenience.)", vbOKOnly + vbInformation + vbApplicationModal, "Unspecified Download Error"
            downloadAllPlugins = False
            Exit Function
        Else
            zLibEnabled = True
        End If
        
        curNumOfFiles = curNumOfFiles + 1
    
    End If
    
    'Next comes EZTW32
    If ScanEnabled = False Then
        downloadSuccessful = downloadPlugin("http://www.tannerhelland.com/photodemon_files/eztw32.pdc", curNumOfFiles, numOfFiles, ezTW32Size, True)
        If downloadSuccessful = False Then
            MsgBox "Due to an unforeseen error, " & PROGRAMNAME & " is postponing plugin downloading for the moment.  Next time you run this application, it will try the download again.  (Apologies for the inconvenience.)", vbOKOnly + vbInformation + vbApplicationModal, "Unspecified Download Error"
            downloadAllPlugins = False
            Exit Function
        Else
            ScanEnabled = True
        End If
        
        curNumOfFiles = curNumOfFiles + 1
    
    End If
            
    'Last is FreeImage
    If FreeImageEnabled = False Then
        downloadSuccessful = downloadPlugin("http://www.tannerhelland.com/photodemon_files/freeimage.pdc", curNumOfFiles, numOfFiles, freeImageSize, True)
        If downloadSuccessful = False Then
            MsgBox "Due to an unforeseen error, " & PROGRAMNAME & " is postponing plugin downloading for the moment.  Next time you run this application, it will try the download again.  (Apologies for the inconvenience.)", vbOKOnly + vbInformation + vbApplicationModal, "Unspecified Download Error"
            downloadAllPlugins = False
            Exit Function
        Else
            FreeImageEnabled = True
        End If
    
    End If
    
    dProgBar.Value = dProgBar.Max
    DoEvents
    
    If hInternetSession Then InternetCloseHandle hInternetSession
    
    lblDownloadInfo.Caption = "All downloads successful.  This screen will automatically close in four seconds."
    
    Dim OT As Single
    OT = Timer
    Do While Timer - OT < 4#
        DoEvents
    Loop
    
    Message "Plugins downloaded successfully.  PhotoDemon is ready to go!  Please load an image (File -> Open)"
    
    Unload Me


End Function

Private Function downloadPlugin(ByVal pluginURL As String, ByVal curNumFile As Long, ByVal maxNumFile As Long, ByVal downloadSize As Long, ByVal toDecompress As Boolean)

    'First, attempt to find the plugin URL; if found, assign it a handle
    lblDownloadInfo.Caption = "Verifying plugin URL..."
    
    Dim hUrl As Long
    hUrl = InternetOpenUrl(hInternetSession, pluginURL, vbNullString, 0, INTERNET_FLAG_RELOAD, 0)

    If hUrl = 0 Then
        MsgBox PROGRAMNAME & " could not locate the plugin server.  Please double-check your Internet connection.  If the problem persists, please try again at another time.", vbCritical + vbApplicationModal + vbOKOnly, "Plugin Server Not Responding"
        If hInternetSession Then InternetCloseHandle hInternetSession
        downloadPlugin = False
        Message "Plugin download postponed."
        Exit Function
    End If
    
    'We need a temporary file to house the image; generate it automatically, using the extension of the original image
    lblDownloadInfo.Caption = "Creating temporary file..."
    Dim tmpFileName As String
    tmpFileName = pluginURL
    StripFilename tmpFileName
    
    Dim tmpFile As String
    If toDecompress = False Then
        StripOffExtension tmpFileName
        tmpFile = PluginPath & tmpFileName & ".dll"
    Else
        tmpFile = PluginPath & tmpFileName
    End If
    
    'Open the temporary file and begin downloading the image to it
    lblDownloadInfo.Caption = "Downloading file " & curNumFile & " of " & maxNumFile & "..."
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
                MsgBox PROGRAMNAME & " lost access to the Internet. Please double-check your Internet connection.  If the problem persists, please try the download again at a later time.", vbCritical + vbApplicationModal + vbOKOnly, "Internet Connection Error"
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
                lblDownloadInfo.Caption = "Downloading file " & curNumFile & " of " & maxNumFile & " (" & totalBytesRead & " of " & downloadSize & " bytes received)..."
            End If
            
            DoEvents
            
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
