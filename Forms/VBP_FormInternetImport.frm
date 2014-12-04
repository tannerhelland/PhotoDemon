VERSION 5.00
Begin VB.Form FormInternetImport 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Download Image"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10050
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
   ScaleHeight     =   156
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   670
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdTextBox txtURL 
      Height          =   315
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   556
      Text            =   "http://"
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   7080
      TabIndex        =   0
      Top             =   1710
      Width           =   1365
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   8550
      TabIndex        =   1
      Top             =   1710
      Width           =   1365
   End
   Begin VB.Label lblCopyrightWarning 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright"
      ForeColor       =   &H00808080&
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   1725
      Width           =   6735
   End
   Begin VB.Label lblBackground 
      Height          =   855
      Left            =   0
      TabIndex        =   4
      Top             =   1560
      Width           =   10095
   End
   Begin VB.Label lblDownloadPath 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "full download path (must begin with ""http://"" or ""ftp://""):"
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
      TabIndex        =   2
      Top             =   360
      Width           =   6090
   End
End
Attribute VB_Name = "FormInternetImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Internet Interface (for importing images directly from a URL)
'Copyright ©2011-2014 by Tanner Helland
'Created: 08/June/12
'Last updated: 03/December/12
'Last update: made some slight modifications to ImportImageFromInternet so it can be used by external callers.
'
'Interface for downloading images directly from the Internet into PhotoDemon.  This code is a heavily
' modified version of publicly available code by Alberto Falossi (http://www.devx.com/vb2themax/Tip/19203).
'
'A number of features have been added to the original version of this code.  The routine checks the file download
' size, and updates the user (via progress bar) on the download progress.  Many checks are in place to protect
' against Internet and download errors.  I'm quite proud of how robust this implementation is, but additional
' testing will be necessary to make sure no possible connectivity errors have been overlooked.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'Import an image from the Internet; all that's required is a valid URL (must be prefaced with http:// or ftp://)
Public Function ImportImageFromInternet(ByVal URL As String) As Boolean

    'First things first - if an invalid URL was provided, exit immediately.
    If Len(URL) = 0 Then
        Message "Image download canceled."
        Exit Function
    End If
    
    'Use the generic download function to retrieve the URL
    Dim downloadedFilename As String
    downloadedFilename = downloadURLToTempFile(URL)
    
    'If the download worked, attempt to load the image.
    If Len(downloadedFilename) <> 0 Then
    
        Dim sFile(0) As String
        sFile(0) = downloadedFilename
        
        Dim tmpFilename As String
        tmpFilename = downloadedFilename
        StripFilename tmpFilename
        
        LoadFileAsNewImage sFile, False, tmpFilename, tmpFilename
        
        'Unique to this particular import is remembering the full filename + extension (because this method of import
        ' actually supplies a file extension, unlike scanning or screen capturing or something else)
        If Not pdImages(g_CurrentImage) Is Nothing Then pdImages(g_CurrentImage).originalFileNameAndExtension = tmpFilename
        
        'Delete the temporary file
        If FileExist(downloadedFilename) Then Kill downloadedFilename
        
        Message "Image download complete. "
        ImportImageFromInternet = True
        
    Else
        ImportImageFromInternet = False
    End If
    
End Function

'Download the contents of a given URL to a temporary file.  Progress reports will be automatically provided via the
' program progress bar.
'
'If successful, the program will return the full path to the temp file used.  If unsuccessful, a blank string will
' be returned.  Use Len(returnString) = 0 to check for failure state.
'
'Note that the calling function is responsible for cleaning up the temp file!
Public Function downloadURLToTempFile(ByVal URL As String) As String

    'Normally changing the cursor is handled by the software processor, but because this function routes
    ' internally, we'll make an exception and change it here. Note that everywhere this function can
    ' terminate (and it's many places - a lot can go wrong while downloading) - the cursor needs to be reset.
    Screen.MousePointer = vbHourglass
    
    'Open an Internet session and assign it a handle
    Dim hInternetSession As Long
    
    Message "Attempting to connect to the Internet..."
    hInternetSession = InternetOpen(App.EXEName, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    
    If hInternetSession = 0 Then
        pdMsgBox "%1 could not establish an Internet connection. Please double-check your connection.  If the problem persists, try downloading the image manually using your Internet browser of choice.  Once downloaded, you may open the file in %1 just like any other image file.", vbExclamation + vbApplicationModal + vbOKOnly, "Internet Connection Error", PROGRAMNAME
        downloadURLToTempFile = ""
        Screen.MousePointer = 0
        Exit Function
    End If
    
    'Using the new Internet session, attempt to find the URL; if found, assign it a handle.
    Message "Verifying image URL (this may take a moment)..."
    
    Dim hUrl As Long
    hUrl = InternetOpenUrl(hInternetSession, URL, vbNullString, 0, INTERNET_FLAG_RELOAD, 0)

    If hUrl = 0 Then
        pdMsgBox "%1 could not locate a valid file at that URL.  Please double-check the path.  If the problem persists, try downloading the file manually using your Internet browser.", vbExclamation + vbApplicationModal + vbOKOnly, "Online File Not Found", PROGRAMNAME
        If hInternetSession Then InternetCloseHandle hInternetSession
        downloadURLToTempFile = ""
        Screen.MousePointer = 0
        Exit Function
    End If
    
    'Check the size of the file to be downloaded...
    Dim downloadSize As Long
    Dim tmpStrBuffer As String
    tmpStrBuffer = String$(1024, 0)
    Call HttpQueryInfo(ByVal hUrl, HTTP_QUERY_CONTENT_LENGTH, ByVal tmpStrBuffer, Len(tmpStrBuffer), 0)
    downloadSize = CLng(Val(tmpStrBuffer))
    SetProgBarVal 0
    
    If downloadSize <> 0 Then SetProgBarMax downloadSize
    
    'We need a temporary file to house the file; generate it automatically, using the extension of the original file.
    Message "Creating temporary file..."
    Dim tmpFilename As String
    tmpFilename = URL
    StripFilename tmpFilename
    makeValidWindowsFilename tmpFilename
    
    Dim tmpFile As String
    tmpFile = g_UserPreferences.GetTempPath & tmpFilename
    
    'Open the temporary file and begin downloading the image to it
    Message "Image URL verified.  Downloading image..."
    Dim fileNum As Integer
    fileNum = FreeFile
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
            If Not chunkOK Then
                pdMsgBox "%1 lost access to the Internet. Please double-check your Internet connection.  If the problem persists, try downloading the file manually using your Internet browser.", vbExclamation + vbApplicationModal + vbOKOnly, "Internet Connection Error", PROGRAMNAME
                If FileExist(tmpFile) Then
                    Close #fileNum
                    Kill tmpFile
                End If
                If hUrl Then InternetCloseHandle hUrl
                If hInternetSession Then InternetCloseHandle hInternetSession
                SetProgBarVal 0
                releaseProgressBar
                downloadURLToTempFile = ""
                Screen.MousePointer = 0
                Exit Function
            End If
   
            'If the file is done, exit this loop
            If numOfBytesRead = 0 Then
                Exit Do
            End If
   
            'If we've made it this far, assume we've received legitimate data.  Place that data into the file.
            Put #fileNum, , Left$(Buffer, numOfBytesRead)
   
            totalBytesRead = totalBytesRead + numOfBytesRead
            
            If downloadSize <> 0 Then
                SetProgBarVal totalBytesRead
                
                'Display a download update in the message area, but do not log it in the debugger (as there may be
                ' many such notifications, and we don't want to inflate the log unnecessarily)
                #If DEBUGMODE = 1 Then
                    Message "Downloading file (%1 of %2 bytes received)...", totalBytesRead, downloadSize, "DONOTLOG"
                #Else
                    Message "Downloading file (%1 of %2 bytes received)...", totalBytesRead, downloadSize
                #End If
                
            End If
            
        'Carry on
        Loop
        
    'Close the temporary file
    Close #fileNum
    
    'Close this URL and Internet session
    If hUrl Then InternetCloseHandle hUrl
    If hInternetSession Then InternetCloseHandle hInternetSession
    
    Message "Download complete. Verifying file integrity..."
    
    'Check to make sure the image downloaded; if the size is unreasonably small, we can assume the site
    ' prevented our download.  (Direct downloads are sometimes treated as hotlinking; similarly, some sites
    ' prevent scraping, which a direct download like this may seem to be.)
    If totalBytesRead < 20 Then
        
        Message "Download canceled.  (Remote server denied access.)"
        
        Dim domainName As String
        domainName = GetDomainName(URL)
        pdMsgBox "Unfortunately, %1 is preventing %2 from directly downloading this image. (Direct downloads are sometimes mistaken as hotlinking by misconfigured servers.)" & vbCrLf & vbCrLf & "You will need to download this file using your Internet browser, then manually load it into %2." & vbCrLf & vbCrLf & "I sincerely apologize for this inconvenience, but unfortunately there is nothing %2 can do about stingy server configurations.  :(", vbCritical + vbApplicationModal + vbOKOnly, "Download Unsuccessful", domainName, PROGRAMNAME
        
        If FileExist(tmpFile) Then Kill tmpFile
        If hUrl Then InternetCloseHandle hUrl
        If hInternetSession Then InternetCloseHandle hInternetSession
        
        SetProgBarVal 0
        releaseProgressBar
        Screen.MousePointer = 0
        
        downloadURLToTempFile = ""
        Exit Function
        
    End If
    
    'If we made it all the way here, the file was downloaded successfully (most likely... with web stuff, it's always
    ' possible that some strange error has occurred, but we have done our due diligence in attempting a download!)
    SetProgBarVal 0
    releaseProgressBar
    Screen.MousePointer = 0
    
    'Return the temp file location
    downloadURLToTempFile = tmpFile

End Function

'CANCEL button
Private Sub CmdCancel_Click()
    Message "Internet import canceled."
    Unload Me
End Sub

'OK Button
Private Sub CmdOK_Click()
    
    'Check to make sure the user followed directions
    Dim fullURL As String
    fullURL = txtURL
    
    If (Left$(fullURL, 7) <> "http://") And (Left$(fullURL, 6) <> "ftp://") Then
        pdMsgBox "This URL is not valid.  Please make sure the URL begins with ""http://"" or ""ftp://"".", vbApplicationModal + vbOKOnly + vbExclamation, "Invalid URL"
        txtURL.selectAll
        Exit Sub
    End If
    
    'If we've made it here, assume the URL is valid
    Me.Visible = False
    
    'Attempt to download the image
    Dim downloadSuccessful As Boolean
    downloadSuccessful = ImportImageFromInternet(fullURL)
    
    'If the download failed, show the user this form (so they can try again).  Otherwise, unload this form.
    If downloadSuccessful = False Then Me.Visible = True Else Unload Me
    
End Sub

'When the form is activated, automatically select the text box for the user.  This makes a quick Ctrl+V possible.
Private Sub Form_Activate()
    txtURL.selectAll
    txtURL.SetFocus
End Sub

'LOAD form
Private Sub Form_Load()

    lblCopyrightWarning.Caption = g_Language.TranslateMessage("Please be respectful of copyrights when downloading images.  Even if an image is available online, it may not be licensed for use outside a specific website. Thanks!")

    Message "Waiting for user input..."
    
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip

End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

