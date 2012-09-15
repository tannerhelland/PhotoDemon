VERSION 5.00
Begin VB.Form FormInternetImport 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Download Image"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8940
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
   ScaleHeight     =   131
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   596
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   6480
      TabIndex        =   1
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7680
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txtURL 
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "http://"
      Top             =   600
      Width           =   8655
   End
   Begin VB.Label lblCopyrightWarning 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"VBP_FormInternetImport.frx":0000
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1305
      Width           =   6255
   End
   Begin VB.Label lblDownloadPath 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Full download path (must begin with ""http://"" or ""ftp://""):"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "FormInternetImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Internet Interface (for importing images directly from a URL)
'Copyright ©2000-2012 by Tanner Helland
'Created: 08/June/12
'Last updated: 15/June/12
'Last update: further hardened error handling.  Image URLs are now checked for invalid Windows filenames
'             (important since we name our temporary file after the URL to ensure that all the image class
'             data populates correctly); invalid characters are now replaced by an underscore.
'
'Interface for downloading images directly from the Internet into PhotoDemon.  This code is a heavily
' modified version of publicly available code by Alberto Falossi (http://www.devx.com/vb2themax/Tip/19203).
'
'A number of features have been added to the original version of this code.  The routine checks the file download
' size, and updates the user (via progress bar) on the download progress.  Many checks are in place to protect
' against Internet and download errors.  I'm quite proud of how robust this implementation is, but additional
' testing will be necessary to make sure no possible connectivity errors have been overlooked.
'
'***************************************************************************

Option Explicit

'Import an image from the Internet; all that's required is a valid URL (must be prefaced with http:// or ftp://)
Public Function ImportImageFromInternet(ByVal URL As String) As Boolean

    'First things first - prompt the user for a URL
    If URL = "" Then
        Message "Image download canceled."
        Exit Function
    End If
    
    'Normally changing the cursor is handled by the software processor, but because this function routes
    ' internally, we'll make an exception and change it here. Note that everywhere this function can
    ' terminate (and it's many places - a lot can go wrong while downloading) - the cursor needs to be reset.
    FormMain.MousePointer = vbHourglass
    
    'Open an Internet session and assign it a handle
    Dim hInternetSession As Long
    
    Message "Attempting to connect to Internet..."
    hInternetSession = InternetOpen(App.EXEName, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    
    If hInternetSession = 0 Then
        MsgBox PROGRAMNAME & " could not establish an Internet connection. Please double-check your connection.  If the problem persists, try downloading the image manually using your Internet browser of choice.  Once downloaded, you may open the file in " & PROGRAMNAME & " just like any other image file.", vbCritical + vbApplicationModal + vbOKOnly, "Internet Connection Error"
        ImportImageFromInternet = False
        FormMain.MousePointer = 0
        Exit Function
    End If
    
    'Using the new Internet session, attempt to find the URL; if found, assign it a handle
    Message "Verifying image URL (this may take a moment)..."
    
    Dim hUrl As Long
    hUrl = InternetOpenUrl(hInternetSession, URL, vbNullString, 0, INTERNET_FLAG_RELOAD, 0)

    If hUrl = 0 Then
        MsgBox PROGRAMNAME & " could not locate a valid image at that URL.  Please double-check the path.  If the problem persists, try downloading the image manually using your Internet browser of choice.  Once downloaded, you may open the file in " & PROGRAMNAME & " just like any other image file.", vbCritical + vbApplicationModal + vbOKOnly, "Online Image Not Found"
        If hInternetSession Then InternetCloseHandle hInternetSession
        ImportImageFromInternet = False
        FormMain.MousePointer = 0
        Exit Function
    End If
    
    'Check the size of the image to be downloaded...
    Dim downloadSize As Long
    Dim tmpStrBuffer As String
    tmpStrBuffer = String$(1024, 0)
    Call HttpQueryInfo(ByVal hUrl, HTTP_QUERY_CONTENT_LENGTH, ByVal tmpStrBuffer, Len(tmpStrBuffer), 0)
    downloadSize = CLng(val(tmpStrBuffer))
    SetProgBarVal 0
    If downloadSize <> 0 Then SetProgBarMax downloadSize
    
    'We need a temporary file to house the image; generate it automatically, using the extension of the original image
    Message "Creating temporary file..."
    Dim tmpFileName As String
    tmpFileName = URL
    StripFilename tmpFileName
    makeValidWindowsFilename tmpFileName
    
    Dim tmpFile As String
    tmpFile = TempPath & tmpFileName
    
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
            If chunkOK = False Then
                MsgBox PROGRAMNAME & " lost access to the Internet. Please double-check your Internet connection.  If the problem persists, try downloading the image manually using your Internet browser of choice.  Once downloaded, you may open the file in " & PROGRAMNAME & " just like any other image file.", vbCritical + vbApplicationModal + vbOKOnly, "Internet Connection Error"
                If FileExist(tmpFile) Then
                    Close #fileNum
                    Kill tmpFile
                End If
                If hUrl Then InternetCloseHandle hUrl
                If hInternetSession Then InternetCloseHandle hInternetSession
                SetProgBarVal 0
                ImportImageFromInternet = False
                FormMain.MousePointer = 0
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
                Message "Image URL verified.  Downloading image (" & totalBytesRead & " of " & downloadSize & " bytes received)..."
            End If
            
            DoEvents
            
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
        MsgBox "Unfortunately, " & domainName & " is preventing " & PROGRAMNAME & " from directly downloading this image. (Direct downloads are sometimes mistaken as hotlinking by misconfigured servers.)" & vbCrLf & vbCrLf & "You will need to download this image using your Internet browser of choice, then manually load it into " & PROGRAMNAME & "." & vbCrLf & vbCrLf & "I sincerely apologize for this inconvenience, but unfortunately there is nothing " & PROGRAMNAME & " can do about stingy server configurations.  :(", vbCritical + vbApplicationModal + vbOKOnly, domainName & " Does Not Allow Direct Downloads"
        If FileExist(tmpFile) Then Kill tmpFile
        If hUrl Then InternetCloseHandle hUrl
        If hInternetSession Then InternetCloseHandle hInternetSession
        SetProgBarVal 0
        ImportImageFromInternet = False
        FormMain.MousePointer = 0
        Exit Function
    End If
    
    'If we've made it this far, it's probably safe to assume that download worked.  Attempt to load the image.
    Dim sFile(0) As String
    sFile(0) = tmpFile
    
    PreLoadImage sFile, False, tmpFileName, tmpFileName
    
    'Unique to this particular import is remembering the full filename + extension (because this method of import
    ' actually supplies a file extension, unlike scanning or screen capturing or something else)
    pdImages(CurrentImage).OriginalFileNameAndExtension = tmpFileName
    
    SetProgBarVal 0
    
    'Delete the temporary file
    If FileExist(tmpFile) Then Kill tmpFile
    
    Message "Image download complete. "
    
    FormMain.MousePointer = 0
    
    ImportImageFromInternet = True
    
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
        MsgBox "This URL is not valid.  Please make sure the URL begins with ""http://"" or ""ftp://.""", vbApplicationModal + vbOKOnly + vbCritical, "Invalid URL"
        AutoSelectText txtURL
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
    AutoSelectText txtURL
End Sub

'LOAD form
Private Sub Form_Load()

    Message "Waiting for user input..."
    
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me

End Sub
