VERSION 5.00
Begin VB.Form FormInternetImport 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Download image"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10050
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
   ScaleHeight     =   179
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   670
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdCommandBarMini cmdBarMini 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   1935
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdTextBox txtURL 
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   556
      Text            =   "http://"
   End
   Begin PhotoDemon.pdLabel lblCopyrightWarning 
      Height          =   615
      Left            =   240
      Top             =   1320
      Width           =   9615
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   ""
      FontSize        =   9
      ForeColor       =   8421504
      Layout          =   1
   End
   Begin PhotoDemon.pdLabel lblDownloadPath 
      Height          =   285
      Left            =   120
      Top             =   360
      Width           =   9720
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "full download path (must begin with ""http://"" or ""ftp://"")"
      FontSize        =   12
      ForeColor       =   4210752
   End
End
Attribute VB_Name = "FormInternetImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Internet Interface (for importing images directly from a URL)
'Copyright 2011-2026 by Tanner Helland
'Created: 08/June/12
'Last updated: 19/July/21
'Last update: move download code elsewhere, so it can be used for non-image purposes
'
'Simple UI for entering a URL to download.  It's expected that most users won't rely on this dialog.
' (Typically, a simple Ctrl+V after copying an image link is sufficient.)  However, this mirrors
' similar functionality in GIMP.
'
'The actual download code doesn't exist here.  Check the Web module for details.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Import an image from the Internet; all that's required is a valid URL (must be prefaced with the protocol, e.g. http:// or similar)
Public Function ImportImageFromInternet(ByRef URL As String) As Boolean

    'First things first - if an invalid URL was provided, exit immediately.
    If (LenB(URL) = 0) Then
        Message "Image download canceled."
        Exit Function
    End If
    
    'Use the generic download function to retrieve the URL
    Dim downloadedFilename As String
    downloadedFilename = Web.DownloadURLToTempFile(URL)
    
    'If the download worked, attempt to load the image.
    If (LenB(downloadedFilename) <> 0) Then
        
        Dim tmpFilename As String
        tmpFilename = Files.FileGetName(downloadedFilename)
        Loading.LoadFileAsNewImage downloadedFilename, tmpFilename, False
        
        'Unique to this particular import is remembering the full filename + extension (because this method of import
        ' actually supplies a file extension, unlike scanning or screen capturing or something else)
        If PDImages.IsImageActive() Then
            PDImages.GetActiveImage.ImgStorage.AddEntry "OriginalFileName", Files.FileGetName(downloadedFilename, True)
            PDImages.GetActiveImage.ImgStorage.AddEntry "OriginalFileExtension", Files.FileGetExtension(downloadedFilename)
        End If
        
        'Delete the temporary file
        Files.FileDeleteIfExists downloadedFilename
        
        Message "Image download complete. "
        ImportImageFromInternet = True
        
    Else
        ImportImageFromInternet = False
    End If
    
End Function

Private Sub cmdBarMini_OKClick()

    'Check to make sure the user followed directions
    Dim fullURL As String
    fullURL = Trim$(txtURL)
    
    If (LCase$(Left$(fullURL, 7)) <> "http://") And (LCase$(Left$(fullURL, 8)) <> "https://") And (LCase$(Left$(fullURL, 6)) <> "ftp://") Then
        PDMsgBox "This URL is not valid.  Please make sure the URL begins with ""http://"" or ""ftp://"".", vbOKOnly Or vbExclamation, "Invalid URL"
        txtURL.SelectAll
        cmdBarMini.DoNotUnloadForm
        Exit Sub
    End If
    
    'If we've made it here, assume the URL is valid
    Me.Visible = False
    
    'Attempt to download the image
    Dim downloadSuccessful As Boolean
    downloadSuccessful = ImportImageFromInternet(fullURL)
    
    'If the download failed, show the user this form (so they can try again).  Otherwise, unload this form.
    If Not downloadSuccessful Then
        Me.Visible = True
        cmdBarMini.DoNotUnloadForm
    End If
    
End Sub

'When the form is activated, automatically select the text box for the user.  This makes a quick Ctrl+V possible.
Private Sub Form_Activate()
    txtURL.SetFocusToEditBox True
End Sub

'LOAD form
Private Sub Form_Load()

    lblCopyrightWarning.Caption = g_Language.TranslateMessage("Please be respectful of copyrights when downloading images.  Thanks!")

    Message "Waiting for user input..."
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub
