VERSION 5.00
Begin VB.Form dialog_AutosaveWarning 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Autosave data detected"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9165
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
   ScaleHeight     =   465
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   611
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3405
      Left            =   3960
      ScaleHeight     =   225
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   330
      TabIndex        =   6
      Top             =   2430
      Width           =   4980
   End
   Begin VB.ListBox lstAutosaves 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3420
      IntegralHeight  =   0   'False
      Left            =   240
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   2430
      Width           =   3615
   End
   Begin PhotoDemon.jcbutton cmdOK 
      Height          =   735
      Left            =   2280
      TabIndex        =   0
      Top             =   6060
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   1296
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
      Caption         =   "Restore selected images  "
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_FormAutosaveWarning.frx":0000
      PictureAlign    =   0
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipType     =   1
   End
   Begin PhotoDemon.jcbutton cmdCancel 
      Default         =   -1  'True
      Height          =   735
      Left            =   5640
      TabIndex        =   5
      Top             =   6060
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   1296
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
      Caption         =   "Discard all images  "
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_FormAutosaveWarning.frx":1052
      PictureAlign    =   0
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipType     =   1
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      X1              =   16
      X2              =   595
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "autosave entries found:"
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
      TabIndex        =   4
      Top             =   2040
      Width           =   2490
   End
   Begin VB.Label lblWarning 
      BackStyle       =   0  'Transparent
      Caption         =   "Warning"
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
      Height          =   645
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   990
      Width           =   8745
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblWarning 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Autosave data found.  Would you like to restore it?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00202020&
      Height          =   525
      Index           =   0
      Left            =   1005
      TabIndex        =   1
      Top             =   330
      Width           =   8055
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "dialog_AutosaveWarning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Autosave (unsafe shutdown) Prompt/Dialog
'Copyright ©2013-2014 by Tanner Helland
'Created: 19/January/14
'Last updated: 21/May/14
'Last update: rewrote the entire dialog against the new Undo/Redo engine
'
'PhotoDemon now provides AutoSave functionality.  If the program terminates unexpectedly, this dialog will be raised,
' which gives the user an option to restore any in-progress image edits.
'
'Images that had been loaded by PhotoDemon but never modified will not be shown.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'The user input from the dialog
Private userAnswer As VbMsgBoxResult

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'Collection of Autosave XML entries found
Private m_numOfXMLFound As Long
Private m_XmlEntries() As AutosaveXML

'When this dialog finally closes, the calling function can use this sub to retrieve the entries the user wants saved.
Friend Sub fillArrayWithSaveResults(ByRef dstArray() As AutosaveXML)
    
    Dim numOfEntriesBeingSaved As Long
    numOfEntriesBeingSaved = 0
    
    'Count how many entries the user is saving
    Dim i As Long
    For i = 0 To lstAutosaves.ListCount - 1
        If lstAutosaves.Selected(i) Then numOfEntriesBeingSaved = numOfEntriesBeingSaved + 1
    Next i
    
    'Prepare the destination array
    ReDim dstArray(0 To numOfEntriesBeingSaved - 1) As AutosaveXML
    
    'Fill the array with all selected entries
    numOfEntriesBeingSaved = 0
    For i = 0 To lstAutosaves.ListCount - 1
        If lstAutosaves.Selected(i) Then
            dstArray(numOfEntriesBeingSaved) = m_XmlEntries(lstAutosaves.ItemData(i))
            numOfEntriesBeingSaved = numOfEntriesBeingSaved + 1
        End If
    Next i
    
End Sub

Public Property Get DialogResult() As VbMsgBoxResult
    DialogResult = userAnswer
End Property

'The ShowDialog routine presents the user with the form.  FormID MUST BE SET in advance of calling this.
Public Sub showDialog()

    'Automatically draw a warning icon using the system icon set
    Dim iconY As Long
    iconY = fixDPI(18)
    If g_UseFancyFonts Then iconY = iconY + fixDPI(2)
    DrawSystemIcon IDI_EXCLAMATION, Me.hDC, fixDPI(22), iconY
    
    'Display a brief explanation of the dialog at the top of the window
    lblWarning(1).Caption = g_Language.TranslateMessage("A previous PhotoDemon session terminated unexpectedly.  Would you like to automatically recover the following autosaved images?")
        
    'Provide a default answer of "restore all images" (in the event that the user clicks the "x" button in the top-right)
    userAnswer = vbOK

    'Apply any custom styles to the form
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip

    'Populate the AutoSave entry list box
    displayAutosaveEntries

    'Display the form
    showPDDialog vbModal, Me, True

End Sub

'If the user cancels, warn them that these image will be lost foreeeever.
Private Sub CmdCancel_Click()

    Dim msgReturn As VbMsgBoxResult
    msgReturn = pdMsgBox("If you exit now, this autosave data will be lost forever.  Are you sure you want to exit?", vbApplicationModal + vbInformation + vbYesNo, "Warning: autosave data will be deleted")
    
    If msgReturn = vbYes Then
        userAnswer = vbNo
        Me.Hide
    End If

End Sub

'OK button
Private Sub CmdOK_Click()
    userAnswer = vbYes
    Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Update the active image preview in the top-right
Private Sub updatePreview(ByVal srcImagePath As String)
    
    'Display a preview of the selected image
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    QuickLoadImageToDIB srcImagePath, tmpDIB
    
    If Not (tmpDIB Is Nothing) And (tmpDIB.getDIBWidth > 0) And (tmpDIB.getDIBHeight > 0) Then
        tmpDIB.renderToPictureBox picPreview
    Else
        picPreview.Picture = LoadPicture("")
        Dim strToPrint As String
        strToPrint = g_Language.TranslateMessage("Preview not available")
        picPreview.CurrentX = (picPreview.ScaleWidth - picPreview.TextWidth(strToPrint)) \ 2
        picPreview.CurrentY = (picPreview.ScaleHeight - picPreview.TextHeight(strToPrint)) \ 2
        picPreview.Print strToPrint
    End If
    
End Sub

'Fill the AutoSave entries list with any images found from the Autosave engine
Private Function displayAutosaveEntries() As Boolean

    'Because we've arrived at this point, we know that the Autosave engine has found at least *some* usable image data.
    ' Our goal now is to present that image data to the user, so they can select which images (if any) they want us
    ' to restore.
    
    'The Image_Autosave_Handler module will already contain a list of all Undo XML files found by the Autosave engine.
    ' It has stored this data in its private m_XmlEntries() array.  We can request a copy of this array as follows:
    Image_Autosave_Handler.getXMLAutosaveEntries m_XmlEntries(), m_numOfXMLFound
    
    'All XML entries will now have been matched up with their latest Undo entry.  Fill the listbox with their data,
    ' ignoring any entries that do not have binary image data attached.
    lstAutosaves.Clear
    
    Dim i As Long
    For i = 0 To m_numOfXMLFound - 1
        lstAutosaves.AddItem m_XmlEntries(i).friendlyName
        lstAutosaves.ItemData(lstAutosaves.newIndex) = i
        lstAutosaves.Selected(lstAutosaves.newIndex) = True
    Next i
    
    'Select the entry at the top of the list by default
    lstAutosaves.ListIndex = 0
    
End Function

Private Sub lstAutosaves_Click()

    'It's a bit ridiculous, but PD always saves a thumbnail of the latest image state to the same Undo path
    ' as the XML file, but with an "asp" extension.  I know what "asp" is usually used for, but in this case,
    ' it means "autosave preview".  The confusing extension also provides a bit of obfuscation about the file's
    ' true contents (PNG data), which never hurts when sticking stuff in the temp folder.
    Dim previewPath As String
    previewPath = m_XmlEntries(lstAutosaves.ItemData(lstAutosaves.ListIndex)).xmlPath & ".asp"
    
    updatePreview previewPath
    
End Sub
