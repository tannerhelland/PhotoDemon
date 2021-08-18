VERSION 5.00
Begin VB.Form FormMetadataBasic 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edit metadata"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   14190
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
   ScaleHeight     =   539
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   946
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdListBox lstGroup 
      Height          =   7095
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3360
      _ExtentX        =   5927
      _ExtentY        =   11245
      Caption         =   "categories"
   End
   Begin PhotoDemon.pdCommandBarMini cmdBarMini 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   7335
      Width           =   14190
      _ExtentX        =   25030
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdContainer pnlGroup 
      Height          =   7095
      Index           =   0
      Left            =   3720
      Top             =   120
      Width           =   10290
      _ExtentX        =   0
      _ExtentY        =   0
      Begin PhotoDemon.pdButtonStrip btsMetadata 
         Height          =   495
         Index           =   0
         Left            =   2760
         TabIndex        =   6
         Top             =   3000
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   873
      End
      Begin PhotoDemon.pdTextBox txtMetadata 
         Height          =   375
         Index           =   0
         Left            =   2760
         TabIndex        =   2
         Top             =   120
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   661
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   375
         Index           =   0
         Left            =   120
         Top             =   120
         Width           =   2535
         _ExtentX        =   5318
         _ExtentY        =   661
         Alignment       =   1
         Caption         =   "document title"
         FontSize        =   11
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   375
         Index           =   1
         Left            =   120
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         Alignment       =   1
         Caption         =   "author"
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   375
         Index           =   2
         Left            =   120
         Top             =   1080
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         Alignment       =   1
         Caption         =   "author title"
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   375
         Index           =   3
         Left            =   120
         Top             =   1560
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         Alignment       =   1
         Caption         =   "description"
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   375
         Index           =   4
         Left            =   120
         Top             =   3000
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         Alignment       =   1
         Caption         =   "rating"
      End
      Begin PhotoDemon.pdTextBox txtMetadata 
         Height          =   375
         Index           =   1
         Left            =   2760
         TabIndex        =   3
         Top             =   600
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   661
      End
      Begin PhotoDemon.pdTextBox txtMetadata 
         Height          =   375
         Index           =   2
         Left            =   2760
         TabIndex        =   4
         Top             =   1080
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   661
      End
      Begin PhotoDemon.pdTextBox txtMetadata 
         Height          =   1335
         Index           =   3
         Left            =   2760
         TabIndex        =   5
         Top             =   1560
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   2355
         Multiline       =   -1  'True
      End
   End
End
Attribute VB_Name = "FormMetadataBasic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Image Metadata Editor
'Copyright 2013-2021 by Tanner Helland
'Created: 27/May/13
'Last updated: 28/June/21
'Last update: split code from the old metadata editor to a new, "simper" interface (similar to Photoshop's)
'
'As of version 6.0, PhotoDemon now provides support for loading and saving image metadata.  What is metadata, you ask?
' See https://en.wikipedia.org/wiki/Metadata#Photographs for more details.
'
'This dialog interacts heavily with the pdMetadata class to present users with a relatively simple interface for
' perusing (and eventually, editing) an image's metadata.
'
'Categories are displayed on the left, and selecting a category pulls up a dedicated panel with metadata entries
' related to that category.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Enum MetadataBasicName
    mbn_DocumentTitle
    mbn_Author
    mbn_AuthorTitle
    mbn_Description
    mbn_Rating
End Enum

#If False Then
    Private Const mbn_DocumentTitle = 0, mbn_Author = 0, mbn_AuthorTitle = 0, mbn_Description = 0, mbn_Rating = 0
#End If

'When the form is loaded, all metadata relevant to controls *on this dialog* is retrieved and
' stored in this local-only format.  If the user clicks OK, this local-only format will be
' translated back into PD's master metadata format.
Private Type MetadataBasic
    mdName As String
    mdOrig As String
    mdNew As String
End Type

Private m_MetadataCount As Long
Private m_Metadata() As MetadataBasic
Private Const INIT_METADATA_COUNT As Long = 16

Private Sub cmdBarMini_OKClick()
    
    'With all metadata updated, notify the central processor that an Undo update is required
    'Process "Edit metadata", False, , UNDO_ImageHeader
    
End Sub

Private Sub Form_Load()
    
    'Initialize a basic metadata collection.  (This will get filled with metadata relevant to
    ' UI elements on the dialog.)
    ReDim m_Metadata(0 To INIT_METADATA_COUNT - 1) As MetadataBasic
    
    Dim i As Long
    
    'Populate available metadata categories
    lstGroup.SetAutomaticRedraws False
    lstGroup.AddItem "description", 0, True
    
    lstGroup.AddItem "camera data", 1
    lstGroup.AddItem "GPS", 2, True
    
    lstGroup.AddItem "IPTC contact", 3
    lstGroup.AddItem "IPTC content", 4
    lstGroup.AddItem "IPTC image", 5
    lstGroup.AddItem "IPTC status", 6
    
    lstGroup.ListIndex = 0
    lstGroup.SetAutomaticRedraws True, True
    
    'A number of metadata-related UI elements also need to be manually populated
    btsMetadata(0).AddItem "none", 0
    For i = 1 To 5
        btsMetadata(0).AddItem CStr(i), i
    Next i
    
    'Prep any other interface components
    
    'Next, initialize all metadata elements.  (This will attempt to retrieve metadata values -
    ' if they exist in the current image - for all UI elements in this dialog.)
    GetInitialMetadata
    
    'Theme the dialog
    ApplyThemeAndTranslations Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub lstGroup_Click()
    
    'Show matching panel
    Dim i As Long
    For i = 0 To lstGroup.ListCount - 1
        If (i < pnlGroup.Count) Then
            pnlGroup(i).Visible = (i = lstGroup.ListIndex)
        End If
    Next i
    
End Sub

'Pull initial metadata values from this image.  This process basically attempts to match up
' on-screen metadata UI elements with any existant metadata inside the current image file.
' (Note that this process is more complicated than you'd think, as different image formats
' store metadata in different places.  ExifTool is a huge help here.)
Private Sub GetInitialMetadata()

End Sub
