VERSION 5.00
Begin VB.Form FormThemeEditor 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Resource editor"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   13260
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
   ScaleHeight     =   579
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   884
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdButton cmdSave 
      Height          =   495
      Left            =   10680
      TabIndex        =   8
      Top             =   7320
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      Caption         =   "force save"
   End
   Begin PhotoDemon.pdButtonStrip btsResourceType 
      Height          =   975
      Left            =   4200
      TabIndex        =   6
      Top             =   3480
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   1720
      Caption         =   "resource type"
   End
   Begin PhotoDemon.pdTextBox txtResourceName 
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   1920
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   661
   End
   Begin PhotoDemon.pdButton cmdResourcePath 
      Height          =   375
      Left            =   12840
      TabIndex        =   4
      Top             =   480
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   661
      Caption         =   "..."
   End
   Begin PhotoDemon.pdTextBox txtResourcePath 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   661
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   375
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   661
      Caption         =   "current resource file"
      FontSize        =   12
   End
   Begin PhotoDemon.pdButton cmdAddResource 
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   7200
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   1085
      Caption         =   "add a new resource"
   End
   Begin PhotoDemon.pdListBox lstResources 
      Height          =   6135
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   10821
      Caption         =   "current resources"
   End
   Begin PhotoDemon.pdCommandBarMini cmdBar 
      Align           =   2  'Align Bottom
      Height          =   795
      Left            =   0
      TabIndex        =   0
      Top             =   7890
      Width           =   13260
      _ExtentX        =   23389
      _ExtentY        =   1402
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   375
      Index           =   1
      Left            =   3960
      Top             =   960
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   661
      Caption         =   "edit current resource"
      FontSize        =   12
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   375
      Index           =   2
      Left            =   4200
      Top             =   1560
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   661
      Caption         =   "resource name"
      FontSize        =   12
   End
   Begin PhotoDemon.pdTextBox txtResourceLocation 
      Height          =   375
      Left            =   4200
      TabIndex        =   7
      Top             =   2880
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   661
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   375
      Index           =   3
      Left            =   4200
      Top             =   2520
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   661
      Caption         =   "resource location"
      FontSize        =   12
   End
End
Attribute VB_Name = "FormThemeEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Resource editor dialog
'Copyright 2016-2016 by Tanner Helland
'Created: 22/August/16
'Last updated: 22/August/16
'Last update: initial build
'
'As of v7.0, PD finally supports visual themes using its internal theming engine.  As part of supporting
' visual themes, various PD controls need to grab image resources at a size and color scheme appropriate for
' the current theme.
'
'This resource editor is designed to help with that task.  It shows icons against backgrounds of different
' themes, and allows me to auto-generate resource files for images just by dropping in new image sizes
' (rather than manually describing all image resources).
'
'At present, PD's original resource file is still required, as all resources have *not* yet been migrated
' to the new format.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Private Enum PD_Resource_Type
    PDRT_Image = 0
    PDRT_Other = 1
End Enum

#If False Then
    Private Const PDRT_Image = 0, PDRT_Other = 1
#End If

Private Type PD_Resource
    ResName As String
    ResFileLocation As String
    ResType As PD_Resource_Type
End Type

Private m_NumOfResources As Long
Private m_Resources() As PD_Resource
Private m_LastResourceIndex As Long

Private m_FSO As pdFSO

Private Sub cmdAddResource_Click()
    
    Dim srcFile As String
    
    Dim cCommonDialog As pdOpenSaveDialog: Set cCommonDialog = New pdOpenSaveDialog
    If cCommonDialog.GetOpenFileName(srcFile, , True, False, , , m_FSO.GetPathOnly(txtResourcePath.Text), "Select resource", , Me.hWnd) Then
        
        If (m_NumOfResources > UBound(m_Resources)) Then ReDim Preserve m_Resources(0 To m_NumOfResources * 2 - 1) As PD_Resource
        
        With m_Resources(m_NumOfResources)
            .ResName = m_FSO.GetFilename(srcFile, True)
            .ResFileLocation = srcFile
            .ResType = PDRT_Image
        End With
        
        lstResources.AddItem m_Resources(m_NumOfResources).ResName
        lstResources.ListIndex = m_NumOfResources
        
        SyncUIAgainstCurrentResource
        
        m_NumOfResources = m_NumOfResources + 1
        
    End If
    
End Sub

Private Sub cmdResourcePath_Click()
    
    Dim srcFile As String
    srcFile = m_FSO.GetFilename(txtResourcePath.Text)
    
    Dim cCommonDialog As pdOpenSaveDialog: Set cCommonDialog = New pdOpenSaveDialog
    If cCommonDialog.GetOpenFileName(srcFile, , False, False, "PD Resource Files (*.pdr)|*.pdr", , m_FSO.GetPathOnly(txtResourcePath.Text), "Select resource file", "pdr", Me.hWnd) Then
        If (Len(srcFile) <> 0) Then
            txtResourcePath.Text = srcFile
            g_UserPreferences.SetPref_String "Themes", "LastResourceFile", srcFile
        End If
    End If
    
End Sub

Private Sub Form_Load()
            
    btsResourceType.AddItem "image", 0
    btsResourceType.AddItem "other", 1
    btsResourceType.ListIndex = 0
    
    Set m_FSO = New pdFSO
    
    'Load the last-edited resource file (if any)
    If g_UserPreferences.DoesValueExist("Themes", "LastResourceFile") Then
        txtResourcePath.Text = g_UserPreferences.GetPref_String("Themes", "LastResourceFile", "")
    Else
        txtResourcePath.Text = ""
    End If
    
    m_NumOfResources = 0
    ReDim m_Resources(0 To 15) As PD_Resource
    
    lstResources.ListIndex = -1
    m_LastResourceIndex = -1
    
    Interface.ApplyThemeAndTranslations Me
    
End Sub

Private Sub lstResources_Click()
    SyncResourceAgainstCurrentUI
    m_LastResourceIndex = lstResources.ListIndex
    SyncUIAgainstCurrentResource
End Sub

Private Sub txtResourceName_LostFocusAPI()

    lstResources.UpdateItem lstResources.ListIndex, txtResourceName.Text
    
    If (lstResources.ListIndex >= 0) Then
        m_Resources(lstResources.ListIndex).ResName = txtResourceName.Text
    End If
    
End Sub

'Prior to changing the current resource index, this function can be called to update the last-selected resource against
' any UI changes the user may have entered.
Private Sub SyncResourceAgainstCurrentUI()

    If (m_LastResourceIndex >= 0) Then
    
        With m_Resources(m_LastResourceIndex)
            .ResName = txtResourceName.Text
            .ResType = btsResourceType.ListIndex
            .ResFileLocation = txtResourceLocation.Text
        End With
    
    End If
    
End Sub

'Whenever the current resource index is changed (e.g. by clicking the left-hand list box), this function can be called
' to update all UI elements against the newly selected resource.
Private Sub SyncUIAgainstCurrentResource()
    
    If (m_LastResourceIndex >= 0) Then
    
        With m_Resources(m_LastResourceIndex)
            txtResourceName.Text = .ResName
            btsResourceType.ListIndex = .ResType
            txtResourceLocation.Text = .ResFileLocation
        End With
    
    End If
    
End Sub
