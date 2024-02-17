VERSION 5.00
Begin VB.Form FormExportLayers 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Layers to files"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8535
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
   ScaleHeight     =   478
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   569
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdButtonStrip btsWhichLayers 
      Height          =   990
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   1746
      Caption         =   "layers to export"
   End
   Begin PhotoDemon.pdCommandBarMini cmdBar 
      Align           =   2  'Align Bottom
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   6555
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   1085
      DontAutoUnloadParent=   -1  'True
   End
   Begin PhotoDemon.pdButton cmdDstFolder 
      Height          =   450
      Left            =   7800
      TabIndex        =   0
      Top             =   555
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   794
      Caption         =   "..."
   End
   Begin PhotoDemon.pdTextBox txtDstFolder 
      Height          =   315
      Left            =   360
      TabIndex        =   1
      Top             =   630
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   556
      Text            =   "automatically generated at run-time"
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   503
      Caption         =   "destination folder"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdTextBox txtPrefix 
      Height          =   315
      Left            =   390
      TabIndex        =   3
      Top             =   3960
      Width           =   3720
      _ExtentX        =   6562
      _ExtentY        =   556
   End
   Begin PhotoDemon.pdButton cmdExportSettings 
      Height          =   735
      Left            =   225
      TabIndex        =   4
      Top             =   5400
      Width           =   8070
      _ExtentX        =   14235
      _ExtentY        =   1296
      Caption         =   "set export settings for this format..."
   End
   Begin PhotoDemon.pdDropDown cboOutputFormat 
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   4560
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   1296
      Caption         =   "file type"
   End
   Begin PhotoDemon.pdButtonStrip btsFilename 
      Height          =   990
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   1746
      Caption         =   "filename"
   End
   Begin PhotoDemon.pdTextBox txtSuffix 
      Height          =   315
      Left            =   4470
      TabIndex        =   8
      Top             =   3960
      Width           =   3720
      _ExtentX        =   6562
      _ExtentY        =   556
   End
   Begin PhotoDemon.pdCheckBox chkSuffix 
      Height          =   330
      Left            =   4440
      TabIndex        =   9
      Top             =   3570
      Width           =   3750
      _ExtentX        =   7990
      _ExtentY        =   582
      Caption         =   "add a suffix to each filename:"
      Value           =   0   'False
   End
   Begin PhotoDemon.pdCheckBox chkPrefix 
      Height          =   330
      Left            =   360
      TabIndex        =   10
      Top             =   3570
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   582
      Caption         =   "add a prefix to each filename:"
      Value           =   0   'False
   End
End
Attribute VB_Name = "FormExportLayers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Export Layers to Files dialog
'Copyright 2024-2024 by Tanner Helland
'Created: 22/January/24
'Last updated: 01/February/24
'Last update: complete initial build
'
'Photoshop provides a useful "File > Export layers to files" tool (earlier versions had this
' tool in the "Scripts" menu).  This feature is a common request in other software, like Paint.NET
' (see https://superuser.com/questions/424033/extracting-layers-from-paint-net).
'
'This dialog provides similar functionality in PhotoDemon.  The user can specify output format,
' visible vs all layers, and a few simple file naming tools.  This feature set is largely copied
' from the same tool in Photoshop.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'After selecting a target file format, the user needs to manually set additional export options (for some formats)
Private m_ExportSettingsSet As Boolean, m_ExportSettingsFormat As String, m_ExportSettingsMetadata As String

Private Sub cboOutputFormat_Click()
        
    'If this format doesn't support export settings, hide the "set export settings" button
    If ImageFormats.IsExportDialogSupported(ImageFormats.GetOutputPDIF(cboOutputFormat.ListIndex)) Then
        m_ExportSettingsSet = False
        m_ExportSettingsFormat = vbNullString
        m_ExportSettingsMetadata = vbNullString
        cmdExportSettings.Visible = True
    Else
        m_ExportSettingsSet = True
        m_ExportSettingsFormat = vbNullString
        m_ExportSettingsMetadata = vbNullString
        cmdExportSettings.Visible = False
    End If
    
End Sub

Private Sub cmdBar_OKClick()
    
    'Make sure the user clicked the "set export options" button for their selected format
    If (Not m_ExportSettingsSet) Then
        PDMsgBox "Before proceeding, you need to click the ""set export settings for this format"" button to specify what export settings you want to use.", vbExclamation Or vbOKOnly, "Export settings required"
        Exit Sub
    End If
    
    'Make sure the destination folder exists
    Dim cFSO As pdFSO
    Set cFSO = New pdFSO
    
    Dim dstFolder As String
    dstFolder = Files.PathAddBackslash(txtDstFolder.Text)
    If (Not Files.PathExists(dstFolder, False)) Then Files.PathCreate dstFolder, True
    
    'Figure out export format
    Dim exportFormat As PD_IMAGE_FORMAT
    exportFormat = ImageFormats.GetOutputPDIF(Me.cboOutputFormat.ListIndex)
    
    'Time to start iterating layers.  Start by figuring out initial and final indices
    Dim idxStart As Long, idxEnd As Long
    Select Case btsWhichLayers.ListIndex
        
        'All layers
        Case 0
            idxStart = 0
            idxEnd = PDImages.GetActiveImage.GetNumOfLayers - 1
        
        'Visible layers only
        Case 1
            idxStart = 0
            idxEnd = PDImages.GetActiveImage.GetNumOfLayers - 1
        
        'Current layer only
        Case 2
            idxStart = PDImages.GetActiveImage.GetActiveLayerIndex
            idxEnd = PDImages.GetActiveImage.GetActiveLayerIndex
        
    End Select
    
    'Make a backup list of layer visibility (as we're going to cheat and simply toggle layer visibility
    ' on the active image).
    Dim backupVisibility() As Boolean
    ReDim backupVisibility(0 To PDImages.GetActiveImage.GetNumOfLayers - 1) As Boolean
    
    Dim i As Long
    For i = 0 To PDImages.GetActiveImage.GetNumOfLayers - 1
        backupVisibility(i) = PDImages.GetActiveImage.GetLayerByIndex(i).GetLayerVisibility()
    Next i
    
    'Start iterating layers and export as we go!
    For i = idxStart To idxEnd
        
        'The current layer may be disqualified from export for various reasons (e.g. it is hidden,
        ' but the user selected "export visible layers only"
        Dim okToExport As Boolean
        okToExport = True
        If (btsWhichLayers.ListIndex = 1) Then okToExport = okToExport And PDImages.GetActiveImage.GetLayerByIndex(i).GetLayerVisibility()
        
        'If this layer is a valid export target, we need to make all other layers invisible, then export the result
        If okToExport Then
            
            'Hide all layers but this one
            Dim j As Long
            For j = 0 To PDImages.GetActiveImage.GetNumOfLayers - 1
                PDImages.GetActiveImage.GetLayerByIndex(i).SetLayerVisibility (i = j)
                PDImages.GetActiveImage.NotifyImageChanged UNDO_LayerHeader, i
            Next j
            
            'Grab a composite copy of the new visibility-adjusted image
            Dim tmpComposite As pdDIB
            PDImages.GetActiveImage.GetCompositedImage tmpComposite
            
            'Prepare a temporary pdImage object to house the exported frame
            Dim tmpImage As pdImage
            Set tmpImage = New pdImage
            
            'In the temporary pdImage object, create a blank layer; this will receive this layer's contents
            Dim newLayerID As Long
            newLayerID = tmpImage.CreateBlankLayer
            tmpImage.GetLayerByID(newLayerID).InitializeNewLayer PDL_Image, , tmpComposite
            tmpImage.UpdateSize
            
            'Assign some default settings to the exported image
            tmpImage.SetAnimated False
            tmpImage.SetDPI PDImages.GetActiveImage.GetDPI(), PDImages.GetActiveImage.GetDPI()
            
            'Generate an output filename
            Dim newFilename As String
            newFilename = GetFinalFilename(i)
            
            'We're now going to loop into the batch process exporter, because it works great for one-off file exports
            Saving.PhotoDemon_BatchSaveImage tmpComposite, newFilename, exportFormat, m_ExportSettingsFormat, m_ExportSettingsMetadata
            
            'Free the temporary image
            Set tmpImage = Nothing
            
        End If
        
    Next i
    
    'Before exiting, restore original layer visibility
    For i = 0 To PDImages.GetActiveImage.GetNumOfLayers - 1
        PDImages.GetActiveImage.GetLayerByIndex(i).SetLayerVisibility backupVisibility(i)
        PDImages.GetActiveImage.NotifyImageChanged UNDO_LayerHeader, i
    Next i
    
End Sub

'Do not pass invalid files or paths to this function.  It does not validate inputs.
Private Function GetFinalFilename(ByVal idxLayer As Long) As String
    
    'The valid filename for this file depends on the user's current settings.  Start by grabbing the layer's current name.
    Dim curLayerName As String
    curLayerName = PDImages.GetActiveImage.GetLayerByIndex(idxLayer).GetLayerName()
    
    
End Function

Private Sub cmdDstFolder_Click()
    Dim folderPath As String
    folderPath = Files.PathBrowseDialog(Me.hWnd, txtDstFolder.Text)
    If (LenB(folderPath) <> 0) Then
        txtDstFolder.Text = Files.PathAddBackslash(folderPath)
        UserPrefs.SetPref_String "Paths", "export-layers", txtDstFolder.Text
    End If
End Sub

Private Sub cmdExportSettings_Click()

    'Convert the current dropdown index into a PD format constant
    Dim saveFormat As PD_IMAGE_FORMAT
    saveFormat = ImageFormats.GetOutputPDIF(cboOutputFormat.ListIndex)
    
    'Not all formats require settings dialogs...
    If ImageFormats.IsExportDialogSupported(saveFormat) Then
        
        'The saving module will now raise a dialog specific to the selected format.
        ' If successful, it will fill the passed settings and metadata strings with XML data
        ' describing the user's chosen settings.
        m_ExportSettingsSet = Saving.GetExportParamsFromDialog(Nothing, saveFormat, m_ExportSettingsFormat, m_ExportSettingsMetadata, False)
        
        'If the user cancels the dialog, exit immediately
        If (Not m_ExportSettingsSet) Then
            m_ExportSettingsSet = False
            m_ExportSettingsFormat = vbNullString
            m_ExportSettingsMetadata = vbNullString
        End If
    
    'Formats that do not support export settings do not need to raise a dialog at all
    Else
        m_ExportSettingsSet = True
        m_ExportSettingsFormat = vbNullString
        m_ExportSettingsMetadata = vbNullString
    End If
    
End Sub

Private Sub Form_Load()
    
    'Load default destination folder.  If previously saved paths are not found, default to the user's current
    ' "save image" path.
    If UserPrefs.DoesValueExist("Paths", "export-layers") Then
        txtDstFolder.Text = UserPrefs.GetPref_String("Paths", "export-layers", UserPrefs.GetPref_String("Paths", "Save Image", vbNullString))
    Else
        txtDstFolder.Text = UserPrefs.GetPref_String("Paths", "Save Image", vbNullString)
    End If
    
    'Users can choose a subset of layers to export
    btsWhichLayers.AddItem "all layers", 0
    btsWhichLayers.AddItem "visible layers", 1
    btsWhichLayers.AddItem "current layer", 2
    btsWhichLayers.ListIndex = 0
    
    'Files support a few different naming schemes
    btsFilename.AddItem "layer name", 0
    btsFilename.AddItem "ascending numbers (1, 2, 3, etc.)", 1
    
    'Populate export file formats, and set the default output format
    m_ExportSettingsSet = False
    Dim i As Long
    For i = 0 To ImageFormats.GetNumOfOutputFormats()
        cboOutputFormat.AddItem ImageFormats.GetOutputFormatDescription(i), i
    Next i
    
    cboOutputFormat.ListIndex = ImageFormats.GetIndexOfOutputPDIF(PDIF_PNG)
    
    'Before the user proceeds, they need to manually set export settings for their chosen format.
    ' (Changing the target format resets this to FALSE.)
    m_ExportSettingsSet = False
    
    'Theme everything
    Interface.ApplyThemeAndTranslations Me
    
End Sub

Private Sub txtPrefix_Change()
    If (Not chkPrefix.Value) Then chkPrefix.Value = True
End Sub

Private Sub txtSuffix_Change()
    If (Not chkSuffix.Value) Then chkSuffix.Value = True
End Sub
