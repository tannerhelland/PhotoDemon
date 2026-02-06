VERSION 5.00
Begin VB.Form FormImageCreateLUT 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Create color lookup table (LUT)"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7065
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
   ScaleHeight     =   514
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   471
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdCheckBox chkAvailable 
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   6240
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   661
      Caption         =   "add this to the ""Adjustments > Color > Color lookup"" tool"
   End
   Begin PhotoDemon.pdDropDown ddQuality 
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   4920
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   661
   End
   Begin PhotoDemon.pdLabel lblSettings 
      Height          =   315
      Left            =   120
      Top             =   3240
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   556
      Caption         =   "options"
      FontSize        =   12
   End
   Begin PhotoDemon.pdCommandBarMini cmdBar 
      Align           =   2  'Align Bottom
      Height          =   855
      Left            =   0
      TabIndex        =   4
      Top             =   6855
      Width           =   7065
      _ExtentX        =   12462
      _ExtentY        =   1508
   End
   Begin PhotoDemon.pdDropDown ddSource 
      Height          =   735
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      _ExtentX        =   11033
      _ExtentY        =   1296
      Caption         =   "base image and layer"
   End
   Begin PhotoDemon.pdDropDown ddSource 
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   6375
      _ExtentX        =   10398
      _ExtentY        =   661
      FontSizeCaption =   11
   End
   Begin PhotoDemon.pdDropDown ddSource 
      Height          =   735
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   6735
      _ExtentX        =   11033
      _ExtentY        =   1296
      Caption         =   "modified image and layer"
   End
   Begin PhotoDemon.pdDropDown ddSource 
      Height          =   375
      Index           =   3
      Left            =   480
      TabIndex        =   3
      Top             =   2520
      Width           =   6375
      _ExtentX        =   10398
      _ExtentY        =   661
      FontSizeCaption =   11
   End
   Begin PhotoDemon.pdSlider sldGridPoints 
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   5400
      Width           =   6255
      _ExtentX        =   16536
      _ExtentY        =   873
      Min             =   2
      Max             =   64
      Value           =   17
      NotchPosition   =   2
      NotchValueCustom=   17
   End
   Begin PhotoDemon.pdTextBox txtDescription 
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   3960
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   661
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   315
      Index           =   0
      Left            =   480
      Top             =   3600
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   556
      Caption         =   "description"
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   315
      Index           =   2
      Left            =   480
      Top             =   4560
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   556
      Caption         =   "grid points"
   End
End
Attribute VB_Name = "FormImageCreateLUT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Create color lookup from differences between two images
'Copyright 2022-2026 by Tanner Helland
'Created: 18/June/22
'Last updated: 24/June/22
'Last update: save *some* last-used settings (but not ones that will break the app, like the layer dropdowns)
'
'This dialog provides a UI for comparing two images, then generating a 3D LUT from their differences.
' This LUT is then exported to an arbitrary LUT file (of the user's choosing) where it can be used to apply
' the color changes between these two images to *any* image.  This allows you to freely reverse-engineer
' proprietary color transforms from any source (e.g. an Instagram filter) - all you need is a "before"
' and "after" image, and this tool handles the rest.
'
'This feature relies heavily on the pdLUT3D class for handling all the actual LUT generation and export work.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Because we don't use a traditional command bar on this dialog, last-used settings must be
' manually handled.  (We need to do this manually anyway because the image- and layer- dropdowns
' are session-specific.)  Note that this *can* be declared WithEvents, but we do not need custom
' save/load setting support in this tool.
Private lastUsedSettings As pdLastUsedSettings
Attribute lastUsedSettings.VB_VarHelpID = -1

Private m_OpenImageIDs As pdStack

'Grid point values for fast/default/extreme modes
Private Const GRID_FAST As Long = 8
Private Const GRID_DEFAULT As Long = 17
Private Const GRID_EXTREME As Long = 32

'Compare two arbitrary layers from two arbitrary images.  All settings must be encoded in a param string.
Public Sub CreateDifferenceLUT(ByRef listOfParameters As String)
    
    Message "Analyzing image..."
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString listOfParameters
    
    Dim srcImageID As Long, srcLayerIdx As Long
    Dim cmpImageID As Long, cmpLayerIdx As Long
    
    Dim dstFilename As String, dstLutFormat As String, dstDescription As String
    Dim numGridPoints As Long
    
    With cParams
        srcImageID = .GetLong("source-image-id", 0, True)
        srcLayerIdx = .GetLong("source-layer-idx", 0, True)
        cmpImageID = .GetLong("compare-image-id", 0, True)
        cmpLayerIdx = .GetLong("compare-layer-idx", 0, True)
        dstFilename = .GetString("filename", vbNullString, True)
        dstLutFormat = .GetString("lut-format", vbNullString, True)
        dstDescription = .GetString("description", vbNullString, True)
        numGridPoints = .GetLong("grid-points", GRID_DEFAULT, True)
    End With
    
    'Ensure all image, layer, and filename references are valid
    Dim srcImage As pdImage, cmpImage As pdImage
    Set srcImage = PDImages.GetImageByID(srcImageID)
    Set cmpImage = PDImages.GetImageByID(cmpImageID)
    If (srcImage Is Nothing) Or (cmpImage Is Nothing) Then Exit Sub
    
    Dim srcLayer As pdLayer, cmpLayer As pdLayer
    Set srcLayer = srcImage.GetLayerByIndex(srcLayerIdx, False)
    Set cmpLayer = cmpImage.GetLayerByIndex(cmpLayerIdx, False)
    If (srcLayer Is Nothing) Or (cmpLayer Is Nothing) Then Exit Sub
    
    'Before and after layers need to be the same size.  This may necessitate destructive changes,
    ' so we need to be careful about using soft vs hard layer references...
    Dim targetWidth As Long, targetHeight As Long
    
    'Start with the base layer
    Dim tmpLayer As pdLayer
    Set tmpLayer = New pdLayer
    tmpLayer.CopyExistingLayer srcLayer
    
    If srcLayer.AffineTransformsActive(True) Then
        tmpLayer.ConvertToNullPaddedLayer srcImage.Width, srcImage.Height, True
        tmpLayer.CropNullPaddedLayer
    End If
    
    Set srcLayer = tmpLayer
    targetWidth = srcLayer.GetLayerDIB.GetDIBWidth()
    targetHeight = srcLayer.GetLayerDIB.GetDIBHeight()
    
    'Repeat above steps for comparison layer, with the added step of resizing to
    ' match base layer dimensions (if necessary; we may also be cropping/enlarging it).
    Set tmpLayer = New pdLayer
    tmpLayer.CopyExistingLayer cmpLayer
    If cmpLayer.AffineTransformsActive(True) Then
        tmpLayer.ConvertToNullPaddedLayer cmpImage.Width, cmpImage.Height, True
        tmpLayer.CropNullPaddedLayer
    End If
    
    Set cmpLayer = tmpLayer
    
    If (cmpLayer.GetLayerDIB.GetDIBWidth <> targetWidth) Or (cmpLayer.GetLayerDIB.GetDIBHeight <> targetHeight) Then
        
        Dim tmpDIB As pdDIB
        Set tmpDIB = New pdDIB
        tmpDIB.CreateFromExistingDIB cmpLayer.GetLayerDIB, srcLayer.GetLayerDIB.GetDIBWidth, srcLayer.GetLayerDIB.GetDIBHeight, GP_IM_Bilinear
        
        'Copy the DIB into a temporary layer object
        Set cmpLayer = New pdLayer
        cmpLayer.CopyExistingLayer tmpLayer
        cmpLayer.SetLayerDIB tmpDIB
        
    End If
    
    'We now have a source and comparison layer that are guaranteed to have matching sizes
    ' (via either crop/padding or resampling).
    Dim baseDIB As pdDIB, cmpDIB As pdDIB
    Set baseDIB = srcLayer.GetLayerDIB
    Set cmpDIB = cmpLayer.GetLayerDIB
    
    'Un-premultiply both images before comparing
    baseDIB.SetAlphaPremultiplication False
    cmpDIB.SetAlphaPremultiplication False
    
    'The pdLUT3D class handles pretty much everything from here.
    
    'Start by constructing a LUT that describes all changes to the current layer (this is the longest part to process)
    Dim cExport As pdLUT3D
    Set cExport = New pdLUT3D
    If cExport.BuildLUTFromTwoDIBs(baseDIB, cmpDIB, numGridPoints, True) Then
        
        Message "Saving file..."
        ProgressBars.SetProgBarVal ProgressBars.GetProgBarMax
        
        'PD can embed copyright strings in LUT files, but for LUTs auto-generated like this, it doesn't
        ' make sense to assign copyright.  Place the user's custom description (if any) into the file,
        ' but explicitly mark the final LUT contents as public domain.
        Dim strCopyright As String
        strCopyright = "freely waived; this file is released into the public domain under the Unlicense (https://unlicense.org/)"
        
        'If the target file already exists, use "safe" file saving (e.g. write the save data to a new file,
        ' and if it's saved successfully, overwrite the original file - this way, if an error occurs mid-save,
        ' the original file remains untouched).
        Dim tmpFilename As String
        If Files.FileExists(dstFilename) Then
            Do
                tmpFilename = dstFilename & Hex$(PDMath.GetCompletelyRandomInt()) & ".pdtmp"
            Loop While Files.FileExists(tmpFilename)
        Else
            tmpFilename = dstFilename
        End If
        
        'Export said LUT to desired format
        Dim saveOK As Boolean
        Select Case dstLutFormat
            Case "cube"
                saveOK = cExport.SaveLUTToFile_Cube(tmpFilename, strCopyright, dstDescription)
            Case "look"
                saveOK = cExport.SaveLUTToFile_look(tmpFilename, strCopyright, dstDescription)
            Case "3dl"
                saveOK = cExport.SaveLUTToFile_3dl(tmpFilename, strCopyright, dstDescription)
        End Select
        
        'If the original file already existed, attempt to replace it now
        If saveOK And Strings.StringsNotEqual(dstFilename, tmpFilename) Then
            saveOK = (Files.FileReplace(dstFilename, tmpFilename) = FPR_SUCCESS)
            If (Not saveOK) Then
                Files.FileDelete tmpFilename
                PDDebug.LogAction "WARNING!  Safe save did not overwrite original file (is it open elsewhere?)"
            End If
        End If
        
        'If the user did *not* save the LUT into PD's LUT folder, we likely want to add it to PD's Color lookup tool
        If cParams.GetBool("add-to-pd", True, True) And saveOK Then
            
            'Check folder before copying (this ensures saving to a subfolder inside the LUT folder is
            ' treated as valid)
            Dim pdLutFolder As String
            pdLutFolder = UserPrefs.GetLUTPath(True)
            If (Not Strings.StringsEqualLeft(dstFilename, pdLutFolder, True)) Then
                Files.FileCopyW dstFilename, pdLutFolder & Files.FileGetName(dstFilename, False)
            End If
            
        End If
        
        ProgressBars.ReleaseProgressBar
        Message "Save complete."
        
    'The source function cannot fail if two valid images (and a valid grid point count) are passed
    '/Else
    End If
    
    Message "Finished."
    ProgressBars.ReleaseProgressBar
    
End Sub

Private Sub cmdBar_OKClick()
    
    'Before exiting, we need the user to supply a filename for the exported LUT.
    ' (If they cancel the common dialog, let them return to this dialog.)
    Dim dstFilename As String, dstLutFormat As String
    If (Not PromptForFilename(dstFilename, dstLutFormat)) Then
        cmdBar.DoNotUnloadForm
        Exit Sub
    End If
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "source-image-id", m_OpenImageIDs.GetInt(ddSource(0).ListIndex)
        .AddParam "source-layer-idx", (ddSource(1).ListCount - 1) - ddSource(1).ListIndex   'Layers are displayed in visual order
        .AddParam "compare-image-id", m_OpenImageIDs.GetInt(ddSource(2).ListIndex)
        .AddParam "compare-layer-idx", (ddSource(3).ListCount - 1) - ddSource(3).ListIndex   'Layers are displayed in visual order
        .AddParam "filename", dstFilename, True
        .AddParam "lut-format", dstLutFormat, True
        .AddParam "description", txtDescription.Text, True
        .AddParam "grid-points", sldGridPoints.Value, True
        .AddParam "add-to-pd", chkAvailable.Value, True
    End With
    
    'Save all last-used settings to file
    If Not (lastUsedSettings Is Nothing) Then
        lastUsedSettings.SaveAllControlValues
        lastUsedSettings.SetParentForm Nothing
    End If
    
    Me.Visible = False
    
    'Note that this feature does *not* generate a new Undo/Redo node, by design
    Process "Create color lookup", , cParams.GetParamString(), UNDO_Nothing
    
End Sub

Private Function PromptForFilename(ByRef dstFilename As String, ByRef dstLutFormat As String) As Boolean

    'Determine an initial folder.  This is easy - just grab the last "3dlut" path from the preferences file.
    Dim initialSaveFolder As String
    initialSaveFolder = UserPrefs.GetLUTPath()
    
    'Build common dialog filter lists
    Dim cdFilter As pdString, cdFilterExtensions As pdString
    Set cdFilter = New pdString
    Set cdFilterExtensions = New pdString
    
    cdFilter.Append "Adobe / IRIDAS (.cube)|*.cube|"
    cdFilterExtensions.Append "cube|"
    cdFilter.Append "Adobe SpeedGrade (.look)|*.look|"
    cdFilterExtensions.Append "look|"
    cdFilter.Append "Autodesk Lustre (.3dl)|*.3dl"
    cdFilterExtensions.Append "3dl"
    
    'Default to cube pending further testing (note common-dialog indices are 1-based)
    Dim cdIndex As Long
    cdIndex = UserPrefs.GetPref_Long("Dialogs", "lut-cdlg-index", 1)
    If (cdIndex < 1) Or (cdIndex > 3) Then cdIndex = 1
    
    'Suggest a file name.  At present, we just reuse the current image's name.
    dstFilename = PDImages.GetActiveImage.ImgStorage.GetEntry_String("OriginalFileName", vbNullString)
    If (LenB(dstFilename) = 0) Then dstFilename = g_Language.TranslateMessage("Color lookup")
    dstFilename = initialSaveFolder & dstFilename
    
    Dim cdTitle As String
    cdTitle = g_Language.TranslateMessage("Export color lookup")
    
    'Prep a common dialog interface
    Dim saveDialog As pdOpenSaveDialog
    Set saveDialog = New pdOpenSaveDialog
    
    PromptForFilename = saveDialog.GetSaveFileName(dstFilename, , True, cdFilter.ToString(), cdIndex, UserPrefs.GetLUTPath, cdTitle, cdFilterExtensions.ToString(), Me.hWnd)
    If PromptForFilename Then
        
        'Update preferences
        UserPrefs.SetLUTPath Files.FileGetPath(dstFilename)
        UserPrefs.SetPref_Long "Dialogs", "lut-cdlg-index", cdIndex
        
        'Convert common-dialog index into a human-readable string
        Select Case cdIndex
            Case 1
                dstLutFormat = "cube"
            Case 2
                dstLutFormat = "look"
            Case 3
                dstLutFormat = "3dl"
        End Select
        
    End If
    
End Function

Private Sub ddQuality_Click()
    Select Case ddQuality.ListIndex
        Case 0
            sldGridPoints.Value = GRID_FAST
        Case 1
            sldGridPoints.Value = GRID_DEFAULT
        Case 2
            sldGridPoints.Value = GRID_EXTREME
        Case Else
            'Do nothing
    End Select
End Sub

Private Sub ddSource_Click(Index As Integer)
    
    Select Case Index
        
        'Base image / Comparison image
        Case 0, 2
            PopulateLayerList Index
            
    End Select
    
End Sub

'Certain actions are done at LOAD time instead of ACTIVATE time to minimize visible flickering
Private Sub Form_Load()

    'Grid size presets
    ddQuality.SetAutomaticRedraws False, False
    ddQuality.Clear
    ddQuality.AddItem "fast", 0
    ddQuality.AddItem "standard", 1
    ddQuality.AddItem "extreme", 2
    ddQuality.AddItem "custom", 3
    ddQuality.SetAutomaticRedraws True
    ddQuality.ListIndex = 1
    
    txtDescription.Text = vbNullString
    
    'Load any last-used settings for this form, but only for controls *below* the layer dropdowns
    Set lastUsedSettings = New pdLastUsedSettings
    lastUsedSettings.SetParentForm Me
    lastUsedSettings.LoadAllControlValues
    
    'Populate both drop-downs with a list of open images
    PDImages.GetListOfActiveImageIDs m_OpenImageIDs
    
    ddSource(0).SetAutomaticRedraws False
    ddSource(0).Clear
    
    ddSource(2).SetAutomaticRedraws False
    ddSource(2).Clear
    
    Dim srcLayerName As String, idxActiveImage As Long
    Dim i As Long
    
    For i = 0 To m_OpenImageIDs.GetNumOfInts - 1
        If (m_OpenImageIDs.GetInt(i) = PDImages.GetActiveImageID) Then idxActiveImage = i
        srcLayerName = Interface.GetWindowCaption(PDImages.GetImageByID(m_OpenImageIDs.GetInt(i)), False, True)
        ddSource(0).AddItem srcLayerName
        ddSource(2).AddItem srcLayerName
    Next i
    
    'Auto-select the currently active image as the modified image, and if another image is available,
    ' set it as the base object
    ddSource(2).ListIndex = idxActiveImage
    
    Dim baseIndex As Long
    If (ddSource(0).ListCount > 1) Then
        baseIndex = idxActiveImage + 1
        If (baseIndex >= ddSource(0).ListCount) Then baseIndex = idxActiveImage - 1
    Else
        baseIndex = idxActiveImage
    End If
    
    ddSource(0).ListIndex = baseIndex
    
    ddSource(0).SetAutomaticRedraws True
    ddSource(2).SetAutomaticRedraws True
    
    'Select an active layer from both drop-downs
    PopulateLayerList 0, True
    PopulateLayerList 2, True
    
    ApplyThemeAndTranslations Me
    
End Sub

Private Sub PopulateLayerList(ByVal srcDropDown As Long, Optional ByVal isInitPopulator As Boolean = False)
    
    Dim ddTarget As Long
    
    'Source image vs comparison image
    If (srcDropDown = 0) Then ddTarget = 1 Else ddTarget = 3
    
    'Clear existing layer lists
    ddSource(ddTarget).SetAutomaticRedraws False
    ddSource(ddTarget).Clear
    
    Dim srcImage As pdImage
    Set srcImage = PDImages.GetImageByID(m_OpenImageIDs.GetInt(ddSource(srcDropDown).ListIndex))
    
    'Populate with layer names (in *descending* order)
    Dim i As Long
    For i = srcImage.GetNumOfLayers - 1 To 0 Step -1
        ddSource(ddTarget).AddItem srcImage.GetLayerByIndex(i).GetLayerName()
    Next i
    
    'Auto-select the currently active layer in either image, unless the two images are identical;
    ' in that case, attempt to select a neighboring layer (if one exists)
    If (srcDropDown = 0) Then
        ddSource(ddTarget).ListIndex = (srcImage.GetNumOfLayers - 1) - srcImage.GetActiveLayerIndex
    Else
        
        'Only one image exists; try to select different layers
        If (ddSource(0).ListIndex = ddSource(2).ListIndex) Then
            
            Dim idxLayer As Long
            idxLayer = ddSource(1).ListIndex
            If (srcImage.GetNumOfLayers > 0) Then
                idxLayer = idxLayer + 1
                If (idxLayer >= ddSource(ddTarget).ListCount) Then
                    idxLayer = ddSource(1).ListIndex - 1
                    If (idxLayer < 0) Then idxLayer = 0
                End If
            End If
            
            ddSource(ddTarget).ListIndex = idxLayer
            
        'Two different images exist; use active layer from both
        Else
            ddSource(ddTarget).ListIndex = (srcImage.GetNumOfLayers - 1) - srcImage.GetActiveLayerIndex
        End If
        
    End If
    
    ddSource(ddTarget).SetAutomaticRedraws True, True
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If (Not lastUsedSettings Is Nothing) Then lastUsedSettings.SetParentForm Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub sldGridPoints_Change()
    If (sldGridPoints.Value = GRID_FAST) Then
        ddQuality.ListIndex = 0
    ElseIf (sldGridPoints.Value = GRID_DEFAULT) Then
        ddQuality.ListIndex = 1
    ElseIf (sldGridPoints.Value = GRID_EXTREME) Then
        ddQuality.ListIndex = 2
    Else
        ddQuality.ListIndex = 3
    End If
End Sub
