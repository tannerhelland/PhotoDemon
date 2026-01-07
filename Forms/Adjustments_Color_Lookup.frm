VERSION 5.00
Begin VB.Form FormColorLookup 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Color lookup"
   ClientHeight    =   7080
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12120
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
   ScaleHeight     =   472
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   808
   Begin PhotoDemon.pdListBox lstLUTs 
      Height          =   3135
      Left            =   6000
      TabIndex        =   5
      Top             =   840
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   5530
      Caption         =   "available LUTs"
   End
   Begin PhotoDemon.pdButton cmdBrowse 
      Height          =   615
      Left            =   6000
      TabIndex        =   4
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1085
      Caption         =   "import LUT file..."
   End
   Begin PhotoDemon.pdDropDown cboBlendMode 
      Height          =   855
      Left            =   6000
      TabIndex        =   3
      Top             =   4920
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1508
      Caption         =   "blend mode"
   End
   Begin PhotoDemon.pdSlider sldIntensity 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   4080
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1244
      Caption         =   "intensity"
      Max             =   100
      Value           =   100
      NotchPosition   =   2
      NotchValueCustom=   100
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   6330
      Width           =   12120
      _ExtentX        =   21378
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdHyperlink lblCollection 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   5910
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   450
      Alignment       =   2
      Caption         =   ""
      RaiseClickEvent =   -1  'True
   End
End
Attribute VB_Name = "FormColorLookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'3D color lookup effect
'Copyright 2020-2026 by Tanner Helland
'Created: 27/October/20
'Last updated: 22/June/22
'Last update: correctly handle importing a LUT with the same name as an existing LUT
'
'For a detailed explanation of 3D color lookup tables and how they work, see the pdLUT3D class.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'LUTs are very expensive to parse.  As such, if the user previews a LUT, we serialize its data
' (in binary format) and store that data locally.  If the user returns to that preview later,
' we can produce a much-faster preview since we can just grab the serialized data.
Private Type LUTCache
    fullPath As String
    filenameOnly As String
    lenDataCompressed As Long
    lenDataUncompressed As Long
    cmpFormat As PD_CompressionFormat
    cachedData() As Byte
End Type

Private m_numOfLUTs As Long
Private m_LUTs() As LUTCache

'To improve performance, we cache LUT data after loading it; this makes subsequent previews *much* faster
Private m_LastLUTIndex As Long

'This class handles all actual LUT duties
Private WithEvents m_LUT As pdLUT3D
Attribute m_LUT.VB_VarHelpID = -1

'To improve preview performance, a persistent preview DIB is cached locally
Private m_EffectDIB As pdDIB

'Apply a hazy, cool color transformation I call an "atmospheric" transform.
Public Sub ApplyColorLookupEffect(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    If (m_LUT Is Nothing) Then Exit Sub
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim fxIntensity As Double, fxPath As String, fxBlend As PD_BlendMode
    
    With cParams
        fxBlend = .GetLong("blendmode", BM_Normal)
        fxPath = .GetString("lut-file", vbNullString)
        fxIntensity = .GetDouble("intensity", 50#)
    End With
    
    If (Not toPreview) Then Message "Applying color lookup..."
    
    'Initialize the effect engine
    Dim tmpSA As SafeArray2D
    EffectPrep.PrepImageData tmpSA, toPreview, dstPic, doNotUnPremultiplyAlpha:=True
    
    If Files.FileExists(fxPath) Then
    
        'For non-previews, set up the progress bar.  (Note that we have to use an integer value,
        ' or taskbar progress updates won't work - this is specifically an OS limitation, as PD's
        ' internal progress bar works just fine with [0, 1] progress values.)
        If (Not toPreview) Then ProgressBars.SetProgBarMax 100#
        
        'Create a copy of the working data
        If (m_EffectDIB Is Nothing) Then Set m_EffectDIB = New pdDIB
        m_EffectDIB.CreateFromExistingDIB workingDIB
        
        'Ensure the requested LUT has been loaded
        If (m_LUT Is Nothing) Then Set m_LUT = New pdLUT3D
        If (m_LUT.GetLUTPath <> fxPath) Then
            If (Not m_LUT.LoadLUTFromFile(fxPath)) Then Exit Sub
        End If
        
        'Apply the LUT to our working DIB copy
        m_LUT.ApplyLUTToDIB m_EffectDIB, (Not toPreview)
        
        'If the effect wasn't canceled, merge the result onto the original working DIB
        ' at the specified strength and blend mode
        If (Not g_cancelCurrentAction) Then
        
            If (Not m_EffectDIB.GetAlphaPremultiplication()) Then m_EffectDIB.SetAlphaPremultiplication True
            
            Dim cCompositor As pdCompositor
            Set cCompositor = New pdCompositor
            cCompositor.QuickMergeTwoDibsOfEqualSize workingDIB, m_EffectDIB, fxBlend, fxIntensity
            
        End If
        
    End If
    
    'Finalize result
    EffectPrep.FinalizeImageData toPreview, dstPic, True
    
End Sub

Private Sub cboBlendMode_Click()
    UpdatePreview
End Sub

Private Sub cmdBar_OKClick()
    Process "Color lookup", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBrowse_Click()
    
    Dim srcFilenames As pdStringStack
    If m_LUT.DisplayLUTLoadDialog(vbNullString, srcFilenames) Then
        
        'We now want to perform the following tasks:
        ' 1) Validate each LUT
        ' 2) If a LUT validates, add it to the list box
        ' 3) Preview the last LUT in the list
        Dim j As Long
        For j = 0 To srcFilenames.GetNumOfStrings - 1
            
            Dim srcFilename As String
            srcFilename = srcFilenames.GetString(j)
            
            'Don't load this LUT at all unless it validates
            If m_LUT.LoadLUTFromFile(srcFilename) Then
                
                'Before doing anything else, copy the source file into the user's LUT folder.
                If (Not Strings.LeftMatches(srcFilename, Files.PathAddBackslash(UserPrefs.GetLUTPath(True)), True)) Then
                    Dim newFilename As String
                    newFilename = Files.PathAddBackslash(UserPrefs.GetLUTPath(True)) & Files.FileGetName(srcFilename, False)
                    If Files.FileExists(newFilename) Then
                        Files.FileReplace newFilename, srcFilename
                    Else
                        Files.FileCopyW srcFilename, newFilename
                    End If
                    srcFilename = newFilename
                End If
                
                'If the file validated, and it has a sub-file, copy the sub-file into the local folder as well
                If (LenB(m_LUT.GetLUTSubPath()) <> 0) Then
                    newFilename = Files.PathAddBackslash(UserPrefs.GetLUTPath(True)) & Files.FileGetName(m_LUT.GetLUTSubPath(), False)
                    If Files.FileExists(newFilename) Then
                        Files.FileReplace newFilename, m_LUT.GetLUTSubPath()
                    Else
                        Files.FileCopyW m_LUT.GetLUTSubPath(), newFilename
                    End If
                End If
                
                'Ensure we have sufficient space for a new entry
                If (m_numOfLUTs <= 0) Then
                    ReDim m_LUTs(0) As LUTCache
                Else
                    If (m_numOfLUTs > UBound(m_LUTs)) Then ReDim Preserve m_LUTs(0 To m_numOfLUTs * 2 - 1) As LUTCache
                End If
                
                Dim targetFilenameOnly As String
                targetFilenameOnly = Files.FileGetName(srcFilename)
                
                'We now need to search the current list of LUTs and figure out where to insert this one.
                Dim i As Long
                i = 0
                
                '(We also need to check to see if a LUT with this name already exists; if it does, we simply
                ' want to update it in-place instead of adding it again.)
                Dim lutAlreadyExists As Boolean: lutAlreadyExists = False
                
                If (m_numOfLUTs > 0) Then
                    
                    'Look for the first entry with a value *less than* or *equal to* the current one
                    For i = 0 To m_numOfLUTs - 1
                        If (Strings.StrCompSortPtr_Filenames(StrPtr(targetFilenameOnly), StrPtr(m_LUTs(i).filenameOnly)) <= 0) Then Exit For
                    Next i
                    
                    'We should check for the entry already existing (maybe the user hand-edited the file,
                    ' or is downloading a new copy from elsewhere)
                    lutAlreadyExists = Strings.StringsEqual(targetFilenameOnly, m_LUTs(i).filenameOnly, True)
                    
                    'If this entry is novel, shift all existing entries to make room for it
                    If (Not lutAlreadyExists) And (i < m_numOfLUTs - 1) Then
                        Dim k As Long
                        For k = m_numOfLUTs - 1 To i + 1 Step -1
                            m_LUTs(k) = m_LUTs(k - 1)
                        Next k
                    End If
                    
                End If
                
                'If this lut already exists, just free its internal cache to force a re-load
                If lutAlreadyExists Then
                    
                    With m_LUTs(i)
                        .fullPath = srcFilename
                        .lenDataCompressed = 0
                        .lenDataUncompressed = 0
                    End With
            
                'Update the backing collection
                Else
                    
                    With m_LUTs(i)
                        .fullPath = srcFilename
                        .filenameOnly = targetFilenameOnly
                        .lenDataCompressed = 0
                        .lenDataUncompressed = 0
                    End With
                    m_numOfLUTs = m_numOfLUTs + 1
                        
                    'Update the listbox too (to ensure it stays synced to the backing collection)
                    lstLUTs.AddItem targetFilenameOnly, i
                    
                End If
                    
                '(Only make this LUT the active one if we're finished processing LUT files.)
                If (j = srcFilenames.GetNumOfStrings - 1) Then lstLUTs.ListIndex = i
                
            'LUT didn't validate!  Notify the user.
            Else
                Dim msgText As String, msgTitle As String
                msgText = g_Language.TranslateMessage("%1 is not a valid 3D lookup table (LUT).", srcFilename)
                msgTitle = g_Language.TranslateMessage("Error")
                PDMsgBox msgText, vbExclamation Or vbOKOnly Or vbApplicationModal, msgTitle
            End If
            
        Next j
            
    End If
    
End Sub

Private Sub Form_Load()
    
    cmdBar.SetPreviewStatus False
    
    'Make sure we have a 3DLUT object to work with
    m_LastLUTIndex = -1
    Set m_LUT = New pdLUT3D
    
    'Add all available LUTs to the list box
    Dim listOfFiles As pdStringStack
    
    Dim lutFolder As String
    lutFolder = UserPrefs.GetLUTPath(True)
    
    If Files.RetrieveAllFiles(lutFolder, listOfFiles, True, False, "cube|look|3dl") Then
        
        'Tell the listbox to use file display mode.
        ' (This will truncate extremely long filenames using OS rules, as necessary.)
        lstLUTs.SetDisplayMode_Files True
        
        'Prep the LUT collection
        Const INIT_LUT_SIZE As Long = 16
        ReDim m_LUTs(0 To INIT_LUT_SIZE - 1) As LUTCache
        m_numOfLUTs = 0
        
        'List all files from the retrieval list
        Dim i As Long, testFile As String, testFilenameOnly As String
        For i = 0 To listOfFiles.GetNumOfStrings - 1
            
            testFile = listOfFiles.GetString(i)
            testFilenameOnly = Files.FileGetName(testFile)
            
            If Files.FileExists(testFile) Then
            
                If (m_numOfLUTs > UBound(m_LUTs)) Then ReDim Preserve m_LUTs(0 To m_numOfLUTs * 2 - 1) As LUTCache
                With m_LUTs(m_numOfLUTs)
                    .fullPath = testFile
                    .filenameOnly = testFilenameOnly
                    .lenDataCompressed = 0
                    .lenDataUncompressed = 0
                End With
                m_numOfLUTs = m_numOfLUTs + 1
                
            End If
            
        Next i
        
        Dim startTime As Currency
        VBHacks.GetHighResTime startTime
        
        'With the list constructed, do a quick insertion sort in case subfolders are used.
        ' (pdStringStack has built-in search functionality, but only for whole strings - and we need to manually
        ' compare filenames *only*, since this list of files may not all come from the same base folder.)
        ' Fortunately, this tends to be fast as Windows file results are often returned semi-sorted as-is.
        Dim j As Long
        Dim tmpSort As LUTCache, searchCont As Boolean
        i = 1
        
        Do While (i < m_numOfLUTs)
        
            tmpSort = m_LUTs(i)
            j = i - 1
            
            'Because VB6 doesn't short-circuit And statements, we have to split this check into separate parts.
            searchCont = False
            If (j >= 0) Then searchCont = (Strings.StrCompSortPtr_Filenames(StrPtr(m_LUTs(j).filenameOnly), StrPtr(tmpSort.filenameOnly)) > 0)
            
            Do While searchCont
                m_LUTs(j + 1) = m_LUTs(j)
                j = j - 1
                searchCont = False
                If (j >= 0) Then searchCont = (Strings.StrCompSortPtr_Filenames(StrPtr(m_LUTs(j).filenameOnly), StrPtr(tmpSort.filenameOnly)) > 0)
            Loop
            
            m_LUTs(j + 1) = tmpSort
            i = i + 1
            
        Loop
        
        PDDebug.LogAction "FYI - sorting LUTs took " & VBHacks.GetTimeDiffNowAsString(startTime)
        
        'Finally, add all entries to the listbox
        lstLUTs.SetAutomaticRedraws False
        For i = 0 To m_numOfLUTs - 1
            lstLUTs.AddItem m_LUTs(i).filenameOnly
        Next i
        lstLUTs.SetAutomaticRedraws True, True
        
    End If
    
    lstLUTs.ListIndex = 0
    
    'Let the user know where their LUT collection resides
    lblCollection.Caption = g_Language.TranslateMessage("Your LUT collection is stored in the ""%1"" folder", UserPrefs.GetLUTPath(True))
    lblCollection.AssignTooltip "click to open this folder in Windows Explorer"
    
    Interface.PopulateBlendModeDropDown cboBlendMode, BM_Normal
    
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub lblCollection_Click()
    Dim filePath As String, shellCommand As String
    filePath = UserPrefs.GetLUTPath(True)
    shellCommand = "explorer.exe """ & filePath & """"
    Shell shellCommand, vbNormalFocus
End Sub

Private Sub lstLUTs_Click()
    UpdatePreview
End Sub

Private Sub m_LUT_ProgressUpdate(ByVal progressValue As Single, cancelOperation As Boolean)

    'Note: because PD uses a high-performance mechanism for updating the progress bar during
    ' long-running events, it's critical that you query for ESC keypresses *BEFORE* attempting
    ' to modify the live progress bar.  (If you do this in reverse order, the progress bar
    ' update may eat any keypress messages.)
    cancelOperation = Interface.UserPressedESC()
    ProgressBars.SetProgBarVal progressValue * 100!
    
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

'Update the preview whenever the combination slider/text control has its value changed
Private Sub sldIntensity_Change()
    UpdatePreview
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then Me.ApplyColorLookupEffect GetLocalParamString(), True, pdFxPreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim curTime As Currency
    VBHacks.GetHighResTime curTime
    
    'Cache the current LUT data, if any exists
    If (m_LastLUTIndex >= 0) And (m_LastLUTIndex < m_numOfLUTs) And (m_LastLUTIndex <> lstLUTs.ListIndex) Then
        
        'Ensure we haven't already cached this LUT
        If (m_LUTs(m_LastLUTIndex).lenDataCompressed = 0) And (m_LUT.GetLUTPath = m_LUTs(m_LastLUTIndex).fullPath) Then
            m_LUTs(m_LastLUTIndex).cmpFormat = cf_Zstd
            m_LUTs(m_LastLUTIndex).lenDataCompressed = m_LUT.Serialize_ToBytes(m_LUTs(m_LastLUTIndex).cachedData, m_LUTs(m_LastLUTIndex).lenDataUncompressed, m_LUTs(m_LastLUTIndex).cmpFormat)
        End If
        
        'PDDebug.LogAction "LUT data cached in " & VBHacks.GetTimeDiffNowAsString(curTime)
        VBHacks.GetHighResTime curTime
        
    End If
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        If (lstLUTs.ListIndex >= 0) Then .AddParam "lut-file", m_LUTs(lstLUTs.ListIndex).fullPath Else .AddParam "lut-file", vbNullString
        .AddParam "intensity", sldIntensity.Value
        .AddParam "blendmode", cboBlendMode.ListIndex
    End With
    
    'Before returning, see if we have cached data for the requested LUT
    If (lstLUTs.ListIndex >= 0) Then
        If (m_LUTs(lstLUTs.ListIndex).lenDataCompressed <> 0) And (lstLUTs.ListIndex <> m_LastLUTIndex) Then
            VBHacks.GetHighResTime curTime
            With m_LUTs(lstLUTs.ListIndex)
                m_LUT.Serialize_FromBytes .cachedData, .lenDataCompressed, .lenDataUncompressed, .cmpFormat
            End With
        End If
    End If
    
    'Remember the current lut index in case we need to cache *it* next
    m_LastLUTIndex = lstLUTs.ListIndex
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
