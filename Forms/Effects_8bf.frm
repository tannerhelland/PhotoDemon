VERSION 5.00
Begin VB.Form FormEffects8bf 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Photoshop (8bf) plugin"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10395
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
   ScaleHeight     =   461
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   693
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdButtonStrip btsPanel 
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   1296
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   6165
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   1323
      DontAutoUnloadParent=   -1  'True
   End
   Begin PhotoDemon.pdContainer pnlOptions 
      Height          =   5175
      Index           =   0
      Left            =   120
      Top             =   960
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   9128
      Begin PhotoDemon.pdLabel lblNoPlugins 
         Height          =   4095
         Left            =   0
         Top             =   120
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   7223
         Caption         =   ""
         FontSize        =   11
         Layout          =   1
      End
      Begin PhotoDemon.pdButton cmdRescan 
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   4440
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   1085
         Caption         =   "scan for new plugins"
      End
      Begin PhotoDemon.pdHyperlink hypAbout 
         Height          =   495
         Left            =   5280
         TabIndex        =   8
         Top             =   4560
         Visible         =   0   'False
         Width           =   4695
         _ExtentX        =   7011
         _ExtentY        =   873
         Alignment       =   1
         Caption         =   "about this plugin"
         RaiseClickEvent =   -1  'True
      End
      Begin PhotoDemon.pdListBox lstPlugins 
         Height          =   4335
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   7858
         Caption         =   "available plugins:"
      End
   End
   Begin PhotoDemon.pdContainer pnlOptions 
      Height          =   5175
      Index           =   1
      Left            =   120
      Top             =   960
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   9128
      Begin PhotoDemon.pdButton cmdFolders 
         Height          =   615
         Index           =   1
         Left            =   7080
         TabIndex        =   3
         Top             =   4440
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   1085
         Caption         =   "add folder..."
      End
      Begin PhotoDemon.pdListBox lstFolders 
         Height          =   2775
         Left            =   0
         TabIndex        =   2
         Top             =   1560
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   4895
      End
      Begin PhotoDemon.pdHyperlink hypPlugins 
         Height          =   375
         Left            =   0
         TabIndex        =   9
         Top             =   600
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   661
         Alignment       =   2
         Caption         =   ""
         RaiseClickEvent =   -1  'True
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   375
         Index           =   0
         Left            =   0
         Top             =   120
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   661
         Caption         =   "default plugin folder:"
         FontSize        =   12
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   375
         Index           =   1
         Left            =   0
         Top             =   1080
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   661
         Caption         =   "additional folders:"
         FontSize        =   12
      End
      Begin PhotoDemon.pdButton cmdFolders 
         Height          =   615
         Index           =   0
         Left            =   0
         TabIndex        =   4
         Top             =   4440
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   1085
         Caption         =   "remove folder"
         Enabled         =   0   'False
      End
   End
   Begin PhotoDemon.pdContainer pnlOptions 
      Height          =   5175
      Index           =   2
      Left            =   120
      Top             =   960
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   9128
      Begin PhotoDemon.pdProgressBar prgUpdate 
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   873
      End
      Begin PhotoDemon.pdLabel lblUpdate 
         Height          =   375
         Left            =   0
         Top             =   240
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   661
         Alignment       =   2
         Caption         =   ""
         FontSize        =   12
      End
   End
End
Attribute VB_Name = "FormEffects8bf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'8bf Plugin Interface Dialog
'Copyright 2021-2025 by Tanner Helland
'Created: 08/February/21
'Last updated: 15/December/25
'Last update: if no 8bf files exist, avoid calling pspihost to avoid crashes on some systems
'             (see https://github.com/tannerhelland/PhotoDemon/issues/716)
'
'In v9.0, PD gained support for hosting 3rd-party 8bf (Photoshop) filter plugins.
'
'These 3rd-party filters represent a problematic workflow, since each plugin's UI is controlled by the
' plugins themselves, so PD has no notification of plugin behavior after "execute-plugin" is invoked.
' (We can sort of infer if OK is pressed if the progress callback is hit, but as you can imagine this
' isn't an ideal place to invoke a bunch of heavy behavior like Undo/Redo flagging.)
'
'Anyway, I mention all this because this dialog breaks a lot of PD conventions in how it handles
' interactions with various program components.  Please do not mimic this behavior elsewhere;
' it is intentionally specific to this very weird, specific use-case.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Number of available plugins (as returned by the 8bf plugin interface), and their categories and
' names (each in their own string stack).  These exist for UI purposes only.
Private Type PD_8bf
    plgCategory As String
    plgCategoryIndex As Long
    plgName As String
    plgNameIndex As Long
    plgPath As String
End Type

Private m_numPlugins As Long, m_Plugins() As PD_8bf

'If we have attempted to execute a plugin, this will be set to TRUE
Private m_PluginLive As Boolean

'If a plugin is canceled (or errors out), we'll restore this dialog so the user can try another one
Private m_PluginCanceled As Boolean

'When the form is first activated, we start building an 8bf collection.  We wait until the form is
' active so that we can display a progress bar.
Private m_FormHasBeenActivated As Boolean

Public Function RestoreDialog() As Boolean
    RestoreDialog = m_PluginCanceled
    m_PluginCanceled = False
End Function

Private Sub btsPanel_Click(ByVal buttonIndex As Long)
    UpdatePanelVisibility
End Sub

Private Sub cmdBar_CancelClick()
    Unload Me
End Sub

'OK button
Private Sub cmdBar_OKClick()
    
    m_PluginCanceled = False
    
    'When OK is clicked, load the selected plugin, then attempt to execute it
    If (lstPlugins.ListIndex >= 0) Then
        
        Dim targetPluginPath As String
        targetPluginPath = m_Plugins(lstPlugins.ListIndex).plgPath
        
        If Plugin_8bf.Load8bf(targetPluginPath) Then
            
            'If a selection is active, retrieve a copy (in 8-bit format) and cache it locally,
            ' then notify pspi of the mask's presence
            Dim pspiMaskOK As Boolean
            If PDImages.GetActiveImage.IsSelectionActive And PDImages.GetActiveImage.MainSelection.IsLockedIn Then
                
                'DISCLAIMER: pspi allows you to pass a selection mask copy on to the target image.
                ' In my experience, however, this doesn't work great.  Many plugins don't support
                ' selection data and there's no way to know this in advance.  As such, I'm inclined
                ' to use our own internal selection masking engine for now, as it provides much more
                ' predictable results.
                'pspiMaskOK = Plugin_8bf.SetMask_CurrentSelectionMask
                
            End If
            
            'Attempt to queue up the current layer as the active image
            Plugin_8bf.SetImage_CurrentWorkingImage pspiMaskOK
            
            'Note that we're attempting to go live
            m_PluginLive = True
            
            'Hide this window
            Me.Visible = False
            
            Message "Waiting for plugin..."
            
            Dim wasPluginCanceled As Boolean
            If Plugin_8bf.Execute8bf(Me.hWnd, wasPluginCanceled) Then
            
                'Plugin ended successfully.
                
                'Finalize the plugin results (e.g. commit the finished effect to the target layer/image)
                EffectPrep.FinalizeImageData ignoreSelection:=pspiMaskOK
                
                'Submit a "fake" processor operation.  This creates an Undo point, among other tasks
                Processor.Process "Photoshop (8bf) plugin", False, vbNullString, UNDO_Layer
                
                'The fake processor call, above, will report faulty timing reports (because it only
                ' tracks the time of the processor function, which is just a dummy call here).
                ' Report time taken manually.
                If g_DisplayTimingReports Then Processor.ReportProcessorTimeTaken Plugin_8bf.GetInitialEffectTimestamp()
                
                'Unload this dialog
                Unload Me
                
            'Plugin may have failed
            Else
                
                m_PluginCanceled = True
                Message "Plugin canceled."
                
                If wasPluginCanceled Then
                    'Fine
                Else
                    PDDebug.LogAction "WARNING: Error with 8bf plugin: " & targetPluginPath
                    'Error or crash
                    
                    'Consider dialog for blacklisting plugin?
                    
                End If
                
            End If
            
            'Free any remaining image and/or plugin resources
            Plugin_8bf.FreeImageResources   'Pointers to our images and/or internal 8bf image structs
            Plugin_8bf.Free8bf              'Plugin itself
            
        Else
            'Warn the user?
        End If
    
    'No plugin selected
    End If
    
End Sub

Private Sub cmdBar_ResetClick()
    'Re-scan for plugins?
End Sub

Private Sub cmdFolders_Click(Index As Integer)
    
    Dim numFoldersAtStart As Long
    numFoldersAtStart = lstFolders.ListCount
    
    'Remove selected folder
    If (Index = 0) Then
        
        If (lstFolders.ListIndex >= 0) Then lstFolders.RemoveItem lstFolders.ListIndex
        cmdFolders(0).Enabled = (lstFolders.ListCount > 0) And (lstFolders.ListIndex >= 0)
        
    'Add a new folder
    ElseIf (Index = 1) Then
        
        Dim initFolder As String
        If (lstFolders.ListIndex >= 0) Then
            initFolder = lstFolders.List(lstFolders.ListIndex)
        Else
            initFolder = UserPrefs.Get8bfPath()
        End If
        
        Dim newFolder As String
        newFolder = Files.PathBrowseDialog(Me.hWnd, initFolder)
        If (LenB(newFolder) <> 0) Then
            If Files.PathExists(newFolder, False) Then lstFolders.AddItem newFolder
        End If
        
    End If
    
    'Always update the saved folder list after changes are made
    If (numFoldersAtStart <> lstFolders.ListCount) Then UpdateSavedFolderList
    
End Sub

Private Sub cmdRescan_Click()
    ScanForPlugins
End Sub

Private Sub Form_Activate()
    
    'On first activation, load the initial plugin collection
    If (Not m_FormHasBeenActivated) Then ScanForPlugins
    m_FormHasBeenActivated = True
    
    'Always default to the plugin collection (not the settings page)
    btsPanel.ListIndex = 0
    
End Sub

'Certain actions are done at LOAD time instead of ACTIVATE time to minimize visible flickering
Private Sub Form_Load()
    
    btsPanel.AddItem "plugins", 0
    btsPanel.AddItem "settings", 1
    btsPanel.ListIndex = 0
    UpdatePanelVisibility
    
    'Retrieve the user's default plugin folder:
    hypPlugins.Caption = UserPrefs.Get8bfPath()
    hypPlugins.AssignTooltip "click to open this folder in Windows Explorer"
    
    'Load any UI resources
    Dim btnImgSize As Long
    btnImgSize = Interface.FixDPI(24)
    cmdFolders(0).AssignImage "file_close", Nothing, btnImgSize, btnImgSize
    cmdFolders(1).AssignImage "generic_add", Nothing, btnImgSize, btnImgSize
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me
    
End Sub

Private Sub UpdatePanelVisibility()
    Dim i As Long
    For i = 0 To btsPanel.ListCount - 1
        pnlOptions(i).Visible = (i = btsPanel.ListIndex)
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub hypAbout_Click()
    Dim targetPluginPath As String
    targetPluginPath = m_Plugins(lstPlugins.ListIndex).plgPath
    If (lstPlugins.ListIndex >= 0) Then Plugin_8bf.ShowAboutDialog targetPluginPath
End Sub

Private Sub hypPlugins_Click()
    Dim filePath As String, shellCommand As String
    filePath = UserPrefs.Get8bfPath()
    shellCommand = "explorer.exe """ & filePath & """"
    Shell shellCommand, vbNormalFocus
End Sub

Private Sub lstFolders_Click()
    cmdFolders(0).Enabled = (lstFolders.ListIndex >= 0)
End Sub

Private Sub lstPlugins_Click()
    hypAbout.Visible = (lstPlugins.ListIndex >= 0)
End Sub

Private Sub ScanForPlugins()
    
    'Clear all existing plugin collections
    lstPlugins.Clear
    Plugin_8bf.ResetPluginCollection
    
    'Switch the UI to "loading" mode
    lblUpdate.Caption = g_Language.TranslateMessage("loading plugin collection...")
    lblUpdate.RequestRefresh
    
    Dim i As Long
    For i = 0 To pnlOptions.Count - 1
        pnlOptions(i).Visible = (i = 2)
    Next i
    
    'Next, we need to build an initial 8bf collection.
        
    'Load any previously saved filter databases (TODO)
    
    'Next, we need to build a list of folders to search for 8bf files.  This consists of PD's
    ' default 8bf folder (always searched), and whatever other folders the user has specified.
    Dim listOfFolders As pdStringStack
    Set listOfFolders = New pdStringStack
    If Files.PathExists(UserPrefs.Get8bfPath(), False) Then listOfFolders.AddString UserPrefs.Get8bfPath()
    
    'Retrieve any previously added folders as well
    RetrieveSavedFolderList
    
    If (lstFolders.ListCount > 0) Then
        For i = 0 To lstFolders.ListCount - 1
            If Files.PathExists(lstFolders.List(i), False) Then listOfFolders.AddString lstFolders.List(i)
        Next i
    End If
    
    'Next, we want to get a quick count of how many 8bf files exist in the target folder(s).
    ' This gives us a useful max value for our scan progress bar.
    Dim listOfFiles As pdStringStack
    For i = 0 To listOfFolders.GetNumOfStrings - 1
        Files.RetrieveAllFiles listOfFolders.GetString(i), listOfFiles, True, False, "8bf"
    Next i
    
    'We now have a (rough) estimate of how many 8bf files we expect to see in the final result.
    ' Note that not all of these may be useable for reasons outside our control (e.g. 64-bit, etc).
    
    'Set the progress bar maximum to the number of 8bf files found
    Dim num8bfCandidates As Long
    num8bfCandidates = listOfFiles.GetNumOfStrings()
    prgUpdate.Max = num8bfCandidates
    
    'Compare filter databases to a quick enumeration of 8bf files in the target folders (TODO).
    ' The goal here is to skip manual enumeration of DLLs we've already seen in the past.
    
    'If no filter database exists, do a first-time enumeration in the default folder.
    
    'UPDATE DEC 2025: previously, I handed off all folders to pspihost here and let it do its thing.
    ' But per https://github.com/tannerhelland/PhotoDemon/issues/716, some users are seeing crashes
    ' even when *zero* 8bf files exist in the target folder.
    '
    'pspihost's code is thorny and apparently no longer maintained, so rather than mess with it,
    ' I'm just going to bypass it completely if no candidate 8bf files exist.
    Dim numPlugins As Long
    If (num8bfCandidates > 0) Then
        
        'DEC 2025: test our own iterator!  Maybe someday it will be robust enough that we don't need
        ' pspihost at all...
        Plugin_8bf.EnumeratePlugins_PD listOfFiles
        
        'Enumerate plugins in all target folder(s)
        For i = 0 To listOfFolders.GetNumOfStrings - 1
            numPlugins = numPlugins + Plugin_8bf.EnumerateAvailable8bf(listOfFolders.GetString(i), prgUpdate)
        Next i
        
        'Sort filters alphabetically (first by category, then by filter name)
        If (numPlugins > 0) Then
            Plugin_8bf.SortAvailable8bf
            
        'No plugins found!  An informative link or explanation would be nice...
        Else
            'TODO
        End If
        
    Else
        numPlugins = 0
    End If
    
    'If any plugins exist, retrieve their categories, names, and paths now
    Dim cat8bf As pdStringStack, name8bf As pdStringStack, path8bf As pdStringStack
    
    If (numPlugins > 0) Then
        m_numPlugins = Plugin_8bf.GetEnumerationResults(cat8bf, name8bf, path8bf)
    Else
        m_numPlugins = 0
    End If
    
    'Prep our internal plugin table to match
    If (m_numPlugins < 0) Then m_numPlugins = 0
    ReDim m_Plugins(0 To m_numPlugins) As PD_8bf
    
    If (m_numPlugins > 0) Then
        For i = 0 To m_numPlugins - 1
            m_Plugins(i).plgCategory = cat8bf.GetString(i)
            m_Plugins(i).plgName = name8bf.GetString(i)
            m_Plugins(i).plgPath = path8bf.GetString(i)
        Next i
    End If
    
    'Populate the list of available plugins
    If (m_numPlugins > 0) Then
        
        Dim addSeparator As Boolean
        
        For i = 0 To numPlugins - 1
            If (i < numPlugins - 1) Then
                addSeparator = Strings.StringsNotEqual(m_Plugins(i).plgCategory, m_Plugins(i + 1).plgCategory, True)
            Else
                addSeparator = False
            End If
            lstPlugins.AddItem m_Plugins(i).plgCategory & " > " & Replace$(m_Plugins(i).plgName, "&&", "&"), i, addSeparator
        Next i
        
    End If
    
    'Regardless of plugin count, hide the "loading" panel and restore the deafult one
    pnlOptions(2).Visible = False
    UpdatePanelVisibility
    lstPlugins.Visible = (m_numPlugins > 0)
    lblNoPlugins.Visible = (m_numPlugins <= 0)
    
    'If no plugins were found, hide the plugin selector and give the user info on how to proceed
    If (m_numPlugins <= 0) Then
    
        lstPlugins.Visible = False
        
        Dim fullCaption As pdString
        Set fullCaption = New pdString
        fullCaption.AppendLine g_Language.TranslateMessage("No plugins found.")
        fullCaption.AppendLineBreak
        fullCaption.AppendLine g_Language.TranslateMessage("Photoshop (8bf) plugins are files with an ""8bf"" extension.  These plugins provide new image effects.  Thousands of 8bf plugins are available online.")
        fullCaption.AppendLineBreak
        fullCaption.AppendLine g_Language.TranslateMessage("PhotoDemon does not ship with 8bf plugins, but if you find plugins online, you can download them and add them to PhotoDemon.  (PhotoDemon supports most 32-bit 8bf plugins.  64-bit plugins are not supported.)")
        fullCaption.AppendLineBreak
        fullCaption.AppendLine g_Language.TranslateMessage("After downloading one or more 8bf files, use the settings button (above) to tell PhotoDemon where to find them.  PhotoDemon will then add them to your Effects collection.")
        
        lblNoPlugins.Caption = fullCaption.ToString()
        
    End If
    
End Sub

'Retrieve the list of previously saved folders from file, and populate the folder list accordingly
Private Sub RetrieveSavedFolderList()
    
    lstFolders.Clear
    
    'See if a saved folder list even exists
    Dim srcFile As String
    srcFile = UserPrefs.GetPresetPath() & "8bfPaths.xml"
    
    If Files.FileExists(srcFile) Then
        
        Dim cStream As pdStream
        Set cStream = New pdStream
        If cStream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, srcFile) Then
            
            Dim cSerialize As pdSerialize
            Set cSerialize = New pdSerialize
            With cSerialize
                
                .SetParamString cStream.ReadString_UTF8(0, True)
                
                Dim numFolders As Long
                numFolders = .GetLong("num-paths")
                
                If (numFolders > 0) Then
                    
                    Dim i As Long, srcPath As String
                    For i = 0 To numFolders - 1
                        srcPath = .GetString("path-" & i)
                        If (LenB(srcPath) > 0) Then lstFolders.AddItem srcPath
                    Next i
                    
                End If
                
            End With
    
        End If
        
    End If
    
End Sub

'Save the user's list of saved folders to file
Private Sub UpdateSavedFolderList()
    
    'We could probably use a more interesting system, but for now, we just dump all custom folders
    ' to a standard PD serialization class.
    Dim dstFile As String
    dstFile = UserPrefs.GetPresetPath() & "8bfPaths.xml"
    
    'Kill the file if it exists
    Files.FileDeleteIfExists dstFile
    
    'Create a new file
    Dim cSerialize As pdSerialize
    Set cSerialize = New pdSerialize
    With cSerialize
        .AddParam "num-paths", lstFolders.ListCount
        If (lstFolders.ListCount > 0) Then
            Dim i As Long
            For i = 0 To lstFolders.ListCount - 1
                .AddParam "path-" & i, lstFolders.List(i)
            Next i
        End If
    End With
    
    Dim cStream As pdStream
    Set cStream = New pdStream
    If cStream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadWrite, dstFile) Then
        cStream.WriteString_UTF8 cSerialize.GetParamString(), True
        cStream.StopStream
    Else
        PDDebug.LogAction "WARNING: couldn't save 8bf paths to file"
    End If
    
End Sub
