VERSION 5.00
Begin VB.Form FormEffects8bf 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Photoshop (8bf) plugins"
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
      Begin PhotoDemon.pdHyperlink hypAbout 
         Height          =   495
         Left            =   0
         Top             =   4590
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   873
         Alignment       =   1
         Caption         =   "show plugin's About dialog"
         RaiseClickEvent =   -1  'True
      End
      Begin PhotoDemon.pdListBox lstPlugins 
         Height          =   4455
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
      Begin PhotoDemon.pdButton cmdAddFolder 
         Height          =   615
         Left            =   7800
         TabIndex        =   3
         Top             =   4440
         Width           =   2175
         _ExtentX        =   3836
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
      Begin PhotoDemon.pdButton cmdRemoveFolder 
         Height          =   615
         Left            =   0
         TabIndex        =   4
         Top             =   4440
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
         Caption         =   "remove folder"
         Enabled         =   0   'False
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
'Copyright 2021-2021 by Tanner Helland
'Created: 08/February/21
'Last updated: 08/February/21
'Last update: initial build
'
'In v9.0, PD gained support for hosting 3rd-party 8bf (Photoshop) filter plugins.
'
'These 3rd-party filters represent a problematic workflow, since each plugin's UI is controlled by the
' plugins themselves, so PD has no notification of plugin behavior after "execute-plugin" is invoked.
' (We can sort of infer if OK is pressed if the progress callback is hit, but as you can image this
' isn't an ideal place to invoke a bunch of heavy behavior like Undo/Redo flagging.)
'
'Anyway, I mention all this because this dialog breaks a lot of PD conventions in how it handles
' interactions with various program components.  Please do not mimic this behavior elsewhere; it is
' intentionally specific to this very weird, specific use-case.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Number of available plugins (as returned by the 8bf plugin interface), and their categories and
' names (each in their own string stack).  These exist for UI purposes only.
Private m_num8bf As Long, m_8bfCategories As pdStringStack, m_8bfNames As pdStringStack

'Paths to individual 8bf files are the only things we actually need to launch plugins
Private m_8bfPaths As pdStringStack

'If we have attempted to execute a plugin, this will be set to TRUE
Private m_PluginLive As Boolean

'If a plugin is canceled (or errors out), we'll restore this dialog so the user can try another one
Private m_PluginCanceled As Boolean

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
        
        If Plugin_8bf.Load8bf(m_8bfPaths.GetString(lstPlugins.ListIndex)) Then
            
            'Attempt to queue up the current layer as the active image
            Plugin_8bf.SetImage_CurrentWorkingImage
            
            'Note that we're attempting to go live
            m_PluginLive = True
            
            'Hide this window
            Me.Visible = False
            
            Message "Waiting for plugin..."
            
            Dim wasPluginCanceled As Boolean
            If Plugin_8bf.Execute8bf(Me.hWnd, wasPluginCanceled) Then
            
                'Plugin ended successfully.
                
                'Free the plugin as it's no longer needed
                Plugin_8bf.Free8bf
                
                'Finalize the plugin results (e.g. commit the finished effect to the target layer/image)
                EffectPrep.FinalizeImageData
                
                'Submit a "fake" processor operation.  This creates an Undo point, among other tasks
                Processor.Process "Photoshop (8bf) filter", False, vbNullString, UNDO_Layer
                
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
                    PDDebug.LogAction "WARNING: Error with 8bf plugin: " & m_8bfPaths.GetString(lstPlugins.ListIndex)
                    'Error or crash
                    
                    'Consider dialog for blacklisting plugin?
                    
                End If
                
            End If
            
        Else
            'Warn the user?
        End If
    
    'No plugin selected
    End If
    
End Sub

Private Sub cmdBar_ResetClick()
    'Re-scan for plugins?
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
    
    'Retrieve any other saved folders (TODO)
    
    'Load any previously saved filter databases (TODO)
    
    'Compare filter databases to a quick enumeration of 8bf files in the target folders (TODO)
    
    'If no filter database exists, do a first-time enumeration in the default folder
    
    '(Also TODO: progress bar?  We can do a quick enum just to get a rough count of 8bf files
    ' in the folder and subfolders, and use that as our "prog bar max")
    
    'TEMPORARY CODE ONLY:
    Dim path8bf As String
    path8bf = UserPrefs.GetProgramPath & "no_sync\8bf\8bf filters\"
    
    'Enumerate plugins
    Dim numPlugins As Long
    numPlugins = Plugin_8bf.EnumerateAvailable8bf(path8bf)
    If (numPlugins > 0) Then
    
        'Sort filters alphabetically (first by category, then by filter name)
        Plugin_8bf.SortAvailable8bf
        
    'No plugins found!  An informative link or explanation would be nice...
    Else
        'TODO
    End If
    
    'If any plugins exist, retrieve their categories, names, and paths now
    numPlugins = Plugin_8bf.GetEnumerationResults(m_8bfCategories, m_8bfNames, m_8bfPaths)
    
    'Populate the list of available plugins
    If (numPlugins > 0) Then
        
        Dim addSeparator As Boolean
        
        Dim i As Long
        For i = 0 To numPlugins - 1
            If (i >= numPlugins) Then
                addSeparator = Strings.StringsNotEqual(m_8bfCategories.GetString(i), m_8bfCategories.GetString(i + 1), True)
            Else
                addSeparator = False
            End If
            lstPlugins.AddItem m_8bfCategories.GetString(i) & " > " & Replace$(m_8bfNames.GetString(i), "&&", "&"), i, addSeparator
        Next i
        
    End If
        
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
    If (lstPlugins.ListIndex >= 0) Then Plugin_8bf.ShowAboutDialog m_8bfPaths.GetString(lstPlugins.ListIndex)
End Sub

Private Sub hypPlugins_Click()
    Dim filePath As String, shellCommand As String
    filePath = UserPrefs.Get8bfPath()
    shellCommand = "explorer.exe """ & filePath & """"
    Shell shellCommand, vbNormalFocus
End Sub
