VERSION 5.00
Begin VB.Form FormHotkeys 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Keyboard shortcuts"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11775
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
   ScaleHeight     =   513
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   785
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdButton cmdAll 
      Height          =   615
      Index           =   0
      Left            =   6240
      TabIndex        =   9
      Top             =   4740
      Width           =   2655
      _ExtentX        =   9551
      _ExtentY        =   1085
      Caption         =   "undo all changes"
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   375
      Index           =   1
      Left            =   6120
      Top             =   4320
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   661
      Caption         =   "all hotkeys"
      FontSize        =   12
   End
   Begin PhotoDemon.pdButton cmdThisHotkey 
      Height          =   615
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   6240
      Width           =   1815
      _ExtentX        =   3413
      _ExtentY        =   1085
      Caption         =   "undo changes"
   End
   Begin PhotoDemon.pdCheckBox chkAutoCapture 
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   5760
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   661
      Caption         =   "try to capture automatically"
      FontSize        =   11
   End
   Begin PhotoDemon.pdDropDown ddKey 
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   4320
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1508
      Caption         =   "this hotkey"
      FontSize        =   11
   End
   Begin PhotoDemon.pdCheckBox chkModifier 
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   5280
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Caption         =   "Ctrl"
      FontSize        =   11
      Value           =   0   'False
   End
   Begin PhotoDemon.pdTextBox txtHotkey 
      Height          =   735
      Left            =   8520
      TabIndex        =   2
      Top             =   2640
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1296
      FontSize        =   12
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   6960
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   1296
      AutoloadLastPreset=   -1  'True
      DontResetAutomatically=   -1  'True
      HideRandomizeButton=   -1  'True
   End
   Begin PhotoDemon.pdTreeviewOD tvMenus 
      Height          =   3975
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   7011
   End
   Begin PhotoDemon.pdCheckBox chkModifier 
      Height          =   375
      Index           =   1
      Left            =   2160
      TabIndex        =   4
      Top             =   5280
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Caption         =   "Alt"
      FontSize        =   11
      Value           =   0   'False
   End
   Begin PhotoDemon.pdCheckBox chkModifier 
      Height          =   375
      Index           =   2
      Left            =   4080
      TabIndex        =   5
      Top             =   5280
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Caption         =   "Shift"
      FontSize        =   11
      Value           =   0   'False
   End
   Begin PhotoDemon.pdButton cmdThisHotkey 
      Height          =   615
      Index           =   1
      Left            =   2160
      TabIndex        =   10
      Top             =   6240
      Width           =   1815
      _ExtentX        =   4895
      _ExtentY        =   1085
      Caption         =   "restore default"
   End
   Begin PhotoDemon.pdButton cmdAll 
      Height          =   615
      Index           =   3
      Left            =   9000
      TabIndex        =   11
      Top             =   5490
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1085
      Caption         =   "export to file..."
   End
   Begin PhotoDemon.pdButton cmdAll 
      Height          =   615
      Index           =   2
      Left            =   6240
      TabIndex        =   12
      Top             =   5490
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1085
      Caption         =   "import from file..."
   End
   Begin PhotoDemon.pdButton cmdThisHotkey 
      Height          =   615
      Index           =   2
      Left            =   4080
      TabIndex        =   13
      Top             =   6240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
      Caption         =   "delete"
   End
   Begin PhotoDemon.pdButton cmdAll 
      Height          =   615
      Index           =   1
      Left            =   9000
      TabIndex        =   14
      Top             =   4740
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1085
      Caption         =   "restore all defaults"
   End
   Begin PhotoDemon.pdButton cmdAll 
      Height          =   615
      Index           =   4
      Left            =   6240
      TabIndex        =   15
      Top             =   6240
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1085
      Caption         =   "generate summary..."
   End
End
Attribute VB_Name = "FormHotkeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Customizable hotkeys dialog
'Copyright 2024-2026 by Tanner Helland
'Created: 09/September/24
'Last updated: 05/November/24
'Last update: final touches!  Almost ready to launch this feature!
'
'This dialog allows the user to customize hotkeys.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Objects retrieved from other engines
Private m_Menus() As PD_MenuEntry, m_NumOfMenus As Long
Private m_Hotkeys() As PD_Hotkey, m_NumOfHotkeys As Long

'To simplify lookup of menus by caption (required for pairing children IDs against parent IDs),
' we use a hash table.
Private m_MenuHash As pdVariantHash

'Menu and hotkey data are merged into this local struct, which is far more convenient for this UI
Private Type PD_HotkeyUI
    hk_TextEn As String
    hk_TextLocalized As String
    hk_ActionID As String
    hk_ParentID As String
    hk_HasChildren As Boolean
    hk_SubmenuLevel As Integer
    hk_NumParents As Long
    hk_KeyCode As Long
    hk_ShiftState As ShiftConstants
    hk_HotkeyText As String
    
    'If the user cancels this dialog (or reverts changes), we can use these backup copies of original hotkey data
    ' to revert everything to a pristine, untouched state.
    hk_BackupKeyCode As Long
    hk_BackupShiftState As Long
    hk_BackupHotkeyText As String
    
    'Similarly, these contain PD's *default* hotkeys for this item.  This may or may not correlate to the "backup" hotkey
    ' (which stores the hotkey when the dialog was loaded, which is the *user's* hotkey).
    hk_DefaultKeyCode As Long
    hk_DefaultShiftState As Long
    hk_DefaultHotkeyText As String
    
    'If this item has a hotkey that is used elsewhere, the duplication will get flagged here.
    hk_DuplicateFound As Boolean
    
End Type

'Menu and hotkey information gets merged into this local array, which is much easier to manage
' against the UI of this dialog
Private m_Items() As PD_HotkeyUI, m_numItems As Long

'Height of each list item in the custom-drawn treeview, in pixels, at 96 DPI
Private Const BLOCKHEIGHT As Long = 32

'Two font objects; one for menus that are allowed to have hotkeys, and one for menus that are not
' (e.g. top-level menus or parent-only menus).
Private m_FontAllowed As pdFont, m_FontDisallowed As pdFont, m_FontHotkey As pdFont

'All rendering is suspended until the form is loaded
Private m_RenderingOK As Boolean

'Retaining hotkey text in an edit box is non-trivial.  Store the last hotkey text here in WM_KEYDOWN,
' and manually restore it in Change/WM_KEYUP
Private m_backupHotkeyText As String, m_backupHotkeyShift As Long, m_backupHotkeyVKCode As Long
Private m_inAutoUpdate As Boolean

'List of all possible hotkeys.  Used to fill the key dropdown in the right-side selector.
Private Type PD_PossibleHotkey
    ph_VKCode As Long
    ph_KeyName As String
    ph_KeyComments As String    'Comments from MSDN; used for debugging only!
End Type

Private m_numPossibleHotkeys As Long, m_idxOtherHotkey As Long, m_idxNoneHotkey As Long, m_idxFinalHotkey As Long
Private m_possibleHotkeys() As PD_PossibleHotkey

'Last-edited hotkey
Private m_idxLastHotkey As Long

'Fast correlation from action ID to item index
Private m_ActionHash As pdVariantHash

'Warning icon, rendered to the treeview when a hotkey is used in more than one place
Private m_WarningIcon As pdDIB

Private Sub chkModifier_Click(Index As Integer)
    UpdateHotkeyManually
End Sub

Private Sub UpdateHotkeyManually()
    
    'When one of the Ctrl/Shift/Alt checkboxes *or* the keycode dropdown are changed, call this function to
    ' relay the change to the underlying hotkey collection, then redraw the treeview.  (Importantly, this function
    ' *won't* fire if those UI elements were updated by changes that originated from the treeview, like the user
    ' typing in a new shortcut when auto-detect is ON.)
    If (Not m_inAutoUpdate) And (tvMenus.ListIndex >= 0) And (ddKey.ListIndex >= 0) Then
        
        m_inAutoUpdate = True
        
        Dim newShiftState As Long
        If chkModifier(0).Value Then newShiftState = newShiftState Or vbCtrlMask
        If chkModifier(1).Value Then newShiftState = newShiftState Or vbAltMask
        If chkModifier(2).Value Then newShiftState = newShiftState Or vbShiftMask
        
        With m_Items(tvMenus.ListIndex)
            .hk_ShiftState = newShiftState
            .hk_KeyCode = m_possibleHotkeys(ddKey.ListIndex).ph_VKCode
            .hk_HotkeyText = AutoTextKeyChange(newShiftState, .hk_KeyCode)
        End With
        
        'Ensure duplicate hotkeys are flagged and marked
        FlagAllDuplicates
        
        'Redraw the listview to reflect the new hotkey
        tvMenus.RequestListRedraw
        
        m_inAutoUpdate = False
        
    End If

End Sub

Private Sub cmdAll_Click(Index As Integer)
    
    Dim i As Long
    
    Select Case Index
        
        'Undo all hotkey changes (this session)
        Case 0
            
            'Because this overwrites everything, ask for change before continuing
            If (Not OkayToOverwriteAll()) Then Exit Sub
            
            For i = 0 To m_numItems - 1
                With m_Items(i)
                    .hk_KeyCode = .hk_BackupKeyCode
                    .hk_HotkeyText = .hk_BackupHotkeyText
                    .hk_ShiftState = .hk_BackupShiftState
                End With
            Next i
            
            'Ensure duplicate hotkeys are flagged and marked
            FlagAllDuplicates
            
            tvMenus.RequestListRedraw
        
        'Restore all hotkey defaults
        Case 1
            
            'Because this overwrites everything, ask for change before continuing
            If (Not OkayToOverwriteAll()) Then Exit Sub
            
            For i = 0 To m_numItems - 1
                With m_Items(i)
                    .hk_KeyCode = .hk_DefaultKeyCode
                    .hk_HotkeyText = .hk_DefaultHotkeyText
                    .hk_ShiftState = .hk_DefaultShiftState
                End With
            Next i
            
            'Ensure duplicate hotkeys are flagged and marked
            FlagAllDuplicates
            
            tvMenus.RequestListRedraw
        
        'Import hotkeys
        Case 2
            
            'Because this overwrites everything, ask for change before continuing
            If (Not OkayToOverwriteAll()) Then Exit Sub
            
            'Disable user input until the dialog closes
            Interface.DisableUserInput
            
            'Determine an initial folder.
            Dim initialFolder As String
            initialFolder = UserPrefs.GetHotkeyPath()
            
            'Build a common dialog filter (only one format is supported for hotkey import/export right now)
            Dim cdFilter As String, cdIndex As Long, cdTitle As String
            cdFilter = g_Language.TranslateMessage("Hotkeys") & " (.xml)|*.xml"
            cdIndex = 1
            cdTitle = g_Language.TranslateMessage("Import hotkeys")
            
            'Prep a common dialog interface
            Dim openDialog As pdOpenSaveDialog
            Set openDialog = New pdOpenSaveDialog
            
            Dim srcFilename As String
            If openDialog.GetOpenFileName(srcFilename, , True, False, cdFilter, cdIndex, initialFolder, cdTitle, "xml", Me.hWnd) Then
                        
                'Update preferences
                UserPrefs.SetHotkeyPath Files.FileGetPath(srcFilename)
                
                'Import hotkeys.  Note that this will overwrite all existing hotkey choices (by design).
                ImportHotkeysFromFile srcFilename
                
                'Ensure duplicate hotkeys are flagged and marked
                FlagAllDuplicates
                
            End If
            
            'Re-enable UI, then redraw the treeview (as hotkeys will have changes)
            Interface.EnableUserInput
            tvMenus.RequestListRedraw
            
        'Export hotkeys
        Case 3
            
            'Disable user input until the dialog closes
            Interface.DisableUserInput
            
            'Determine an initial folder.  This is easy - just grab the last "profile" path from the preferences file.
            Dim initialSaveFolder As String
            initialSaveFolder = Files.PathAddBackslash(UserPrefs.GetHotkeyPath())
            
            'Build a common dialog filter list
            Dim cdFilterExtensions As String
            cdFilter = g_Language.TranslateMessage("Hotkeys") & " (.xml)|*.xml"
            cdFilterExtensions = "xml"
            cdIndex = 1
            cdTitle = g_Language.TranslateMessage("Export hotkeys")
            
            'Suggest a file name.
            Dim dstFilename As String
            dstFilename = g_Language.TranslateMessage("Hotkeys")
            dstFilename = initialSaveFolder & dstFilename
            
            'Display a common save dialog
            Dim saveDialog As pdOpenSaveDialog
            Set saveDialog = New pdOpenSaveDialog
            If saveDialog.GetSaveFileName(dstFilename, , True, cdFilter, cdIndex, initialSaveFolder, cdTitle, cdFilterExtensions, Me.hWnd) Then
                
                'Update preferences, then export to the user's requested file.
                UserPrefs.SetHotkeyPath Files.FileGetPath(dstFilename)
                ExportHotkeysToFile dstFilename
                
            End If
            
            'Re-enable UI
            Interface.EnableUserInput
            
        'Generate summary file (html)
        Case 4
            
            'Disable user input until the dialog closes
            Interface.DisableUserInput
            
            'Determine an initial folder.  This is easy - just grab the last "profile" path from the preferences file.
            initialSaveFolder = Files.PathAddBackslash(OS.SpecialFolder(CSIDL_COMMON_DOCUMENTS))
            
            'Build a common dialog filter list
            cdFilter = "HTML (.html)|*.html"
            cdFilterExtensions = "html"
            
            cdIndex = 1
            
            'Suggest a file name.
            dstFilename = g_Language.TranslateMessage("PhotoDemon Hotkeys")
            dstFilename = initialSaveFolder & dstFilename
            
            cdTitle = g_Language.TranslateMessage("Export hotkeys")
            
            'Display a common save dialog, then export on a success
            Set saveDialog = New pdOpenSaveDialog
            If saveDialog.GetSaveFileName(dstFilename, , True, cdFilter, cdIndex, initialSaveFolder, cdTitle, cdFilterExtensions, Me.hWnd) Then
                GenerateSummaryFile dstFilename
                Web.OpenURL dstFilename
            End If
            
            'Re-enable UI
            Interface.EnableUserInput
            
    End Select
        
End Sub

'Load all hotkeys from an XML file.  This will erase all existing hotkey choices, by design.
Private Sub ImportHotkeysFromFile(ByRef srcFile As String)
    
    Dim cXML As pdXML
    Set cXML = New pdXML
    If cXML.LoadXMLFile(srcFile) Then FillHotkeyArrayFromXML cXML
    
    'If the user previously selected an item in the treeview, update it now
    If (tvMenus.ListIndex >= 0) Then
        
        'Similarly, ensure the left-side manual edit controls are updated correctly.
        m_inAutoUpdate = True
        AutoTextKeyChange m_Items(tvMenus.ListIndex).hk_ShiftState, m_Items(tvMenus.ListIndex).hk_KeyCode
        If Me.txtHotkey.Visible Then Me.txtHotkey.Text = m_Items(tvMenus.ListIndex).hk_HotkeyText
        m_inAutoUpdate = False
        
    End If
    
End Sub

Private Function FillHotkeyArrayFromXML(ByRef cXML As pdXML) As Boolean
    
    'Ensure the XML object actually holds hotkey data
    If cXML.IsPDDataType("hotkeys") Then
        
        'Wipe all existing hotkey data
        Dim i As Long
        For i = 0 To m_numItems - 1
            m_Items(i).hk_KeyCode = 0
            m_Items(i).hk_ShiftState = 0
            m_Items(i).hk_HotkeyText = vbNullString
        Next i
        
        'Get a list of all "hotkey" entries from the XML file
        Dim hotkeyTags() As Long
        If cXML.FindAllTagLocations(hotkeyTags, "hotkey") Then
        
            On Error GoTo BadHotkey
            
            'Iterate hotkey entries, loading as we go
            For i = LBound(hotkeyTags) To UBound(hotkeyTags)
                
                Dim hkActionID As String
                Const HOTKEY_CODE_ACTION As String = "action"
                hkActionID = cXML.GetUniqueTag_String(HOTKEY_CODE_ACTION, vbNullString, hotkeyTags(i))
                If (LenB(hkActionID) <> 0) Then
                
                    'Find the matching action ID for this command
                    Dim idxTarget As Long
                    If m_ActionHash.GetItemByKey(hkActionID, idxTarget) Then
                    
                        'Pull the shift state and keycode from this entry and store them in this hotkey
                        Dim newShiftState As ShiftConstants
                        newShiftState = 0
                        
                        Const HOTKEY_CODE_CTRL As String = "ctrl", HOTKEY_CODE_ALT As String = "alt", HOTKEY_CODE_SHIFT As String = "shift"
                        If (cXML.GetUniqueTag_Long(HOTKEY_CODE_CTRL, 0, hotkeyTags(i)) <> 0) Then newShiftState = newShiftState Or vbCtrlMask
                        If (cXML.GetUniqueTag_Long(HOTKEY_CODE_ALT, 0, hotkeyTags(i)) <> 0) Then newShiftState = newShiftState Or vbAltMask
                        If (cXML.GetUniqueTag_Long(HOTKEY_CODE_SHIFT, 0, hotkeyTags(i)) <> 0) Then newShiftState = newShiftState Or vbShiftMask
                        m_Items(idxTarget).hk_ShiftState = newShiftState
                        
                        Const HOTKEY_CODE_TAG As String = "key-id"
                        m_Items(idxTarget).hk_KeyCode = cXML.GetUniqueTag_Long(HOTKEY_CODE_TAG, 0, hotkeyTags(i))
                        
                        'Finally, populate text for this key combo
                        m_Items(idxTarget).hk_HotkeyText = GetHotkeyNameFromKeys(newShiftState, m_Items(idxTarget).hk_KeyCode)
                        
                    'Not sure what to do here... maybe an "other, non-menu" category someday?
                    End If
                
                '/action exists for this hotkey
                End If
                
BadHotkey:
            Next i
            
            On Error GoTo 0
            
        '/at least one hotkey found
        End If
        
    '/file contains valid hotkey data
    End If

    FillHotkeyArrayFromXML = True
    
End Function

'Export the current hotkey collection to file
Private Sub ExportHotkeysToFile(ByRef dstFile As String, Optional ByVal stripDuplicatesFirst As Boolean = False)
    Dim cXML As pdXML
    If GetCurrentHotkeysAsXML(cXML, stripDuplicatesFirst) Then
        If Files.FileExists(dstFile) Then Files.FileDelete dstFile
        cXML.WriteXMLToFile dstFile
    End If
End Sub

'Convert the current hotkey collection (including all edits) to an XML object.  The result can be dumped out to file,
' or stored as a preset (or other string-based object).
Private Function GetCurrentHotkeysAsXML(ByRef cXML As pdXML, Optional ByVal stripDuplicatesFirst As Boolean = False) As Boolean
    
    If (cXML Is Nothing) Then Set cXML = New pdXML
    cXML.PrepareNewXML "hotkeys"
    cXML.WriteBlankLine
    
    Dim hotkeysWritten As pdVariantHash, actionsWritten As pdVariantHash
    Set hotkeysWritten = New pdVariantHash: Set actionsWritten = New pdVariantHash
    
    'Iterate all hotkeys, and write any with non-zero keycodes out to file
    Dim i As Long
    For i = 0 To m_numItems - 1
        If (m_Items(i).hk_KeyCode <> 0) Then
            
            Dim okToWrite As Boolean, meaninglessReturn As Variant
            okToWrite = True
            
            'Generate a string version of this keycode, and only proceed if we *haven't* already written out
            ' this keycode before.
            Dim strKeyCode As String
            strKeyCode = Trim$(Str$(m_Items(i).hk_KeyCode))
            
            Const STR_ZERO As String = "0", STR_ONE As String = "1"
            Dim strCtrl As String, strAlt As String, strShift As String
            strCtrl = IIf((m_Items(i).hk_ShiftState And vbCtrlMask) <> 0, STR_ONE, STR_ZERO)
            strAlt = IIf((m_Items(i).hk_ShiftState And vbAltMask) <> 0, STR_ONE, STR_ZERO)
            strShift = IIf((m_Items(i).hk_ShiftState And vbShiftMask) <> 0, STR_ONE, STR_ZERO)
            
            If stripDuplicatesFirst Then
                okToWrite = (Not actionsWritten.GetItemByKey(m_Items(i).hk_ActionID, meaninglessReturn))
                okToWrite = okToWrite And (Not hotkeysWritten.GetItemByKey(strKeyCode & strCtrl & strAlt & strShift, meaninglessReturn))
            End If
            
            If okToWrite Then
                
                Const HOTKEY_TAG As String = "hotkey"
                cXML.WriteTag HOTKEY_TAG, vbNullString, doNotCloseTag:=True
                
                Const HOTKEY_CODE_TAG As String = "key-id"
                cXML.WriteTag HOTKEY_CODE_TAG, strKeyCode
                
                Const HOTKEY_CODE_CTRL As String = "ctrl", HOTKEY_CODE_ALT As String = "alt", HOTKEY_CODE_SHIFT As String = "shift"
                cXML.WriteTag HOTKEY_CODE_CTRL, strCtrl
                cXML.WriteTag HOTKEY_CODE_ALT, strAlt
                cXML.WriteTag HOTKEY_CODE_SHIFT, strShift
                
                Const HOTKEY_CODE_ACTION As String = "action"
                cXML.WriteTag HOTKEY_CODE_ACTION, m_Items(i).hk_ActionID
                
                cXML.CloseTag HOTKEY_TAG
                
                'Add this as a "written" hotkey, and tag it both by action *and* hotkey.
                ' (The menu manager automatically handles resolving duplicates by action, and duplicates by *hotkey* break things.)
                actionsWritten.AddItem m_Items(i).hk_ActionID, meaninglessReturn
                hotkeysWritten.AddItem strKeyCode & strCtrl & strAlt & strShift, meaninglessReturn
                
            End If
                
        End If
    Next i
    
    'Add a last line to make it slightly prettier ;)
    cXML.WriteBlankLine
    GetCurrentHotkeysAsXML = True
    
End Function

Private Sub cmdBar_AddCustomPresetData()
    
    Dim cXML As pdXML
    If GetCurrentHotkeysAsXML(cXML, False) Then
        cmdBar.AddPresetData "hotkeys-as-xml", cXML.ReturnCurrentXMLString(True)
    End If
    
End Sub

Private Sub cmdBar_BeforeResetClick(ByRef cancelReset As Boolean)
    
    'Because this overwrites everything, ask for change before continuing
    If (Not OkayToOverwriteAll()) Then cancelReset = True
    
End Sub

Private Sub cmdBar_ExtraValidations()
    
    'PD can automatically resolve duplicate hotkeys, but the user should be warned (since this may produce unexpected results).
    FlagAllDuplicates
    
    Dim dupesFound As Boolean: dupesFound = False
    Dim i As Long, idxFirstDuplicate As Long
    For i = 0 To m_numItems - 1
        If m_Items(i).hk_DuplicateFound Then
            dupesFound = True
            idxFirstDuplicate = i
            Exit For
        End If
    Next i
    
    If dupesFound Then
        
        Dim txtWarning As pdString, txtTitle As String
        txtTitle = g_Language.TranslateMessage("Warning")
        
        Set txtWarning = New pdString
        txtWarning.AppendLine g_Language.TranslateMessage("One or more hotkeys is currently assigned to multiple actions.")
        txtWarning.AppendLineBreak
        txtWarning.AppendLine g_Language.TranslateMessage("If you proceed, PhotoDemon will only keep the first occurrence of any duplicated hotkeys.")
        txtWarning.AppendLineBreak
        txtWarning.AppendLine g_Language.TranslateMessage("Press OK to proceed.")
        txtWarning.AppendLine g_Language.TranslateMessage("Press cancel to keep editing hotkeys.")
        
        Dim userAction As VbMsgBoxResult
        userAction = PDMsgBox(txtWarning.ToString, vbExclamation Or vbOKCancel Or vbApplicationModal, txtTitle)
        
        If (userAction <> vbOK) Then
            
            'The user wants to go back to editing.  Auto-select the first duplicate in the list for them.
            tvMenus.ListIndex = idxFirstDuplicate
            
            cmdBar.ValidationFailed
            Exit Sub
            
        End If
        
    End If
    
End Sub

Private Sub cmdBar_OKClick()
    
    'Look at ExtraValidations() to see the things this dialog validates before allowing an OK press to continue.
    'TODO: validate that something has changed?
    
    'Start by writing the current collection out to file.  (Any duplicates, if they exist, will be forcibly resolved now.)
    Dim dstFile As String
    dstFile = Hotkeys.GetNameOfHotkeyFile()
    ExportHotkeysToFile dstFile, True
    
    'We now need to notify some outside entities that hotkeys have changed.
    
    'Notify various UI elements that hotkeys are out of date, and need to be re-loaded
    Hotkeys.EraseHotkeyCollection
    Menus.NotifyHotkeysChanged
    
    'Now ask the hotkey manager to re-load all hotkey data from our freshly saved file.  It will also notify the
    ' menu module of all changes.
    Hotkeys.LoadAllHotkeys
    
    'Before exiting, some UI elements need to be redrawn (as their text will have changed)
    Menus.UpdateAgainstCurrentTheme True
    toolbar_Toolbox.UpdateAgainstCurrentTheme
    
End Sub

Private Sub cmdBar_ReadCustomPresetData()
    
    Dim cXML As pdXML: Set cXML = New pdXML
    If cXML.LoadXMLFromString(cmdBar.RetrievePresetData("hotkeys-as-xml")) Then
        FillHotkeyArrayFromXML cXML
    End If
    
    FlagAllDuplicates
    tvMenus.RequestListRedraw
    
End Sub

Private Sub cmdBar_ResetClick()
    
    'Restore everything to its original state (e.g. when the dialog was loaded)
    Dim i As Long
    For i = 0 To m_numItems - 1
        With m_Items(i)
            .hk_KeyCode = .hk_BackupKeyCode
            .hk_HotkeyText = .hk_BackupHotkeyText
            .hk_ShiftState = .hk_BackupShiftState
        End With
    Next i
    
    'Ensure duplicate hotkeys are flagged and marked (failsafe only)
    FlagAllDuplicates
    
    'Toggle the current listindex to ensure a redraw and correct checkbox values on the lower-left
    Dim idxOld As Long
    idxOld = tvMenus.ListIndex
    tvMenus.ListIndex = -1
    tvMenus.ListIndex = idxOld
    tvMenus.RequestListRedraw
    
    'Select the (none) entry as relevant
    If (tvMenus.ListIndex = -1) Then ddKey.ListIndex = m_idxNoneHotkey
    
End Sub

Private Sub cmdThisHotkey_Click(Index As Integer)
    
    If (tvMenus.ListIndex >= 0) Then
        
        Select Case Index
            
            'Undo any changes to this hotkey
            Case 0
                With m_Items(tvMenus.ListIndex)
                    .hk_KeyCode = .hk_BackupKeyCode
                    .hk_HotkeyText = .hk_BackupHotkeyText
                    .hk_ShiftState = .hk_BackupShiftState
                End With
                
                'Ensure duplicate hotkeys are flagged and marked
                FlagAllDuplicates
                
            'Reset this hotkey to PD's default hotkey
            Case 1
                With m_Items(tvMenus.ListIndex)
                    .hk_KeyCode = .hk_DefaultKeyCode
                    .hk_HotkeyText = .hk_DefaultHotkeyText
                    .hk_ShiftState = .hk_DefaultShiftState
                End With
                
                'Ensure duplicate hotkeys are flagged and marked
                FlagAllDuplicates
        
            'Delete this hotkey
            Case 2
                With m_Items(tvMenus.ListIndex)
                    .hk_KeyCode = 0
                    .hk_HotkeyText = vbNullString
                    .hk_ShiftState = 0
                End With
                
                'Ensure duplicate hotkeys are flagged and marked.
                ' (In this case, we're mostly interested in *clearing* duplicate flags, as they may have been resolved
                ' by erasing this hotkey!)
                FlagAllDuplicates
                
        End Select
        
        tvMenus.RequestListRedraw
        
    End If
    
End Sub

Private Sub ddKey_Click()
    UpdateHotkeyManually
End Sub

Private Sub Form_Load()
    
    'No hotkeys have been edited yet
    m_idxLastHotkey = -1
    
    'Retrieve a copy of all menus (including hierarchies and attributes) from the menu manager
    m_NumOfMenus = Menus.GetCopyOfAllMenus(m_Menus)
    
    'We will add all menus (by a hierarchical ID) to a hash table so we can quickly move between IDs and array indices.
    Set m_MenuHash = New pdVariantHash
    
    'Similarly, we will add all action IDs to a hash table so we can lookup specific actions quickly
    Set m_ActionHash = New pdVariantHash
    
    'Retrieve a copy of all hotkeys from the hotkey manager
    m_NumOfHotkeys = Hotkeys.GetCopyOfAllHotkeys(m_Hotkeys)
    
    'There will (typically? always?) be fewer hotkeys than there are menu/action targets.  To simplify
    ' correlating between action IDs and hotkey indices, build a quick dictionary.
    Dim cHotkeys As pdVariantHash
    Set cHotkeys = New pdVariantHash
    
    Dim i As Long
    If (m_NumOfHotkeys > 0) Then
        For i = 0 To m_NumOfHotkeys - 1
            cHotkeys.AddItem m_Hotkeys(i).hkAction, i
        Next i
    End If
    
    'Turn off automatic redraws in the treeview object
    tvMenus.SetAutomaticRedraws False
    tvMenus.ListItemHeight = Interface.FixDPI(BLOCKHEIGHT)
    
    'Iterate the menu collection, and pair each menu with a hotkey against its relevant hotkey partner
    ReDim m_Items(0 To m_NumOfMenus - 1) As PD_HotkeyUI
    m_numItems = 0
    
    For i = 0 To m_NumOfMenus - 1
        
        'Ignore separators and a few other "special" menus that require special handling (e.g. "File > Open recent")
        Const ID_DASH As String = "-"
        
        Dim okToProcessMenu As Boolean
        okToProcessMenu = True
        okToProcessMenu = okToProcessMenu And (m_Menus(i).me_Name <> ID_DASH)
        okToProcessMenu = okToProcessMenu And (m_Menus(i).me_Name <> "file_openrecent")
        okToProcessMenu = okToProcessMenu And (m_Menus(i).me_Name <> "tools_developers")
        okToProcessMenu = okToProcessMenu And (m_Menus(i).me_Name <> "tools_viewdebuglog")
        okToProcessMenu = okToProcessMenu And (m_Menus(i).me_Name <> "tools_themeeditor")
        okToProcessMenu = okToProcessMenu And (m_Menus(i).me_Name <> "tools_themepackage")
        okToProcessMenu = okToProcessMenu And (m_Menus(i).me_Name <> "tools_standalonepackage")
        okToProcessMenu = okToProcessMenu And (m_Menus(i).me_Name <> "effects_developertest")
        
        If okToProcessMenu Then
            
            With m_Items(m_numItems)
                
                'Before doing anything with this menu, add it to a hash table (so we can quickly correlate between
                ' menu positions *in the menu bar* (which is hierarchical) and menu positions *in this array*).
                Dim mnuID As String
                mnuID = GetMenuPositionID(i)
                m_MenuHash.AddItem mnuID, i
                
                'Start by copying over the menu data we can use as-is (like localizations)
                .hk_ActionID = m_Menus(i).me_Name
                .hk_HasChildren = m_Menus(i).me_HasChildren
                .hk_TextEn = m_Menus(i).me_TextEn
                .hk_TextLocalized = m_Menus(i).me_TextTranslated
                
                'Save this action and index in a fast lookup table
                m_ActionHash.AddItem .hk_ActionID, m_numItems
                
                'If a hotkey exists for this menu's action, retrieve it and add it
                ' (and make backups of these *original* hotkeys, so we can revert them if the user doesn't like later changes)
                Dim idxHotkey As Variant
                If cHotkeys.GetItemByKey(.hk_ActionID, idxHotkey) Then
                    .hk_KeyCode = m_Hotkeys(idxHotkey).hkKeyCode
                    .hk_BackupKeyCode = .hk_KeyCode
                    .hk_ShiftState = m_Hotkeys(idxHotkey).hkShiftState
                    .hk_BackupShiftState = .hk_ShiftState
                    .hk_HotkeyText = m_Menus(i).me_HotKeyTextTranslated
                    .hk_BackupHotkeyText = .hk_HotkeyText
                End If
                
                'Finally, if this is not a top-level menu, retrieve the ID of this menu's *parent* menu
                If (m_Menus(i).me_SubMenu >= 0) Then
                    Dim idxParent As Variant
                    m_MenuHash.GetItemByKey GetMenuParentPositionID(i), idxParent
                    .hk_ParentID = m_Menus(idxParent).me_Name
                    .hk_NumParents = 1
                    If (m_Menus(i).me_SubSubMenu >= 0) Then .hk_NumParents = 2
                End If
                
                'PDDebug.LogAction .hk_ActionID & ", " & .hk_ParentID & ", " & .hk_HasChildren & ", " & .hk_NumParents
                
                'Add this menu item to the treeview
                tvMenus.AddItem .hk_ActionID, .hk_TextLocalized, .hk_ParentID, (.hk_SubmenuLevel = 0)
                
                'Advance to the next mappable menu index
                m_numItems = m_numItems + 1
                
            End With
            
        '/Ignore separators
        End If
        
    Next i
    
    'Next, we need to manually add toolbox commands
    AddToolboxActions cHotkeys
    
    'Now we can pull all of PD's default hotkeys and correlate those with the current menu and tool collection.
    Dim defaultHotkeys() As PD_Hotkey, numDefaultHotkeys As Long
    numDefaultHotkeys = Hotkeys.GetCopyOfAllHotkeys(defaultHotkeys, True)
    If (numDefaultHotkeys > 0) And (m_numItems > 0) Then
        For i = 0 To numDefaultHotkeys - 1
            If m_ActionHash.GetItemByKey(defaultHotkeys(i).hkAction, idxHotkey) Then
                With m_Items(idxHotkey)
                    .hk_DefaultKeyCode = defaultHotkeys(i).hkKeyCode
                    .hk_DefaultShiftState = defaultHotkeys(i).hkShiftState
                    .hk_DefaultHotkeyText = GetHotkeyNameFromKeys(.hk_DefaultShiftState, .hk_DefaultKeyCode)
                End With
            End If
        Next i
    End If
    
    'Ensure any duplicate hotkeys are flagged and marked
    FlagAllDuplicates
    
    'Initialize font renderers for the custom treeview
    Set m_FontAllowed = New pdFont
    m_FontAllowed.SetFontBold True
    m_FontAllowed.SetFontSize 12
    m_FontAllowed.CreateFontObject
    m_FontAllowed.SetTextAlignment vbLeftJustify
    
    Set m_FontDisallowed = New pdFont
    m_FontDisallowed.SetFontBold False
    m_FontDisallowed.SetFontSize 12
    m_FontDisallowed.CreateFontObject
    m_FontDisallowed.SetTextAlignment vbLeftJustify
    
    Set m_FontHotkey = New pdFont
    m_FontHotkey.SetFontBold False
    m_FontHotkey.SetFontSize 12
    m_FontHotkey.CreateFontObject
    m_FontHotkey.SetTextAlignment vbLeftJustify
    
    'Add all possible hotkeys to the dropdown
    GeneratePossibleHotkeys
    
    'Load some icons to various toolbars
    Dim buttonImgSize As Long
    buttonImgSize = Interface.FixDPI(24)
    cmdThisHotkey(0).AssignImage "edit_undo", , buttonImgSize, buttonImgSize, g_Themer.GetGenericUIColor(UI_IconMonochrome)
    cmdThisHotkey(1).AssignImage "generic_reset", , buttonImgSize, buttonImgSize, g_Themer.GetGenericUIColor(UI_IconMonochrome)
    cmdThisHotkey(2).AssignImage "file_close", , buttonImgSize, buttonImgSize, g_Themer.GetGenericUIColor(UI_IconMonochrome)
    cmdAll(0).AssignImage "edit_undo", , buttonImgSize, buttonImgSize, g_Themer.GetGenericUIColor(UI_IconMonochrome)
    cmdAll(1).AssignImage "generic_reset", , buttonImgSize, buttonImgSize, g_Themer.GetGenericUIColor(UI_IconMonochrome)
    cmdAll(2).AssignImage "file_open", , buttonImgSize, buttonImgSize
    cmdAll(3).AssignImage "file_saveas", , buttonImgSize, buttonImgSize
    
    'Apply custom themes
    Interface.ApplyThemeAndTranslations Me, True, False
    
    '*Now* allow the treeview to render itself
    m_RenderingOK = True
    tvMenus.SetAutomaticRedraws True, True
    
End Sub

'Manually add toolbox shortcuts to the bottom of the list
Private Sub AddToolboxActions(ByRef cHotkeys As pdVariantHash)
    
    If (m_numItems > UBound(m_Items)) Then ReDim Preserve m_Items(0 To m_numItems * 2 - 1) As PD_HotkeyUI
    
    'Add a top-level "toolbox tools" item
    Dim idxToolboxActions As Long
    idxToolboxActions = m_numItems
    
    With m_Items(m_numItems)
        
        .hk_ActionID = "toolbox-tools"
        .hk_HasChildren = True
        .hk_TextEn = "Toolbox tools"
        .hk_TextLocalized = g_Language.TranslateMessage("Toolbox tools")
        .hk_ParentID = vbNullString
        
        'Save this action and index in a fast lookup table
        m_ActionHash.AddItem .hk_ActionID, m_numItems
        
        'Add this to the treeview, then advance
        tvMenus.AddItem .hk_ActionID, .hk_TextLocalized, .hk_ParentID, (.hk_SubmenuLevel = 0)
        
        'Advance to the next mappable menu index
        m_numItems = m_numItems + 1
        
    End With
    
    'Now manually add all toolbox actions
    Dim toolNames As pdStringStack, toolActions As pdStringStack
    toolbar_Toolbox.GetListOfToolNamesAndActions toolNames, toolActions
    
    Dim i As Long
    For i = 0 To toolNames.GetNumOfStrings - 1
        AddOneToolboxAction toolNames.GetString(i), toolActions.GetString(i), cHotkeys
    Next i
    
End Sub

Private Sub AddOneToolboxAction(ByRef toolName As String, ByRef toolAction As String, ByRef cHotkeys As pdVariantHash)

    If (m_numItems > UBound(m_Items)) Then ReDim Preserve m_Items(0 To m_numItems * 2 - 1) As PD_HotkeyUI
    
    'Add a top-level "toolbox tools" item
    With m_Items(m_numItems)
        
        .hk_ActionID = toolAction
        .hk_HasChildren = False
        
        'Mirror localized name across both english *and* localized text (in this case, we don't need the English text for anything)
        .hk_TextEn = toolName
        .hk_TextLocalized = toolName
        .hk_ParentID = "toolbox-tools"
        .hk_SubmenuLevel = 1
        
        'Save this action and index in a fast lookup table
        m_ActionHash.AddItem .hk_ActionID, m_numItems
        
        'If a hotkey exists for this menu's action, retrieve it and add it
        ' (and make backups of these *original* hotkeys, so we can revert them if the user doesn't like later changes)
        Dim idxHotkey As Variant
        If cHotkeys.GetItemByKey(.hk_ActionID, idxHotkey) Then
            .hk_KeyCode = m_Hotkeys(idxHotkey).hkKeyCode
            .hk_BackupKeyCode = .hk_KeyCode
            .hk_ShiftState = m_Hotkeys(idxHotkey).hkShiftState
            .hk_BackupShiftState = .hk_ShiftState
            .hk_HotkeyText = GetHotkeyNameFromKeys(.hk_ShiftState, .hk_KeyCode)
            .hk_BackupHotkeyText = .hk_HotkeyText
        End If
        
        'Add this to the treeview, then advance
        tvMenus.AddItem .hk_ActionID, .hk_TextLocalized, .hk_ParentID, (.hk_SubmenuLevel = 0)
        
        'Advance to the next mappable menu index
        m_numItems = m_numItems + 1
        
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Interface.ReleaseFormTheming Me
End Sub

'Return a unique hash table ID for a given menu
Private Function GetMenuPositionID(ByVal idxMenu As Long) As String
    
    Const ID_SEPARATOR As String = "-"
    
    With m_Menus(idxMenu)
        
        If (.me_TopMenu >= 0) Then
            GetMenuPositionID = Trim$(Str$((.me_TopMenu)))
            If (.me_SubMenu >= 0) Then GetMenuPositionID = GetMenuPositionID & ID_SEPARATOR & Trim$(Str$((.me_SubMenu)))
            If (.me_SubSubMenu >= 0) Then GetMenuPositionID = GetMenuPositionID & ID_SEPARATOR & Trim$(Str$((.me_SubSubMenu)))
        Else
            GetMenuPositionID = vbNullString
        End If
        
    End With
    
End Function

'Return a unique hash table ID for a given menu's parent (if one exists).  Returns a null-string if no parent exists.
Private Function GetMenuParentPositionID(ByVal idxMenu As Long) As String

    Const ID_SEPARATOR As String = "-"
    
    With m_Menus(idxMenu)
        If (.me_TopMenu >= 0) Then
            If (.me_SubMenu >= 0) Then GetMenuParentPositionID = Trim$(Str$((.me_TopMenu)))
            If (.me_SubSubMenu >= 0) Then GetMenuParentPositionID = GetMenuParentPositionID & ID_SEPARATOR & Trim$(Str$((.me_SubMenu)))
        End If
    End With
    
End Function

Private Sub tvMenus_Click()
    
    'Failsafe only
    If (tvMenus.ListIndex < 0) Then
        HideEditBox
        Exit Sub
    End If
    
    m_inAutoUpdate = True
    
    'Do not allow hotkeys on menu items with children
    Dim hotkeyEditingAllowed As Boolean
    hotkeyEditingAllowed = (Not m_Items(tvMenus.ListIndex).hk_HasChildren)
    
    'Manual hotkey controls at the bottom of the screen mirror the availability of hotkeys for this command
    If (Not hotkeyEditingAllowed) Then
        chkModifier(0).Value = False
        chkModifier(1).Value = False
        chkModifier(2).Value = False
        ddKey.ListIndex = m_idxNoneHotkey
    End If
    
    chkModifier(0).Enabled = hotkeyEditingAllowed
    chkModifier(1).Enabled = hotkeyEditingAllowed
    chkModifier(2).Enabled = hotkeyEditingAllowed
    ddKey.Enabled = hotkeyEditingAllowed
    
    'Note that the user can also disallow auto key capture
    If hotkeyEditingAllowed Then
        
        'Ensure the hotkey editors at the bottom of the screen reflect this item's current hotkey (if any)
        Dim curShiftState As Long, curKeyCode As Long
        curShiftState = m_Items(tvMenus.ListIndex).hk_ShiftState
        curKeyCode = m_Items(tvMenus.ListIndex).hk_KeyCode
        chkModifier(0).Value = (curShiftState And vbCtrlMask) = vbCtrlMask
        chkModifier(1).Value = (curShiftState And vbAltMask) = vbAltMask
        chkModifier(2).Value = (curShiftState And vbShiftMask) = vbShiftMask
        
        Dim i As Long, keyFound As Boolean
        For i = 0 To m_idxFinalHotkey - 1
            If (m_possibleHotkeys(i).ph_VKCode = curKeyCode) Then
                keyFound = True
                ddKey.ListIndex = i
                Exit For
            End If
        Next i
        
        If (Not keyFound) Then
            If (curKeyCode <= 0) Then
                ddKey.ListIndex = m_idxNoneHotkey
            Else
                ddKey.ListIndex = m_idxOtherHotkey
            End If
        End If
        
        'Automatic hotkey capture can be toggled by the user
        If Me.chkAutoCapture.Value Then
            
            'To figure out where to position the text box, we need to query the underlying tree support object for details
            Dim tmpTreeSupport As pdTreeSupport
            Set tmpTreeSupport = tvMenus.AccessUnderlyingTreeSupport()
            
            '...including where its child treeview_view is positioned
            Dim lbViewRectF As RectF
            CopyMemoryStrict VarPtr(lbViewRectF), tvMenus.GetListBoxRectFPtr, LenB(lbViewRectF)
            
            '...and the selected treeview item itself
            Dim tmpTreeItem As PD_TreeItem, tmpScrollX As Long, tmpScrollY As Long
            tmpTreeSupport.GetRenderingItem tvMenus.ListIndex, tmpTreeItem, tmpScrollX, tmpScrollY
            
            'Use data from these to figure out where the edit box should go
            Dim ebRectF As RectF
            ebRectF.Left = (tvMenus.GetLeft + lbViewRectF.Left + tmpTreeItem.captionRect.Left + tmpTreeItem.captionRect.Width) - Interface.FixDPI(200)
            ebRectF.Top = tvMenus.GetTop + lbViewRectF.Top + tmpTreeItem.captionRect.Top + Interface.FixDPI(3) - tmpScrollY
            ebRectF.Width = Interface.FixDPI(192)
            ebRectF.Height = tmpTreeItem.captionRect.Height - Interface.FixDPI(4)
            
            'Position it and fill it with the hotkey for the current tree item.
            ' (Note that the backup hotkey text *must* be set first - see the edit box _Change event for details.)
            Me.txtHotkey.Text = m_Items(tvMenus.ListIndex).hk_HotkeyText
            Me.txtHotkey.SetPositionAndSize ebRectF.Left, ebRectF.Top, ebRectF.Width, ebRectF.Height
            Me.txtHotkey.Visible = True
            Me.txtHotkey.ZOrder 0
            Me.txtHotkey.SetFocusToEditBox True
            
        End If
        
        'Note this as the last-edited hotkey, and update all data backups to match.
        ' (We'll restore these if the user enters an invalid hotkey, like Ctrl+Shift+[nothing])
        m_idxLastHotkey = tvMenus.ListIndex
        m_backupHotkeyText = m_Items(tvMenus.ListIndex).hk_HotkeyText
        m_backupHotkeyVKCode = m_Items(tvMenus.ListIndex).hk_KeyCode
        m_backupHotkeyShift = m_Items(tvMenus.ListIndex).hk_ShiftState
        
    End If
    m_inAutoUpdate = False
    
End Sub

'Render an item into the treeview
Private Sub tvMenus_DrawListEntry(ByVal bufferDC As Long, ByVal itemIndex As Long, ByRef itemID As String, ByVal itemIsSelected As Boolean, ByVal itemIsHovered As Boolean, ByVal ptrToItemRectF As Long, ByVal ptrToCaptionRectF As Long, ByVal ptrToControlRectF As Long)
    
    If (bufferDC = 0) Then Exit Sub
    If (Not m_RenderingOK) Then Exit Sub
    
    'Retrieve the boundary region for this list entry
    Dim tmpRectF As RectF
    CopyMemoryStrict VarPtr(tmpRectF), ptrToCaptionRectF, 16&
    
    Dim offsetY As Single, offsetX As Single
    offsetX = tmpRectF.Left
    offsetY = tmpRectF.Top + Interface.FixDPI(1)
    
    'Hotkeys get a fixed (at 96-dpi) 192 pixels to display their key combo.  If menu text overflows this boundary,
    ' it will be truncated with ellipses.
    Dim leftOffsetHotkey As Long
    leftOffsetHotkey = tmpRectF.Left + tmpRectF.Width - (Interface.FixDPI(192))
    
    'If this item has been selected, draw the background with the system's current selection color
    Dim curFont As pdFont
    If m_Items(itemIndex).hk_HasChildren Then Set curFont = m_FontDisallowed Else Set curFont = m_FontAllowed
    
    If itemIsSelected Then
        curFont.SetFontColor g_Themer.GetGenericUIColor(UI_TextClickableSelected)
        m_FontHotkey.SetFontColor g_Themer.GetGenericUIColor(UI_TextClickableSelected)
    Else
        curFont.SetFontColor g_Themer.GetGenericUIColor(UI_TextClickableUnselected, , , itemIsHovered)
        m_FontHotkey.SetFontColor g_Themer.GetGenericUIColor(UI_TextClickableUnselected, , , itemIsHovered)
    End If
    
    'Prepare the rendering text
    Dim drawString As String
    drawString = m_Items(itemIndex).hk_TextLocalized
    
    'Render the text
    If (LenB(drawString) <> 0) Then
        curFont.AttachToDC bufferDC
        curFont.FastRenderTextWithClipping offsetX, offsetY + Interface.FixDPI(4), leftOffsetHotkey - tmpRectF.Left, tmpRectF.Height, drawString, True, False, False
        curFont.ReleaseFromDC
    End If
    
    'Next, solve for the on-screen size of the hotkey text
    Dim keyCodeToUse As Long, keyComboText As String
    keyCodeToUse = m_Items(itemIndex).hk_KeyCode
    keyComboText = m_Items(itemIndex).hk_HotkeyText
    
    If (keyCodeToUse <> 0) Then
        
        'Right-align the hotkey text in the drop-down area, with a little padding
        m_FontHotkey.AttachToDC bufferDC
        m_FontHotkey.FastRenderText leftOffsetHotkey, offsetY + Interface.FixDPI(4), keyComboText
        m_FontHotkey.ReleaseFromDC
        
        'If this hotkey has been flagged as a duplicate, render a warning icon
        If m_Items(itemIndex).hk_DuplicateFound Then
            
            Dim icoSize As Long
            icoSize = Interface.FixDPI(20)
            
            If (m_WarningIcon Is Nothing) Then
                Set m_WarningIcon = New pdDIB
                If (Not IconsAndCursors.LoadResourceToDIB("generic_warning", m_WarningIcon, icoSize, icoSize, 0)) Then Set m_WarningIcon = Nothing
            End If
            
            If (Not m_WarningIcon Is Nothing) Then
                m_WarningIcon.AlphaBlendToDC bufferDC, 255, leftOffsetHotkey - icoSize - Interface.FixDPI(12), tmpRectF.Top + (tmpRectF.Height - icoSize) \ 2
            End If
            
        End If
            
    End If
    
End Sub

Private Sub tvMenus_MouseOver(ByVal itemIndex As Long, itemTextEn As String)
    
    'If the current treeview item has a duplicate hotkey, warn the user.
    If (itemIndex >= 0) And (itemIndex < m_numItems) Then
        If (m_Items(itemIndex).hk_KeyCode <> 0) And m_Items(itemIndex).hk_DuplicateFound Then
            
            Dim toolMsg As String
            toolMsg = g_Language.TranslateMessage("Please use this hotkey on just one action:") & vbCrLf
            
            'Start by adding the *current* item to the list
            Dim hkItemName As String
            hkItemName = m_Items(itemIndex).hk_TextLocalized
            
            Dim idxSource As Long, idxTarget As Variant
            If (m_Items(itemIndex).hk_NumParents > 0) Then
                
                idxSource = itemIndex
                Do While m_ActionHash.GetItemByKey(m_Items(idxSource).hk_ParentID, idxTarget)
                    If (idxTarget = idxSource) Then Exit Do     'Failsafe only
                    hkItemName = m_Items(idxTarget).hk_TextLocalized & " > " & hkItemName
                    If (m_Items(idxTarget).hk_NumParents > 0) Then
                        idxSource = idxTarget
                    Else
                        Exit Do
                    End If
                Loop
                
            End If
            
            'Append this menu to the list of "duplicates"
            toolMsg = toolMsg & vbCrLf & hkItemName & " " & g_Language.TranslateMessage("(this item)")
            
            Dim i As Long
            For i = 0 To m_numItems - 1
                If (i <> itemIndex) Then
                    If (m_Items(i).hk_KeyCode = m_Items(itemIndex).hk_KeyCode) And (m_Items(i).hk_ShiftState = m_Items(itemIndex).hk_ShiftState) Then
                        
                        'This item uses the same hotkey as the target item.  Generate a string for it.
                        hkItemName = m_Items(i).hk_TextLocalized
                        If (m_Items(i).hk_NumParents > 0) Then
                            
                            idxSource = i
                            Do While m_ActionHash.GetItemByKey(m_Items(idxSource).hk_ParentID, idxTarget)
                                If (idxTarget = idxSource) Then Exit Do     'Failsafe only
                                hkItemName = m_Items(idxTarget).hk_TextLocalized & " > " & hkItemName
                                If (m_Items(idxTarget).hk_NumParents > 0) Then
                                    idxSource = idxTarget
                                Else
                                    Exit Do
                                End If
                            Loop
                            
                        End If
                        
                        'Append this menu to the list of "duplicates"
                        toolMsg = toolMsg & vbCrLf & hkItemName
                        
                    End If
                End If
            Next i
            
            Dim toolTitle As String
            toolTitle = g_Language.TranslateMessage("Hotkeys must be unique.")
            tvMenus.AssignTooltip toolMsg, toolTitle, True
        Else
            tvMenus.AssignTooltip vbNullString, vbNullString, False
        End If
    Else
        tvMenus.AssignTooltip vbNullString, vbNullString, False
    End If
    
End Sub

Private Sub tvMenus_ScrollOccurred()
    HideEditBox
End Sub

Private Sub txtHotkey_KeyDown(ByVal Shift As ShiftConstants, ByVal vKey As Long, preventFurtherHandling As Boolean)
    
    'Failsafe only
    If (tvMenus.ListIndex < 0) Then Exit Sub
    
    'Prevent circular updates
    m_inAutoUpdate = True
    
    'Get a text representation of the hotkey as it currently appears, and reflect that in the edit box
    Dim hkAsText As String
    hkAsText = AutoTextKeyChange(Shift, vKey)
    txtHotkey.Text = hkAsText
    
    'Only update the stored keycode on non-ctrl/alt/shift presses.
    If (vKey <> VK_SHIFT) And (vKey <> VK_ALT) And (vKey <> VK_CONTROL) Then
        
        'Build a string for Ctrl/Alt/Shift, and ensure the checkboxes at the bottom reflect the current state.
        With m_Items(tvMenus.ListIndex)
            .hk_HotkeyText = hkAsText
            .hk_ShiftState = Shift
            .hk_KeyCode = vKey
        End With
        
    End If
    
    preventFurtherHandling = True
    
End Sub

Private Sub txtHotkey_KeyPress(ByVal Shift As ShiftConstants, ByVal vKey As Long, preventFurtherHandling As Boolean)
    preventFurtherHandling = True
End Sub

Private Sub txtHotkey_KeyUp(ByVal Shift As ShiftConstants, ByVal vKey As Long, preventFurtherHandling As Boolean)
    
    preventFurtherHandling = True
    m_inAutoUpdate = False
    
    'Display a warning symbol if this hotkey is already used elsewhere
    FlagAllDuplicates
    tvMenus.RequestListRedraw
    
End Sub

Private Sub txtHotkey_LostFocusAPI()
    
    'To avoid issues at startup, manually check visibility (this will always be *true* after startup)
    If txtHotkey.Visible Then
        
        'Failsafe checks only
        If (m_idxLastHotkey >= 0) And (Not m_inAutoUpdate) Then
            
            'Make sure the user entered a valid hotkey
            If (m_Items(m_idxLastHotkey).hk_KeyCode = 0) Then
                
                'The user didn't enter a hotkey.  Restore the previous selection, if any.
                With m_Items(m_idxLastHotkey)
                    .hk_HotkeyText = m_backupHotkeyText
                    .hk_KeyCode = m_backupHotkeyVKCode
                    .hk_ShiftState = m_backupHotkeyShift
                End With
                
                'Because shift/ctrl/alt modifiers have changed, we also need to reset the on-screen UI for these keys
                If (tvMenus.ListIndex = m_idxLastHotkey) Then AutoTextKeyChange m_backupHotkeyShift, m_backupHotkeyVKCode
                
            End If
            
            'Flag duplicates again, just in case
            FlagAllDuplicates
            
        End If
        
        'Hide the textbox and request a redraw to ensure on-screen state matches any changes made via edit box
        txtHotkey.Visible = False
        tvMenus.RequestListRedraw
        
    End If
    
End Sub

'Update the bottom "manual" controls to reflect current keystate.
' Returns: a string reflecting the hotkey key state.
Private Function AutoTextKeyChange(ByVal Shift As ShiftConstants, ByVal vKey As Long) As String
    
    Dim newText As String
    
    If ((Shift And vbCtrlMask) <> 0) Then
        newText = newText & Hotkeys.GetGenericMenuText(cmt_Ctrl) & "+"
        Me.chkModifier(0).Value = True
    Else
        Me.chkModifier(0).Value = False
    End If
    If ((Shift And vbAltMask) <> 0) Then
        newText = newText & Hotkeys.GetGenericMenuText(cmt_Alt) & "+"
        Me.chkModifier(1).Value = True
    Else
        Me.chkModifier(1).Value = False
    End If
    If ((Shift And vbShiftMask) <> 0) Then
        newText = newText & Hotkeys.GetGenericMenuText(cmt_Shift) & "+"
        Me.chkModifier(2).Value = True
    Else
        Me.chkModifier(2).Value = False
    End If
    
    'Retrieve the system name for this key, but *only* if it's not a Ctrl/Alt/Shift modifier.
    If (vKey <> VK_SHIFT) And (vKey <> VK_ALT) And (vKey <> VK_CONTROL) Then
        
        Dim keyName As String
        keyName = Hotkeys.GetCharFromKeyCode(vKey)
        newText = newText & keyName
        
        Dim i As Long, keyFound As Boolean
        For i = 0 To m_idxOtherHotkey - 1
            If (m_possibleHotkeys(i).ph_VKCode = vKey) Then
                keyFound = True
                Exit For
            End If
        Next i
        
        If keyFound Then
            Me.ddKey.ListIndex = i
        Else
            If (vKey > 0) Then
                Me.ddKey.ListIndex = m_idxOtherHotkey
            Else
                Me.ddKey.ListIndex = m_idxNoneHotkey
            End If
        End If
        
    End If
    
    AutoTextKeyChange = newText
    
End Function

'Hide the hotkey edit box (if visible).  Make sure you commit any pending hotkey changes *before* calling this.
' (NOTE: previously, this function would also commit pending hotkey changes.  This has since been removed in favor
'  of committing on KeyDown; this allows me to immediately display a warning symbol for duplicate hotkeys.)
Private Sub HideEditBox()
    If (Not txtHotkey.Visible) Then Exit Sub
    txtHotkey.Visible = False
End Sub

Private Sub GeneratePossibleHotkeys()
    
    'This list is manually generated from https://learn.microsoft.com/en-us/windows/win32/inputdev/virtual-key-codes
    AddPossibleHotkey &H8, "BACKSPACE key"
    AddPossibleHotkey &H9, "TAB key"
    AddPossibleHotkey &HC, "CLEAR key"
    AddPossibleHotkey &HD, "ENTER key"
    AddPossibleHotkey &H13, "PAUSE key"
    AddPossibleHotkey &H14, "CAPS LOCK key"
    AddPossibleHotkey &H1B, "ESC key"
    AddPossibleHotkey &H20, "SPACEBAR"
    AddPossibleHotkey &H21, "PAGE UP key"
    AddPossibleHotkey &H22, "PAGE DOWN key"
    AddPossibleHotkey &H23, "END key"
    AddPossibleHotkey &H24, "HOME key"
    AddPossibleHotkey &H25, "LEFT ARROW key"
    AddPossibleHotkey &H26, "UP ARROW key"
    AddPossibleHotkey &H27, "RIGHT ARROW key"
    AddPossibleHotkey &H28, "DOWN ARROW key"
    AddPossibleHotkey &H29, "SELECT key"
    AddPossibleHotkey &H2A, "PRINT key"
    AddPossibleHotkey &H2B, "EXECUTE key"
    AddPossibleHotkey &H2C, "PRINT SCREEN key"
    AddPossibleHotkey &H2D, "INS key"
    AddPossibleHotkey &H2E, "DEL key"
    AddPossibleHotkey &H2F, "HELP key"
    AddPossibleHotkey &H30, "0 key"
    AddPossibleHotkey &H31, "1 key"
    AddPossibleHotkey &H32, "2 key"
    AddPossibleHotkey &H33, "3 key"
    AddPossibleHotkey &H34, "4 key"
    AddPossibleHotkey &H35, "5 key"
    AddPossibleHotkey &H36, "6 key"
    AddPossibleHotkey &H37, "7 key"
    AddPossibleHotkey &H38, "8 key"
    AddPossibleHotkey &H39, "9 key"
    AddPossibleHotkey &H41, "A key"
    AddPossibleHotkey &H42, "B key"
    AddPossibleHotkey &H43, "C key"
    AddPossibleHotkey &H44, "D key"
    AddPossibleHotkey &H45, "E key"
    AddPossibleHotkey &H46, "F key"
    AddPossibleHotkey &H47, "G key"
    AddPossibleHotkey &H48, "H key"
    AddPossibleHotkey &H49, "I key"
    AddPossibleHotkey &H4A, "J key"
    AddPossibleHotkey &H4B, "K key"
    AddPossibleHotkey &H4C, "L key"
    AddPossibleHotkey &H4D, "M key"
    AddPossibleHotkey &H4E, "N key"
    AddPossibleHotkey &H4F, "O key"
    AddPossibleHotkey &H50, "P key"
    AddPossibleHotkey &H51, "Q key"
    AddPossibleHotkey &H52, "R key"
    AddPossibleHotkey &H53, "S key"
    AddPossibleHotkey &H54, "T key"
    AddPossibleHotkey &H55, "U key"
    AddPossibleHotkey &H56, "V key"
    AddPossibleHotkey &H57, "W key"
    AddPossibleHotkey &H58, "X key"
    AddPossibleHotkey &H59, "Y key"
    AddPossibleHotkey &H5A, "Z key"
    AddPossibleHotkey &H5B, "Left Windows key"
    AddPossibleHotkey &H5C, "Right Windows key"
    AddPossibleHotkey &H5D, "Applications key"
    AddPossibleHotkey &H5F, "Computer Sleep key"
    AddPossibleHotkey &H60, "Numeric keypad 0 key"
    AddPossibleHotkey &H61, "Numeric keypad 1 key"
    AddPossibleHotkey &H62, "Numeric keypad 2 key"
    AddPossibleHotkey &H63, "Numeric keypad 3 key"
    AddPossibleHotkey &H64, "Numeric keypad 4 key"
    AddPossibleHotkey &H65, "Numeric keypad 5 key"
    AddPossibleHotkey &H66, "Numeric keypad 6 key"
    AddPossibleHotkey &H67, "Numeric keypad 7 key"
    AddPossibleHotkey &H68, "Numeric keypad 8 key"
    AddPossibleHotkey &H69, "Numeric keypad 9 key"
    AddPossibleHotkey &H6A, "Multiply key"
    AddPossibleHotkey &H6B, "Add key"
    AddPossibleHotkey &H6C, "Separator key"
    AddPossibleHotkey &H6D, "Subtract key"
    AddPossibleHotkey &H6E, "Decimal key"
    AddPossibleHotkey &H6F, "Divide key"
    AddPossibleHotkey &H70, "F1 key"
    AddPossibleHotkey &H71, "F2 key"
    AddPossibleHotkey &H72, "F3 key"
    AddPossibleHotkey &H73, "F4 key"
    AddPossibleHotkey &H74, "F5 key"
    AddPossibleHotkey &H75, "F6 key"
    AddPossibleHotkey &H76, "F7 key"
    AddPossibleHotkey &H77, "F8 key"
    AddPossibleHotkey &H78, "F9 key"
    AddPossibleHotkey &H79, "F10 key"
    AddPossibleHotkey &H7A, "F11 key"
    AddPossibleHotkey &H7B, "F12 key"
    AddPossibleHotkey &H7C, "F13 key"
    AddPossibleHotkey &H7D, "F14 key"
    AddPossibleHotkey &H7E, "F15 key"
    AddPossibleHotkey &H7F, "F16 key"
    AddPossibleHotkey &H80, "F17 key"
    AddPossibleHotkey &H81, "F18 key"
    AddPossibleHotkey &H82, "F19 key"
    AddPossibleHotkey &H83, "F20 key"
    AddPossibleHotkey &H84, "F21 key"
    AddPossibleHotkey &H85, "F22 key"
    AddPossibleHotkey &H86, "F23 key"
    AddPossibleHotkey &H87, "F24 key"
    AddPossibleHotkey &H90, "NUM LOCK key"
    AddPossibleHotkey &H91, "SCROLL LOCK key"
    AddPossibleHotkey &H92, "OEM specific"
    AddPossibleHotkey &H93, "OEM specific"
    AddPossibleHotkey &H94, "OEM specific"
    AddPossibleHotkey &H95, "OEM specific"
    AddPossibleHotkey &H96, "OEM specific"
    AddPossibleHotkey &HA6, "Browser Back key"
    AddPossibleHotkey &HA7, "Browser Forward key"
    AddPossibleHotkey &HA8, "Browser Refresh key"
    AddPossibleHotkey &HA9, "Browser Stop key"
    AddPossibleHotkey &HAA, "Browser Search key"
    AddPossibleHotkey &HAB, "Browser Favorites key"
    AddPossibleHotkey &HAC, "Browser Start and Home key"
    AddPossibleHotkey &HAD, "Volume Mute key"
    AddPossibleHotkey &HAE, "Volume Down key"
    AddPossibleHotkey &HAF, "Volume Up key"
    AddPossibleHotkey &HB0, "Next Track key"
    AddPossibleHotkey &HB1, "Previous Track key"
    AddPossibleHotkey &HB2, "Stop Media key"
    AddPossibleHotkey &HB3, "Play/Pause Media key"
    AddPossibleHotkey &HB4, "Start Mail key"
    AddPossibleHotkey &HB5, "Select Media key"
    AddPossibleHotkey &HB6, "Start Application 1 key"
    AddPossibleHotkey &HB7, "Start Application 2 key"
    AddPossibleHotkey &HBA, "Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the ;: key"
    AddPossibleHotkey &HBB, "For any country/region, the + key"
    AddPossibleHotkey &HBC, "For any country/region, the , key"
    AddPossibleHotkey &HBD, "For any country/region, the - key"
    AddPossibleHotkey &HBE, "For any country/region, the . key"
    AddPossibleHotkey &HBF, "Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the /? key"
    AddPossibleHotkey &HC0, "Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the `~ key"
    AddPossibleHotkey &HDB, "Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the [{ key"
    AddPossibleHotkey &HDC, "Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the \\| key"
    AddPossibleHotkey &HDD, "Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the ]} key"
    AddPossibleHotkey &HDE, "Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the '""key"
    AddPossibleHotkey &HDF, "Used for miscellaneous characters; it can vary by keyboard."
    AddPossibleHotkey &HE1, "OEM specific"
    AddPossibleHotkey &HE2, "The <> keys on the US standard keyboard, or the \\| key on the non-US 102-key keyboard"
    AddPossibleHotkey &HE3, "OEM specific"
    AddPossibleHotkey &HE4, "OEM specific"
    AddPossibleHotkey &HE6, "OEM specific"
    AddPossibleHotkey &HE9, "OEM specific"
    AddPossibleHotkey &HEA, "OEM specific"
    AddPossibleHotkey &HEB, "OEM specific"
    AddPossibleHotkey &HEC, "OEM specific"
    AddPossibleHotkey &HED, "OEM specific"
    AddPossibleHotkey &HEE, "OEM specific"
    AddPossibleHotkey &HEF, "OEM specific"
    AddPossibleHotkey &HF1, "OEM specific"
    AddPossibleHotkey &HF2, "OEM specific"
    AddPossibleHotkey &HF3, "OEM specific"
    AddPossibleHotkey &HF4, "OEM specific"
    AddPossibleHotkey &HF5, "OEM specific"
    AddPossibleHotkey &HF6, "Attn key"
    AddPossibleHotkey &HF7, "CrSel key"
    AddPossibleHotkey &HF8, "ExSel key"
    AddPossibleHotkey &HFA, "Play key"
    AddPossibleHotkey &HFB, "Zoom key"
    AddPossibleHotkey &HFD, "PA1 key"
    AddPossibleHotkey &HFE, "Clear key"
    
    'Do a quick insertion sort.  NAmes Points are likely to be somewhat close to sorted, as e.g. A-Z are added in order.
    Dim tmpSortKey As PD_PossibleHotkey, searchCont As Boolean
    
    Dim i As Long, j As Long
    i = 1
    
    Do While (i < m_numPossibleHotkeys)
        tmpSortKey = m_possibleHotkeys(i)
        j = i - 1
        
        'Because VB6 doesn't short-circuit And statements, we split this check into separate parts.
        searchCont = False
        If (j >= 0) Then searchCont = (Strings.StrCompSortPtr(StrPtr(m_possibleHotkeys(j).ph_KeyName), StrPtr(tmpSortKey.ph_KeyName)) > 0)
        
        Do While searchCont
            m_possibleHotkeys(j + 1) = m_possibleHotkeys(j)
            j = j - 1
            searchCont = False
            If (j >= 0) Then searchCont = (Strings.StrCompSortPtr(StrPtr(m_possibleHotkeys(j).ph_KeyName), StrPtr(tmpSortKey.ph_KeyName)) > 0)
        Loop
        
        m_possibleHotkeys(j + 1) = tmpSortKey
        i = i + 1
        
    Loop
    
    'Manually add an "other" key to the list, which we'll use for miscellaneous keypresses that the keyboard driver can't name
    AddPossibleHotkey &HFF, "(other)", "(other)"
    m_idxOtherHotkey = m_numPossibleHotkeys - 1
    
    '...and a "none" hotkey
    AddPossibleHotkey 0, "(none)", "(none)"
    m_idxNoneHotkey = m_numPossibleHotkeys - 1
    
    m_idxFinalHotkey = m_numPossibleHotkeys - 1
    ReDim Preserve m_possibleHotkeys(0 To m_idxFinalHotkey) As PD_PossibleHotkey
    
    'Add all items to the on-screen dropdown
    For i = 0 To m_numPossibleHotkeys - 1
        ddKey.AddItem m_possibleHotkeys(i).ph_KeyName
    Next i
    
    'Select the (none) entry
    ddKey.ListIndex = m_idxNoneHotkey
    
End Sub

Private Sub AddPossibleHotkey(ByVal vkCode As Long, Optional ByRef keyComments As String = vbNullString, Optional ByRef manualKeyName As String = vbNullString)
    
    If (m_numPossibleHotkeys = 0) Then
        Const INIT_POSSIBLE_HOTKEYS As Long = 64
        ReDim m_possibleHotkeys(0 To INIT_POSSIBLE_HOTKEYS) As PD_PossibleHotkey
    End If
    
    If (g_Language Is Nothing) Then Exit Sub
    
    'See if this key exists on this keyboard (null names mean key doesn't exist, typically)
    Dim keyName As String, keyNameExtended As String
    If (LenB(manualKeyName) > 0) Then
        keyName = manualKeyName
    Else
        
        Select Case vkCode
            
            'Some unreadable chars have to be manually entered
            Case 8
                keyName = g_Language.TranslateMessage("Backspace")
            Case 9
                keyName = g_Language.TranslateMessage("Tab")
            Case &H1B
                keyName = g_Language.TranslateMessage("Escape")
            
            'Other ones can be pulled from the keyboard driver
            Case Else
                Hotkeys.GetCharFromKeyCode vkCode, outKeyName:=keyName, outKeyNameExtended:=keyNameExtended
                
        End Select
        
    End If
    
    'Ignore blank names (those are likely keys that do not exist)
    If (LenB(keyName) > 0) Then
        
        'Iterate previous entries and skip duplicates.  (OEMs may use OEM-specific keycodes to duplicate standard keycodes.)
        If (m_numPossibleHotkeys > 0) Then
            Dim i As Long
            For i = 0 To m_numPossibleHotkeys - 1
                If Strings.StringsEqual(keyName, m_possibleHotkeys(i).ph_KeyName) Then
                    m_possibleHotkeys(i).ph_VKCode = PDMath.Min2Int(vkCode, m_possibleHotkeys(i).ph_VKCode)
                    Exit Sub
                End If
            Next i
        End If
        
        With m_possibleHotkeys(m_numPossibleHotkeys)
            .ph_VKCode = vkCode
            .ph_KeyName = keyName
            .ph_KeyComments = keyComments
        End With
        
        m_numPossibleHotkeys = m_numPossibleHotkeys + 1
        If (m_numPossibleHotkeys > UBound(m_possibleHotkeys)) Then ReDim Preserve m_possibleHotkeys(0 To m_numPossibleHotkeys * 2 - 1) As PD_PossibleHotkey
        
    End If
    
    'Repeat previous steps for extended key name
    If (LenB(keyNameExtended) > 0) Then
        
        If (m_numPossibleHotkeys > 0) Then
            For i = 0 To m_numPossibleHotkeys - 1
                If Strings.StringsEqual(keyNameExtended, m_possibleHotkeys(i).ph_KeyName) Then
                    m_possibleHotkeys(i).ph_VKCode = PDMath.Min2Int(vkCode, m_possibleHotkeys(i).ph_VKCode)
                    Exit Sub
                End If
            Next i
        End If
        
        With m_possibleHotkeys(m_numPossibleHotkeys)
            .ph_VKCode = vkCode
            .ph_KeyName = keyNameExtended
            .ph_KeyComments = keyComments
        End With
        
        m_numPossibleHotkeys = m_numPossibleHotkeys + 1
        If (m_numPossibleHotkeys > UBound(m_possibleHotkeys)) Then ReDim Preserve m_possibleHotkeys(0 To m_numPossibleHotkeys * 2 - 1) As PD_PossibleHotkey
        
    End If
    
End Sub

Private Function GetHotkeyNameFromKeys(ByVal Shift As ShiftConstants, ByVal vKey As Long) As String
    
    Dim newText As String
    If ((Shift And vbCtrlMask) <> 0) Then newText = newText & Hotkeys.GetGenericMenuText(cmt_Ctrl) & "+"
    If ((Shift And vbAltMask) <> 0) Then newText = newText & Hotkeys.GetGenericMenuText(cmt_Alt) & "+"
    If ((Shift And vbShiftMask) <> 0) Then newText = newText & Hotkeys.GetGenericMenuText(cmt_Shift) & "+"
    
    GetHotkeyNameFromKeys = newText & Hotkeys.GetCharFromKeyCode(vKey)
    
End Function

Private Sub GenerateSummaryFile(ByRef dstFile As String)
        
    Files.FileDeleteIfExists dstFile
    
    'This string will be large, so use a string builder object
    Dim cExport As pdString
    Set cExport = New pdString
    
    'Start by populating the string builder with some boilerplate HTML
    Const HTML_BP_BASE64 As String = "PCFET0NUWVBFIGh0bWw+DQo8aHRtbCBsYW5nPSIlMSI+DQo8aGVhZD4NCjxtZXRhIGNoYXJzZXQ9InV0Zi04Ij4NCjx0aXRsZT4lMjwvdGl0bGU+DQoJPHN0eWxlPg"
    Const HTML_BP_BASE64_2 As String = "dGFibGUgew0KICAgICAgICAgICAgd2lkdGg6IDEwMCU7DQogICAgICAgICAgICBib3JkZXItY29sbGFwc2U6IGNvbGxhcHNlOw0KICAgICAgICAgICAgbWFyZ2luOiAyNXB4IGF1dG8gMDsNCiAgICAgICAgICAgIGZvbnQtc2l6ZTogMWVtOw0KICAgICAgICAgICAgZm9udC1mYW1pbHk6IHNhbnMtc2VyaWY7DQogICAgICAgICAgICBtYXgtd2lkdGg6IDEyMDBweDsNCgkJCW1pbi13aWR0aDogNDAwcHg7DQogICAgICAgICAgICBib3gtc2hhZG93OiAwIDAgMjBweCByZ2JhKDAsIDAsIDAsIDAuMTUpOw0KICAgICAgICB9DQogICAgICAgIHRoLCB0ZCB7DQogICAgICAgICAgICBwYWRkaW5nOiA4cHggMTVweDsNCiAgICAgICAgICAgIGJvcmRlcjogMXB4IHNvbGlkICNkZGQ7DQogICAgICAgICAgICB0ZXh0LWFsaWduOiBsZWZ0Ow0KICAgICAgICB9DQogICAgICAgIHRoZWFkIHsNCiAgICAgICAgICAgIGJhY2tncm91bmQtY29sb3I6ICMlMTsNCiAgICAgICAgICAgIGNvbG9yOiAjZmZmZmZmOw0KICAgICAgICB9DQogICAgICAgIHRib2R5IHRyOm50aC1jaGlsZChldmVuKSB7DQogICAgICAgICAgICBiYWNrZ3JvdW5kLWNvbG9yOiAjZjRmNGY0Ow0KICAgICAgICB9DQogICAgICAgIHRib2R5IHRyOmxhc3Qtb2YtdHlwZSB7DQogICAgICAgICAgICBib3JkZXItYm90dG9tOiA0cHggc29saWQgIyUxOw0KICAgICAgICB9"
    Const HTML_BP_BASE64_3 As String = "aDIgew0KCQkJbWFyZ2luOiAxNXB4IGF1dG8gMjVweDsNCiAgICAgICAgICAgIGZvbnQtc2l6ZTogMi41ZW07DQogICAgICAgICAgICBmb250LWZhbWlseTogc2Fucy1zZXJpZjsNCgkJCXRleHQtYWxpZ246IGNlbnRlcjsNCgkJfQ"
    Const HTML_BP_BASE64_4 As String = "DQogICAgPC9zdHlsZT4NCjwvaGVhZD4NCiAgPGJvZHk+"
    
    Dim utf8Bytes() As Byte
    Strings.BytesFromBase64 utf8Bytes, HTML_BP_BASE64
    
    Dim initHTML As String
    initHTML = Strings.StringFromUTF8(utf8Bytes)
    
    'Replace some placeholders with current user settings
    initHTML = Replace$(initHTML, "%1", g_Language.GetCurrentLanguage(False))
    initHTML = Replace$(initHTML, "%2", g_Language.TranslateMessage("PhotoDemon Hotkeys"))
    
    'Append remaining style bits
    Strings.BytesFromBase64 utf8Bytes, HTML_BP_BASE64_2
    initHTML = initHTML & vbCrLf & Replace$(Strings.StringFromUTF8(utf8Bytes), "%1", Colors.GetHexStringFromRGB(g_Themer.GetGenericUIColor(UI_Accent)))
    Strings.BytesFromBase64 utf8Bytes, HTML_BP_BASE64_3
    initHTML = initHTML & Strings.StringFromUTF8(utf8Bytes)
    Strings.BytesFromBase64 utf8Bytes, HTML_BP_BASE64_4
    initHTML = initHTML & Strings.StringFromUTF8(utf8Bytes)
    
    cExport.Append initHTML
    Erase utf8Bytes
    initHTML = vbNullString
    
    'Title
    cExport.AppendLine "<h2>" & g_Language.TranslateMessage("PhotoDemon Hotkeys") & "</h2>"
    
    'Append a minimalist header
    cExport.AppendLine "<table>"
    cExport.AppendLine "<thead><tr><th>" & g_Language.TranslateMessage("Command") & "</th><th>" & g_Language.TranslateMessage("Hotkey") & "</th></thead><tbody>"
    
    'Now, append all menus and hotkeys
    Dim i As Long
    For i = 0 To m_numItems - 1
        Const HTML_TABLE_ROW As String = "<tr>"
        cExport.AppendLine HTML_TABLE_ROW
        With m_Items(i)
            Const HTML_TABLE_CELL As String = "<td>"
            Const HTML_TABLE_CELL_END As String = "</td>"
            cExport.Append HTML_TABLE_CELL
            If (.hk_NumParents > 0) Then
                Dim j As Long
                For j = 0 To .hk_NumParents
                    Const HTML_SPACER As String = "&nbsp;&nbsp;&nbsp;&nbsp;"
                    cExport.Append HTML_SPACER
                Next j
            End If
            cExport.Append .hk_TextLocalized
            cExport.AppendLine HTML_TABLE_CELL_END
            cExport.Append HTML_TABLE_CELL
            If (.hk_KeyCode <> 0) Then cExport.Append .hk_HotkeyText
            cExport.AppendLine HTML_TABLE_CELL_END
        End With
        Const HTML_TABLE_ROW_END As String = "</tr>"
        cExport.AppendLine HTML_TABLE_ROW_END
    Next i
    
    'Close any open tags
    cExport.AppendLine "</tbody></table></body></html>"
    
    'Write the text out to file
    Files.FileSaveAsText cExport.ToString(), dstFile, True, False
    
End Sub

'Search the current hotkey list for duplicates.  *IMPORTANTLY*, this function only marks a duplicate if it...
' 1) has the same hotkey as idxSource, and...
' 2) the matched item(s) map to DIFFERENT ACTIONS.
'
'If two menus map to the same underlying action (e.g. the Adjustents > Curves, and Adjustments > Color > Curves menus),
' it is not just fine, but *expected* for them to have the same hotkey.  (And in fact, the menu editor handles
' these mappings automagically.)  So these cases are *not* returned as duplicates.
'
'RETURNS: TRUE if invalid duplicates were found; FALSE if no invalid duplicates were found.
'
'The passed stack will hold indices of any matches.  Its value is indeterminate on a FALSE return.
Private Function FindDuplicates(ByVal idxSource As Long, ByRef dstIdxList As pdStack) As Boolean
    
    FindDuplicates = False
    
    If (dstIdxList Is Nothing) Then
        Set dstIdxList = New pdStack
    Else
        dstIdxList.ResetStack 2
    End If
    
    Dim targetHotkey As String
    targetHotkey = m_Items(idxSource).hk_HotkeyText
    
    Dim i As Long
    For i = 0 To m_numItems - 1
        If (i <> idxSource) Then
            If Strings.StringsEqual(targetHotkey, m_Items(i).hk_HotkeyText, True) Then
                
                'Compare action IDs; we only care if these are *mismatched*
                If Strings.StringsNotEqual(m_Items(idxSource).hk_ActionID, m_Items(i).hk_ActionID, True) Then
                    FindDuplicates = True
                    dstIdxList.AddInt i
                End If
                
            End If
        End If
    Next i
    
End Function

'Correctly set all "this hotkey is used elsewhere" flags in m_Items().  This uses a naive double-loop, and could easily be
' accelerated with hash tables.
Private Sub FlagAllDuplicates()
    
    Dim i As Long, j As Long
    For i = 0 To m_numItems - 1
        m_Items(i).hk_DuplicateFound = False
    Next i
    
    For i = 0 To m_numItems - 1
        If (m_Items(i).hk_KeyCode <> 0) Then
            For j = i + 1 To m_numItems - 1
                If (i <> j) Then
                    If IsThisAnInvalidDuplicate(i, j) Then
                        m_Items(i).hk_DuplicateFound = True
                        m_Items(j).hk_DuplicateFound = True
                    End If
                End If
            Next j
        End If
    Next i
    
End Sub

'Compare two items, and return TRUE if they share a hotkey, but map to different actions.
Private Function IsThisAnInvalidDuplicate(ByVal srcIndex1 As Long, ByVal srcIndex2 As Long) As Boolean
    
    'Ignore requests on the same index!
    If (srcIndex1 <> srcIndex2) Then
        
        'Hotkey text must be correctly filled for this function to work
        If Strings.StringsEqual(m_Items(srcIndex1).hk_HotkeyText, m_Items(srcIndex2).hk_HotkeyText, True) Then
            
            'Compare action IDs; we only care if these are *mismatched*.  (Sometimes, two different menus map to the same action;
            ' e.g. this can happen in the Adjustments menu for top-level shortcuts and formally organized ops - these items share
            ' the same hotkey *by design* because they do the *same thing*.)
            IsThisAnInvalidDuplicate = Strings.StringsNotEqual(m_Items(srcIndex1).hk_ActionID, m_Items(srcIndex2).hk_ActionID, True)
            
        End If
    End If
    
End Function

'Present a generic dialog that warns all hotkeys are about to be overwritten.  We can show this before any "reset" action.
Private Function OkayToOverwriteAll() As Boolean
    Dim msgWarn As String
    msgWarn = g_Language.TranslateMessage("This will overwrite any hotkey edits.  Are you sure you want to continue?")
    OkayToOverwriteAll = (PDMsgBox(msgWarn, vbYesNoCancel Or vbExclamation Or vbApplicationModal, "Warning") = vbYes)
End Function
