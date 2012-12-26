VERSION 5.00
Begin VB.Form FormPluginManager 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " PhotoDemon Plugin Manager"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10770
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
   ScaleHeight     =   410
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   718
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   7920
      TabIndex        =   16
      Top             =   5520
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   9240
      TabIndex        =   15
      Top             =   5520
      Width           =   1245
   End
   Begin VB.ListBox lstPlugins 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   4890
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label lblDisable 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Disable pngnq-s9"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C07031&
      Height          =   240
      Index           =   3
      Left            =   9015
      MouseIcon       =   "VBP_FormPluginManager.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   4245
      Width           =   1470
   End
   Begin VB.Label lblDisable 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Disable EZTwain"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C07031&
      Height          =   240
      Index           =   2
      Left            =   9075
      MouseIcon       =   "VBP_FormPluginManager.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   3165
      Width           =   1410
   End
   Begin VB.Label lblDisable 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Disable zLib"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C07031&
      Height          =   240
      Index           =   1
      Left            =   9480
      MouseIcon       =   "VBP_FormPluginManager.frx":02A4
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   2085
      Width           =   1005
   End
   Begin VB.Label lblDisable 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Disable FreeImage"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C07031&
      Height          =   240
      Index           =   0
      Left            =   8880
      MouseIcon       =   "VBP_FormPluginManager.frx":03F6
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   1005
      Width           =   1605
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Caption         =   "installed, enabled, and up to date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   270
      Index           =   3
      Left            =   4380
      MouseIcon       =   "VBP_FormPluginManager.frx":0548
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   4560
      Width           =   3255
   End
   Begin VB.Label lblInterfaceSubheader 
      AutoSize        =   -1  'True
      Caption         =   "status:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Index           =   3
      Left            =   3600
      MouseIcon       =   "VBP_FormPluginManager.frx":069A
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   4560
      Width           =   675
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Caption         =   "installed, enabled, and up to date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   270
      Index           =   2
      Left            =   4380
      MouseIcon       =   "VBP_FormPluginManager.frx":07EC
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   3480
      Width           =   3255
   End
   Begin VB.Label lblInterfaceSubheader 
      AutoSize        =   -1  'True
      Caption         =   "status:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Index           =   2
      Left            =   3600
      MouseIcon       =   "VBP_FormPluginManager.frx":093E
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   3480
      Width           =   675
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Caption         =   "installed, enabled, and up to date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   270
      Index           =   1
      Left            =   4380
      MouseIcon       =   "VBP_FormPluginManager.frx":0A90
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   2400
      Width           =   3255
   End
   Begin VB.Label lblInterfaceSubheader 
      AutoSize        =   -1  'True
      Caption         =   "status:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Index           =   1
      Left            =   3600
      MouseIcon       =   "VBP_FormPluginManager.frx":0BE2
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   2400
      Width           =   675
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Caption         =   "installed, enabled, and up to date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   270
      Index           =   0
      Left            =   4380
      MouseIcon       =   "VBP_FormPluginManager.frx":0D34
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   1320
      Width           =   3255
   End
   Begin VB.Label lblInterfaceSubheader 
      AutoSize        =   -1  'True
      Caption         =   "status:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Index           =   0
      Left            =   3600
      MouseIcon       =   "VBP_FormPluginManager.frx":0E86
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   1320
      Width           =   675
   End
   Begin VB.Label lblInterfaceTitle 
      AutoSize        =   -1  'True
      Caption         =   "FreeImage"
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
      Left            =   3360
      MouseIcon       =   "VBP_FormPluginManager.frx":0FD8
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   960
      Width           =   1125
   End
   Begin VB.Label lblInterfaceTitle 
      AutoSize        =   -1  'True
      Caption         =   "EZTwain"
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
      Index           =   2
      Left            =   3360
      MouseIcon       =   "VBP_FormPluginManager.frx":112A
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   3120
      Width           =   915
   End
   Begin VB.Label lblInterfaceTitle 
      AutoSize        =   -1  'True
      Caption         =   "zLib"
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
      Index           =   1
      Left            =   3360
      MouseIcon       =   "VBP_FormPluginManager.frx":127C
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2040
      Width           =   420
   End
   Begin VB.Label lblInterfaceTitle 
      AutoSize        =   -1  'True
      Caption         =   "pngnq-s9"
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
      Index           =   3
      Left            =   3360
      MouseIcon       =   "VBP_FormPluginManager.frx":13CE
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   4200
      Width           =   1005
   End
   Begin VB.Label lblPluginStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GOOD"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000B909&
      Height          =   285
      Left            =   5460
      TabIndex        =   1
      Top             =   240
      Width           =   690
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "current plugin status:"
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
      Left            =   3120
      TabIndex        =   0
      Top             =   240
      Width           =   2265
   End
End
Attribute VB_Name = "FormPluginManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Plugin Manager
'Copyright ©2011-2012 by Tanner Helland
'Created: 21/December/12
'Last updated: 26/December/12
'Last update: finished initial build
'
'Dialog for presenting the user data related to the currently installed plugins.
'
'I seriously considered merging this form with the main Preferences (now Options) dialog, but there
' are simply too many settings present.  Rather than clutter up the main Preferences dialog with
' plugin-related settings, I have moved those all here.
'
'In the future, I suppose this could be merged with the plugin updater to form one happy plugin
' handler, but for now it makes sense to make them both available (and to keep them separate).
'
'***************************************************************************

Option Explicit

'Green and red hues for use with our GOOD and BAD labels
Private Const GOODCOLOR As Long = 49152 'RGB(0,192,0)
Private Const BADCOLOR As Long = 192    'RGB(192,0,0)

'Much of the version-checking code used in this form was derived from http://allapi.mentalis.org/apilist/GetFileVersionInfo.shtml
' Many thanks to those authors for their work on demystifying some of these more obscure API calls
Private Type VS_FIXEDFILEINFO
   dwSignature As Long
   dwStrucVersionl As Integer     ' e.g. = &h0000 = 0
   dwStrucVersionh As Integer     ' e.g. = &h0042 = .42
   dwFileVersionMSl As Integer    ' e.g. = &h0003 = 3
   dwFileVersionMSh As Integer    ' e.g. = &h0075 = .75
   dwFileVersionLSl As Integer    ' e.g. = &h0000 = 0
   dwFileVersionLSh As Integer    ' e.g. = &h0031 = .31
   dwProductVersionMSl As Integer ' e.g. = &h0003 = 3
   dwProductVersionMSh As Integer ' e.g. = &h0010 = .1
   dwProductVersionLSl As Integer ' e.g. = &h0000 = 0
   dwProductVersionLSh As Integer ' e.g. = &h0031 = .31
   dwFileFlagsMask As Long        ' = &h3F for version "0.42"
   dwFileFlags As Long            ' e.g. VFF_DEBUG Or VFF_PRERELEASE
   dwFileOS As Long               ' e.g. VOS_DOS_WINDOWS16
   dwFileType As Long             ' e.g. VFT_DRIVER
   dwFileSubtype As Long          ' e.g. VFT2_DRV_KEYBOARD
   dwFileDateMS As Long           ' e.g. 0
   dwFileDateLS As Long           ' e.g. 0
End Type
Private Declare Function GetFileVersionInfo Lib "version" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "version" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "version" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal length As Long)

'This array will contain the full version strings of our various plugins
Dim vString(0 To 3) As String

'If the user presses "cancel", we need to restore the previous enabled/disabled values
Dim pEnabled(0 To 3) As Boolean

Private Sub CollectVersionInfo(ByVal FullFileName As String, ByVal strIndex As Long)
   
   Dim Filename As String, Directory As String
   Dim StrucVer As String, FileVer As String, ProdVer As String
   Dim FileFlags As String, FileOS As String, FileType As String, FileSubType As String

   Dim rc As Long, lDummy As Long, sBuffer() As Byte
   Dim lBufferLen As Long, lVerPointer As Long, udtVerBuffer As VS_FIXEDFILEINFO
   Dim lVerbufferLen As Long

   '*** Get size ****
   lBufferLen = GetFileVersionInfoSize(FullFileName, lDummy)
   If lBufferLen < 1 Then
      'MsgBox "No Version Info available!"
      Exit Sub
   End If

   '**** Store info to udtVerBuffer struct ****
   ReDim sBuffer(lBufferLen)
   rc = GetFileVersionInfo(FullFileName, 0&, lBufferLen, sBuffer(0))
   rc = VerQueryValue(sBuffer(0), "\", lVerPointer, lVerbufferLen)
   MoveMemory udtVerBuffer, lVerPointer, Len(udtVerBuffer)

   '**** Determine Structure Version number - NOT USED ****
   StrucVer = Trim(Format$(udtVerBuffer.dwStrucVersionh)) & "." & Trim(Format$(udtVerBuffer.dwStrucVersionl))

   '**** Determine File Version number ****
   FileVer = Trim(Format$(udtVerBuffer.dwFileVersionMSh)) & "." & Trim(Format$(udtVerBuffer.dwFileVersionMSl)) & "." & Trim(Format$(udtVerBuffer.dwFileVersionLSh)) & "." & Trim(Format$(udtVerBuffer.dwFileVersionLSl))

   '**** Determine Product Version number ****
   ProdVer = Trim(Format$(udtVerBuffer.dwProductVersionMSh)) & "." & Trim(Format$(udtVerBuffer.dwProductVersionMSl)) & "." & Trim(Format$(udtVerBuffer.dwProductVersionLSh)) & "." & Trim(Format$(udtVerBuffer.dwProductVersionLSl))

   vString(strIndex) = ProdVer

End Sub

'CANCEL button
Private Sub cmdCancel_Click()
    
    'Restore the original values for enabled or disabled plugins
    imageFormats.FreeImageEnabled = pEnabled(0)
    zLibEnabled = pEnabled(1)
    ScanEnabled = pEnabled(2)
    imageFormats.pngnqEnabled = pEnabled(3)
    
    Unload Me
End Sub

'OK button
Private Sub cmdOK_Click()
    
    Me.Visible = False
        
    'Write all enabled/disabled plugin changes to the INI file
    If imageFormats.FreeImageEnabled Then
        userPreferences.SetPreference_Boolean "Plugin Preferences", "ForceFreeImageDisable", False
    Else
        userPreferences.SetPreference_Boolean "Plugin Preferences", "ForceFreeImageDisable", True
    End If
            
    'zLib
    If zLibEnabled Then
        userPreferences.SetPreference_Boolean "Plugin Preferences", "ForceZLibDisable", False
    Else
        userPreferences.SetPreference_Boolean "Plugin Preferences", "ForceZLibDisable", True
    End If
        
    'EZTwain
    If ScanEnabled Then
        userPreferences.SetPreference_Boolean "Plugin Preferences", "ForceEZTwainDisable", False
    Else
        userPreferences.SetPreference_Boolean "Plugin Preferences", "ForceEZTwainDisable", True
    End If
        
    'pngnq-s9
    If imageFormats.pngnqEnabled Then
        userPreferences.SetPreference_Boolean "Plugin Preferences", "ForcePngnqDisable", False
    Else
        userPreferences.SetPreference_Boolean "Plugin Preferences", "ForcePngnqDisable", True
    End If
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    'Populate the left-hand list box with all relevant plugins
    lstPlugins.Clear
    lstPlugins.AddItem "Overview", 0
    lstPlugins.AddItem "FreeImage", 1
    lstPlugins.AddItem "zLib", 2
    lstPlugins.AddItem "EZTwain", 3
    lstPlugins.AddItem "pngnq-s9", 4
    
    lstPlugins.ListIndex = 0
    
    'Now, check version numbers of each plugin.  This is more complicated than it needs to be, on account of
    ' each plugin having its own unique mechanism for version-checking, but I have wrapped these various functions
    ' inside fairly standard wrapper calls.
    CollectAllVersionNumbers
    
    'We now have a collection of version numbers for our various plugins.  Let's use those to populate our
    ' "good/bad" labels for each plugin.
    UpdatePluginLabels
    
    'Remember which plugins the user has enabled or disabled
    pEnabled(0) = imageFormats.FreeImageEnabled
    pEnabled(1) = zLibEnabled
    pEnabled(2) = ScanEnabled
    pEnabled(3) = imageFormats.pngnqEnabled
    
    'Apply visual styles
    makeFormPretty Me
    
End Sub

'Assuming version numbers have been successfully retrieved, this function can be called to update the
' green/red plugin label display on the main panel.
Private Sub UpdatePluginLabels()
    
    Dim pluginStatus As Boolean
    
    'FreeImage
    pluginStatus = popPluginLabel(0, "FreeImage", "3.15.4", isFreeImageAvailable, imageFormats.FreeImageEnabled)
    
    'zLib
    pluginStatus = pluginStatus And popPluginLabel(1, "zLib", "1.2.5", isZLibAvailable, zLibEnabled)
    
    'EZTwain
    pluginStatus = pluginStatus And popPluginLabel(2, "EZTwain", "1.18.0", isEZTwainAvailable, ScanEnabled)
    
    'pngnq-s9
    pluginStatus = pluginStatus And popPluginLabel(3, "pngnq-s9", "2.0.1", isPngnqAvailable, imageFormats.pngnqEnabled)
    
    If pluginStatus Then
        lblPluginStatus.ForeColor = GOODCOLOR
        lblPluginStatus.Caption = "GOOD"
    Else
        lblPluginStatus.ForeColor = BADCOLOR
        lblPluginStatus.Caption = "problems detected"
    End If
    
End Sub

'Retrieve all relevant plugin version numbers and store them in the vString() array
Private Sub CollectAllVersionNumbers()

    'Start by analyzing plugin file metadata for version information.  This works for FreeImage and zLib (but
    ' do it for all four, just in case).
    If isFreeImageAvailable Then CollectVersionInfo PluginPath & "freeimage.dll", 0 Else vString(0) = "none"
    If isZLibAvailable Then CollectVersionInfo PluginPath & "zlibwapi.dll", 1 Else vString(1) = "none"
    If isEZTwainAvailable Then CollectVersionInfo PluginPath & "eztw32.dll", 2 Else vString(2) = "none"
    If isPngnqAvailable Then CollectVersionInfo PluginPath & "pngnq-s9.exe", 3 Else vString(3) = "none"
    
    'Special techniques are required for for EZTwain and pngnq-s9.
    
    'The EZTwain DLL provides its own version-checking function
    If isEZTwainAvailable Then vString(2) = getEZTwainVersion Else vString(2) = "none"
    
    If isPngnqAvailable Then vString(3) = getPngnqVersion Else vString(3) = "none"
    
    'Remove trailing build numbers from the version strings
    Dim i As Long
    For i = 0 To 3
        If vString(i) <> "none" Then StripOffExtension vString(i)
    Next i

End Sub

'Given a plugin's availability, expected version, and index on this form, populate the relevant labels associated with it.
' This function will return TRUE if the plugin is in good status, FALSE if it isn't (for any reason)
Private Function popPluginLabel(ByVal curPlugin As Long, ByRef pluginName As String, ByRef expectedVersion As String, ByVal isAvailable As Boolean, ByVal isDisabled As Boolean) As Boolean
        
    'Is this plugin present on the machine?
    If isAvailable Then
    
        'If present, has it been forcibly disabled?
        If Not isDisabled Then
            lblStatus(curPlugin).Caption = "installed"
            lblDisable(curPlugin).Caption = "disable " & pluginName
            
            'If this plugin is present and enabled, does its version match what we expect?
            If StrComp(vString(curPlugin), expectedVersion, vbTextCompare) = 0 Then
                lblStatus(curPlugin).Caption = lblStatus(curPlugin).Caption & " and up to date"
                lblStatus(curPlugin).ForeColor = GOODCOLOR
                popPluginLabel = True
                
            'Version mismatch
            Else
                lblStatus(curPlugin).Caption = lblStatus(curPlugin).Caption & ", but incorrect version (" & vString(0) & " found, " & expectedVersion & " expected)"
                lblStatus(curPlugin).ForeColor = BADCOLOR
                popPluginLabel = False
            End If
            
        'Plugin is disabled
        Else
            lblStatus(curPlugin).Caption = "installed, but disabled by user"
            lblStatus(curPlugin).ForeColor = BADCOLOR
            lblDisable(curPlugin).Caption = "enable " & pluginName
            popPluginLabel = False
        End If
        
    'Plugin is not present on the machine
    Else
        lblStatus(curPlugin).Caption = "missing"
        lblStatus(curPlugin).ForeColor = BADCOLOR
        lblDisable(curPlugin).Visible = False
        popPluginLabel = False
    End If
    
End Function

'The user is now allowed to selectively disable/enable various plugins.  This can be used to test certain program
' parameters, or to force certain behaviors.
Private Sub lblDisable_Click(Index As Integer)

    Select Case Index
    
        'FreeImage
        Case 0
            imageFormats.FreeImageEnabled = Not imageFormats.FreeImageEnabled
            
        'zLib
        Case 1
            zLibEnabled = Not zLibEnabled
            
        'EZTwain
        Case 2
            ScanEnabled = Not ScanEnabled
            
        'pngnq-s9
        Case 3
            imageFormats.pngnqEnabled = Not imageFormats.pngnqEnabled
            
    End Select
    
    'Update the various labels to match the new situation
    UpdatePluginLabels

End Sub
