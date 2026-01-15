VERSION 5.00
Begin VB.Form dialog_ExportDDS 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   6585
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12135
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
   Icon            =   "File_Save_DDS.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   439
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   809
   Begin PhotoDemon.pdButtonStrip btsMipMaps 
      Height          =   975
      Left            =   6120
      TabIndex        =   4
      Top             =   3960
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1720
      Caption         =   "mipmaps"
   End
   Begin PhotoDemon.pdDropDown ddFilter 
      Height          =   495
      Left            =   8880
      TabIndex        =   6
      Top             =   5310
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   873
   End
   Begin PhotoDemon.pdSlider sldMipMaps 
      Height          =   495
      Left            =   6120
      TabIndex        =   5
      Top             =   5280
      Width           =   2655
      _ExtentX        =   10186
      _ExtentY        =   873
      Min             =   1
      Max             =   64
      Value           =   2
      DefaultValue    =   2
   End
   Begin PhotoDemon.pdListBox lstFormat 
      Height          =   2895
      Left            =   6120
      TabIndex        =   3
      Top             =   240
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   5106
      Caption         =   "format"
   End
   Begin PhotoDemon.pdCheckBox chkLivePreview 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   5400
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   661
      Caption         =   "preview quality changes"
      FontSize        =   11
      Value           =   0   'False
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5835
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   1323
      DontAutoUnloadParent=   -1  'True
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   5145
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9075
   End
   Begin PhotoDemon.pdButtonStrip btsCompression 
      Height          =   975
      Left            =   6120
      TabIndex        =   8
      Top             =   3600
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1720
      Caption         =   "compression"
   End
   Begin PhotoDemon.pdButtonStrip btsDither 
      Height          =   975
      Left            =   6120
      TabIndex        =   7
      Top             =   3360
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1720
      Caption         =   "dithering"
   End
End
Attribute VB_Name = "dialog_ExportDDS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'DDS Export Dialog
'Copyright 2025-2026 by Tanner Helland
'Created: 14/May/25
'Last updated: 14/May/25
'Last update: initial build
'
'Dialog for presenting the user a number of options related to DDS exporting.  This feature
' relies on a 3rd-party library for operation (currently DirectXTex); this export dialog is
' not accessible if the required 3rd-party library isn't available.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'This form can (and should!) be notified of the image being exported.  The only exception to this rule is invoking
' the dialog from the batch process dialog, as no image is associated with that preview.
Private m_SrcImage As pdImage

'A composite of the current image, 32-bpp, fully composited.  This is only regenerated if the source image changes.
Private m_CompositedImage As pdDIB

'Clean, modified source preview image, in PNG format.  (This only needs to be created when the preview source
' image changes - e.g. if the user zooms or scrolls the preview control.)
Private m_PreviewImagePath As String, m_PreviewImageBackup As pdDIB

'OK or CANCEL result
Private m_UserDialogAnswer As VbMsgBoxResult

'Final format-specific XML packet, with all format-specific settings defined as tag+value pairs
Private m_FormatParamString As String

'Final metadata XML packet, with all metadata settings defined as tag+value pairs
Private m_MetadataParamString As String

'DDS format names and IDs; we display names, but pass corresponding IDs to DirectXTex
Private m_listOfNames As pdStringStack, m_listOfIDs As pdStringStack

'Value of the last "is alpha supported" parameter.  When this changes, we need to generate a new
' "before" PNG image (because it changes compositing against a solid background).
Private m_LastAlphaState As PD_BOOL

'The user's answer is returned via this property
Public Function GetDialogResult() As VbMsgBoxResult
    GetDialogResult = m_UserDialogAnswer
End Function

Public Function GetFormatParams() As String
    GetFormatParams = m_FormatParamString
End Function

Public Function GetMetadataParams() As String
    GetMetadataParams = m_MetadataParamString
End Function

Private Sub btsDither_Click(ByVal buttonIndex As Long)
    UpdatePreview
End Sub

Private Sub btsMipMaps_Click(ByVal buttonIndex As Long)
    'Mipmaps don't change the generated preview
    ReflowInterface
End Sub

Private Sub chkLivePreview_Click()
    UpdatePreview
End Sub

Private Sub cmdBar_CancelClick()
    m_UserDialogAnswer = vbCancel
    Me.Visible = False
End Sub

Private Sub cmdBar_OKClick()

    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    'Add the selected format
    cParams.AddParam "dds-format", m_listOfIDs.GetString(Me.lstFormat.ListIndex), True
    
    'Some block-compression algorithms support additional settings
    cParams.AddParam "dds-bc-settings", GetBCSettings(), True
    
    'Mipmaps correspond to text values passed to texconv
    Dim numMipMaps As Long
    If (btsMipMaps.ListIndex = 0) Or (btsMipMaps.ListIndex = 1) Then
        numMipMaps = btsMipMaps.ListIndex
    Else
        numMipMaps = sldMipMaps.Value
    End If
    cParams.AddParam "dds-mipmaps", numMipMaps, True, True
    If (ddFilter.ListIndex >= 0) Then cParams.AddParam "dds-mipmap-filter", ddFilter.List(ddFilter.ListIndex)
    
    m_FormatParamString = cParams.GetParamString()
    
    'If ExifTool someday supports metadata embedding for this format, you can add a metadata manager here
    m_MetadataParamString = vbNullString
    
    'Free resources that are no longer required
    Set m_CompositedImage = Nothing
    Set m_SrcImage = Nothing
    If (LenB(m_PreviewImagePath) > 0) Then Files.FileDeleteIfExists m_PreviewImagePath
    
    'Hide but *DO NOT UNLOAD* the form.  The dialog manager needs to retrieve the setting strings before unloading us
    m_UserDialogAnswer = vbOK
    Me.Visible = False
    
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub ddFilter_Click()
    'Mipmap filtering doesn't affect the live preview
End Sub

Private Sub Form_Load()
    
    m_LastAlphaState = PD_BOOL_UNKNOWN
    
    btsDither.AddItem "none", 0
    btsDither.AddItem "perceptual", 1
    btsDither.AddItem "uniform", 2
    btsDither.ListIndex = 1
    
    btsCompression.AddItem "fast"
    btsCompression.AddItem "medium"
    btsCompression.AddItem "slow"
    btsCompression.ListIndex = 1
    
    btsMipMaps.AddItem "all", 0
    btsMipMaps.AddItem "none", 1
    btsMipMaps.AddItem "custom", 2
    btsMipMaps.ListIndex = 0
    
    ddFilter.SetAutomaticRedraws False
    ddFilter.Clear
    ddFilter.AddItem "point"
    ddFilter.AddItem "linear"
    ddFilter.AddItem "cubic"
    ddFilter.AddItem "fant"
    ddFilter.AddItem "box"
    ddFilter.AddItem "triangle"
    ddFilter.ListIndex = 0
    ddFilter.SetAutomaticRedraws True, True
    
    chkLivePreview.AssignTooltip "This image format is very computationally intensive.  On older or slower PCs, you may want to disable live previews."
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'Ensure we release any temp files on exit
    Files.FileDeleteIfExists m_PreviewImagePath
    
    'Release subclassing form themer
    ReleaseFormTheming Me
    
End Sub

Private Sub lstFormat_Click()
        
    ReflowInterface
        
    If (Not m_listOfIDs Is Nothing) Then
    
        Dim curAlphaState As PD_BOOL
        curAlphaState = Plugin_DDS.DoesFormatSupportAlpha(m_listOfIDs.GetString(lstFormat.ListIndex))
        If (curAlphaState <> m_LastAlphaState) Then
            m_LastAlphaState = curAlphaState
            UpdatePreviewSource
        End If
    
    Else
        UpdatePreviewSource
    End If
    
    UpdatePreview
    
End Sub

Private Sub pdFxPreview_ViewportChanged()
    UpdatePreviewSource
    UpdatePreview
End Sub

'The ShowDialog routine presents the user with this form.
Public Sub ShowDialog(Optional ByRef srcImage As pdImage = Nothing)

    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    m_UserDialogAnswer = vbCancel
    
    'Make sure that the proper cursor is set
    Screen.MousePointer = 0
    Message "Waiting for user to specify export options... "
    
    'Populate the (rather large) list of export settings
    lstFormat.SetAutomaticRedraws False
    lstFormat.Clear
    
    Plugin_DDS.GetListOfFormatNamesAndIDs m_listOfNames, m_listOfIDs
    
    Dim i As Long
    For i = 0 To m_listOfNames.GetNumOfStrings - 1
        lstFormat.AddItem m_listOfNames.GetString(i), i
    Next i

    'TODO: see if the image was originally a DDS file; if it was, auto-supply the same format
    ' (if possible)
    lstFormat.ListIndex = 0
    
    lstFormat.SetAutomaticRedraws True, True
    
    'Make a copy of the composited image; it takes time to composite layers, so we don't want to redo this except
    ' when absolutely necessary.
    Set m_SrcImage = srcImage
    If ((m_SrcImage Is Nothing) Or (Not Plugin_DDS.IsDirectXTexAvailable())) Then
        Interface.ShowDisabledPreviewImage pdFxPreview
    Else
        m_SrcImage.GetCompositedImage m_CompositedImage, True
        pdFxPreview.NotifyNonStandardSource m_CompositedImage.GetDIBWidth, m_CompositedImage.GetDIBHeight
    End If
    
    'Update the preview
    UpdatePreviewSource
    UpdatePreview True
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True
    Interface.SetFormCaptionW Me, g_Language.TranslateMessage("%1 options", "DDS")
    If (Not g_WindowManager Is Nothing) Then g_WindowManager.SetFocusAPI cmdBar.hWnd
    ReflowInterface
    
    'Display the dialog
    ShowPDDialog vbModal, Me, True
    
End Sub

'When a parameter changes that requires a new source DIB for the preview (e.g. changing the background composite color),
' call this function to generate a new preview DIB.  Note that you *do not* need to call this function for format-specific
' changes (like quality, subsampling, etc).
Private Sub UpdatePreviewSource()

    If Not (m_CompositedImage Is Nothing) Then
        
        'Because the user can change the preview viewport, we can't guarantee that the preview region hasn't changed
        ' since the last preview.  Prep a new preview now.
        Dim tmpSafeArray As SafeArray2D
        EffectPrep.PreviewNonStandardImage tmpSafeArray, m_CompositedImage, pdFxPreview, (m_LastAlphaState = PD_BOOL_FALSE)
        
        If (m_LastAlphaState = PD_BOOL_FALSE) Then workingDIB.CompositeBackgroundColor 255, 255, 255
        
        If (m_PreviewImageBackup Is Nothing) Then Set m_PreviewImageBackup = New pdDIB
        m_PreviewImageBackup.CreateFromExistingDIB workingDIB
        
        'Save a copy of the source image to file, in PNG format.  (PD's current AVIF encoder
        ' works as a command-line tool; we need to pass it a source PNG file.)
        If (LenB(m_PreviewImagePath) > 0) Then Files.FileDeleteIfExists m_PreviewImagePath
        m_PreviewImagePath = OS.UniqueTempFilename(customExtension:="png")
        If (Not Saving.QuickSaveDIBAsPNG(m_PreviewImagePath, workingDIB, False, True)) Then
            InternalError "UpdatePreviewSource", "couldn't save preview png"
        End If
        
    End If
    
End Sub

Private Sub UpdatePreview(Optional ByVal forceUpdate As Boolean = False)
    
    Const funcName As String = "UpdatePreview"
    
    If ((cmdBar.PreviewsAllowed Or forceUpdate) And Plugin_DDS.IsDirectXTexAvailable() And (Not m_SrcImage Is Nothing)) Then
        
        'Make sure the preview source is up-to-date
        If (workingDIB Is Nothing) Then UpdatePreviewSource
        If (workingDIB Is Nothing) Then Exit Sub
        
        'Because previews are so intensive to generate, this dialog provides a toggle so the user
        ' can suspend real-time previews.
        If chkLivePreview.Value Then
            
            'Now perform the (ugly) dance of workingDIB > PNG > DDS > PNG > workingDIB.
            ' (Note that the first workingDIB > PNG step was performed by UpdatePreviewSource,
            ' and the resulting file is stored in m_PreviewImagePath.)
            
            'Start by generating temporary filenames for intermediary files
            Dim tmpFilenameIntermediary As String, tmpFilenameDDS As String

            'Pull a format name from the list box
            Dim ddsFormatName As String
            If (Me.lstFormat.ListIndex >= 0) And (Not m_listOfIDs Is Nothing) Then ddsFormatName = m_listOfIDs.GetString(lstFormat.ListIndex)
            
            'Shell directxtex, and request it to convert the preview PNG to DDS
            If Plugin_DDS.ConvertStandardImageToDDS(m_PreviewImagePath, tmpFilenameDDS, ddsFormatName, dxTex_BC:=GetBCSettings()) Then
                
                'Immediately shell it again, but this time, ask it to convert the DDS it just made
                ' back into a PNG
                If Plugin_DDS.ConvertDDStoStandardImage(tmpFilenameDDS, tmpFilenameIntermediary, False) Then
                    
                    'We are done with the DDS; kill it
                    Files.FileDeleteIfExists tmpFilenameDDS
                    
                    'Load the finished PNG *back* into a pdDIB object
                    If Loading.QuickLoadImageToDIB(tmpFilenameIntermediary, workingDIB, False, False, True) Then
                        
                        'We are done with the intermediary image; kill it
                        Files.FileDeleteIfExists tmpFilenameIntermediary
                        
                        'Display the final result
                        workingDIB.SetAlphaPremultiplication True, True
                        FinalizeNonstandardPreview pdFxPreview, True
                        
                    Else
                        InternalError funcName, "couldn't load finished PNG to pdDIB"
                    End If
                
                Else
                    InternalError funcName, "couldn't convert DDS back to PNG"
                End If
            
            Else
                InternalError funcName, "couldn't save DDS"
            End If
        
        'Live previews are disabled; just mirror the original image to the screen
        Else
            workingDIB.CreateFromExistingDIB m_PreviewImageBackup
            FinalizeNonstandardPreview pdFxPreview, False
        End If
                
    Else
        If (Not Plugin_DDS.IsDirectXTexAvailable()) Then InternalError funcName, "dds library broken"
    End If

End Sub

Private Sub InternalError(ByRef funcName As String, ByRef errMsg As String)
    PDDebug.LogAction "WARNING! Problem in dialog_ExportDDS." & funcName & ": " & errMsg
End Sub

Private Sub ReflowInterface()
    
    Dim ySpacing As Long
    ySpacing = Interface.FixDPI(6)
    
    Dim yOffset As Long
    yOffset = Me.lstFormat.GetTop + Me.lstFormat.GetHeight + ySpacing
    
    'Dithering is only available for BC1-3
    btsDither.SetTop yOffset
    btsDither.Visible = IsDitherAvailable()
    If btsDither.Visible Then yOffset = yOffset + btsDither.GetHeight + ySpacing
    
    btsCompression.SetTop yOffset
    btsCompression.Visible = IsVariableCompressionAvailable()
    If btsCompression.Visible Then yOffset = yOffset + btsCompression.GetHeight + ySpacing
    
    'Mipmaps are always available
    Me.btsMipMaps.SetTop yOffset
    yOffset = yOffset + Me.btsMipMaps.GetHeight + ySpacing
    sldMipMaps.Visible = (btsMipMaps.ListIndex = 2)
    ddFilter.Visible = sldMipMaps.Visible
    
    'Mipmaps use defaults unless "custom" is selected
    If sldMipMaps.Visible Then sldMipMaps.SetTop yOffset
    If ddFilter.Visible Then ddFilter.SetTop yOffset + Interface.FixDPI(2)
    
End Sub

'Some settings are only available for certain formats
Private Function IsDitherAvailable() As Boolean
    IsDitherAvailable = (lstFormat.ListIndex < 6)
End Function

Private Function IsVariableCompressionAvailable() As Boolean
    IsVariableCompressionAvailable = (lstFormat.ListIndex = 12) Or (lstFormat.ListIndex = 13)
End Function

'Some block-compression algorithms support additional features
Private Function GetBCSettings() As String
    
    If IsDitherAvailable() Then
        
        Select Case btsDither.ListIndex
            
            'None
            Case 0
                GetBCSettings = vbNullString
            
            'Perceptual
            Case 1
                GetBCSettings = "d"
            
            'Uniform
            Case 2
                GetBCSettings = "ud"
            
        End Select
    
    ElseIf IsVariableCompressionAvailable() Then
    
        Select Case btsCompression.ListIndex
            
            'Fast
            Case 0
                GetBCSettings = "q"
                
            'Medium
            Case 1
                GetBCSettings = vbNullString
                
            'Slow
            Case 2
                GetBCSettings = "x"
                
        End Select
    
    Else
        GetBCSettings = vbNullString
    End If
    
End Function
