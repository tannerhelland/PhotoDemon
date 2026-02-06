VERSION 5.00
Begin VB.Form FormThemeEditor 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Resource editor"
   ClientHeight    =   10170
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13260
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
   ScaleHeight     =   678
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   884
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdCheckBox chkSort 
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   8280
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      Caption         =   "sort before saving"
      Value           =   0   'False
   End
   Begin PhotoDemon.pdCheckBox chkCustomMenuColor 
      Height          =   375
      Left            =   4200
      TabIndex        =   17
      Top             =   6600
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   661
      Caption         =   "use custom menu color"
      Value           =   0   'False
   End
   Begin PhotoDemon.pdLabel lblExport 
      Height          =   375
      Left            =   4200
      Top             =   8760
      Visible         =   0   'False
      Width           =   8895
      _ExtentX        =   16113
      _ExtentY        =   661
      Caption         =   ""
      FontSize        =   12
   End
   Begin PhotoDemon.pdCheckBox chkDelete 
      Height          =   375
      Left            =   9600
      TabIndex        =   15
      Top             =   8280
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   661
      Caption         =   "mark resource for deletion"
   End
   Begin PhotoDemon.pdButton cmdExport 
      Height          =   615
      Left            =   240
      TabIndex        =   14
      Top             =   8640
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1085
      Caption         =   "export finished resource file"
   End
   Begin PhotoDemon.pdButtonStrip btsBackcolor 
      Height          =   615
      Left            =   9600
      TabIndex        =   13
      Top             =   7200
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1085
   End
   Begin PhotoDemon.pdColorSelector csLight 
      Height          =   855
      Left            =   4200
      TabIndex        =   11
      Top             =   5640
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1508
      Caption         =   "light theme color"
      FontSize        =   10
   End
   Begin PhotoDemon.pdPictureBox picPreview 
      Height          =   2535
      Left            =   9600
      Top             =   4560
      Width           =   3495
      _ExtentX        =   0
      _ExtentY        =   0
   End
   Begin PhotoDemon.pdCheckBox chkColoration 
      Height          =   375
      Left            =   4200
      TabIndex        =   9
      Top             =   5160
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   661
      Caption         =   "run-time coloration"
   End
   Begin PhotoDemon.pdButton cmdSave 
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   7560
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   873
      Caption         =   "force save resource package"
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
      Left            =   12720
      TabIndex        =   4
      Top             =   480
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   661
      Caption         =   "..."
   End
   Begin PhotoDemon.pdTextBox txtResourcePath 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   661
      FontSize        =   12
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
      Left            =   240
      TabIndex        =   2
      Top             =   6840
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1085
      Caption         =   "add a new resource"
   End
   Begin PhotoDemon.pdListBox lstResources 
      Height          =   5775
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   10186
      Caption         =   "current resources"
   End
   Begin PhotoDemon.pdCommandBarMini cmdBar 
      Align           =   2  'Align Bottom
      Height          =   795
      Left            =   0
      TabIndex        =   0
      Top             =   9375
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
      Width           =   8295
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
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   375
      Index           =   4
      Left            =   4200
      Top             =   4680
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   661
      Caption         =   "image resource properties:"
      FontSize        =   12
   End
   Begin PhotoDemon.pdColorSelector csDark 
      Height          =   855
      Left            =   6960
      TabIndex        =   12
      Top             =   5640
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1508
      Caption         =   "dark theme color"
      FontSize        =   10
   End
   Begin PhotoDemon.pdButton cmdResItemPath 
      Height          =   375
      Left            =   12600
      TabIndex        =   16
      Top             =   2880
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   661
      Caption         =   "..."
   End
   Begin PhotoDemon.pdColorSelector csMenu 
      Height          =   735
      Left            =   4200
      TabIndex        =   18
      Top             =   7080
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1296
      Caption         =   "custom menu color"
      FontSize        =   10
   End
End
Attribute VB_Name = "FormThemeEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Resource editor dialog
'Copyright 2016-2026 by Tanner Helland
'Created: 22/August/16
'Last updated: 14/April/22
'Last update: replace lingering picture box with pdPictureBox
'
'As of v7.0, PD finally supports visual themes using its internal theming engine.  As part of supporting
' visual themes, various PD controls need access to image resources at a size and color scheme appropriate
' for the current theme.
'
'This resource editor is designed to help with that task.
'
'At present, PD's original resource file is still required, as all resources have *not* yet been migrated
' to the new format.
'
'Also, please note that this dialog is absolutely *not* meant for external use.  It is for PD developers, only.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Enum PD_Resource_Type
    PDRT_Image = 0
    PDRT_Other = 1
End Enum

#If False Then
    Private Const PDRT_Image = 0, PDRT_Other = 1
#End If

Private Type PD_Resource
    ResourceName As String
    ResFileLocation As String
    ResType As PD_Resource_Type
    ResColorLight As Long
    ResColorDark As Long
    ResColorMenu As Long
    ResSupportsColoration As Boolean
    ResCustomMenuColor As Boolean
    MarkedForDeletion As Boolean    'Resource deletion is very primitive at present; it may not work as expected
End Type

Private m_NumOfResources As Long
Private m_Resources() As PD_Resource
Private m_LastResourceIndex As Long

Private m_FSO As pdFSO
Private m_PreviewDIBOriginal As pdDIB, m_PreviewDIB As pdDIB

Private m_SuspendUpdates As Boolean

'Current preview image, if any
Private m_previewImage As pdDIB

Private Sub btsBackcolor_Click(ByVal buttonIndex As Long)
    UpdatePreview
End Sub

Private Sub btsResourceType_LostFocusAPI()
    SyncResourceAgainstCurrentUI
End Sub

Private Sub chkColoration_Click()
    SyncResourceAgainstCurrentUI
    UpdatePreview
End Sub

Private Sub chkDelete_Click()
    SyncResourceAgainstCurrentUI
End Sub

Private Sub cmdAddResource_Click()
    
    Dim srcFile As String
    
    Dim cCommonDialog As pdOpenSaveDialog: Set cCommonDialog = New pdOpenSaveDialog
    If cCommonDialog.GetOpenFileName(srcFile, , True, False, , , m_FSO.FileGetPath(txtResourcePath.Text), "Select resource", , Me.hWnd) Then
        
        If (m_NumOfResources > UBound(m_Resources)) Then ReDim Preserve m_Resources(0 To m_NumOfResources * 2 - 1) As PD_Resource
        
        With m_Resources(m_NumOfResources)
            .ResourceName = Files.FileGetName(srcFile, True)
            .ResFileLocation = srcFile
            If Loading.QuickLoadImageToDIB(srcFile, m_PreviewDIBOriginal, False, False) Then .ResType = PDRT_Image
        End With
        
        lstResources.AddItem m_Resources(m_NumOfResources).ResourceName
        lstResources.ListIndex = m_NumOfResources
        
        m_NumOfResources = m_NumOfResources + 1
        
        SyncUIAgainstCurrentResource
        
    End If
    
End Sub

Private Sub cmdBar_OKClick()
    SaveWorkingFile
End Sub

'Export the current resource collection to an actual resource file.  This is a one-way conversion.
Private Sub cmdExport_Click()
    
    'At present, resources are saved to the App/PhotoDemon/Themes subfolder.  This resource file is
    ' ultimately compiled into the .exe, so this location should be considered "temporary" only.
    ' (PD will preferentially load this resource file, if it exists, instead of the embedded .exe
    ' copy - this is done as a convenience for testing, and should not be relied upon in a stable
    ' release!)
    Dim targetResFile As String
    If (LenB(txtResourcePath.Text) <> 0) Then
        
        'Provide minimal UI feedback
        lblExport.Visible = True
        lblExport.Caption = "Prepping resource file..."
        
        'Keep the existing filename, but strip the extension and replace it with "PDRC"
        ' (for... PhotoDemon Resource Collection, I guess?)
        targetResFile = UserPrefs.GetThemePath & Files.FileGetName(txtResourcePath.Text, True) & ".pdrc"
        Files.FileDeleteIfExists targetResFile
        
        'Prep a pdPackage
        Dim cPackage As pdPackageChunky
        Set cPackage = New pdPackageChunky
        cPackage.StartNewPackage_File targetResFile, packageID:="PDRS"
        
        'Images are actually stored in their own two "special" packages.  Why?  Better compression.
        ' PD's image resources each require two separate streams: image headers (XML describing image contents)
        ' and pixel buffers.  Rather than individually compressing a bunch of tiny standalone files, we instead
        ' split XML bits and pixel bits into separate, dedicated packages, and compress each *entire* collection
        ' at once.  This maximizes compression by sharing tables across files, and the end result is 20-30%
        ' smaller than compressing each file individually.
        Dim cImageHeaders As pdPackageChunky
        Set cImageHeaders = New pdPackageChunky
        cImageHeaders.StartNewPackage_Memory packageID:="IMGH"
        
        Dim cImagePixels As pdPackageChunky
        Set cImagePixels = New pdPackageChunky
        cImagePixels.StartNewPackage_Memory packageID:="IMGP"
        
        'We're also going to use a quick trick to significantly reduce file size of bitmap data.
        ' In our icons, we force all transparency values to be a multiple of 5.  This reduces net entropy
        ' by 80% (vs normal 8-bit data), and since we typically resize these icons to a tiny fraction of
        ' their original size, there's basically no difference in final visual quality.
        '
        '(As a hard number for file size reduction, on a test resource file with 8 icons, the resource file
        ' size drops from 16.5 kb to 6 kb thanks to this silly trick.)
        '
        'You can pick an interval larger than five if you want an even larger reduction, but at some point
        ' it starts interfering with antialiasing quality, so be cautious.
        Dim cmpLookup() As Byte
        ReDim cmpLookup(0 To 255) As Byte
        
        Dim x As Long, y As Long
        For x = 0 To 255
            cmpLookup(x) = ((x + 2) \ 5) * 5
        Next x
        
        'Start adding resources.  Resources are stored in a predefined format that describes how the icons are
        ' to be treated at load-time.  (Generally speaking, we apply specific post-processing based on the
        ' current theme and/or request information from the caller.)  Some icons come pre-colored, and as such,
        ' they obey different rules.  This must all be stored in the resource file.
        Dim i As Long
        Dim cXML As pdSerialize: Set cXML = New pdSerialize
        Dim tmpDIB As pdDIB, tmpDIBSize As Long, tmpDIBPointer As Long
        
        For i = 0 To m_NumOfResources - 1
            
            lblExport.Caption = "Writing resource #" & CStr(i + 1) & " of " & CStr(m_NumOfResources)
            lblExport.RequestRefresh
            
            'Prep the XML packet for this resource.  For image-type entries, this stores things like the original
            ' resource image size (w/h), coloration behavior, and any other special instructions.
            If (m_Resources(i).ResType = PDRT_Image) Then
                
                Dim cCount As Long, testPalette() As RGBQuad, testPixels() As Byte
                cXML.Reset
                
                'Load the source image to a temporary DIB (so we can query various image attributes)
                If Loading.QuickLoadImageToDIB(m_Resources(i).ResFileLocation, tmpDIB, False, False, True) Then
                    
                    Dim useImagePalette As Boolean: useImagePalette = False
                    
                    'Write the bare amount of information required to reconstruct the image at run-time
                    cXML.AddParam "w", tmpDIB.GetDIBWidth, True, True
                    cXML.AddParam "h", tmpDIB.GetDIBHeight, True, True
                    cXML.AddParam "bpp", tmpDIB.GetDIBColorDepth, True, True
                    
                    If m_Resources(i).ResSupportsColoration Then
                        
                        cXML.AddParam "rt-clr", "True", True, True
                        cXML.AddParam "clr-l", m_Resources(i).ResColorLight, True, True
                        cXML.AddParam "clr-d", m_Resources(i).ResColorDark, True, True
                        If m_Resources(i).ResCustomMenuColor Then
                            cXML.AddParam "rt-clrmenu", "True", True, True
                            cXML.AddParam "clr-m", m_Resources(i).ResColorMenu, True, True
                        End If
                        
                    Else
                    
                        cXML.AddParam "rt-clr", "False", True, True
                        
                        'See how many colors this DIB has.  If it's 256 or less, we can write it to file
                        ' using a palette.
                        cCount = DIBs.GetDIBAs8bpp_RGBA(tmpDIB, testPalette, testPixels)
                        
                        'If this image only has [0, 255] unique RGBA entries, we can store it as a paletted image
                        ' and conserve a bunch of file space!
                        If (cCount <= 256) Then
                            useImagePalette = True
                            cXML.AddParam "uses-palette", "True", True, True
                            cXML.AddParam "palette-size", cCount, True, True
                            Debug.Print "Palette candidate found: " & cCount & " - " & m_Resources(i).ResFileLocation
                        End If
                        
                    End If
                    
                    'Write this data to the first half of the node. (Note that zstd is always used to compress headers.)
                    'cPackage.AddNodeDataFromString nodeIndex, True, cXML.ReturnCurrentXMLString, resCompFormat, Compression.GetMaxCompressionLevel(resCompFormat)
                    Dim tmpXmlBytes() As Byte, lenXmlBytes As Long
                    Strings.UTF8FromString cXML.GetParamString(), tmpXmlBytes, lenXmlBytes, , True
                    cImageHeaders.AddChunk_NameValuePair "NAME", m_Resources(i).ResourceName, "DATA", VarPtr(tmpXmlBytes(0)), lenXmlBytes, cf_None
                    
                    'Write the actual bitmap data to the second half of the node.  Note that we use two
                    ' different strategies here.
                    ' 1) If this resource does *not* support run-time coloration, store it like a normal DIB
                    ' 2) If this resource *does* support run-time coloration, just store the alpha channel.
                    '    (Color values will be plugged-in at run-time.)
                    If m_Resources(i).ResSupportsColoration Then
                        
                        Dim tmpBytes() As Byte
                        If DIBs.RetrieveTransparencyTable(tmpDIB, tmpBytes) Then
                        
                            'Apply our previously calculated lookup table to the transparency bytes
                            For y = 0 To tmpDIB.GetDIBHeight - 1
                            For x = 0 To tmpDIB.GetDIBWidth - 1
                                tmpBytes(x, y) = cmpLookup(tmpBytes(x, y))
                            Next x
                            Next y
                            
                            cImagePixels.AddChunk_NameValuePair "NAME", m_Resources(i).ResourceName, "DATA", VarPtr(tmpBytes(0, 0)), tmpDIB.GetDIBWidth * tmpDIB.GetDIBHeight, cf_None
                            
                        End If
                    
                    Else
                    
                        'Paletted images are stored differently
                        If useImagePalette Then
                        
                            'Build a combined palette + image bytes array
                            Dim totalData() As Byte, totalDataSize As Long
                            totalDataSize = (cCount * 4) + (tmpDIB.GetDIBWidth * tmpDIB.GetDIBHeight)
                            ReDim totalData(0 To totalDataSize - 1) As Byte
                            CopyMemoryStrict VarPtr(totalData(0)), VarPtr(testPalette(0)), 4 * cCount
                            CopyMemoryStrict VarPtr(totalData(4 * cCount)), VarPtr(testPixels(0, 0)), tmpDIB.GetDIBWidth * tmpDIB.GetDIBHeight
                            
                            cImagePixels.AddChunk_NameValuePair "NAME", m_Resources(i).ResourceName, "DATA", VarPtr(totalData(0)), totalDataSize, cf_None
                            
                        Else
                            PDDebug.LogAction "WARNING!  A palette was not detected for source image (" & m_Resources(i).ResFileLocation & ") - revisit to improve compression ratio"
                            tmpDIB.RetrieveDIBPointerAndSize tmpDIBPointer, tmpDIBSize
                            cImagePixels.AddChunk_NameValuePair "NAME", m_Resources(i).ResourceName, "DATA", tmpDIBPointer, tmpDIBSize, cf_None
                        End If
                        
                    End If
                    
                End If
                
            '"Other" type resources are simply saved as-is, byte for byte; it's up to the caller to
            ' interpret the information manually at run-time.
            ElseIf (m_Resources(i).ResType = PDRT_Other) Then
                Dim otherResBytes() As Byte
                If Files.FileLoadAsByteArray(m_Resources(i).ResFileLocation, otherResBytes) Then
                    cPackage.AddChunk_NameValuePair "NAME", m_Resources(i).ResourceName, "DATA", VarPtr(otherResBytes(0)), UBound(otherResBytes) + 1, cf_Zstd, Compression.GetMaxCompressionLevel(cf_Zstd)
                Else
                    PDDebug.LogAction "WARNING!  Other resource (" & m_Resources(i).ResFileLocation & ") wasn't loaded successfully."
                End If
            End If
            
        Next i
        
        lblExport.Caption = "Finalizing package..."
        lblExport.RequestRefresh
        
        'Add the final image packages
        Dim tmpStream As pdStream
        cImageHeaders.FinishPackage tmpStream
        cPackage.AddChunk_NameValuePair "NAME", "final_img_headers", "DATA", tmpStream.Peek_PointerOnly(0, tmpStream.GetStreamSize()), tmpStream.GetStreamSize(), cf_Zstd, Compression.GetMaxCompressionLevel(cf_Zstd)
        
        Set tmpStream = Nothing
        cImagePixels.FinishPackage tmpStream
        cPackage.AddChunk_NameValuePair "NAME", "final_img_pixels", "DATA", tmpStream.Peek_PointerOnly(0, tmpStream.GetStreamSize()), tmpStream.GetStreamSize(), cf_Zstd, Compression.GetMaxCompressionLevel(cf_Zstd)
        
        'With the package complete, write it out to file!
        cPackage.FinishPackage
        
        lblExport.Caption = "Resource export complete."
        lblExport.RequestRefresh
        
    End If

End Sub

Private Sub cmdExport_LostFocusAPI()
    lblExport.Visible = False
End Sub

Private Sub cmdResItemPath_Click()

    Dim srcFile As String
    srcFile = Files.FileGetName(txtResourceLocation.Text)
    
    Dim cCommonDialog As pdOpenSaveDialog: Set cCommonDialog = New pdOpenSaveDialog
    If cCommonDialog.GetOpenFileName(srcFile, , True, False, , , m_FSO.FileGetPath(txtResourceLocation.Text), "Select resource item", , Me.hWnd) Then
        If (LenB(srcFile) <> 0) Then
            txtResourceLocation.Text = srcFile
            SyncResourceAgainstCurrentUI
            UpdatePreview
        End If
    End If
    
End Sub

Private Sub cmdResourcePath_Click()
    
    Dim srcFile As String
    srcFile = Files.FileGetName(txtResourcePath.Text)
    
    Dim cCommonDialog As pdOpenSaveDialog: Set cCommonDialog = New pdOpenSaveDialog
    If cCommonDialog.GetOpenFileName(srcFile, , False, False, "PD Resource Files (*.pdr)|*.pdr", , m_FSO.FileGetPath(txtResourcePath.Text), "Select resource file", "pdr", Me.hWnd) Then
        If (LenB(srcFile) <> 0) Then
            txtResourcePath.Text = srcFile
            UserPrefs.SetPref_String "Themes", "LastResourceFile", srcFile
            LoadResourceFromFile
        End If
    End If
    
End Sub

Private Sub cmdSave_Click()
    SaveWorkingFile
End Sub

Private Sub SaveWorkingFile()

    Dim okayToProceed As Boolean: okayToProceed = True
    
    'If the user isn't editing an existing file, prompt them for a filename
    If (LenB(Trim$(txtResourcePath.Text)) = 0) Then
    
        Dim srcFile As String
        
        Dim cCommonDialog As pdOpenSaveDialog: Set cCommonDialog = New pdOpenSaveDialog
        If cCommonDialog.GetSaveFileName(srcFile, , True, "PD Resource Files (*.pdr)|*.pdr", , UserPrefs.GetThemePath, "Save resource file", "pdr", Me.hWnd) Then
            If (LenB(srcFile) <> 0) Then
                txtResourcePath.Text = srcFile
                UserPrefs.SetPref_String "Themes", "LastResourceFile", srcFile
                okayToProceed = True
            Else
                okayToProceed = False
            End If
        Else
            okayToProceed = False
        End If
        
    End If
    
    If okayToProceed Then
    
        'Save a copy of the current XML information in XML format.  (Note that this is different from *compiling*
        ' the resource file, as you'd expect.)
        Dim cXML As pdXML: Set cXML = New pdXML
        cXML.PrepareNewXML "pdResource"
        cXML.WriteTag "ResourceCount", m_NumOfResources
        cXML.WriteTag "LastEditedResource", m_LastResourceIndex
        
        Dim numResourcesWritten As Long: numResourcesWritten = 0
        Dim i As Long
        
        'Make a local copy of the resource collection.  We may need to sort the collection before writing it
        ' out to file, and we don't want to use our in-progress copy for that (as it needs to be synched to
        ' the list box order).
        Dim tmpResources() As PD_Resource
        ReDim tmpResources(0 To m_NumOfResources - 1) As PD_Resource
        For i = 0 To m_NumOfResources - 1
            tmpResources(i) = m_Resources(i)
        Next i
        
        'If requested, sort the resources prior to writing them
        If chkSort.Value Then
        
            Dim tmpSort As PD_Resource
            Dim j As Long
            
            For i = 0 To m_NumOfResources - 1
            For j = 0 To m_NumOfResources - 1
                If (StrComp(tmpResources(i).ResourceName, tmpResources(j).ResourceName, vbBinaryCompare) < 0) Then
                    tmpSort = tmpResources(i)
                    tmpResources(i) = tmpResources(j)
                    tmpResources(j) = tmpSort
                End If
            Next j
            Next i
        
        End If
        
        For i = 0 To m_NumOfResources - 1
            
            If (Not tmpResources(i).MarkedForDeletion) Then
            
                cXML.WriteTag CStr(numResourcesWritten + 1), vbNullString, True
                
                With tmpResources(i)
                    cXML.WriteTag "Name", .ResourceName
                    cXML.WriteTag "FileLocation", .ResFileLocation
                    cXML.WriteTag "Type", .ResType
                    cXML.WriteTag "SupportsColoration", .ResSupportsColoration
                    If .ResSupportsColoration Then
                        cXML.WriteTag "ColorLight", .ResColorLight
                        cXML.WriteTag "ColorDark", .ResColorDark
                    End If
                    cXML.WriteTag "SupportsCustomMenuColor", .ResCustomMenuColor
                    If .ResCustomMenuColor Then cXML.WriteTag "ColorMenu", .ResColorMenu
                End With
                
                cXML.CloseTag CStr(numResourcesWritten + 1)
                
                numResourcesWritten = numResourcesWritten + 1
                
            End If
            
        Next i
        
        'Update the final tag count with the amount of resources we actually wrote to file
        cXML.UpdateTag "ResourceCount", numResourcesWritten
        
        If (Not cXML.WriteXMLToFile(txtResourcePath.Text)) Then Debug.Print "WARNING!  Save to file failed!!"
    
    End If
    
End Sub

Private Sub csLight_ColorChanged()
    If (Not m_SuspendUpdates) Then
        SyncResourceAgainstCurrentUI
        m_SuspendUpdates = True
        btsBackcolor.ListIndex = 0
        m_SuspendUpdates = False
        UpdatePreview
    End If
End Sub

Private Sub csDark_ColorChanged()
    If (Not m_SuspendUpdates) Then
        SyncResourceAgainstCurrentUI
        m_SuspendUpdates = True
        btsBackcolor.ListIndex = 1
        m_SuspendUpdates = False
        UpdatePreview
    End If
End Sub

Private Sub csMenu_ColorChanged()
    If (Not m_SuspendUpdates) Then
        SyncResourceAgainstCurrentUI
        m_SuspendUpdates = True
        btsBackcolor.ListIndex = 2
        m_SuspendUpdates = False
        UpdatePreview
    End If
End Sub

Private Sub Form_Load()
            
    btsResourceType.AddItem "image", 0
    btsResourceType.AddItem "other", 1
    btsResourceType.ListIndex = 0
    
    btsBackcolor.AddItem "light", 0
    btsBackcolor.AddItem "dark", 1
    btsBackcolor.AddItem "menu", 2
    If (g_Themer.GetCurrentThemeClass = PDTC_Dark) Then btsBackcolor.ListIndex = 1 Else btsBackcolor.ListIndex = 0
    
    Set m_FSO = New pdFSO
    
    'Load the last-edited resource file (if any)
    If UserPrefs.DoesValueExist("Themes", "LastResourceFile") Then
        txtResourcePath.Text = UserPrefs.GetPref_String("Themes", "LastResourceFile", vbNullString)
        LoadResourceFromFile
    Else
        txtResourcePath.Text = vbNullString
        
        m_NumOfResources = 0
        ReDim m_Resources(0 To 15) As PD_Resource
        
        lstResources.ListIndex = -1
        m_LastResourceIndex = -1
        
    End If
    
    Interface.ApplyThemeAndTranslations Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub lstResources_Click()
    SyncResourceAgainstCurrentUI
    m_LastResourceIndex = lstResources.ListIndex
    SyncUIAgainstCurrentResource
End Sub

Private Sub picPreview_DrawMe(ByVal targetDC As Long, ByVal ctlWidth As Long, ByVal ctlHeight As Long)
    If (Not m_previewImage Is Nothing) Then m_previewImage.AlphaBlendToDC targetDC
End Sub

Private Sub txtResourceLocation_LostFocusAPI()
    SyncResourceAgainstCurrentUI
End Sub

Private Sub txtResourceName_LostFocusAPI()
    lstResources.UpdateItem lstResources.ListIndex, txtResourceName.Text
    lstResources.SetAutomaticRedraws True, True
    SyncResourceAgainstCurrentUI
End Sub

Private Sub LoadResourceFromFile()
    
    Dim cXML As pdXML: Set cXML = New pdXML
    If cXML.LoadXMLFile(txtResourcePath.Text) Then
        If cXML.IsPDDataType("pdResource") Then
        
            m_NumOfResources = cXML.GetUniqueTag_Long("ResourceCount", 0)
            
            If (m_NumOfResources > 0) Then
                
                ReDim m_Resources(0 To m_NumOfResources - 1) As PD_Resource
                
                lstResources.Clear
                
                Dim i As Long, tagPos As Long
                tagPos = 1
                
                For i = 0 To m_NumOfResources - 1
                    
                    tagPos = cXML.GetLocationOfTag(CStr(i + 1), tagPos)
                    If (tagPos > 0) Then
                        
                        With m_Resources(i)
                            .ResourceName = cXML.GetUniqueTag_String("Name", vbNullString, tagPos)
                            .ResFileLocation = cXML.GetUniqueTag_String("FileLocation", vbNullString, tagPos)
                            .ResType = cXML.GetUniqueTag_Long("Type", 0, tagPos)
                            .ResSupportsColoration = cXML.GetUniqueTag_Boolean("SupportsColoration", False, tagPos)
                            If .ResSupportsColoration Then
                                .ResColorLight = cXML.GetUniqueTag_Long("ColorLight", 0, tagPos)
                                .ResColorDark = cXML.GetUniqueTag_Long("ColorDark", 0, tagPos)
                            End If
                            .ResCustomMenuColor = cXML.GetUniqueTag_Boolean("SupportsCustomMenuColor", False, tagPos)
                            If .ResCustomMenuColor Then .ResColorMenu = cXML.GetUniqueTag_Long("ColorMenu", 0, tagPos)
                            .MarkedForDeletion = False
                        End With
                        
                        lstResources.AddItem m_Resources(i).ResourceName
                        
                    End If
                    
                Next i
                
                m_LastResourceIndex = cXML.GetUniqueTag_Long("LastEditedResource")
                If (m_LastResourceIndex > lstResources.ListCount - 1) Then m_LastResourceIndex = lstResources.ListCount - 1
                SyncUIAgainstCurrentResource
                
                lstResources.ListIndex = m_LastResourceIndex
                
            End If
        
        End If
    End If
    
End Sub

'Prior to changing the current resource index, this function can be called to update the last-selected resource against
' any UI changes the user may have entered.
Private Sub SyncResourceAgainstCurrentUI()

    If (m_LastResourceIndex >= 0) And (Not m_SuspendUpdates) Then
    
        With m_Resources(m_LastResourceIndex)
        
            .ResourceName = txtResourceName.Text
            .ResType = btsResourceType.ListIndex
            .ResFileLocation = txtResourceLocation.Text
            If (.ResType = PDRT_Image) Then .ResSupportsColoration = chkColoration.Value
            If .ResSupportsColoration Then
                .ResColorLight = csLight.Color
                .ResColorDark = csDark.Color
            End If
            If (.ResType = PDRT_Image) Then .ResCustomMenuColor = chkCustomMenuColor.Value
            If .ResCustomMenuColor Then .ResColorMenu = csMenu.Color
            
            'To delete a resource, you have to click the delete button, save the resource file,
            ' then exit and re-enter the dialog.  (Sorry; deletion is not really meant to be used often.)
            .MarkedForDeletion = chkDelete.Value
            
        End With
    
    End If
    
End Sub

'Whenever the current resource index is changed (e.g. by clicking the left-hand list box), this function can be called
' to update all UI elements against the newly selected resource.
Private Sub SyncUIAgainstCurrentResource()
    
    If (m_LastResourceIndex >= 0) Then
        
        m_SuspendUpdates = True
        
        With m_Resources(m_LastResourceIndex)
            txtResourceName.Text = .ResourceName
            btsResourceType.ListIndex = .ResType
            txtResourceLocation.Text = .ResFileLocation
            If .ResSupportsColoration Then
                chkColoration.Value = True
                csLight.Color = .ResColorLight
                csDark.Color = .ResColorDark
            Else
                chkColoration.Value = False
            End If
            If .ResCustomMenuColor Then
                chkCustomMenuColor.Value = True
                csMenu.Color = .ResColorMenu
            Else
                chkCustomMenuColor.Value = False
            End If
            
            chkDelete.Value = .MarkedForDeletion
            
            m_SuspendUpdates = False
            
            'Image resources get a live preview
            If (.ResType = PDRT_Image) Then UpdatePreview
            
        End With
    
    End If
    
End Sub

'Paint a preview of the current resource image, with any coloration settings applied
Private Sub UpdatePreview()
    
    On Error GoTo PreviewError
    
    If (m_NumOfResources <= 0) Then Exit Sub
    
    If (Not m_SuspendUpdates) Then
    
        If (m_Resources(m_LastResourceIndex).ResType = PDRT_Image) Then
            
            Dim newColor As Long
            If (btsBackcolor.ListIndex = 0) Then
                Colors.GetColorFromString "#ffffff", newColor, ColorHex
            ElseIf (btsBackcolor.ListIndex = 1) Then
                Colors.GetColorFromString "#313131", newColor, ColorHex
            ElseIf (btsBackcolor.ListIndex = 2) Then
                Colors.GetColorFromString "#ffffff", newColor, ColorHex
            End If
            
            'Prep a back buffer with the specified background color
            If (m_previewImage Is Nothing) Then Set m_previewImage = New pdDIB
            m_previewImage.CreateBlank picPreview.GetWidth, picPreview.GetHeight, 32, newColor, 255
            
            If Loading.QuickLoadImageToDIB(m_Resources(m_LastResourceIndex).ResFileLocation, m_PreviewDIBOriginal, False, False, True) Then
            
                'If coloration is supported, apply it now
                If m_Resources(m_LastResourceIndex).ResSupportsColoration Then
                    
                    If (m_PreviewDIB Is Nothing) Then Set m_PreviewDIB = New pdDIB
                    m_PreviewDIB.CreateFromExistingDIB m_PreviewDIBOriginal
                    
                    If (btsBackcolor.ListIndex = 0) Then
                        DIBs.ColorizeDIB m_PreviewDIB, csLight.Color
                    ElseIf (btsBackcolor.ListIndex = 1) Then
                        DIBs.ColorizeDIB m_PreviewDIB, csDark.Color
                    ElseIf (btsBackcolor.ListIndex = 2) Then
                        DIBs.ColorizeDIB m_PreviewDIB, csMenu.Color
                    End If
                    m_PreviewDIB.RenderToDIB m_previewImage, False, True, True
                    
                'If coloration is *not* supported, just render the preview image as-is
                Else
                    m_PreviewDIBOriginal.RenderToDIB m_previewImage, False, True, True
                End If
                
                'Because the preview will appear on-screen, paint a border around it
                Dim cPenBorder As pd2DPen
                Set cPenBorder = New pd2DPen
                cPenBorder.SetPenWidth 1!
                cPenBorder.SetPenLineJoin P2_LJ_Miter
                If (Not g_Themer Is Nothing) Then cPenBorder.SetPenColor g_Themer.GetGenericUIColor(UI_GrayNeutral)
                
                Dim cSurface As pd2DSurface
                Set cSurface = New pd2DSurface
                cSurface.WrapSurfaceAroundPDDIB m_previewImage
                cSurface.SetSurfaceAntialiasing P2_AA_None
                PD2D.DrawRectangleI cSurface, cPenBorder, 0, 0, m_previewImage.GetDIBWidth - 1, m_previewImage.GetDIBHeight - 1
                Set cPenBorder = Nothing
                Set cSurface = Nothing
                
                'To reflect the new image, tequest a redraw from the target picture box
                picPreview.RequestRedraw True
                
            End If
            
        End If
        
    End If
    
PreviewError:

End Sub
