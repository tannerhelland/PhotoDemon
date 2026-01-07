VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdPSPBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon PSP (PaintShop Pro) Block Container
'Copyright 2020-2026 by Tanner Helland
'Created: 30/December/20
'Last updated: 02/January/21
'Last update: add helper function for human-readable block names
'
'This class describes a single "block" inside a JASC/Corel Paint Shop Pro image file.
' It has been custom-built for PhotoDemon, with a special emphasis on parsing performance.
'
'Unless otherwise noted, all code in this class is my original work.  I've based my work off the
' "official" PSP spec at this URL (link good as of December 2020):
' ftp://ftp.corel.com/pub/documentation/PSP/
'
'Older PSP specs were also useful.  You may be able to find them here (link good as of December 2020);
' look for files with names like "psp8spec.pdf":
' http://www.telegraphics.com.au/svn/pspformat/trunk
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Initially, blocks are noted by three key attributes:
' 1) their original offset in the parent file
' 2) their "ID" (a hard-coded PSP-specific WORD)
' 3) block size (which does *NOT* include the 4-byte offset and 2-byte ID of the block)
Private Type PSPBlock
    blkOffset As Long
    blkID As PSPBlockID
    blkLength As Long
End Type

Private m_BlockHeader As PSPBlock

Friend Function GetBlockID() As PSPBlockID
    GetBlockID = m_BlockHeader.blkID
End Function

Friend Function GetBlockIDAsString() As String
    
    Select Case m_BlockHeader.blkID
        Case PSP_IMAGE_BLOCK
            GetBlockIDAsString = "General Image Attributes Block (main)"
        Case PSP_CREATOR_BLOCK
            GetBlockIDAsString = "Creator Data Block (main)"
        Case PSP_COLOR_BLOCK
            GetBlockIDAsString = "Color Palette Block (main and sub)"
        Case PSP_LAYER_START_BLOCK
            GetBlockIDAsString = "Layer Bank Block (main)"
        Case PSP_LAYER_BLOCK
            GetBlockIDAsString = "Layer Block (sub)"
        Case PSP_CHANNEL_BLOCK
            GetBlockIDAsString = "Channel Block (sub)"
        Case PSP_SELECTION_BLOCK
            GetBlockIDAsString = "Selection Block (main)"
        Case PSP_ALPHA_BANK_BLOCK
            GetBlockIDAsString = "Alpha Bank Block (main)"
        Case PSP_ALPHA_CHANNEL_BLOCK
            GetBlockIDAsString = "Alpha Channel Block (sub)"
        Case PSP_COMPOSITE_IMAGE_BLOCK
            GetBlockIDAsString = "Composite Image Block (sub)"
        Case PSP_EXTENDED_DATA_BLOCK
            GetBlockIDAsString = "Extended Data Block (main)"
        Case PSP_TUBE_BLOCK
            GetBlockIDAsString = "Picture Tube Data Block (main)"
        Case PSP_ADJUSTMENT_EXTENSION_BLOCK
            GetBlockIDAsString = "Adjustment Layer Block (sub)"
        Case PSP_VECTOR_EXTENSION_BLOCK
            GetBlockIDAsString = "Vector Layer Block (sub)"
        Case PSP_SHAPE_BLOCK
            GetBlockIDAsString = "Vector Shape Block (sub)"
        Case PSP_PAINTSTYLE_BLOCK
            GetBlockIDAsString = "Paint Style Block (sub)"
        Case PSP_COMPOSITE_IMAGE_BANK_BLOCK
            GetBlockIDAsString = "Composite Image Bank (main)"
        Case PSP_COMPOSITE_ATTRIBUTES_BLOCK
            GetBlockIDAsString = "Composite Image Attr. (sub)"
        Case PSP_JPEG_BLOCK
            GetBlockIDAsString = "JPEG Image Block (sub)"
        Case PSP_LINESTYLE_BLOCK
            GetBlockIDAsString = "Line Style Block (sub)"
        Case PSP_TABLE_BANK_BLOCK
            GetBlockIDAsString = "Table Bank Block (main)"
        Case PSP_TABLE_BLOCK
            GetBlockIDAsString = "Table Block (sub)"
        Case PSP_PAPER_BLOCK
            GetBlockIDAsString = "Vector Table Paper Block (sub)"
        Case PSP_PATTERN_BLOCK
            GetBlockIDAsString = "Vector Table Pattern Block (sub)"
        Case PSP_GRADIENT_BLOCK
            GetBlockIDAsString = "Vector Table Gradient Block (not used)"
        Case PSP_GROUP_EXTENSION_BLOCK
            GetBlockIDAsString = "Group Layer Block (sub)"
        Case PSP_MASK_EXTENSION_BLOCK
            GetBlockIDAsString = "Mask Layer Block (sub)"
        Case PSP_BRUSH_BLOCK
            GetBlockIDAsString = "Brush Data Block (main)"
        Case PSP_ART_MEDIA_BLOCK
            GetBlockIDAsString = "Art Media Layer Block (main)"
        Case PSP_ART_MEDIA_MAP_BLOCK
            GetBlockIDAsString = "Art Media Layer map data Block (main)"
        Case PSP_ART_MEDIA_TILE_BLOCK
            GetBlockIDAsString = "Art Media Layer map tile Block (main)"
        Case PSP_ART_MEDIA_TEXTURE_BLOCK
            GetBlockIDAsString = "Art Media Layer map texture Block (main)"
        Case PSP_COLORPROFILE_BLOCK
            GetBlockIDAsString = "ICC Color profile block"
        Case Else
            GetBlockIDAsString = "(unknown)"
    End Select
    
End Function

Friend Function GetBlockLength() As Long
    GetBlockLength = m_BlockHeader.blkLength
End Function

Friend Function GetBlockOffset() As Long
    GetBlockOffset = m_BlockHeader.blkOffset
End Function

Friend Sub InitializeBlock(ByVal srcBlkOffset As Long, ByVal srcBlkID As PSPBlockID, ByVal srcBlkLength As Long)
    m_BlockHeader.blkOffset = srcBlkOffset
    m_BlockHeader.blkID = srcBlkID
    m_BlockHeader.blkLength = srcBlkLength
End Sub
