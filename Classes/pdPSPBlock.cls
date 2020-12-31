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
'Copyright 2020-2020 by Tanner Helland
'Created: 30/December/20
'Last updated: 30/December/20
'Last update: start work on initial build
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
    Debug.Print "Found block: ", srcBlkOffset, srcBlkID, srcBlkLength
End Sub
