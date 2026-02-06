VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdPSDLayerInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon PSD (PhotoShop Image) Layer Info Container and Parser
'Copyright 2019-2026 by Tanner Helland
'Created: 15/January/19
'Last updated: 24/January/19
'Last update: split layer info storage into a separate class, as we need to use it from multiple places
'
'This class contains layer-subdata pulled from a PSD file.  It is populated by two possible places: a parent
' pdPSD instance, or a parent pdPSDLayer instance.  It has no purpose outside of a PSD parsing context;
' for layer handling inside PhotoDemon, refer to the pdLayer class.
'
'All code in this class is my original work.  It is based off the "official" Adobe spec at this URL
' (link good as of January 2019):
' https://www.adobe.com/devnet-apps/photoshop/fileformatashtml/#50577409_72092
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'PSD files contain a *lot* of information.  To aid debugging, you can activate "verbose" output; this will
' dump all kinds of diagnostic information to the debug log.  (Note that other PSD classes have their own
' version of this constant.)
Private Const PSD_DEBUG_VERBOSE As Boolean = False

'Layers can include additional optional data.  This comes in a variety of shapes and sizes,
' and while we can successfully parse it all, we don't make use of every possible descriptor.
Private Type PSD_AdditionalInfo
    aiSignature As String * 4
    aiKey As String * 4
    aiDataLength As Long
    aiDataBytes() As Byte
End Type

Private m_AdditionalInfo() As PSD_AdditionalInfo
Private m_AdditionalInfoCount As Long

Friend Function DoesKeyExist(ByVal srcKey As String) As Boolean
    DoesKeyExist = (GetIndexOfKey(srcKey) >= 0)
End Function

Friend Function GetInfoCount() As Long
    GetInfoCount = m_AdditionalInfoCount
End Function

Private Function GetIndexOfKey(ByRef srcKey As String) As Long
    
    GetIndexOfKey = -1
    If (m_AdditionalInfoCount <= 0) Then Exit Function
    
    Dim i As Long
    For i = 0 To m_AdditionalInfoCount - 1
        If (m_AdditionalInfo(i).aiKey = srcKey) Then
            GetIndexOfKey = i
            Exit Function
        End If
    Next i
    
End Function

Friend Function GetStreamForKey(ByVal srcKey As String) As pdStream
    
    Dim kIndex As Long
    kIndex = GetIndexOfKey(srcKey)
    If (kIndex >= 0) Then
        Set GetStreamForKey = New pdStream
        With m_AdditionalInfo(kIndex)
            GetStreamForKey.StartStream PD_SM_ExternalPtrBacked, PD_SA_ReadOnly, , .aiDataLength, VarPtr(.aiDataBytes(0))
        End With
        Exit Function
    End If
    
End Function

'NOTE: the passed stream *must* point at a valid block within the PSD file, or this function will fail.
Friend Function ParseAdditionalLayerInfo(ByRef srcStream As pdStream, ByRef warningStack As pdStringStack, ByVal imageIsPSB As Boolean, ByVal finalPointerPos As Long, Optional ByVal infoIsInGlobalArea As Boolean = False) As PD_PSDResult
    
    ParseAdditionalLayerInfo = psd_Success
    
    'Reset internal storage
    ReDim m_AdditionalInfo(0 To 3) As PSD_AdditionalInfo
    m_AdditionalInfoCount = 0
    
    Dim sigCheck As String
    
    Do While (srcStream.GetPosition() < finalPointerPos)
        
        'Verify the signature
        sigCheck = srcStream.ReadString_ASCII(4)
        
        If (sigCheck <> "8BIM") And (sigCheck <> "8B64") Then
            warningStack.AddString "ParseAdditionalLayerInfo found an unknown additional segment signature: " & sigCheck
            ParseAdditionalLayerInfo = psd_Warning
        Else
            
            If (m_AdditionalInfoCount > UBound(m_AdditionalInfo)) Then ReDim Preserve m_AdditionalInfo(0 To m_AdditionalInfoCount * 2 - 1) As PSD_AdditionalInfo
            With m_AdditionalInfo(m_AdditionalInfoCount)
                
                .aiKey = srcStream.ReadString_ASCII(4)
                .aiDataLength = srcStream.ReadLong_BE()
                
                'PSB has another 4 bytes of length here BUT ONLY FOR CERTAIN KEYS
                If imageIsPSB Then
                
                    'Per the spec, in a PSB file "...the following keys have a length count of 8 bytes:
                    ' LMsk, Lr16, Lr32, Layr, Mt16, Mt32, Mtrn, Alph, FMsk, lnk2, FEid, FXid, PxSD.
                    If (.aiKey = "LMsk") Or (.aiKey = "Lr16") Or (.aiKey = "Lr32") Or (.aiKey = "Layr") Then .aiDataLength = srcStream.ReadLong_BE()
                    If (.aiKey = "Mt16") Or (.aiKey = "Mt32") Or (.aiKey = "Mtrn") Or (.aiKey = "Alph") Then .aiDataLength = srcStream.ReadLong_BE()
                    If (.aiKey = "FMsk") Or (.aiKey = "lnk2") Or (.aiKey = "FEid") Or (.aiKey = "FXid") Then .aiDataLength = srcStream.ReadLong_BE()
                    If (.aiKey = "PxSD") Then .aiDataLength = srcStream.ReadLong_BE()
                    
                End If
                
                If PSD_DEBUG_VERBOSE Then PDDebug.LogAction "Additional layer info found: " & .aiKey & " (" & .aiDataLength & " bytes)"
                If (.aiDataLength > 0) Then srcStream.ReadBytes .aiDataBytes, .aiDataLength, True
                
                'Per the spec, length data should be "rounded up to an even byte count."  Unfortunately, like many
                ' things in the spec, this does not appear to be true in practice.  Quoting from the developers
                ' of Paint.NET's PSD plugin (https://github.com/PsdPlugin/PsdPlugin/blob/068c3c6e1629c47a4904fc6f98d4ebc6473ca25d/PsdFile/Layers/LayerInfo.cs):
                
                '// Documentation states that the length is even-padded.  Actually:
                '//   1. Most keys have 4-padded lengths.
                '//   2. However, some keys (LMsk) have even-padded lengths.
                '//   3. Other keys (Txt2, Lr16, Lr32) have unpadded lengths.
                '//
                '// Photoshop writes data that is always 4-padded, even when the stated
                '// length is not a multiple of 4.  The length mismatch seems to occur
                '// only on global layer info.  We do not read extra padding in other
                '// cases because third-party programs are likely to follow the spec.
                
                'Many thanks to those developers for cracking this problem and sharing their solution.
                If infoIsInGlobalArea Then
                    If ((.aiDataLength Mod 4) <> 0) Then srcStream.SetPosition 4 - (.aiDataLength Mod 4), FILE_CURRENT
                End If
                
            End With
            
            m_AdditionalInfoCount = m_AdditionalInfoCount + 1
            
        End If
        
    Loop
    
    If PSD_DEBUG_VERBOSE Then PDDebug.LogAction "pdPSDLayerInfo found " & m_AdditionalInfoCount & " additional info records"

End Function
