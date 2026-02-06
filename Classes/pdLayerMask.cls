VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdLayerMask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Layer Mask class
'Copyright 2022-2026 by Tanner Helland
'Created: 24/March/22
'Last updated: 24/March/22
'Last update: initial build
'
'This class is a WIP.  Layer masks are not fully supported by PhotoDemon yet, so this class's interfaces
' are likely to change in the future.  (At present, this class primarily exists to persist layer masks
' to/from PSD files.)
'
'All source code in this file is licensed under a modified BSD license. This means you may use the code in your own
' projects IF you provide attribution. For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'TRUE if mask data exists; FALSE otherwise.  If FALSE, do not attempt to query mask bytes.
Private m_MaskDataExists As Boolean

'Boundary of the rect
Private m_MaskRectF As RectF

'Actual mask bytes.  Dimensions are guaranteed to match GetMaskWidth/Height() values.
Private m_MaskBytes() As Byte

'PSD files have a nice option where you can declare the area around the image as opaque or transparent,
' giving you a way to swiftly track "cut-out" or "spot" masks with minimal file size.  Because the value
' is encoded as a byte, you could hypothetically use it to make everything outside the mask e.g. "half-opaque",
' but I've never seen this behavior in-the-wild and it may not even be supported by Photoshop, so we
' currently lock the value to [0] or [max] when exporting to PSD.
Private m_OpacityOutsideMask As Single

'If the mask came from an outside file format (like PSD), it may behave differently than a normal PhotoDemon mask.
' As of 2022, this is relevant for layer group masks because PD does not support layer groups yet.  These flag(s)
' can be used to note special mask properties that affect the way this mask is handled internally.
Private m_MaskBelongsToPSGroup As Boolean

'Returns TRUE on successful clone, FALSE otherwise; FALSE usually means the source mask doesn't exist
Friend Function CloneExistingMask(ByRef srcMask As pdLayerMask) As Boolean

    'Undefined behavior?  We can either:
    ' 1) Do not modify the existing mask on a failed clone, or...
    ' 2) Erase the current mask on a failed clone
    '
    'Right now we do (2), but you can do (1) by simply omitting the .Reset command below
    If (srcMask Is Nothing) Then
        Me.Reset
        CloneExistingMask = False
    Else
        
        CloneExistingMask = True
        
        'If the source mask doesn't exist, we'll reset ourselves to match;
        ' otherwise, clone all pertinent bits of information from the source.
        If srcMask.DoesMaskExist() Then
            
            'Clone mask header
            With m_MaskRectF
                .Left = srcMask.GetMaskLeft
                .Top = srcMask.GetMaskTop
                .Width = srcMask.GetMaskWidth
                .Height = srcMask.GetMaskHeight
                
                'Clone mask itself
                If (.Width > 0) And (.Height > 0) Then
                    ReDim m_MaskBytes(0 To .Width - 1, 0 To .Height - 1) As Byte
                    If (srcMask.GetPtrToMaskBytes <> 0) Then VBHacks.CopyMemoryStrict VarPtr(m_MaskBytes(0, 0)), srcMask.GetPtrToMaskBytes, .Width * .Height
                Else
                    ReDim m_MaskBytes(0) As Byte
                End If
                
            End With
            
            'Clone any other mask settings
            Me.SetOpacityOutsideMask srcMask.GetOpacityOutsideMask
        
        'If the source mask doesn't exist, that's fine - just reset our internal data to match
        Else
            Me.Reset
        End If
        
    End If

End Function

Friend Function DoesMaskExist() As Boolean
    DoesMaskExist = m_MaskDataExists
End Function

Friend Function EstimateRAMUsage() As Double
    EstimateRAMUsage = m_MaskRectF.Width * m_MaskRectF.Height
End Function

Friend Function GetMaskLeft() As Long
    GetMaskLeft = Int(m_MaskRectF.Left + 0.5!)
End Function

Friend Function GetMaskTop() As Long
    GetMaskTop = Int(m_MaskRectF.Top + 0.5!)
End Function

Friend Function GetMaskWidth() As Long
    GetMaskWidth = Int(m_MaskRectF.Width + 0.5!)
End Function

Friend Function GetMaskHeight() As Long
    GetMaskHeight = Int(m_MaskRectF.Height + 0.5!)
End Function

Friend Sub Reset()
    
    m_MaskDataExists = False
    
    With m_MaskRectF
        .Left = 0!
        .Top = 0!
        .Width = 0!
        .Height = 0!
    End With
    
    Erase m_MaskBytes
    
    m_OpacityOutsideMask = 0!
    
End Sub

Friend Function IsMaskFlag_PSGroup() As Boolean
    IsMaskFlag_PSGroup = m_MaskBelongsToPSGroup
End Function

Friend Sub SetMaskFlag_PSGroup()
    m_MaskBelongsToPSGroup = True
End Sub

Friend Sub SetMaskRect(ByVal mLeft As Long, ByVal mTop As Long, ByVal mWidth As Long, ByVal mHeight As Long)
    With m_MaskRectF
        .Left = mLeft
        .Top = mTop
        .Width = mWidth
        .Height = mHeight
    End With
End Sub

'Opacity is always handled on the [0.0, 100.0] scale
Friend Function GetOpacityOutsideMask() As Single
    GetOpacityOutsideMask = m_OpacityOutsideMask
End Function

Friend Sub SetOpacityOutsideMask(ByVal newOpacity As Single)
    m_OpacityOutsideMask = newOpacity
End Sub

'Do *NOT* call without first querying for non-zero mask width/height.
Friend Function GetPtrToMaskBytes() As Long
    If (Me.GetMaskWidth <> 0) And (Me.GetMaskHeight <> 0) Then GetPtrToMaskBytes = VarPtr(m_MaskBytes(0, 0))
End Function

'Set new mask data.  A non-zero pointer, if provided, will be used to auto-populate the mask array
Friend Sub SetMaskBytes(ByVal newMaskWidth As Long, ByVal newMaskHeight As Long, Optional ByVal ptrToSrcData As Long = 0)
    If (newMaskWidth > 0) And (newMaskHeight > 0) Then
        ReDim m_MaskBytes(0 To newMaskWidth - 1, 0 To newMaskHeight - 1) As Byte
        If (ptrToSrcData <> 0) Then VBHacks.CopyMemoryStrict VarPtr(m_MaskBytes(0, 0)), ptrToSrcData, newMaskWidth * newMaskHeight
        m_MaskDataExists = True
    Else
        m_MaskDataExists = False
    End If
End Sub
