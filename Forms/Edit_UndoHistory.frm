VERSION 5.00
Begin VB.Form FormUndoHistory 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Undo history"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9135
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
   ScaleHeight     =   423
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   609
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7590
      TabIndex        =   4
      Top             =   5670
      Width           =   1365
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&Restore selected state"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Top             =   5670
      Width           =   3645
   End
   Begin VB.PictureBox picBuffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      FillColor       =   &H80000007&
      ForeColor       =   &H80000008&
      Height          =   4485
      Left            =   240
      ScaleHeight     =   297
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   551
      TabIndex        =   2
      Top             =   600
      Width           =   8295
   End
   Begin VB.VScrollBar vsBuffer 
      Height          =   4425
      LargeChange     =   32
      Left            =   8520
      TabIndex        =   1
      Top             =   600
      Width           =   330
   End
   Begin VB.Label lblBackground 
      Height          =   855
      Left            =   0
      TabIndex        =   5
      Top             =   5520
      Width           =   11535
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "available image states"
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
      Left            =   240
      TabIndex        =   0
      Top             =   150
      Width           =   2310
   End
End
Attribute VB_Name = "FormUndoHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Undo History dialog
'Copyright 2014-2015 by Tanner Helland
'Created: 14/July/14
'Last updated: 14/July/14
'Last update: initial build
'
'This is a first draft of a functional Undo History browser for PD.  Most applications provide this as a floating
' toolbar, but because that would require some complicated UI work (including integration into PD's window manager),
' I'm postponing such an implementation until after we've gotten the browser working first.
'
'All previous image states, including selections, are available for restoration.
'
'Obviously, this dialog interacts heavily with the pdUndo class, as only the undo manager has access to the full
' Undo/Redo stack, including detailed information like process IDs, Undo file types, etc.
'
'When the user selects a point for restoration, the Undo/Redo manager handles the actual work of restoring the image
' to that point.  This dialog simply presents the list to the user, and returns a clicked index position to pdUndo.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'This array will contain the contents of the current Undo stack, as copied from the pdUndo class
Dim undoEntries() As undoEntry

'Total number of Undo entries, and index of the current Undo entry.
Dim numOfUndos As Long, curUndoIndex As Long

'Height of each Undo content block
Private Const BLOCKHEIGHT As Long = 53

'An outside class provides access to mousewheel events for scrolling the filter view
Private WithEvents cMouseEvents As pdInputMouse
Attribute cMouseEvents.VB_VarHelpID = -1
Private WithEvents cKeyEvents As pdInputKeyboard
Attribute cKeyEvents.VB_VarHelpID = -1

'Extra variables for custom list rendering
Dim bufferDIB As pdDIB
Dim m_BufferWidth As Long, m_BufferHeight As Long

'Two font objects; one for names and one for descriptions.  (Two are needed because they have different sizes and colors,
' and it is faster to cache these values rather than constantly recreating them on a single pdFont object.)
Dim firstFont As pdFont, secondFont As pdFont

'A primary and secondary color for font rendering
Dim primaryColor As Long, secondaryColor As Long

'The currently selected and currently hovered undo entry
Dim curBlock As Long, curBlockHover As Long

'Redraw the current list of undo entries
Private Sub redrawUndoList()
        
    Dim scrollOffset As Long
    scrollOffset = vsBuffer.Value
    
    bufferDIB.createBlank picBuffer.ScaleWidth, picBuffer.ScaleHeight
    
    Dim i As Long
    For i = 0 To numOfUndos - 1
        renderUndoBlock i, 0, FixDPI(i * BLOCKHEIGHT) - scrollOffset - FixDPI(2)
    Next i
    
    'Copy the buffer to the main form
    BitBlt picBuffer.hDC, 0, 0, m_BufferWidth, m_BufferHeight, bufferDIB.getDIBDC, 0, 0, vbSrcCopy
    picBuffer.Picture = picBuffer.Image
    picBuffer.Refresh
    
End Sub

'Render an individual "block" for a given filter (including name, description, color)
Private Sub renderUndoBlock(ByVal blockIndex As Long, ByVal offsetX As Long, ByVal offsetY As Long)

    'Only draw the current block if it will be visible
    If ((offsetY + FixDPI(BLOCKHEIGHT)) > 0) And (offsetY < m_BufferHeight) Then
    
        offsetY = offsetY + FixDPI(2)
        
        Dim linePadding As Long
        linePadding = FixDPI(2)
    
        Dim mHeight As Single
        Dim tmpRect As RECTL
        Dim hBrush As Long
        
        'If this filter has been selected, draw the background with the system's current selection color
        If blockIndex = curBlock Then
        
            SetRect tmpRect, offsetX, offsetY, m_BufferWidth, offsetY + FixDPI(BLOCKHEIGHT)
            hBrush = CreateSolidBrush(ConvertSystemColor(vbHighlight))
            FillRect bufferDIB.getDIBDC, tmpRect, hBrush
            DeleteObject hBrush
            
            'Also, color the fonts with the matching highlighted text color (otherwise they won't be readable)
            firstFont.SetFontColor ConvertSystemColor(vbHighlightText)
            secondFont.SetFontColor ConvertSystemColor(vbHighlightText)
        
        Else
            firstFont.SetFontColor primaryColor
            secondFont.SetFontColor secondaryColor
        End If
        
        'If the current filter is highlighted but not selected, simply render the border with a highlight
        If (blockIndex <> curBlock) And (blockIndex = curBlockHover) Then
            SetRect tmpRect, offsetX, offsetY, m_BufferWidth, offsetY + FixDPI(BLOCKHEIGHT)
            hBrush = CreateSolidBrush(ConvertSystemColor(vbHighlight))
            FrameRect bufferDIB.getDIBDC, tmpRect, hBrush
            DeleteObject hBrush
        End If
        
        Dim drawString As String
        drawString = ""
        
        If (blockIndex + 1) = curUndoIndex Then drawString = "* "
        drawString = drawString & blockIndex & " - " & g_Language.TranslateMessage(undoEntries(blockIndex).processID)
        
        'Render the thumbnail for this entry onto its block
        Dim thumbWidth As Long
        thumbWidth = offsetX + FixDPI(4) + undoEntries(blockIndex).thumbnailSmall.getDIBWidth
        undoEntries(blockIndex).thumbnailSmall.alphaBlendToDC bufferDIB.getDIBDC, 255, offsetX + FixDPI(4), offsetY + ((FixDPI(BLOCKHEIGHT) - undoEntries(blockIndex).thumbnailSmall.getDIBHeight) \ 2)
            
        'Render the index and name fields
        firstFont.AttachToDC bufferDIB.getDIBDC
        firstFont.FastRenderText thumbWidth + FixDPI(16) + offsetX, offsetY + FixDPI(4), drawString
        firstFont.ReleaseFromDC
                
        'Below that, add the description text
        mHeight = firstFont.GetHeightOfString(drawString) + linePadding
        drawString = getStringForUndoType(undoEntries(blockIndex).undoType, undoEntries(blockIndex).undoLayerID)
        
        secondFont.AttachToDC bufferDIB.getDIBDC
        secondFont.FastRenderText thumbWidth + FixDPI(16) + offsetX, offsetY + FixDPI(4) + mHeight, drawString
        secondFont.ReleaseFromDC
        
    End If

End Sub

Private Function getStringForUndoType(ByVal typeOfUndo As PD_UNDO_TYPE, Optional ByVal layerID As Long = 0) As String

    Dim newText As String
    
    Select Case typeOfUndo
    
        Case UNDO_EVERYTHING
            newText = ""
            
        Case UNDO_IMAGE, UNDO_IMAGE_VECTORSAFE, UNDO_IMAGEHEADER
            newText = ""
            
        Case UNDO_LAYER, UNDO_LAYER_VECTORSAFE, UNDO_LAYERHEADER
            If Not (pdImages(g_CurrentImage).getLayerByID(layerID) Is Nothing) Then
                newText = pdImages(g_CurrentImage).getLayerByID(layerID).getLayerName()
            Else
                newText = ""
            End If
        
        Case UNDO_SELECTION
            newText = g_Language.TranslateMessage("selection shape shown")
        
    End Select
    
    getStringForUndoType = newText

End Function

Private Sub cKeyEvents_KeyDownCustom(ByVal Shift As ShiftConstants, ByVal vkCode As Long, markEventHandled As Boolean)

    'Up and down arrows navigate the list
    If (vkCode = VK_UP) Or (vkCode = VK_DOWN) Then
    
        If (vkCode = VK_UP) Then
            curBlock = curBlock - 1
            If curBlock < 0 Then curBlock = numOfUndos - 1
        End If
        
        If (vkCode = VK_DOWN) Then
            curBlock = curBlock + 1
            If curBlock >= numOfUndos Then curBlock = 0
        End If
        
        'Calculate a new vertical scroll position so that the selected filter appears on-screen
        Dim newScrollOffset As Long
        newScrollOffset = curBlock * FixDPI(BLOCKHEIGHT)
        If newScrollOffset > vsBuffer.Max Then newScrollOffset = vsBuffer.Max
        vsBuffer.Value = newScrollOffset
        
        'Redraw the custom filter list
        redrawUndoList
        
    End If

End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOK_Click()
    
    Me.Visible = False
    Process "Undo history", , buildParams(curBlock + 1), UNDO_NOTHING
    Unload Me
    
End Sub

'When the mouse leaves the filter box, remove any hovered entries and redraw
Private Sub cMouseEvents_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    curBlockHover = -1
    redrawUndoList
End Sub

Private Sub cMouseEvents_MouseWheelVertical(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)

    'Vertical scrolling - only trigger it if the vertical scroll bar is actually visible
    If vsBuffer.Visible Then
  
        If scrollAmount < 0 Then
            
            If vsBuffer.Value + vsBuffer.LargeChange > vsBuffer.Max Then
                vsBuffer.Value = vsBuffer.Max
            Else
                vsBuffer.Value = vsBuffer.Value + vsBuffer.LargeChange
            End If
            
            curBlockHover = getUndoAtPosition(x, y)
            redrawUndoList
        
        ElseIf scrollAmount > 0 Then
            
            If vsBuffer.Value - vsBuffer.LargeChange < vsBuffer.Min Then
                vsBuffer.Value = vsBuffer.Min
            Else
                vsBuffer.Value = vsBuffer.Value - vsBuffer.LargeChange
            End If
            
            curBlockHover = getUndoAtPosition(x, y)
            redrawUndoList
            
        End If
        
    End If

End Sub

Private Sub Form_Activate()
    
    'Apply translations and visual themes
    MakeFormPretty Me
    
    'Redraw the undo list
    redrawUndoList
    
End Sub

Private Sub Form_Load()
    
    'Enable mousewheel scrolling for the filter box
    Set cMouseEvents = New pdInputMouse
    cMouseEvents.addInputTracker picBuffer.hWnd, True, , , True
    cMouseEvents.addInputTracker Me.hWnd
    cMouseEvents.setSystemCursor IDC_HAND
    
    'Enable some key events as well
    Set cKeyEvents = New pdInputKeyboard
    cKeyEvents.createKeyboardTracker "Undo History picBuffer", picBuffer.hWnd, VK_UP, VK_DOWN
    
    'Create a background buffer the same size as the buffer picture box
    Set bufferDIB = New pdDIB
    bufferDIB.createBlank picBuffer.ScaleWidth, picBuffer.ScaleHeight
    
    'Initialize a few other variables now (for performance reasons)
    m_BufferWidth = picBuffer.ScaleWidth
    m_BufferHeight = picBuffer.ScaleHeight
    
    'Initialize a custom font object for names
    primaryColor = RGB(64, 64, 64)
    Set firstFont = New pdFont
    firstFont.SetFontColor primaryColor
    firstFont.SetFontBold True
    firstFont.SetFontSize 12
    firstFont.CreateFontObject
    firstFont.SetTextAlignment vbLeftJustify
    
    '...and a second custom font object for descriptions
    secondaryColor = RGB(92, 92, 92)
    Set secondFont = New pdFont
    secondFont.SetFontColor secondaryColor
    secondFont.SetFontBold False
    secondFont.SetFontSize 10
    secondFont.CreateFontObject
    secondFont.SetTextAlignment vbLeftJustify
    
    'Retrieve a copy of all Undo data from the current image's undo manager
    pdImages(g_CurrentImage).undoManager.copyUndoStack numOfUndos, curUndoIndex, undoEntries
    
    'Select the current undo state by default
    curBlock = curUndoIndex - 1
    curBlockHover = -1
    
    'Determine if the vertical scrollbar needs to be visible or not
    Dim maxListSize As Long
    maxListSize = FixDPIFloat(BLOCKHEIGHT) * numOfUndos - 1
    
    vsBuffer.Value = 0
    If maxListSize < picBuffer.ScaleHeight Then
        vsBuffer.Visible = False
    Else
        vsBuffer.Visible = True
        vsBuffer.Max = maxListSize - picBuffer.ScaleHeight
        
        'We also want to calculate an ideal position for the vertical scroll bar, so that the current image state
        ' is displayed in the center of the box by default.  (This gives the user a chance to see several actions
        ' above and below the current state.)
        Dim idealPosition As Long
        idealPosition = curBlock * FixDPIFloat(BLOCKHEIGHT) - ((picBuffer.ScaleHeight - FixDPIFloat(BLOCKHEIGHT)) / 2)
        
        If idealPosition < vsBuffer.Max Then
            If idealPosition < 0 Then idealPosition = 0
            vsBuffer.Value = idealPosition
        Else
            vsBuffer.Value = vsBuffer.Max
        End If
        
    End If
    
    vsBuffer.Height = picBuffer.Height
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
      
    'Unload the mouse tracker
    Set cMouseEvents = Nothing
    ReleaseFormTheming Me
        
End Sub

Private Sub picBuffer_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    curBlock = getUndoAtPosition(x, y)
    redrawUndoList
    
End Sub

Private Sub picBuffer_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    curBlockHover = getUndoAtPosition(x, y)
    redrawUndoList
    
End Sub

'Given mouse coordinates over the buffer picture box, return the filter at that location
Private Function getUndoAtPosition(ByVal x As Long, ByVal y As Long) As Long
    
    Dim vOffset As Long
    vOffset = vsBuffer.Value
    
    getUndoAtPosition = (y + vOffset) \ FixDPI(BLOCKHEIGHT)
    
End Function

Private Sub vsBuffer_Change()
    redrawUndoList
End Sub

Private Sub vsBuffer_Scroll()
    redrawUndoList
End Sub
