Attribute VB_Name = "Menu_Icons"
'***************************************************************************
'Specialized Icon Handler
'Copyright �2011-2012 by Tanner Helland
'Created: 24/June/12
'Last updated: 06/August/12
'Last update: Added CreateCustomFormIcon, which sets the icon of an MDI child form to match its contained image
'
'Because VB6 doesn't provide many mechanisms for working with icons, I've had to manually add a number of
' icon-related functions to PhotoDemon.  First is a way to add icons/bitmaps to menus, as originally written
' by Leandro Ascierto.  Menu icons are extracted from a resource file (where they're stored in PNG format) and
' rendered to the menu at run-time.  See the clsMenuImage class for details on how this works.
' (A link to Leandro's original project can also be found there.)
'
'NOTE: Because the Windows XP version of Leandro's code utilizes potentially dirty subclassing,
' PhotoDemon automatically disables menu icons while running in the IDE on Windows XP.  Compile the project to see icons.
' (Windows Vista and 7 use a different mechanism, so menu icons are enabled in the IDE, and menu icons appear on all
' versions of Windows when compiled.)
'
'Also in this module is a heavily modified version of Paul Turcksin's "Icon Handlemaker" project, which I've modified
' to convert bitmaps to icons on the fly (the "CreateCustomFormIcon" sub).  PhotoDemon uses this to dynamically change
' the icon of its MDI child forms.  To see Paul's original project, please visit this PSC link:
' http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=60600&lngWId=1
'
'***************************************************************************

Option Explicit

'API calls for building an icon at run-time
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateIconIndirect Lib "user32" (icoInfo As ICONINFO) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long
'Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lppictDesc As pictDesc, riid As Guid, ByVal fown As Long, ipic As IPicture) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long

'Types required by the above API calls
Private Type BITMAP
   bmType As Long
   bmWidth As Long
   bmHeight As Long
   bmWidthBytes As Long
   bmPlanes As Integer
   bmBitsPixel As Integer
   bmBits As Long
End Type

Private Type ICONINFO
   fIcon As Boolean
   xHotspot As Long
   yHotspot As Long
   hBMMask As Long
   hBMColor As Long
   End Type

'This array will be used to store our icon handles so we can delete them on program exit
Dim numOfIcons As Long
Dim iconHandles() As Long

'The types and constants below (commented out) can be used to generate an icon object for use within VB

'Private Type Guid
'   Data1 As Long
'   Data2 As Integer
'   Data3 As Integer
'   Data4(7) As Byte
'   End Type

'Private Type pictDesc
'   cbSizeofStruct As Long
'   picType As Long
'   hImage As Long
'End Type

'Constants required by the icon-related API calls
'Private Const PICTYPE_BITMAP = 1
'Private Const PICTYPE_ICON = 3

'API call for manually setting a 32-bit icon to a form (as opposed to Form.Icon = ...)
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'clsMenuImage does the heavy lifting for inserting icons into menus
Dim cMenuImage As clsMenuImage


'Load all the menu icons from PhotoDemon's embedded resource file
Public Sub LoadMenuIcons()

    Set cMenuImage = New clsMenuImage

    With cMenuImage
    
        'Disable menu icon drawing if on Windows XP and uncompiled (to prevent subclassing crashes on unclean IDE breaks)
        If (Not .IsWindowVistaOrLater) And (IsProgramCompiled = False) Then Exit Sub
        
        .Init FormMain.hWnd, 16, 16
        
        .AddImageFromStream LoadResData("OPENIMG", "CUSTOM")    '0
        .AddImageFromStream LoadResData("OPENREC", "CUSTOM")    '1
        .AddImageFromStream LoadResData("IMPORT", "CUSTOM")     '2
        .AddImageFromStream LoadResData("SAVE", "CUSTOM")       '3
        .AddImageFromStream LoadResData("SAVEAS", "CUSTOM")     '4
        .AddImageFromStream LoadResData("CLOSE", "CUSTOM")      '5
        .AddImageFromStream LoadResData("BCONVERT", "CUSTOM")   '6
        .AddImageFromStream LoadResData("PRINT", "CUSTOM")      '7
        .AddImageFromStream LoadResData("SCANNER", "CUSTOM")    '8
        .AddImageFromStream LoadResData("DOWNLOAD", "CUSTOM")   '9
        .AddImageFromStream LoadResData("SCREENCAP", "CUSTOM")  '10
        .AddImageFromStream LoadResData("FRXIMPORT", "CUSTOM")  '11
        .AddImageFromStream LoadResData("UNDO", "CUSTOM")       '12
        .AddImageFromStream LoadResData("REDO", "CUSTOM")       '13
        .AddImageFromStream LoadResData("REPEAT", "CUSTOM")     '14
        .AddImageFromStream LoadResData("COPY", "CUSTOM")       '15
        .AddImageFromStream LoadResData("PASTE", "CUSTOM")      '16
        .AddImageFromStream LoadResData("CLEAR", "CUSTOM")      '17
        .AddImageFromStream LoadResData("PREFERENCES", "CUSTOM") '18
        .AddImageFromStream LoadResData("RESIZE", "CUSTOM")     '19
        .AddImageFromStream LoadResData("ROTATECW", "CUSTOM")   '20
        .AddImageFromStream LoadResData("ROTATECCW", "CUSTOM")  '21
        .AddImageFromStream LoadResData("ROTATE180", "CUSTOM")  '22
        .AddImageFromStream LoadResData("FLIP", "CUSTOM")       '23
        .AddImageFromStream LoadResData("MIRROR", "CUSTOM")     '24
        .AddImageFromStream LoadResData("PDWEBSITE", "CUSTOM")  '25
        .AddImageFromStream LoadResData("FEEDBACK", "CUSTOM")   '26
        .AddImageFromStream LoadResData("ABOUT", "CUSTOM")      '27
        .AddImageFromStream LoadResData("FITWINIMG", "CUSTOM")  '28
        .AddImageFromStream LoadResData("FITONSCREEN", "CUSTOM") '29
        .AddImageFromStream LoadResData("TILEHOR", "CUSTOM")    '30
        .AddImageFromStream LoadResData("TILEVER", "CUSTOM")    '31
        .AddImageFromStream LoadResData("CASCADE", "CUSTOM")    '32
        .AddImageFromStream LoadResData("ARNGICONS", "CUSTOM")  '33
        .AddImageFromStream LoadResData("MINALL", "CUSTOM")     '34
        .AddImageFromStream LoadResData("RESTOREALL", "CUSTOM")     '35
        .AddImageFromStream LoadResData("OPENMACRO", "CUSTOM")  '36
        .AddImageFromStream LoadResData("RECORD", "CUSTOM")     '37
        .AddImageFromStream LoadResData("RECORDSTOP", "CUSTOM") '38
        .AddImageFromStream LoadResData("BUG", "CUSTOM") '39
        .AddImageFromStream LoadResData("FAVORITE", "CUSTOM") '40
        
        
        'File Menu
        .PutImageToVBMenu 0, 0, 0       'Open Image
        .PutImageToVBMenu 1, 1, 0       'Open recent
        .PutImageToVBMenu 2, 2, 0       'Import
        .PutImageToVBMenu 3, 4, 0       'Save
        .PutImageToVBMenu 4, 5, 0       'Save As...
        .PutImageToVBMenu 5, 7, 0       'Close...
        .PutImageToVBMenu 6, 9, 0       'Batch conversion
        .PutImageToVBMenu 7, 11, 0      'Print
        
        '--> Import Sub-Menu
        'NOTE: the specific menu values will be different if the scanner plugin (eztw32.dll) isn't found.
        If ScanEnabled = True Then
            .PutImageToVBMenu 8, 0, 0, 2       'Scan Image
            .PutImageToVBMenu 9, 3, 0, 2       'Download Image
            .PutImageToVBMenu 10, 5, 0, 2      'Capture Screen
            .PutImageToVBMenu 11, 7, 0, 2      'Import from FRX
        Else
            .PutImageToVBMenu 9, 0, 0, 2       'Download Image
            .PutImageToVBMenu 10, 2, 0, 2      'Capture Screen
            .PutImageToVBMenu 11, 4, 0, 2      'Import from FRX
        End If
        
        'Edit Menu
        .PutImageToVBMenu 12, 0, 1      'Undo
        .PutImageToVBMenu 13, 1, 1      'Redo
        .PutImageToVBMenu 14, 2, 1      'Repeat Last Action
        .PutImageToVBMenu 15, 4, 1      'Copy
        .PutImageToVBMenu 16, 5, 1      'Paste
        .PutImageToVBMenu 17, 6, 1      'Empty Clipboard
        .PutImageToVBMenu 18, 8, 1      'Program Preferences
        
        'Image Menu
        .PutImageToVBMenu 19, 0, 2      'Resize
        .PutImageToVBMenu 20, 0, 2, 3     'Rotate Clockwise (rotate submenu)
        .PutImageToVBMenu 21, 1, 2, 3     'Rotate Counter-clockwise (rotate submenu)
        .PutImageToVBMenu 22, 2, 2, 3     'Rotate 180 (rotate submenu)
        .PutImageToVBMenu 23, 4, 2      'Flip
        .PutImageToVBMenu 24, 5, 2      'Mirror
        
        'Macro Menu
        .PutImageToVBMenu 36, 0, 5     'Open Macro
        .PutImageToVBMenu 37, 2, 5     'Start Recording
        .PutImageToVBMenu 38, 3, 5     'Stop Recording
        
        'Window Menu
        .PutImageToVBMenu 28, 0, 6     'Fit Window to Image
        .PutImageToVBMenu 29, 1, 6     'Fit on Screen
        .PutImageToVBMenu 30, 3, 6     'Tile Horizontally
        .PutImageToVBMenu 31, 4, 6     'Tile Vertically
        .PutImageToVBMenu 32, 5, 6     'Cascade
        .PutImageToVBMenu 33, 6, 6     'Arrange Icons
        .PutImageToVBMenu 34, 8, 6     'Minimize All
        .PutImageToVBMenu 35, 9, 6     'Restore All
        
        'Help Menu
        .PutImageToVBMenu 40, 0, 7     'Donate
        .PutImageToVBMenu 25, 2, 7     'Visit the PhotoDemon website
        .PutImageToVBMenu 26, 3, 7     'Submit Feedback
        .PutImageToVBMenu 39, 4, 7     'Submit Bug
        .PutImageToVBMenu 27, 6, 7     'About PD
    
    End With

End Sub

'Create a custom form icon for an MDI child form (using the image stored in the back buffer of imgForm)
'Again, thanks to Paul Turcksin for the original draft of this code.
Public Sub CreateCustomFormIcon(ByRef imgForm As FormImage)

    'Generating an icon requires many variables; see below for specific comments on each one
    Dim bitmapData As BITMAP
    Dim iWidth As Long
    Dim iHeight As Long
    Dim srcDC As Long
    Dim oldSrcObj As Long
    Dim MonoDC As Long
    Dim MonoBmp As Long
    Dim oldMonoObj As Long
    Dim InvertDC As Long
    Dim InvertBmp As Long
    Dim oldInvertObj As Long
    Dim cBkColor As Long
    Dim maskClr As Long
    Dim icoInfo As ICONINFO
   
    'The icon can be drawn at any size, but 16x16 is how it will (typically) end up on the form.  I use 32x32 here
    ' in order to get slightly higher quality stretching during the resampling phase.
    Dim icoSize As Long
    icoSize = 32

    'Draw the form's backbuffer image to an in-memory picture
    Dim aspectRatio As Single
    aspectRatio = CSng(imgForm.BackBuffer.ScaleWidth) / CSng(imgForm.BackBuffer.ScaleHeight)
    
    'The target icon's width and height, x and y positioning
    Dim tIcoWidth As Single, tIcoHeight As Single, tX As Single, tY As Single
    
    'If the form is wider than it is tall...
    If aspectRatio > 1 Then
        
        'Determine proper sizes and (x, y) positioning so the icon will be centered
        tIcoWidth = icoSize
        tIcoHeight = icoSize * (1 / aspectRatio)
        tX = 0
        tY = (icoSize - tIcoHeight) / 2
        
    Else
    
        'Same thing, but with the math adjusted for images taller than they are wide
        tIcoHeight = icoSize
        tIcoWidth = icoSize * aspectRatio
        tY = 0
        tX = (icoSize - tIcoWidth) / 2
        
    End If
    
    'Resize the picture box that will receive the first draft of the icon
    imgForm.picIcon.Width = icoSize
    imgForm.picIcon.Height = icoSize
    
    'Because we'll be shrinking the image dramatically, set StretchBlt to use resampling
    SetStretchBltMode imgForm.picIcon.hDC, STRETCHBLT_HALFTONE
    
    'Render the bitmap that will ultimately be converted into an icon
    StretchBlt imgForm.picIcon.hDC, CLng(tX), CLng(tY), CLng(tIcoWidth), CLng(tIcoHeight), imgForm.BackBuffer.hDC, 0, 0, imgForm.BackBuffer.ScaleWidth, imgForm.BackBuffer.ScaleHeight, vbSrcCopy
    imgForm.picIcon.Picture = imgForm.picIcon.Image
   
    'Now that we have a first draft to work from, start preparing the data types required by the icon API calls
    GetObject imgForm.picIcon.Picture.Handle, Len(bitmapData), bitmapData

    With bitmapData
        iWidth = .bmWidth
        iHeight = .bmHeight
    End With
   
    'Create a copy of the original image; this will be used to generate a mask (necessary if the image isn't square-shaped)
    srcDC = CreateCompatibleDC(0&)
    oldSrcObj = SelectObject(srcDC, imgForm.picIcon.Picture.Handle)
   
    'If the image isn't square-shaped, the backcolor of the first draft image will need to be made transparent
    If tIcoWidth < icoSize Or tIcoHeight < icoSize Then
        maskClr = imgForm.picIcon.BackColor
    Else
        maskClr = 0
    End If
   
    'Generate two masks.  First, a monochrome mask.
    MonoDC = CreateCompatibleDC(0&)
    MonoBmp = CreateCompatibleBitmap(MonoDC, iWidth, iHeight)
    oldMonoObj = SelectObject(MonoDC, MonoBmp)
    cBkColor = GetBkColor(srcDC)
    SetBkColor srcDC, maskClr
    BitBlt MonoDC, 0, 0, iWidth, iHeight, srcDC, 0, 0, vbSrcCopy
    SetBkColor srcDC, cBkColor
    
    'Second, an AND mask
    InvertDC = CreateCompatibleDC(0&)
    InvertBmp = CreateCompatibleBitmap(imgForm.hDC, iWidth, iHeight)
    oldInvertObj = SelectObject(InvertDC, InvertBmp)
    BitBlt InvertDC, 0, 0, iWidth, iHeight, srcDC, 0, 0, vbSrcCopy
    SetBkColor InvertDC, vbBlack
    SetTextColor InvertDC, vbWhite
    BitBlt InvertDC, 0, 0, iWidth, iHeight, MonoDC, 0, 0, vbSrcAnd
  
    'We no longer need our copy of the original image, so free up that memory
    SelectObject srcDC, oldSrcObj
    DeleteDC srcDC
    'We can also free up the temporary DCs used to generate our two masks
    SelectObject MonoDC, oldMonoObj
    SelectObject InvertDC, oldInvertObj

    'Populate the icon header
    With icoInfo
      .fIcon = True
      .xHotspot = icoSize
      .yHotspot = icoSize
      .hBMMask = MonoBmp
      .hBMColor = InvertBmp
    End With
      
    'Render the icon to a handle
    Dim generatedIcon As Long
    generatedIcon = CreateIconIndirect(icoInfo)
    
    'Clear out our temporary masks (whose info are now embedded in the icon itself)
    DeleteObject icoInfo.hBMMask
    DeleteObject icoInfo.hBMColor
    DeleteDC MonoDC
    DeleteDC InvertDC
   
    'Use the API to assign this new icon to the specified MDI child form
    SendMessageLong imgForm.hWnd, &H80, 0, generatedIcon
    
    'Store this icon in our running list, so we can destroy it when the program is closed
    addIconToList generatedIcon
   
    'The chunk of code below will generate an actual icon object for use within VB.  I don't use this mechanism because
    ' VB will internally convert the icon to 256-colors before assigning it to the form. <sigh>  Rather than do that,
    ' I use an alternate API call above to assign the new icon in its transparent, full color glory.
    
    'Dim iGuid As Guid
    'With iGuid
    '   .Data1 = &H20400
    '   .Data4(0) = &HC0
    '   .Data4(7) = &H46
    'End With
    
    'Dim pDesc As pictDesc
    'With pDesc
    '   .cbSizeofStruct = Len(pDesc)
    '   .picType = PICTYPE_ICON
    '   .hImage = generatedIcon
    'End With
    
    'Dim icoObject As IPicture
    'OleCreatePictureIndirect pDesc, iGuid, 1, icoObject
    
    'imgForm.Icon = icoObject
   
End Sub

'Needs to be run only once, at the start of the program
Public Sub initializeIconHandler()
    numOfIcons = 0
End Sub

'Add another icon reference to the list
Private Sub addIconToList(ByVal hIcon As Long)

    ReDim Preserve iconHandles(0 To numOfIcons) As Long
    iconHandles(numOfIcons) = hIcon
    numOfIcons = numOfIcons + 1

End Sub

'Remove all icons generated since the program launched
Public Sub destroyAllIcons()

    Dim i As Long
    For i = 0 To numOfIcons - 1
        DestroyIcon iconHandles(i)
    Next i
    
    Erase iconHandles

End Sub
