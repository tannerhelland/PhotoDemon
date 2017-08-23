VERSION 5.00
Begin VB.UserControl pdCanvasView 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
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
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "pdCanvasView.ctx":0000
End
Attribute VB_Name = "pdCanvasView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon "CanvasView" User Control (e.g. the primary component of pdCanvas)
'Copyright 2002-2017 by Tanner Helland
'Created: 29/November/02
'Last updated: 16/February/16
'Last update: migrate the main view portions of pdCanvas into this control, which will greatly simplify paint tool integration
'
'In 2013, PD's canvas was rebuilt as a dedicated user control, and instead of each image maintaining its own canvas inside
' separate, dedicated windows (which required a *ton* of code to keep in sync with the main PD window), a single canvas was
' integrated directly into the main window, and shared by all windows.
'
'In 2016, we refined this further, embedding this unique pdCanvasView object inside the larger pdCanvas object.  pdCanvas has
' manage a lot of things - scroll bars, buttons, a full status bar - and by migrating the main viewer portion into its own
' usercontrol, we gained a lot of flexibility over performance and code separation.  This is especially important as part of
' improving on-canvas tool responsiveness.
'
'To really understand how this control operates, you'll need to examine pdCanvas, as it ultimately deals with the many mouse
' and key events we potentially raise.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Because VB focus events are wonky, especially when we use CreateWindow within a UC, this control raises its own
' specialized focus events.  If you need to track focus, use these instead of the default VB functions.
Public Event GotFocusAPI()
Public Event LostFocusAPI()

'This control bubbles every input event up to its parent control (pdCanvas).  Please do not add any tool-specific handling
' to this control instance; that level of decision-making should happen inside pdCanvas itself.
Public Event MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
Public Event MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
Public Event MouseHover(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
Public Event MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
Public Event MouseUpCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal clickEventAlsoFiring As Boolean, ByVal timeStamp As Long)
Public Event MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
Public Event MouseWheelVertical(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)
Public Event MouseWheelHorizontal(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)
Public Event MouseWheelZoom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal zoomAmount As Double)
Public Event ClickCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
Public Event DoubleClickCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
Public Event KeyDownCustom(ByVal Shift As ShiftConstants, ByVal vkCode As Long, ByRef markEventHandled As Boolean)
Public Event KeyUpCustom(ByVal Shift As ShiftConstants, ByVal vkCode As Long, ByRef markEventHandled As Boolean)
Public Event AppCommand(ByVal cmdID As AppCommandConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)

'To improve performance, external functions can request that we ignore any refresh requests.
' IMPORTANT NOTE: IT IS IMPERATIVE THAT YOU SET THIS VALUE CORRECTLY.  If you set it and forget to release it,
' the canvas will be locked in an inactive state.
Private m_SuspendRedraws As Boolean

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since attempted to wrap these into a single master control support class.
Private WithEvents ucSupport As pdUCSupport
Attribute ucSupport.VB_VarHelpID = -1

'Local list of themable colors.  This list includes all potential colors used by this class, regardless of state change
' or internal control settings.  The list is updated by calling the UpdateColorList function.
' (Note also that this list does not include variants, e.g. "BorderColor" vs "BorderColor_Hovered".  Variant values are
'  automatically calculated by the color management class, and they are retrieved by passing boolean modifiers to that
'  class, rather than treating every imaginable variant as a separate constant.)
Private Enum PDCANVAS_COLOR_LIST
    [_First] = 0
    PDC_Background = 0
    [_Last] = 0
    [_Count] = 1
End Enum

'Color retrieval and storage is handled by a dedicated class; this allows us to optimize theme interactions,
' without worrying about the details locally.
Private m_Colors As pdThemeColors

Public Function GetControlType() As PD_ControlType
    GetControlType = pdct_CanvasView
End Function

Public Function GetControlName() As String
    GetControlName = UserControl.Extender.Name
End Function

'The Enabled property is a bit unique; see http://msdn.microsoft.com/en-us/library/aa261357%28v=vs.60%29.aspx
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
    UserControl.Enabled = newValue
    PropertyChanged "Enabled"
End Property

Public Property Get hWnd()
    hWnd = UserControl.hWnd
End Property

'This is one of the few controls in PD that manages a persistent DC.  This is a deliberate decision, so we can maintain expensive
' per-DC settings like color management.
Public Property Get hDC()
    hDC = ucSupport.GetBackBufferDC
End Property

Public Function GetCanvasWidth() As Long
    GetCanvasWidth = ucSupport.GetControlWidth
End Function

Public Function GetCanvasHeight() As Long
    GetCanvasHeight = ucSupport.GetControlHeight
End Function

'To support high-DPI settings properly, we expose some specialized move+size functions
Public Function GetLeft() As Long
    GetLeft = ucSupport.GetControlLeft
End Function

Public Sub SetLeft(ByVal newLeft As Long)
    ucSupport.RequestNewPosition newLeft, , True
End Sub

Public Function GetTop() As Long
    GetTop = ucSupport.GetControlTop
End Function

Public Sub SetTop(ByVal newTop As Long)
    ucSupport.RequestNewPosition , newTop, True
End Sub

Public Function GetWidth() As Long
    GetWidth = ucSupport.GetControlWidth
End Function

Public Sub SetWidth(ByVal newWidth As Long)
    ucSupport.RequestNewSize newWidth, , True
End Sub

Public Function GetHeight() As Long
    GetHeight = ucSupport.GetControlHeight
End Function

Public Sub SetHeight(ByVal newHeight As Long)
    ucSupport.RequestNewSize , newHeight, True
End Sub

Public Sub SetPositionAndSize(ByVal newLeft As Long, ByVal newTop As Long, ByVal newWidth As Long, ByVal newHeight As Long)
    ucSupport.RequestFullMove newLeft, newTop, newWidth, newHeight, True
End Sub

'Erase the current canvas.  If no images are loaded (which is really the only time this function should be called,
' we'll render the generic "please load an image" icon onto the background.
Public Sub ClearCanvas()
    
    'Request the back buffer DC, and ask the support module to erase any existing rendering for us.
    Dim bufferDC As Long, bWidth As Long, bHeight As Long
    bufferDC = ucSupport.GetBackBufferDC(True, m_Colors.RetrieveColor(PDC_Background))
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
        
    'If no images have been loaded, draw a "load image" placeholder atop the empty background.
    If (g_OpenImageCount = 0) And MainModule.IsProgramRunning() And (Not g_ProgramShuttingDown) Then
        
        Dim placeholderImageSize As Long
        placeholderImageSize = 256
        
        Dim iconLoadAnImage As pdDIB
        LoadResourceToDIB "generic_imageplaceholder", iconLoadAnImage, placeholderImageSize, placeholderImageSize

        Dim notifyFont As pdFont
        Set notifyFont = New pdFont
        notifyFont.SetFontFace g_InterfaceFont

        'Set the font size dynamically.  en-US gets a larger size; other languages, whose text may be longer, use a smaller one.
        If (Not g_Language Is Nothing) Then
            If g_Language.TranslationActive Then notifyFont.SetFontSize 13 Else notifyFont.SetFontSize 14
        Else
            notifyFont.SetFontSize 14
        End If

        notifyFont.SetFontBold False
        notifyFont.SetFontColor RGB(41, 43, 54)
        notifyFont.SetTextAlignment vbCenter

        'Create the font and attach it to our temporary DIB's DC
        notifyFont.CreateFontObject
        notifyFont.AttachToDC bufferDC

        If (Not iconLoadAnImage Is Nothing) Then

            Dim modifiedHeight As Long
            modifiedHeight = bHeight + (iconLoadAnImage.GetDIBHeight / 2) + FixDPI(24)

            Dim loadImageMessage As String
            If Not (g_Language Is Nothing) Then
                loadImageMessage = g_Language.TranslateMessage("Drag an image onto this space to begin editing." & vbCrLf & vbCrLf & "You can also use the Open Image button on the left," & vbCrLf & "or the File > Open and File > Import menus.")
            End If
            notifyFont.DrawCenteredText loadImageMessage, bWidth, modifiedHeight

            'Just above the text instructions, add a generic image icon
            iconLoadAnImage.AlphaBlendToDC bufferDC, 255, (bWidth - iconLoadAnImage.GetDIBWidth) / 2, (modifiedHeight / 2) - (iconLoadAnImage.GetDIBHeight) - FixDPI(20)
            
        End If

        notifyFont.ReleaseFromDC
        Set notifyFont = Nothing
        Set iconLoadAnImage = Nothing
        
    End If
    
    ucSupport.RequestRepaint
    
End Sub

'Is the mouse over the canvas view right now?
Public Function IsMouseOverCanvasView()
    IsMouseOverCanvasView = ucSupport.IsMouseInside()
End Function

'Certain criteria prevent the user from interacting with this canvas object (e.g. no images being loaded).  External functions
' *must* check this value before attempting to render to the canvas.
Public Function IsCanvasInteractionAllowed() As Boolean
    
    'By default, canvas interaction is allowed
    IsCanvasInteractionAllowed = True
    
    'Now, check a bunch of states that might indicate canvas interactions should not be allowed
    
    'If the main form is disabled, exit
    If (Not FormMain.Enabled) Then IsCanvasInteractionAllowed = False
        
    'If user input has been forcibly disabled by some other part of the program, exit
    If g_DisableUserInput Then IsCanvasInteractionAllowed = False
    
    'If no images have been loaded, exit
    If (g_OpenImageCount = 0) Then IsCanvasInteractionAllowed = False
    
    'If our own internal redraw suspension flag is set, exit
    If m_SuspendRedraws Then IsCanvasInteractionAllowed = False
    
    'If any of the previous checks succeeded, exit immediately
    If (Not IsCanvasInteractionAllowed) Then Exit Function
    
    'If there is no active images or valid layers, canvas interactions are also disallowed.  (This is primarily a failsafe check.)
    If (pdImages(g_CurrentImage) Is Nothing) Then
        IsCanvasInteractionAllowed = False
    Else
        If (Not pdImages(g_CurrentImage).IsActive) Then IsCanvasInteractionAllowed = False
        If (pdImages(g_CurrentImage).GetNumOfLayers = 0) Then IsCanvasInteractionAllowed = False
    End If
    
    'If the central processor is active, exit
    If Processor.IsProgramBusy Then IsCanvasInteractionAllowed = False
    
End Function

Public Sub RequestCursor_System(Optional ByVal standardCursorType As SystemCursorConstant = IDC_DEFAULT)
    ucSupport.RequestCursor standardCursorType
End Sub

Public Sub RequestCursor_PNG(ByVal pngResourceName As String, Optional ByVal cursorHotspotX As Long = 0, Optional ByVal cursorHotspotY As Long = 0)
    ucSupport.RequestCursor_PNG pngResourceName, cursorHotspotX, cursorHotspotY
End Sub

'External functions can request an immediate redraw.  Please don't abuse this - it should really only be used when some
' UI element needs to be updated independent of PD's normal refresh cycles.
' TODO: decide if we should expose the "repaint immediately" functionality... I have mixed feelings about this, and if
' we can avoid it, so much the better.
Public Sub RequestRedraw(Optional ByVal repaintImmediately As Boolean = False)
    If (Not g_ProgramShuttingDown) Then ucSupport.RequestRepaint repaintImmediately
End Sub

'For some tool actions, it may be helpful to move the cursor for the user.  Call this function to forcibly set a cursor position.
' (Note that this function will automatically handle the translation to screen coordinates, so passed coordinates should be
'  relative to the canvas itself - e.g. to the control's window.)
Public Sub SetCursorToCanvasPosition(ByVal canvasX As Double, ByVal canvasY As Double)
    ucSupport.RequestMousePosition canvasX, canvasY
End Sub

'Use these functions to forcibly prevent the canvas from redrawing itself.  REDRAWS WILL NOT HAPPEN AGAIN UNTIL YOU RESTORE ACCESS!
Public Function GetRedrawSuspension() As Boolean
    GetRedrawSuspension = m_SuspendRedraws
End Function

Public Sub SetRedrawSuspension(ByVal newRedrawValue As Boolean)
    m_SuspendRedraws = newRedrawValue
End Sub

'Manually request standard-rate or high-rate mouse tracking.  (Drawing tools support high-rate tracking.)
Public Sub SetMouseInput_HighRes(ByVal newState As Boolean)
    ucSupport.RequestHighResMouseInput newState
End Sub

Public Sub SetMouseInput_AutoDrop(ByVal newState As Boolean)
    ucSupport.RequestAutoDropMouseMessages newState
End Sub

Public Function GetNumMouseEventsPending() As Long
    GetNumMouseEventsPending = ucSupport.GetNumMouseEventsPending()
End Function

Public Function GetNextMouseMovePoint(ByVal ptrToDstMMP As Long) As Boolean
    Dim tmpMMP As MOUSEMOVEPOINT
    GetNextMouseMovePoint = ucSupport.GetNextMouseMovePoint(tmpMMP)
    CopyMemory ByVal ptrToDstMMP, ByVal VarPtr(tmpMMP), LenB(tmpMMP)
End Function

Private Sub ucSupport_AppCommand(ByVal cmdID As AppCommandConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    RaiseEvent AppCommand(cmdID, Shift, x, y)
End Sub

Private Sub ucSupport_ClickCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    RaiseEvent ClickCustom(Button, Shift, x, y)
End Sub

Private Sub ucSupport_DoubleClickCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    RaiseEvent DoubleClickCustom(Button, Shift, x, y)
End Sub

Private Sub ucSupport_GotFocusAPI()
    RaiseEvent GotFocusAPI
End Sub

Private Sub ucSupport_KeyDownCustom(ByVal Shift As ShiftConstants, ByVal vkCode As Long, markEventHandled As Boolean)
    RaiseEvent KeyDownCustom(Shift, vkCode, markEventHandled)
End Sub

Private Sub ucSupport_KeyUpCustom(ByVal Shift As ShiftConstants, ByVal vkCode As Long, markEventHandled As Boolean)
    RaiseEvent KeyUpCustom(Shift, vkCode, markEventHandled)
End Sub

Private Sub ucSupport_LostFocusAPI()
    RaiseEvent LostFocusAPI
End Sub

Private Sub ucSupport_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    RaiseEvent MouseDownCustom(Button, Shift, x, y, timeStamp)
End Sub

Private Sub ucSupport_MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    'If no images have been loaded, reset the cursor
    If (g_OpenImageCount = 0) Then ucSupport.RequestCursor IDC_DEFAULT
    
    RaiseEvent MouseEnter(Button, Shift, x, y)
    
End Sub

Private Sub ucSupport_MouseHover(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    RaiseEvent MouseHover(Button, Shift, x, y)
End Sub

Private Sub ucSupport_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    RaiseEvent MouseLeave(Button, Shift, x, y)
End Sub

Private Sub ucSupport_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    RaiseEvent MouseMoveCustom(Button, Shift, x, y, timeStamp)
End Sub

Private Sub ucSupport_MouseUpCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal clickEventAlsoFiring As Boolean, ByVal timeStamp As Long)
    RaiseEvent MouseUpCustom(Button, Shift, x, y, clickEventAlsoFiring, timeStamp)
End Sub

Private Sub ucSupport_MouseWheelHorizontal(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)
    RaiseEvent MouseWheelHorizontal(Button, Shift, x, y, scrollAmount)
End Sub

Private Sub ucSupport_MouseWheelVertical(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)
    RaiseEvent MouseWheelVertical(Button, Shift, x, y, scrollAmount)
End Sub

Private Sub ucSupport_MouseWheelZoom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal zoomAmount As Double)
    RaiseEvent MouseWheelZoom(Button, Shift, x, y, zoomAmount)
End Sub

Private Sub ucSupport_RepaintRequired(ByVal updateLayoutToo As Boolean)
    
    'If no images are loaded, repaint ourselves automatically
    If (g_OpenImageCount = 0) And MainModule.IsProgramRunning() Then
        Me.ClearCanvas
    Else
        Debug.Print "Main viewport requested its own redraw, likely due to a buffer size change."
        ViewportEngine.Stage1_InitializeBuffer pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    End If
    
End Sub

Private Sub UserControl_Initialize()

    'Initialize a master user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd, True
    ucSupport.RequestExtraFunctionality True, True, True
    ucSupport.SpecifyRequiredKeys VK_LEFT, VK_UP, VK_RIGHT, VK_DOWN, VK_DELETE, VK_INSERT, VK_TAB, VK_SPACE, VK_ESCAPE, VK_BACK
    
    'Prep the color manager and load default colors
    Set m_Colors = New pdThemeColors
    Dim colorCount As PDCANVAS_COLOR_LIST: colorCount = [_Count]
    m_Colors.InitializeColorList "PDCanvas", colorCount
    If (Not MainModule.IsProgramRunning()) Then UpdateColorList
    
    'Update the control size parameters at least once
    UpdateControlLayout
    
End Sub

'(This code is copied from FormMain's OLEDragDrop event - please mirror any changes there, or even better, stop being lazy
' and write a universal drag/drop handler!)
Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'Make sure the form is available (e.g. a modal form hasn't stolen focus)
    If (Not g_AllowDragAndDrop) Then Exit Sub
    
    'Use the external function (in the clipboard handler, as the code is roughly identical to clipboard pasting)
    ' to load the OLE source.
    g_Clipboard.LoadImageFromDragDrop Data, Effect, True
    
End Sub

'(This code is copied from FormMain's OLEDragOver event - please mirror any changes there, or even better, stop being lazy
' and write a universal drag/drop handler!)
Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)

    'PD supports a lot of potential drop sources these days.  These values are defined and addressed by the main
    ' clipboard handler, as Drag/Drop and clipboard actions share a ton of similar code.
    If g_Clipboard.IsObjectDragDroppable(Data) And g_AllowDragAndDrop Then
        Effect = vbDropEffectCopy And Effect
    Else
        Effect = vbDropEffectNone
    End If
    
End Sub

'Whenever a control property changes that affects control size or layout (including internal changes, like caption adjustments),
' call this function to recalculate the control's internal layout
Private Sub UpdateControlLayout()
    
    'Retrieve DPI-aware control dimensions from the support class
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
    
End Sub

'Primary rendering function.  Note that ucSupport handles a number of rendering duties (like maintaining a back buffer for us).
Private Sub RedrawBackBuffer()
    
    'Request the back buffer DC, and ask the support module to erase any existing rendering for us.
    Dim bufferDC As Long, bWidth As Long, bHeight As Long
    bufferDC = ucSupport.GetBackBufferDC(False)
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight

    'Paint the final result to the screen, as relevant
    ucSupport.RequestRepaint
    
End Sub

'Before this control does any painting, we need to retrieve relevant colors from PD's primary theming class.  Note that this
' step must also be called if/when PD's visual theme settings change.
Private Sub UpdateColorList()
    m_Colors.LoadThemeColor PDC_Background, "Background", IDE_GRAY
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog.
Public Sub UpdateAgainstCurrentTheme(Optional ByVal hostFormhWnd As Long = 0)
    If ucSupport.ThemeUpdateRequired Then
        UpdateColorList
        UserControl.BackColor = m_Colors.RetrieveColor(PDC_Background, Me.Enabled)
        If (g_OpenImageCount = 0) Then Me.ClearCanvas
        If MainModule.IsProgramRunning() Then NavKey.NotifyControlLoad Me, hostFormhWnd
        If MainModule.IsProgramRunning() Then ucSupport.UpdateAgainstThemeAndLanguage
    End If
End Sub
