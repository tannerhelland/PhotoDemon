VERSION 5.00
Begin VB.UserControl smartResize 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   1755
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6810
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   117
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   454
   ToolboxBitmap   =   "smartResize.ctx":0000
   Begin PhotoDemon.smartCheckBox chkRatio 
      Height          =   480
      Left            =   4080
      TabIndex        =   0
      Top             =   255
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   847
      Caption         =   "lock aspect ratio"
      Value           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.textUpDown tudWidth 
      Height          =   405
      Left            =   840
      TabIndex        =   1
      Top             =   0
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   714
      Min             =   1
      Max             =   32767
      Value           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.textUpDown tudHeight 
      Height          =   405
      Left            =   840
      TabIndex        =   2
      Top             =   630
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   714
      Min             =   1
      Max             =   32767
      Value           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblWidth 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "width:"
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
      Left            =   0
      TabIndex        =   7
      Top             =   30
      Width           =   675
   End
   Begin VB.Label lblHeight 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "height:"
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
      Left            =   0
      TabIndex        =   6
      Top             =   660
      Width           =   750
   End
   Begin VB.Label lblWidthUnit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "pixels"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   480
      Left            =   2130
      TabIndex        =   5
      Top             =   30
      Width           =   855
   End
   Begin VB.Label lblHeightUnit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "pixels"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   2130
      TabIndex        =   4
      Top             =   660
      Width           =   855
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   192
      X2              =   257
      Y1              =   58
      Y2              =   58
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   192
      X2              =   257
      Y1              =   11
      Y2              =   11
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   256
      X2              =   256
      Y1              =   12
      Y2              =   58
   End
   Begin VB.Label lblAspectRatio 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "new aspect ratio will be"
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
      Left            =   0
      TabIndex        =   3
      Top             =   1245
      Width           =   2490
   End
End
Attribute VB_Name = "smartResize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'Image Resize User Control
'Copyright ©2001-2014 by Tanner Helland
'Created: 6/12/01 (original resize dialog), 24/Jan/14 (conversion to user control)
'Last updated: 24/Jan/14
'Last update: initial conversion of resize UI to dedicated user control
'
'Many tools in PD relate to resizing: image size, canvas size, (soon) layer size, content-aware rescaling,
' perhaps a more advanced autocrop tool, plus dedicated resize options in the batch converter...
'
'Rather than develop custom resize UIs for all these scenarios, I finally asked myself: why not use a single
' resize-centric user control? As an added bonus, that would allow me to update the user control to extend new
' capabilities to all of PD's resize tools.
'
'Thus this UC was born.  All resize-related dialogs in the project now use it, and any added features can now
' automatically propagate across them.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'This object provides a single raised event:
' - Change (which triggers when a size value is updated)
Public Event Change(newWidth As Long, newHeight As Long)

Private WithEvents mFont As StdFont
Attribute mFont.VB_VarHelpID = -1

'Store a copy of the original width/height values we are passed
Private initWidth As Long, initHeight As Long

'Used for maintaining ratios when the check box is clicked
Private wRatio As Double, hRatio As Double

'Used to prevent infinite recursion as updates to one text box force updates to the other
Private allowedToUpdateWidth As Boolean, allowedToUpdateHeight As Boolean

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'Font handling is a bit specialized for user controls; see http://msdn.microsoft.com/en-us/library/aa261313%28v=vs.60%29.aspx
Public Property Get Font() As StdFont
    Set Font = mFont
End Property

Public Property Set Font(mNewFont As StdFont)
    With mFont
        .Bold = mNewFont.Bold
        .Italic = mNewFont.Italic
        .Name = mNewFont.Name
        .Size = mNewFont.Size
    End With
    PropertyChanged "Font"
End Property

'When the control's font is changed, this sub will be fired; make sure all child controls have their fonts changed here.
Private Sub mFont_FontChanged(ByVal PropertyName As String)
    Set UserControl.Font = mFont
    Set lblWidth.Font = UserControl.Font
    Set lblHeight.Font = UserControl.Font
    Set lblAspectRatio.Font = UserControl.Font
    Set tudWidth.Font = UserControl.Font
    Set tudHeight.Font = UserControl.Font
    Set lblWidthUnit.Font = UserControl.Font
    Set lblHeightUnit.Font = UserControl.Font
    Set chkRatio.Font = UserControl.Font
End Sub

'Width and height can be retrieved from these properties
Public Property Get imgWidth() As Long
    imgWidth = tudWidth
End Property

Public Property Let imgWidth(newWidth As Long)
    tudWidth = newWidth
    syncDimensions True
End Property

Public Property Get imgHeight() As Long
    imgHeight = tudHeight
End Property

Public Property Let imgHeight(newHeight As Long)
    tudHeight = newHeight
    syncDimensions False
End Property

'Before using this control, dialogs MUST call this function to notify the control of the initial width/height values
' they want to use.  We cannot do this automatically as some dialogs determine this by the current image's dimensions
' (e.g. resize) while others may do it when no images are loaded (e.g. batch process).
Public Sub setInitialDimensions(ByVal srcWidth As Long, ByVal srcHeight As Long)

    'Store local copies
    initWidth = srcWidth
    initHeight = srcHeight
    
    'To prevent aspect ratio changes to one box resulting in recursion-type changes to the other, we only
    ' allow one box at a time to be updated.
    allowedToUpdateWidth = True
    allowedToUpdateHeight = True
    
    'Establish aspect ratios
    wRatio = initWidth / initHeight
    hRatio = initHeight / initWidth
    
    'Display the initial width/height
    tudWidth = srcWidth
    tudHeight = srcHeight

End Sub

Private Sub UserControl_Initialize()

    'When compiled, manifest-themed controls need to be further subclassed so they can have transparent backgrounds.
    If g_IsProgramCompiled And g_IsThemingEnabled And g_IsVistaOrLater Then
        g_Themer.requestContainerSubclass UserControl.hWnd
    End If
    
    allowedToUpdateWidth = True
    allowedToUpdateHeight = True
    
    'Prepare a font object for use
    Set mFont = New StdFont
    Set UserControl.Font = mFont

End Sub

Private Sub UserControl_InitProperties()

    Set mFont = UserControl.Font
    mFont.Name = "Tahoma"
    mFont.Size = 10
    mFont_FontChanged ("")

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        Set Font = .ReadProperty("Font", Ambient.Font)
    End With
    
End Sub

Private Sub UserControl_Show()

    'Translate various bits of UI text at run-time
    If g_UserModeFix Then
        
        lblWidth.Caption = g_Language.TranslateMessage("width") & ": "
        lblHeight.Caption = g_Language.TranslateMessage("height") & ": "
        lblWidth.Refresh
        lblHeight.Refresh
        
        'Because the width and height labels are autosized, they will now have grown or shrunk per their translation.
        ' Re-align all subsequent controls to match this new layout.
        
        'Start by finding the longer caption of the two; this will serve as our alignment baseline
        Dim hOffset As Long
        hOffset = lblWidth.Width
        
        If lblHeight.Width > hOffset Then hOffset = lblHeight.Width
        
        'Add in a bit of padding
        hOffset = hOffset + lblWidth.Left + fixDPI(8)
        
        'Move the pixel entry boxes over
        tudWidth.Left = hOffset
        tudHeight.Left = hOffset
        
        'Now repeat the steps above for all other controls, e.g. cascade all subsequent controls to the right,
        ' using the left-most control nearest them as a guide.
        hOffset = hOffset + tudWidth.Width + fixDPI(8)
        
        lblWidthUnit.Left = hOffset
        lblHeightUnit.Left = hOffset
        
    End If

End Sub

'If "Lock Image Aspect Ratio" is selected, these two routines keep all values in sync
Private Sub tudHeight_Change()
    syncDimensions False
End Sub

Private Sub tudWidth_Change()
    syncDimensions True
End Sub

'If the preserve aspect ratio button is pressed, update the height box to reflect the image's current aspect ratio
Private Sub ChkRatio_Click()
    syncDimensions True
End Sub

Private Sub UserControl_Terminate()
    
    'When the control is terminated, release the subclassing used for transparent backgrounds
    If g_IsProgramCompiled And g_IsThemingEnabled And g_IsVistaOrLater Then g_Themer.releaseContainerSubclass UserControl.hWnd
    
End Sub

'When one dimension is updated, call this to synchronize the other (as necessary) and/or the aspect ratio
Private Sub syncDimensions(ByVal useWidthAsSource As Boolean)

    If useWidthAsSource Then
    
        'When changing width, do not also update height unless "preserve aspect ratio" is checked
        If CBool(chkRatio) And allowedToUpdateHeight Then
            allowedToUpdateWidth = False
            tudHeight = Int((tudWidth * hRatio) + 0.5)
            allowedToUpdateWidth = True
        End If
    
    Else
    
        'When changing height, do not also update width unless "preserve aspect ratio" is checked
        If CBool(chkRatio) And allowedToUpdateWidth Then
            allowedToUpdateHeight = False
            tudWidth = Int((tudHeight * wRatio) + 0.5)
            allowedToUpdateHeight = True
        End If
    
    End If
    
    updateAspectRatio
    
    RaiseEvent Change(tudWidth, tudHeight)

End Sub

'This control displays an approximate aspect ratio for the selected dimensions.  This can be helpful when
' trying to select new width/height values for a specific application with a set aspect ratio (e.g. 16:9 screens).
Private Sub updateAspectRatio()

    Dim wholeNumber As Double, Numerator As Double, Denominator As Double
    
    If tudWidth.IsValid And tudHeight.IsValid Then
        convertToFraction tudWidth / tudHeight, wholeNumber, Numerator, Denominator, 4, 99.9
        
        'Aspect ratios are typically given in terms of base 10 if possible, so change values like 8:5 to 16:10
        If CLng(Denominator) = 5 Then
            Numerator = Numerator * 2
            Denominator = Denominator * 2
        End If
        
        lblAspectRatio.Caption = g_Language.TranslateMessage("new aspect ratio will be %1:%2", Numerator, Denominator)
    End If

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    'Store all associated properties
    With PropBag
        .WriteProperty "Font", mFont, "Tahoma"
    End With

End Sub
