VERSION 5.00
Begin VB.Form FormCurves 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Curves"
   ClientHeight    =   8205
   ClientLeft      =   -15
   ClientTop       =   225
   ClientWidth     =   13095
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
   ScaleHeight     =   547
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   873
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   7455
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   1323
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
   Begin VB.ComboBox cboHistogram 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   8040
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   5850
      Width           =   4815
   End
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5160
      Left            =   6000
      ScaleHeight     =   344
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   464
      TabIndex        =   2
      Top             =   120
      Width           =   6960
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.smartCheckBox chkGrid 
      Height          =   480
      Left            =   6240
      TabIndex        =   7
      Top             =   6360
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   847
      Caption         =   "display grid"
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
   Begin PhotoDemon.smartCheckBox chkDiagonal 
      Height          =   480
      Left            =   6240
      TabIndex        =   8
      Top             =   6840
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   847
      Caption         =   "display original curve (diagonal line)"
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
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "histogram overlay:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   1
      Left            =   6240
      TabIndex        =   5
      Top             =   5910
      Width           =   1605
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "additional options:"
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
      Index           =   0
      Left            =   6000
      TabIndex        =   4
      Top             =   5400
      Width           =   1980
   End
   Begin VB.Label lblExplanation 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1440
      Left            =   240
      TabIndex        =   3
      Top             =   5910
      Width           =   5535
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "FormCurves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Image Curves Adjustment Dialog
'Copyright ©2008-2014 by Tanner Helland
'Created: sometime 2008
'Last updated: 03/December/13
'Last update: store curve nodes as relative values rather than absolute ones.  This fixes an extremely rare error when
'              the user has stored curve presets (or last-used settings), changes their monitor DPI, then re-loads PD.
'
'Standard luminosity adjustment via curves.  This dialog is based heavily on similar tools in other photo editors, but
' with a few neat options of its own.  The curve rendering area has received a great deal of attention; small touches
' like full-AA, dynamic node highlighting, and background histogram are nice improvements over other Curves tools.  I
' have also used some trickery with the picture box that handles the curve edit area - note that the edit area sits
' well within the borders of the picture box.  This is necessary so that nodes at the edge of the histogram are not
' cut-off by the picture box boundaries.  Even when highlighted, nodes at the edges are fully rendered.
'
'As the on-dialog instructions state, the LMB can be used to add new nodes or drag existing nodes.  RMB will delete
' nodes.  There is no hard-coded upper limit on nodes, but because each horizontal pixel can only belong to a single
' node, nodes will be automatically removed if the count exceeds the pixel width of the curve box.  (Never gonna happen,
' but I coded against it anyway!)
'
'The function that actually applies the curve to the image is fully ParamString compatible, meaning this function
' works beautifully with the macro tool despite the complex parameters it requires.  I have also heavily optimized the
' render function to make it extremely quick, and it is currently comparable to brightness/contrast adjustment in speed
' (e.g. VERY FAST!).
'
'At present, only the image's luminance curve can be adjusted.  I've debated adding color curves as well, but I have my
' doubts about the utility of this.  It complicates the interface greatly to add those features, and at what benefit to
' the end-user?  If research shows a good technical reason for RGB curves as well, I'll consider adding it later.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Floating-point coordinate type
Private Type fPoint
    pX As Double
    pY As Double
End Type

'This array will store all curve control nodes, including those added by the user at run-time
Private numOfNodes As Long
Private cNodes() As fPoint

'Track mouse status between MouseDown and MouseMove events
Private isMouseDown As Boolean

'Currently selected node in the workspace area
Private selectedNode As Long

'How close to a node the user must click to select that node
Private Const mouseAccuracy As Byte = 6

'Two additional arrays are needed to generate the cubic spline used for the curve function
Private p() As Double
Private u() As Double

'The final curve is used to fill this array, which will contain the actual spline points for each location
' in the spline.  It will be dynamically resized to match the width of the curve preview picture box.
Private cResults() As Double

'It is difficult to see the results of the curve if they lie directly on the preview box border.  To circumvent this
' problem, we render the curve dialog to the center of the picture box, with this value specifying the size of the
' blank border used.
Private Const previewBorder As Long = 10

'These five arrays will hold histogram data for the current image.  They are filled when the form is activated, and
' not modified again unless the form is unloaded and reopened.
Private hData() As Double
Private hDataLog() As Double
Private hMax() As Double
Private hMaxLog() As Double
Private hMaxPosition() As Byte

'An image of the current image histogram is drawn once each for regular and logarithmic, then stored to these DIBs.
Private hDIB As pdDIB, hLogDIB As pdDIB

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

Private Sub cboHistogram_Click()
    updatePreview
End Sub

Private Sub chkDiagonal_Click()
    updatePreview
End Sub

Private Sub chkGrid_Click()
    updatePreview
End Sub

'Apply a curve to an image's luminance values
' Input: a list of 256 values, one for each luminance point in the image
Public Sub ApplyCurveToImage(ByVal listOfPoints As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)

    If Not toPreview Then Message "Applying new luminance curve to image..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    
    prepImageData tmpSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
    
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Color variables
    Dim r As Long, g As Long, b As Long
    
    'Take the list of curve points we were passed (in string format) and convert them into a numeric array.
    Dim cHistogram(0 To 255) As Long
    
    Dim cParams As pdParamString
    Set cParams = New pdParamString
    cParams.setParamString listOfPoints
    
    For x = 0 To 255
        cHistogram(x) = cParams.GetDouble(x + 1) * 255
    Next x
    
    'Our curves correction can be easily applied using a look-up table
    Dim newGamma(0 To 255) As Byte
    Dim tmpGamma As Double
    
    For x = 0 To 255
    
        tmpGamma = CDbl(x) / 255
        
        'This 'if' statement is necessary to match a weird trend with Photoshop's Curves dialog.  For darker gamma
        ' values, Photoshop increases the force of the gamma conversion.  This is good for emphasizing subtle dark
        ' shades that the human eye doesn't normally pick up... I think.  If this 'if' statement is removed and
        ' only the TRUE condition is kept, the function will yield more mathematically correct results.
        If cHistogram(x) <= (256 - x) Then
            tmpGamma = tmpGamma ^ (1 / ((256 - x) / (cHistogram(x) + 1)))
        Else
            tmpGamma = tmpGamma ^ ((1 / ((256 - x) / (cHistogram(x) + 1))) ^ 1.5)
        End If
        
        tmpGamma = tmpGamma * 255
        
        If tmpGamma > 255 Then tmpGamma = 255
        If tmpGamma < 0 Then tmpGamma = 0
        
        newGamma(x) = CByte(tmpGamma)
        
    Next x
        
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
                
        'Assign the new values to each color channel
        ImageData(QuickVal + 2, y) = newGamma(r)
        ImageData(QuickVal + 1, y) = newGamma(g)
        ImageData(QuickVal, y) = newGamma(b)
        
    Next y
        If Not toPreview Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x
            End If
        End If
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
        
End Sub

'Nodes from the Curves dialog must be manually added to the preset file when requested.  This event will be raised
' whenever the command bar needs custom data from us.
Private Sub cmdBar_AddCustomPresetData()
    
    'Write the number of nodes to file
    cmdBar.addPresetData "NodeCount", CStr(numOfNodes)
    
    'Next, place all node data in one giant string.
    ' UPDATE 03 Dec 2014: instead of storing absolute coordinates, store relative ones per the size of the
    '                     curve box.  This fixes an extremely rare error when the user changes DPI for their
    '                     monitor while having a previously stored set of curve coordinates.
    Dim nodeString As String
    nodeString = ""
    
    Dim nodeBoxWidth As Long, nodeBoxHeight As Long
    nodeBoxWidth = picDraw.ScaleWidth
    nodeBoxHeight = picDraw.ScaleHeight
    
    Dim i As Long
    For i = 1 To numOfNodes
        nodeString = nodeString & CStr(cNodes(i).pX / nodeBoxWidth) & "," & CStr(cNodes(i).pY / nodeBoxHeight)
        If i < numOfNodes Then nodeString = nodeString & "|"
    Next i
    
    cmdBar.addPresetData "NodeData", nodeString
    
End Sub

'Randomizing the curves array is a bit more complicated than normal tools, because we have to randomize it ourselves.
Private Sub cmdBar_RandomizeClick()

    Randomize Timer

    'Initialize the control to somewhere between 3 and 6 points
    numOfNodes = Int(Rnd * 4) + 3
    ReDim cNodes(0 To numOfNodes) As fPoint
    
    'Start by equally spacing the nodes
    Dim i As Long
    For i = 0 To numOfNodes
        cNodes(i).pX = (i - 1) * ((picDraw.ScaleWidth - previewBorder * 2) / (numOfNodes - 1))
        cNodes(i).pY = (picDraw.ScaleHeight - previewBorder * 2) - (cNodes(i).pX / (picDraw.ScaleWidth - previewBorder * 2)) * (picDraw.ScaleHeight - previewBorder * 2)
        cNodes(i).pX = cNodes(i).pX + previewBorder
        cNodes(i).pY = cNodes(i).pY + previewBorder
    Next i
    
    'Finally, move all nodes a random amount up or down, left or right
    For i = 0 To numOfNodes
        
        cNodes(i).pX = cNodes(i).pX + Int(-20 + Rnd * 41)
        If cNodes(i).pX < previewBorder Then cNodes(i).pX = previewBorder
        If cNodes(i).pX > (picDraw.ScaleWidth - previewBorder) Then cNodes(i).pX = (picDraw.ScaleWidth - previewBorder)
        
        cNodes(i).pY = cNodes(i).pY + Int(-40 + Rnd * 81)
        If cNodes(i).pY < previewBorder Then cNodes(i).pY = previewBorder
        If cNodes(i).pY > (picDraw.ScaleHeight - previewBorder) Then cNodes(i).pY = (picDraw.ScaleHeight - previewBorder)
        
    Next i
    
End Sub

'When a preset is loaded from file, we need to retrieve the custom curve information alongside it
Private Sub cmdBar_ReadCustomPresetData()
    
    'Retrieve the number of nodes in this preset
    Dim tmpString As String
    tmpString = cmdBar.retrievePresetData("NodeCount")
    numOfNodes = CLng(tmpString)
    
    'Using that as our guide, repopulate the cNodes array
    ReDim cNodes(0 To numOfNodes) As fPoint
    
    'Retrieve the string that contains the node coordinates
    tmpString = cmdBar.retrievePresetData("NodeData")
    
    'With the help of a paramString class, parse out individual coordinates into the cNodes array
    Dim cParams As pdParamString
    Set cParams = New pdParamString
    cParams.setParamString Replace(tmpString, ",", "|")
    
    'UPDATE 03 Dec 2014: instead of storing absolute coordinates, we now store relative ones per the size of
    '                    the curve box.  This fixes an extremely rare error when the user changes DPI for
    '                    their monitor while having a previously stored set of curve coordinates.
    Dim nodeBoxWidth As Long, nodeBoxHeight As Long
    nodeBoxWidth = picDraw.ScaleWidth
    nodeBoxHeight = picDraw.ScaleHeight
    
    Dim i As Long
    For i = 1 To numOfNodes
        
        'Retrieve this node's x and y values
        cNodes(i).pX = cParams.GetDouble((i - 1) * 2 + 1)
        cNodes(i).pY = cParams.GetDouble((i - 1) * 2 + 2)
        
        'Old preset values may store the node values as absolutes rather than relatives.  Check for this, and
        ' adjust node values accordingly.
        If cNodes(i).pX > 1 Then
        
            If cNodes(i).pX > nodeBoxWidth Then cNodes(i).pX = nodeBoxWidth
            If cNodes(i).pY > nodeBoxHeight Then cNodes(i).pY = nodeBoxHeight
        
        Else
        
            cNodes(i).pX = cNodes(i).pX * nodeBoxWidth
            cNodes(i).pY = cNodes(i).pY * nodeBoxHeight
        
        End If
                
    Next i
    
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_OKClick()
    Process "Curves", , getCurvesParamString()
End Sub

'Reset the curve to three points in a straight line
Private Sub cmdBar_ResetClick()

    resetCurvePoints
    
    'Also, reset will automatically select the first entry in a combo box.  In this case, we actually want the 1st one.
    cboHistogram.ListIndex = 1
    
End Sub

Private Sub Form_Activate()
    
    'Populate the explanation label
    Dim addInstructions As String
    addInstructions = ""
    addInstructions = g_Language.TranslateMessage("instructions:")
    addInstructions = addInstructions & vbCrLf
    addInstructions = addInstructions & "  + " & g_Language.TranslateMessage("left-click to add new nodes or drag existing nodes")
    addInstructions = addInstructions & vbCrLf
    addInstructions = addInstructions & "  + " & g_Language.TranslateMessage("right-click to remove nodes")
    
    lblExplanation.Caption = addInstructions
        
    'If translations are active, the translated text may not fit the explanation label.  Automatically adjust it to fit.
    fitWordwrapLabel lblExplanation, Me
    
    'Mark the mouse as not being down
    isMouseDown = False
    
    'Fill the histogram arrays
    fillHistogramArrays hData, hDataLog, hMax, hMaxLog, hMaxPosition
    
    'Initialize the background histogram image DIBs
    Set hDIB = New pdDIB
    Set hLogDIB = New pdDIB
    hDIB.createBlank picDraw.ScaleWidth - (previewBorder * 2) - 1, picDraw.ScaleHeight - (previewBorder * 2) - 1
    hLogDIB.createFromExistingDIB hDIB
    
    'Build a look-up table of x-positions for the histogram data
    Dim hLookupX(0 To 255) As Double
    Dim i As Long
    
    For i = 0 To 255
        hLookupX(i) = (CDbl(i) / 255) * hDIB.getDIBWidth
    Next i
    
    'Render the luminance histogram data to each DIB (one for regular, one for logarithmic)
    Dim yMax As Double
    yMax = 0.9 * hDIB.getDIBHeight
    
    For i = 1 To 255
        GDIPlusDrawLineToDC hDIB.getDIBDC, hLookupX(i - 1), hDIB.getDIBHeight - (hData(3, i - 1) / hMax(3)) * yMax, hLookupX(i), hDIB.getDIBHeight - (hData(3, i) / hMax(3)) * yMax, RGB(192, 192, 192), 255
        GDIPlusDrawLineToDC hLogDIB.getDIBDC, hLookupX(i - 1), hDIB.getDIBHeight - (hDataLog(3, i - 1) / hMaxLog(3)) * yMax, hLookupX(i), hDIB.getDIBHeight - (hDataLog(3, i) / hMaxLog(3)) * yMax, RGB(192, 192, 192), 255
    Next i
    
    'Beneath each line, add an even lighter "filled" version of the line
    For i = 0 To 255
        GDIPlusDrawLineToDC hDIB.getDIBDC, hLookupX(i), hDIB.getDIBHeight - (hData(3, i) / hMax(3)) * yMax - 1, hLookupX(i), hDIB.getDIBHeight, RGB(192, 192, 192), 128
        GDIPlusDrawLineToDC hLogDIB.getDIBDC, hLookupX(i), hDIB.getDIBHeight - (hDataLog(3, i) / hMaxLog(3)) * yMax - 1, hLookupX(i), hDIB.getDIBHeight, RGB(192, 192, 192), 128
    Next i
        
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
    cmdBar.markPreviewStatus True
    updatePreview
    
End Sub

Private Sub Form_Load()

    cmdBar.markPreviewStatus False
    
    'Populate the histogram display drop-down
    cboHistogram.Clear
    cboHistogram.AddItem " none", 0
    cboHistogram.AddItem " standard", 1
    cboHistogram.AddItem " logarithmic", 2
    cboHistogram.ListIndex = 1
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Redraw the on-screen preview of the transformed image
Private Sub updatePreview()
    
    If cmdBar.previewsAllowed Then
    
        'Start by generating a list of points that correspond to the cubic spline used for the curve
        fillResultsArray
        
        'Redraw the preview box
        redrawPreviewBox
        
        'Redraw the image effect preview
        ApplyCurveToImage getCurvesParamString(), True, fxPreview
        
    End If
    
End Sub

'Assuming that cResults has been filled by calling fillResultsArray, this function will convert the curve into
' a list of histogram points, in PD string parameter format.
Private Function getCurvesParamString() As String
    
    Dim paramString As String
    paramString = ""

    Dim i As Long
    
    'The histogram array will be filled with a list of values in the range [0.0, 1.0]
    Dim cHistogram(0 To 255) As Double
    Dim cEntry As Long
    For i = 0 To 255
        cEntry = previewBorder + (CDbl(i) / 255) * (picDraw.ScaleWidth - previewBorder * 2)
        cHistogram(i) = (cResults(cEntry) - previewBorder) / (picDraw.ScaleHeight - previewBorder * 2)
    Next i

    'We now need to convert the histogram array into a "|"-delimited string that can be passed through the
    ' software processor.  Generate it automatically.
    For i = 0 To 255
        paramString = paramString & CStr(cHistogram(i))
        If i < 255 Then paramString = paramString & "|"
    Next i
    
    getCurvesParamString = paramString
    
End Function

Private Sub redrawPreviewBox()

    If Not cmdBar.previewsAllowed Then Exit Sub

    picDraw.Picture = LoadPicture("")
    
    'Start by copying the proper histogram image into the picture box
    Select Case cboHistogram.ListIndex
    
        'No histogram
        Case 0
        
        'Normal histogram
        Case 1
            BitBlt picDraw.hDC, previewBorder + 1, previewBorder + 1, hDIB.getDIBWidth, hDIB.getDIBHeight, hDIB.getDIBDC, 0, 0, vbSrcCopy
        
        'Logarithmic histogram
        Case 2
            BitBlt picDraw.hDC, previewBorder + 1, previewBorder + 1, hDIB.getDIBWidth, hDIB.getDIBHeight, hLogDIB.getDIBDC, 0, 0, vbSrcCopy
        
    End Select
    
    'Next, draw a grid that separates the image into 16 segments; this helps orient the user, and it also provides a
    ' border for the drawing area (important since that area sits well within the picture box itself).
    picDraw.DrawWidth = 1
    picDraw.ForeColor = RGB(172, 172, 172)
    
    Dim i As Long
    Dim loopUpperLimit As Long
    
    If CBool(chkGrid) Then loopUpperLimit = 4 Else loopUpperLimit = 1
    
    For i = 0 To loopUpperLimit
        picDraw.Line (previewBorder + (i / loopUpperLimit) * (picDraw.ScaleWidth - previewBorder * 2), previewBorder)-(previewBorder + (i / loopUpperLimit) * (picDraw.ScaleWidth - previewBorder * 2), picDraw.ScaleHeight - previewBorder)
        picDraw.Line (previewBorder, previewBorder + (i / loopUpperLimit) * (picDraw.ScaleHeight - previewBorder * 2))-(picDraw.ScaleWidth - previewBorder, previewBorder + (i / loopUpperLimit) * (picDraw.ScaleHeight - previewBorder * 2))
    Next i
    
    'Next, draw a diagonal per the user's request
    If CBool(chkDiagonal) Then
        GDIPlusDrawLineToDC picDraw.hDC, previewBorder, picDraw.ScaleHeight - previewBorder, picDraw.ScaleWidth - previewBorder, previewBorder, RGB(127, 127, 127), 127
    End If
    
    'Use the previously created spline array (cResults) to draw the cubic spline onto picDraw, while using GDI+ for antialiasing
    For i = previewBorder + 1 To picDraw.ScaleWidth - previewBorder
        GDIPlusDrawLineToDC picDraw.hDC, i, cResults(i), i - 1, cResults(i - 1), RGB(0, 0, 255), 192, 2
    Next i
    
    'Next, render the spline control points.
    Dim circRadius As Long
    circRadius = 8
    
    Dim circAlpha As Long
    circAlpha = 190
    
    'The curves function requires an input of 256 points - one for each level of the histogram.
    'NOTE: this function requires fillResultsArray() to have been called immediately prior.  Otherwise, the
    '       cResults array will not contain the entries necessary to generate a parameter list.
    For i = 1 To numOfNodes
        GDIPlusDrawEllipseToDC picDraw.hDC, cNodes(i).pX - (circRadius / 2), cNodes(i).pY - (circRadius / 2), circRadius, circRadius, RGB(32, 32, 255), True
    Next i
    
    'Render a special highlight around the currently selected node
    If selectedNode > 0 Then
        GDIPlusDrawCanvasCircle picDraw.hDC, cNodes(selectedNode).pX, cNodes(selectedNode).pY, circRadius, circAlpha
    End If
    
    'Lock the image, force a refresh, and our work here is done
    picDraw.Picture = picDraw.Image
    picDraw.Refresh
    
End Sub

Private Sub picDraw_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'If the mouse is over a point, mark it as the active point
    selectedNode = checkClick(x, y)
    
    'Different actions are initiated for left vs right clicks (left to add/move points, right to remove)
    If Button = vbLeftButton Then
    
        isMouseDown = True
        
        'If this click was not over an existing point, add a new one to the point list!
        If selectedNode = -1 Then
        
            'Find the appropriate location in the array for this point.
            Dim i As Long
            
            Dim pointFound As Long
            pointFound = -1
            
            For i = 0 To numOfNodes
                If cNodes(i).pX > x Then
                    pointFound = i
                    Exit For
                End If
            Next i
        
            numOfNodes = numOfNodes + 1
            ReDim Preserve cNodes(0 To numOfNodes) As fPoint
            
            'If a neighboring point was found, use that location to insert the new point
            If pointFound > -1 Then
                
                'Shift all points "above" this one to the right
                For i = numOfNodes To pointFound + 1 Step -1
                    cNodes(i).pX = cNodes(i - 1).pX
                    cNodes(i).pY = cNodes(i - 1).pY
                Next i
                
                'Store the new point
                cNodes(pointFound).pX = x
                cNodes(pointFound).pY = y
                
                'Make sure the new point falls within acceptable boundaries
                If cNodes(pointFound).pX < previewBorder Then cNodes(pointFound).pX = previewBorder
                If cNodes(pointFound).pX > picDraw.ScaleWidth - previewBorder Then cNodes(pointFound).pX = picDraw.ScaleWidth - previewBorder
                If cNodes(pointFound).pY < previewBorder Then cNodes(pointFound).pY = previewBorder
                If cNodes(pointFound).pY > picDraw.ScaleHeight - previewBorder Then cNodes(pointFound).pY = picDraw.ScaleHeight - previewBorder
                
                'Perform a fail-safe check of the array to make sure there are no duplicate x-values
                For i = numOfNodes To 1 Step -1
                    If cNodes(i).pX = cNodes(i - 1).pX Then cNodes(i - 1).pX = cNodes(i - 1).pX - 1
                Next i
                
                'And finally, perform an additional fail-safe to remove any x-values that now occur outside acceptable boundaries
                ' (e.g. points pushed off the left of the curve)
                For i = numOfNodes To 1 Step -1
                    If cNodes(i).pX < previewBorder Then deleteCurveNode i
                Next i
                
                'Mark this node as the currently selected one
                selectedNode = pointFound
            
            'If no neighboring point was found, this point should be inserted at the end of the curve
            Else
                cNodes(numOfNodes).pX = x
                cNodes(numOfNodes).pY = y
                selectedNode = numOfNodes
            End If
            
            'Request a full redraw of the curve
            updatePreview
        
        End If
        
    'On right-clicks, remove the selected point
    ElseIf Button = vbRightButton Then
    
        'Only erase a point if one was actually clicked; then request a redraw
        If selectedNode > -1 Then
            deleteCurveNode selectedNode
            updatePreview
        End If
        
    End If
    
End Sub

'Delete the specified node from the curve.  This function assumes that the passed nodeIndex is a valid entry.
Private Sub deleteCurveNode(ByVal nodeIndex As Long)

    'Only erase a node if more than two nodes will be left after the operation
    If numOfNodes > 2 Then
    
        'Start by shifting all nodes "above" the current one to the left
        Dim i As Long
        For i = nodeIndex To numOfNodes - 1
            cNodes(i).pX = cNodes(i + 1).pX
            cNodes(i).pY = cNodes(i + 1).pY
        Next i
        
        'Reduce the point count and resize the main point array
        numOfNodes = numOfNodes - 1
        ReDim Preserve cNodes(0 To numOfNodes) As fPoint
    
        selectedNode = -1
    
    End If

End Sub

Private Sub picDraw_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    'If the mouse is *not* down, indicate to the user that points can be moved
    If Not isMouseDown Then
        
        'If the user is close to a knot, change the mousepointer to 'move'
        If checkClick(x, y) > -1 Then
            If picDraw.MousePointer <> 5 Then picDraw.MousePointer = 5
            selectedNode = checkClick(x, y)
        Else
            If picDraw.MousePointer <> 0 Then picDraw.MousePointer = 0
            selectedNode = -1
        End If
        
        'Redraw just the preview box, with the selected node highlighted
        fillResultsArray
        redrawPreviewBox
            
    'If the mouse *is* down, move the current point and redraw the preview
    Else
    
        If selectedNode > 0 Then
        
            cNodes(selectedNode).pX = x
            cNodes(selectedNode).pY = y
            
            'Perform basic bounds-checking.  Points are not allowed to cross over each other, and they cannot lie
            ' outside the bounds of the curve preview box.
            If selectedNode < numOfNodes Then
                If cNodes(selectedNode).pX >= cNodes(selectedNode + 1).pX Then cNodes(selectedNode).pX = cNodes(selectedNode + 1).pX - 1
            End If
            
            'Because legitimate points start at index position 1, we don't need to worry about "if selectedNode > 0"
            ' as that statement is already handled at the top of this segment.
            If cNodes(selectedNode).pX <= cNodes(selectedNode - 1).pX Then
                cNodes(selectedNode).pX = cNodes(selectedNode - 1).pX + 1
            End If
            
            If cNodes(selectedNode).pX < previewBorder Then cNodes(selectedNode).pX = previewBorder
            If cNodes(selectedNode).pX > picDraw.ScaleWidth - previewBorder Then cNodes(selectedNode).pX = picDraw.ScaleWidth - previewBorder
            If cNodes(selectedNode).pY < previewBorder Then cNodes(selectedNode).pY = previewBorder
            If cNodes(selectedNode).pY > picDraw.ScaleHeight - previewBorder Then cNodes(selectedNode).pY = picDraw.ScaleHeight - previewBorder
            
            updatePreview
            
        Else
            fillResultsArray
            redrawPreviewBox
        End If
    
    End If

End Sub

Private Sub picDraw_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    isMouseDown = False
    selectedNode = -1
End Sub

'Simple distance routine to see if a location on the picture box is near an existing point
Private Function checkClick(ByVal x As Long, ByVal y As Long) As Long
    
    Dim pDist As Double
    Dim i As Long
    
    For i = 1 To numOfNodes
        pDist = pDistance(x, y, cNodes(i).pX, cNodes(i).pY)
        
        'If we're close to an existing point, return the index of that point
        If pDist < mouseAccuracy Then
            checkClick = i
            Exit Function
        End If
        
    Next i
    
    'Returning -1 says we're not close to an existing point
    checkClick = -1
    
End Function

'Simple distance formula here - we use this to calculate if the user has clicked on (or near) a point
Private Function pDistance(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Double
    pDistance = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
End Function

'Original required spline function:
Private Function getCurvePoint(ByVal i As Long, ByVal v As Double) As Double
    Dim t As Double
    'derived curve equation (which uses p's and u's for coefficients)
    t = (v - cNodes(i).pX) / u(i)
    getCurvePoint = t * cNodes(i + 1).pY + (1 - t) * cNodes(i).pY + u(i) * u(i) * (f(t) * p(i + 1) + f(1 - t) * p(i)) / 6#
End Function

'Original required spline function:
Private Function f(ByRef x As Double) As Double
        f = x * x * x - x
End Function

'Original required spline function:
Private Sub SetPandU()
    
    Dim i As Long
    Dim d() As Double
    Dim w() As Double
    ReDim d(0 To numOfNodes) As Double
    ReDim w(0 To numOfNodes) As Double
    
    'Routine to compute the parameters of our cubic spline.  Based on equations derived from some basic facts...
    'Each segment must be a cubic polynomial.  Curve segments must have equal first and second derivatives
    'at knots they share.  General algorithm taken from a book which has long since been lost.
    
    'The math that derived this stuff is pretty messy...  expressions are isolated and put into
    'arrays.  we're essentially trying to find the values of the second derivative of each polynomial
    'at each knot within the curve.  That's why theres only N-2 p's (where N is # points).
    'later, we use the p's and u's to calculate curve points...

    For i = 2 To numOfNodes - 1
        d(i) = 2 * (cNodes(i + 1).pX - cNodes(i - 1).pX)
    Next
    For i = 1 To numOfNodes - 1
        u(i) = cNodes(i + 1).pX - cNodes(i).pX
    Next
    For i = 2 To numOfNodes - 1
        w(i) = 6# * ((cNodes(i + 1).pY - cNodes(i).pY) / u(i) - (cNodes(i).pY - cNodes(i - 1).pY) / u(i - 1))
    Next
    For i = 2 To numOfNodes - 2
        w(i + 1) = w(i + 1) - w(i) * u(i) / d(i)
        d(i + 1) = d(i + 1) - u(i) * u(i) / d(i)
    Next
    p(1) = 0#
    For i = numOfNodes - 1 To 2 Step -1
        p(i) = (w(i) - u(i) * p(i + 1)) / d(i)
    Next
    p(numOfNodes) = 0#
    
End Sub

'By default, three points are provided: one at each corner, and one in the middle of the diagonal
Private Sub resetCurvePoints()

    numOfNodes = 3
    ReDim cNodes(0 To numOfNodes) As fPoint
    
    Dim i As Long
    For i = 0 To numOfNodes
        cNodes(i).pX = (i - 1) * ((picDraw.ScaleWidth - previewBorder * 2) / (numOfNodes - 1))
        cNodes(i).pY = (picDraw.ScaleHeight - previewBorder * 2) - (cNodes(i).pX / (picDraw.ScaleWidth - previewBorder * 2)) * (picDraw.ScaleHeight - previewBorder * 2)
        cNodes(i).pX = cNodes(i).pX + previewBorder
        cNodes(i).pY = cNodes(i).pY + previewBorder
    Next i

End Sub

'Generates a spline from the current set of control points, and fills the results array with the relevant values
Private Sub fillResultsArray()

    ReDim p(0 To numOfNodes) As Double
    ReDim u(0 To numOfNodes) As Double

    'Clear the results array and reset the max/min variables
    ReDim cResults(-1 To picDraw.ScaleWidth) As Double
    
    Dim i As Long
    For i = -1 To picDraw.ScaleWidth
        cResults(i) = -1
    Next i
    
    Dim minX As Double, maxX As Double
    minX = picDraw.ScaleWidth
    maxX = -1
    
    'Now run a loop through the knots, calculating spline values as we go
    SetPandU
    Dim xPos As Long, yPos As Single
    
    For i = 1 To numOfNodes - 1
        For xPos = cNodes(i).pX To cNodes(i + 1).pX
            yPos = getCurvePoint(i, xPos)
            If xPos < minX Then minX = xPos
            If xPos > maxX Then maxX = xPos
            If yPos > picDraw.ScaleHeight - previewBorder Then yPos = picDraw.ScaleHeight - previewBorder
            If yPos < previewBorder Then yPos = previewBorder
            cResults(xPos) = yPos
        Next xPos
    Next i

    'cResults() now contains the y-coordinate of the spline for every x-coordinate in picDraw that falls between the
    ' initial point and the final point.  Points outside this range are treated as flat lines with values matching
    ' the nearest end point, and we fill those values now.
    For i = previewBorder - 1 To minX - 1
        cResults(i) = cResults(minX)
    Next i
    For i = picDraw.ScaleWidth - previewBorder To maxX + 1 Step -1
        cResults(i) = cResults(maxX)
    Next i
    
    'cResults is now complete.  Its primary dimension is the width of the picture box, and each entry in the array
    ' contains the y-value of the spline at that x-position.  This can be used to easily render the spline on-screen,
    ' and also to apply the curve to the image.

End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub


