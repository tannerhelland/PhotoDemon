VERSION 5.00
Begin VB.Form toolpanel_Measure 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   1515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16650
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
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
   ScaleHeight     =   101
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1110
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.pdButton cmdAction 
      Height          =   480
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   873
      Caption         =   "swap points"
   End
   Begin PhotoDemon.pdLabel lblMeasure 
      Height          =   255
      Index           =   0
      Left            =   4080
      Top             =   30
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   450
      Alignment       =   1
      Caption         =   "distance:"
   End
   Begin PhotoDemon.pdLabel lblMeasure 
      Height          =   255
      Index           =   1
      Left            =   4080
      Top             =   390
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   450
      Alignment       =   1
      Caption         =   "angle:"
   End
   Begin PhotoDemon.pdLabel lblMeasure 
      Height          =   255
      Index           =   2
      Left            =   4080
      Top             =   750
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   450
      Alignment       =   1
      Caption         =   "width:"
   End
   Begin PhotoDemon.pdLabel lblMeasure 
      Height          =   255
      Index           =   3
      Left            =   4080
      Top             =   1110
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   450
      Alignment       =   1
      Caption         =   "height:"
   End
   Begin PhotoDemon.pdLabel lblValue 
      Height          =   255
      Index           =   0
      Left            =   5640
      Top             =   30
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   450
      Caption         =   "0"
   End
   Begin PhotoDemon.pdLabel lblValue 
      Height          =   255
      Index           =   1
      Left            =   5640
      Top             =   390
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   450
      Caption         =   "0"
   End
   Begin PhotoDemon.pdLabel lblValue 
      Height          =   255
      Index           =   2
      Left            =   5640
      Top             =   750
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   450
      Caption         =   "0"
   End
   Begin PhotoDemon.pdLabel lblValue 
      Height          =   255
      Index           =   3
      Left            =   5640
      Top             =   1110
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   450
      Caption         =   "0"
   End
   Begin PhotoDemon.pdLabel lblMeasure 
      Height          =   255
      Index           =   4
      Left            =   7440
      Top             =   30
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   450
      Alignment       =   1
      Caption         =   "distance:"
   End
   Begin PhotoDemon.pdLabel lblMeasure 
      Height          =   255
      Index           =   5
      Left            =   7440
      Top             =   390
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   450
      Alignment       =   1
      Caption         =   "angle:"
   End
   Begin PhotoDemon.pdLabel lblMeasure 
      Height          =   255
      Index           =   6
      Left            =   7440
      Top             =   750
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   450
      Alignment       =   1
      Caption         =   "width:"
   End
   Begin PhotoDemon.pdLabel lblMeasure 
      Height          =   255
      Index           =   7
      Left            =   7440
      Top             =   1110
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   450
      Alignment       =   1
      Caption         =   "height:"
   End
   Begin PhotoDemon.pdLabel lblValue 
      Height          =   255
      Index           =   4
      Left            =   9000
      Top             =   30
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   450
      Caption         =   "0"
   End
   Begin PhotoDemon.pdLabel lblValue 
      Height          =   255
      Index           =   5
      Left            =   9000
      Top             =   390
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   450
      Caption         =   "0"
   End
   Begin PhotoDemon.pdLabel lblValue 
      Height          =   255
      Index           =   6
      Left            =   9000
      Top             =   750
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   450
      Caption         =   "0"
   End
   Begin PhotoDemon.pdLabel lblValue 
      Height          =   255
      Index           =   7
      Left            =   9000
      Top             =   1110
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   450
      Caption         =   "0"
   End
   Begin PhotoDemon.pdButton cmdAction 
      Height          =   480
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   847
      Caption         =   "straighten image to angle"
   End
   Begin PhotoDemon.pdButton cmdAction 
      Height          =   480
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   847
      Caption         =   "straighten layer to angle"
   End
End
Attribute VB_Name = "toolpanel_Measure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Measurement Tool Panel
'Copyright 2013-2018 by Tanner Helland
'Created: 11/July/18
'Last updated: 12/July/18
'Last update: wrap up initial build
'
'PD's measurement tool is pretty straightforward: measure the distance and angle between two points,
' and relay those values to the user.  Can't beat that for simplicity!
'
'As an added convenience to the user, we also provide options for auto-straightening the image to
' match the current measurement angle.  This is great for visually aligning horizontal or vertical
' elements in a photo.  (And yes - it works for both horizontal *and* vertical lines, and it solves
' for this automagically!)
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Localized text is cached once, at theming time
Private m_NullTextString As String, m_StringsInitialized As Boolean

'The value of all controls on this form are saved and loaded to file by this class
Private WithEvents lastUsedSettings As pdLastUsedSettings
Attribute lastUsedSettings.VB_VarHelpID = -1

'Call to update the current measurement text.
Public Sub UpdateUIText()
    
    If (Not m_StringsInitialized) Then Exit Sub
    
    Dim i As Long
    Dim firstPoint As PointFloat, secondPoint As PointFloat
    Dim measurementUnitText As String
    measurementUnitText = Units.GetNameOfUnit(mu_Pixels)
    
    If Tools_Measure.GetFirstPoint(firstPoint) And Tools_Measure.GetSecondPoint(secondPoint) Then
        
        'Allow point swapping
        cmdAction(0).Enabled = True
        
        'Distance
        Dim measureValue As Double
        If Tools_Measure.GetDistanceInPx(measureValue) Then
            lblValue(0).Caption = Format$(measureValue, "#.0") & " " & measurementUnitText
        Else
            lblValue(0).Caption = m_NullTextString
        End If
        
        'Angle
        If Tools_Measure.GetAngleInDegrees(measureValue) Then
            cmdAction(1).Enabled = (measureValue > 0.001)
            cmdAction(2).Enabled = (measureValue > 0.001)
            measureValue = Abs(measureValue)
            If (measureValue > 90#) Then measureValue = (180# - measureValue)
            lblValue(1).Caption = Format$(measureValue, "#.00") & " " & ChrW(&HB0)
        Else
            cmdAction(1).Enabled = False
            cmdAction(2).Enabled = False
            lblValue(1).Caption = m_NullTextString
        End If
        
        'Width
        lblValue(2).Caption = Format$(Abs(firstPoint.x - secondPoint.x), "#") & " " & measurementUnitText
        
        'Height
        lblValue(3).Caption = Format$(Abs(firstPoint.y - secondPoint.y), "#") & " " & measurementUnitText
        
        'If the current statusbar/ruler unit is something *other* than pixels, display a second set of
        ' measurement values, in said unit.
        If (FormMain.MainCanvas(0).GetRulerUnit <> mu_Pixels) Then
            
            Dim newUnit As PD_MeasurementUnit
            newUnit = FormMain.MainCanvas(0).GetRulerUnit()
            measurementUnitText = Units.GetNameOfUnit(newUnit)
            
            'Ensure the display elements are visible
            If (Not lblValue(4).Visible) Then
                For i = 4 To 7
                    lblMeasure(i).Visible = True
                    lblValue(i).Visible = True
                Next i
            End If
            
            'Repeat the same steps that we used for pixels, but this time, perform an additional conversion
            ' into the target unit space
            If Tools_Measure.GetDistanceInPx(measureValue) Then
                lblValue(4).Caption = Format$(Units.ConvertPixelToOtherUnit(newUnit, measureValue, pdImages(g_CurrentImage).GetDPI), "#.0##") & " " & measurementUnitText
            Else
                lblValue(4).Caption = m_NullTextString
            End If
            
            'Angle
            If Tools_Measure.GetAngleInDegrees(measureValue) Then
                measureValue = Abs(measureValue)
                If (measureValue > 90#) Then measureValue = (180# - measureValue)
                lblValue(5).Caption = Format$(measureValue, "#.00") & " " & ChrW(&HB0)
            Else
                lblValue(5).Caption = m_NullTextString
            End If
            
            'Width
            lblValue(6).Caption = Format$(Units.ConvertPixelToOtherUnit(newUnit, Abs(firstPoint.x - secondPoint.x), pdImages(g_CurrentImage).GetDPI), "#.0##") & " " & measurementUnitText
            
            'Height
            lblValue(7).Caption = Format$(Units.ConvertPixelToOtherUnit(newUnit, Abs(firstPoint.y - secondPoint.y), pdImages(g_CurrentImage).GetDPI), "#.0##") & " " & measurementUnitText
        
        'If the current unit is "pixels", hide the extra info area
        Else
            
            If lblValue(4).Visible Then
                For i = 4 To 7
                    lblMeasure(i).Visible = False
                    lblValue(i).Visible = False
                Next i
            End If
        
        End If
        
    'If a measurement isn't available, blank all labels and disable certain buttons
    Else
        
        For i = 0 To 7
            lblValue(i).Caption = m_NullTextString
        Next i
        
        For i = 4 To 7
            lblMeasure(i).Visible = False
            lblValue(i).Visible = False
        Next i
        
        For i = cmdAction.lBound To cmdAction.UBound
            cmdAction(i).Enabled = False
        Next i
        
    End If
    
End Sub

Private Sub cmdAction_Click(Index As Integer)

    Select Case Index
    
        'Swap points
        Case 0
            Tools_Measure.SwapPoints
        
        'Straighten image to angle
        Case 1
            Tools_Measure.StraightenImageToMatch
            
        'Straighten layer to angle
        Case 2
            Tools_Measure.StraightenLayerToMatch
    
    End Select

End Sub

Private Sub Form_Load()

    Tools.SetToolBusyState True
    
    'Load any last-used settings for this form
    Set lastUsedSettings = New pdLastUsedSettings
    lastUsedSettings.SetParentForm Me
    lastUsedSettings.LoadAllControlValues
    
    Tools.SetToolBusyState False
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'Save all last-used settings to file
    If Not (lastUsedSettings Is Nothing) Then
        lastUsedSettings.SaveAllControlValues
        lastUsedSettings.SetParentForm Nothing
    End If

End Sub

'Updating against the current theme accomplishes a number of things:
' 1) All user-drawn controls are redrawn according to the current g_Themer settings.
' 2) All tooltips and captions are translated according to the current language.
' 3) ApplyThemeAndTranslations is called, which redraws the form itself according to any theme and/or system settings.
'
'This function is called at least once, at Form_Load, but can be called again if the active language or theme changes.
Public Sub UpdateAgainstCurrentTheme()
    
    m_NullTextString = g_Language.TranslateMessage("n/a")
    m_StringsInitialized = True
    
    'Start by redrawing the form according to current theme and translation settings.  (This function also takes care of
    ' any common controls that may still exist in the program.)
    ApplyThemeAndTranslations Me
    
    'As language settings may have changed, we now need to redraw the current UI text
    Me.UpdateUIText

End Sub
