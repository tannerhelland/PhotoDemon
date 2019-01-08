VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pd2DSurfaceCairo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon 2D Rendering Surface class - using Cairo instead of GDI+
'Copyright 2018-2018 by Tanner Helland
'Created: 25/June/18
'Last updated: 28/June/18
'Last update: continue migrating features from the base pd2DSurface class
'
'In the pd2D drawing model, "Surfaces" are objects onto which you can paint.  At present, there are two types of surfaces:
'
' 1) Wrappers around existing DCs.  This is help for UI painting, as wrapping a DC is effectively instantaneous
'    (compared to making a copy of the DC's bitmap, painting to that, then flipping back to the DC when finished)
'
' 2) Bitmap (raster) surfaces.  The creation and destruction of such surfaces is handled internally, using a pdDIB object.
'
'Regardless of which surface type you use, you can always access an hDC for external GDI interop purposes.
'
'All source code in this file is licensed under a modified BSD license. This means you may use the code in your own
' projects IF you provide attribution. For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Current cairo surface handle; 0 if the class has not been wrapped/created yet
Private m_hSurface As Long

'When wrapping a surface, PD automatically creates a drawing context.  This allows the class to work similarly
' to a GDI+ Graphics object (where drawing settings are stored alongside the image).
Private m_hContext As Long

'Individual context/surface settings
Private m_Operator As Cairo_Operator
Private m_Antialias As Cairo_Antialias

Friend Function GetContextHandle() As Long
    GetContextHandle = m_hContext
End Function

Friend Function GetOperator() As Long
    GetOperator = m_Operator
End Function

Friend Function GetSurfaceHandle() As Long
    GetSurfaceHandle = m_hSurface
End Function

Friend Sub SetAntialias(ByVal newAA As Cairo_Antialias)
    m_Antialias = newAA
    If (m_hContext <> 0) Then Plugin_Cairo.Context_SetAntialias m_hContext, m_Antialias
End Sub

Friend Sub SetOperator(ByVal newOperator As Cairo_Operator)
    m_Operator = newOperator
    If (m_hContext <> 0) Then Plugin_Cairo.Context_SetOperator m_hContext, newOperator
End Sub

Friend Function WrapAroundPDDIB(ByRef srcDIB As pdDIB) As Boolean

    If (m_hSurface <> 0) Or (m_hContext <> 0) Then DeleteSurface
    m_hSurface = Plugin_Cairo.GetCairoSurfaceFromPDDib(srcDIB, m_hContext)
    WrapAroundPDDIB = (m_hSurface <> 0) And (m_hContext <> 0)
    
    'On a successful wrap, relay any set parameters
    If WrapAroundPDDIB Then
        Plugin_Cairo.Context_SetOperator m_hContext, m_Operator
        Plugin_Cairo.Context_SetAntialias m_hContext, m_Antialias
    End If
    
End Function

Friend Function WrapAroundDC(ByVal dstDC As Long) As Boolean

    If (m_hSurface <> 0) Or (m_hContext <> 0) Then DeleteSurface
    m_hSurface = Plugin_Cairo.WrapCairoSurfaceAroundDC(dstDC, m_hContext)
    WrapAroundDC = (m_hSurface <> 0) And (m_hContext <> 0)
    
    'On a successful wrap, relay any set parameters
    If WrapAroundDC Then
        Plugin_Cairo.Context_SetOperator m_hContext, m_Operator
        Plugin_Cairo.Context_SetAntialias m_hContext, m_Antialias
    End If
    
End Function

Friend Sub DeleteSurface()
    If (m_hSurface <> 0) Then Plugin_Cairo.FreeCairoSurface m_hSurface
    If (m_hContext <> 0) Then Plugin_Cairo.FreeCairoContext m_hContext
End Sub

Private Sub Class_Initialize()
    m_Operator = co_Over
    m_Antialias = ca_NONE
End Sub

Private Sub Class_Terminate()
    DeleteSurface
End Sub
