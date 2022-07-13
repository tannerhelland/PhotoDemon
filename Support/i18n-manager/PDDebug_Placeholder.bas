Attribute VB_Name = "PDDebug"
'Placeholder module to allow me to use PhotoDemon code in non-PhotoDemon projects

Option Explicit

Public Enum PD_DebugMessages
    PDM_Normal = 0
    PDM_User_Message = 1
    PDM_Mem_Report = 2
    PDM_HDD_Report = 3
    PDM_Processor = 4
    PDM_External_Lib = 5
    PDM_Startup_Message = 6
    PDM_Timer_Report = 7
End Enum

#If False Then
    Private Const PDM_Normal = 0, PDM_User_Message = 1, PDM_Mem_Report = 2, PDM_HDD_Report = 3, PDM_Processor = 4, PDM_External_Lib = 5, PDM_Startup_Message = 6, PDM_Timer_Report = 7
#End If

'Dummy placeholder
Public Sub LogAction(Optional ByVal actionString As String = vbNullString, Optional ByVal debugMsgType As PD_DebugMessages = PDM_Normal, Optional ByVal suspendMemoryAutoUpdate As Boolean = False)
    Debug.Print actionString
End Sub

