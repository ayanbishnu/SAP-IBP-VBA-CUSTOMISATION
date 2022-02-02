Attribute VB_Name = "EpmAfterRefresh"
'SAP SE, 2018-11-28
'this is sample code and provided "AS IS"
'
' Developer: "Ayan Bishnu" 09/01/2022: Custom Function to repeat header attributes only against new planning combination.
' Developer: "Ayan Bishnu" 09/01/2022: Jira Issue # SIIBP-1942.
' Developer: "Ayan Bishnu" 22/01/2022: This Function module contains the logic to toggle between standard format and the
' Developer: "Ayan Bishnu" 09/01/2022: repeat header format based on  checkbox status in the runtime after every refresh.
'

Function AFTER_REFRESH() As Boolean
'This function is automatically called by the add-in after any
're-rendering of the planning view 'e.g. like after Edit View,
'Refresh, Save Data, Simulate...

' Developer: "Ayan Bishnu" 22/01/2022: Checks if the checkbox exist in the active planning view ? If no it add the checkbox.

On Error Resume Next
If ActiveSheet.OLEObjects("NewCheckBox").Object Is Nothing Then
Call AddCheckBox

'Developer: "Ayan Bishnu" 09/01/2022: UGLY HACK To remove the white line showing up due to fill color.

ActiveSheet.Shapes.Range(Array("NewCheckBox")).Select
ActiveSheet.Shapes("NewCheckBox").Fill.Visible = msoFalse
Else
End If
On Error GoTo 0

AFTER_REFRESH = True

If ActiveSheet.OLEObjects("NewCheckBox").Object.Value = True Then
Call RB_Formatheader
ElseIf ActiveSheet.OLEObjects("NewCheckBox").Object.Value = False Then
End If

'Developer: "Ayan Bishnu" 09/01/2022: Standard SAP code to refresh the IBP embedded charts.

On Error GoTo Ignore:
    Application.Run "SAP_IBP_Chart.xlam!AFTER_REFRESH"

Ignore:

End Function




