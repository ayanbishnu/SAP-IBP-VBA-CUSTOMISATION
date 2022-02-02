Attribute VB_Name = "RB_FORMAT"
Sub RB_Formatheader()
'
' Developer: "Ayan Bishnu" 09/01/2022: Custom Function to repeat header attributes only against new planning combination.
' Developer: "Ayan Bishnu" 09/01/2022: Jira Issue # SIIBP-1942
' Developer: "Ayan Bishnu" 22/01/2022: This sub of RB_FORMAT module contains the logic of the custom formatting with selective row header
'
Dim ibpAutomationObject As Object
    Dim rptID As String
    Dim topLeftCell As String
    Dim KFRow As Integer
    Dim KFCol As Integer
    Dim KFCell As Range
    Dim Allrows As Range
    Dim KFArray
    Dim sheetName As String
    Dim KFGroupRows As Integer
    Dim KFHeader As String
    Dim rowCount As Integer
    Dim bottomRightCell As String
    Dim groupRow As Integer
    Dim planningViewRange As String
    
' Developer: "Ayan Bishnu" 22/01/2022: Checks to see if the active workbook in connected to a IBP Planning area
    
    If ibpAutomationObject Is Nothing Then Set ibpAutomationObject = Application.COMAddIns("IBPXLClient.Connect").Object
    rptID = ibpAutomationObject.GetActiveReportName(ActiveSheet)
    
    If rptID = "" Then
    
    MsgBox ("Please log-on and select a planning view.")
    Exit Sub
    
    End If

    sheetName = Application.ActiveSheet.name

' Developer: "Ayan Bishnu" 22/01/2022: Start of the actual row header formatting logic

    s = Range("j6").Value
    rows("5:5").Select

    selection.AutoFilter
    Range("J5").Select

    ActiveSheet.Range("A:xfd").AutoFilter Field:=10, Criteria1:="<>" & s, Operator:=xlFilterValues

    Range("A7:I1048576").Select
    
    Range(selection, selection.End(xlDown)).Select
    Range(selection, selection.End(xlToLeft)).Select
    selection.ClearContents
    selection.AutoFilter
    
    Range("J6").Select

' Developer: "Ayan Bishnu" 22/01/2022: End of the actual row header formatting logic
    
End Sub
