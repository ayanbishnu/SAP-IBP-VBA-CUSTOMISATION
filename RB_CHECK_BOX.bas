Attribute VB_Name = "RB_CHECK_BOX"
Sub AddCheckBox()
'
' Developer: "Ayan Bishnu" 09/01/2022: Custom Function to repeat header attributes only against new planning combination.
' Developer: "Ayan Bishnu" 09/01/2022: Jira Issue # SIIBP-1942
' Developer: "Ayan Bishnu" 22/01/2022: This sub of RB_CHECK_BOX module contains the code to create a checkbox in the runtime
'
Dim c As Range
Dim myCBX As OLEObject
Dim wks As Worksheet
Dim rngCB As Range
Dim strCap As String
Dim ForeColor As OLE_COLOR

' Developer: "Ayan Bishnu" 22/01/2022: Start of the activeX checkbox creation

Set wks = ActiveSheet
Set rngCB = wks.Range("J2")
strCap = "REPEAT ROW HEADERS ON"


For Each c In rngCB
  With c
    Set myCBX = ActiveSheet.OLEObjects.add(ClassType:="Forms.CheckBox.1", _
        Left:=400.75, Top:=43, Width:=160, Height:=22.5)
  End With
  With myCBX
        .name = "NewCheckBox"
        .Object.Caption = strCap
        .Object.Font.Bold = True
        .Object.ForeColor = RGB(255, 255, 255)
        .Object.BackColor = RGB(0, 26, 114)
        .Object.Value = True
        .LinkedCell = Cells(2, 10).Address(0, 0)
'        .Object.TripleState = True
'        Call RB_Formatheader
'        Call NewCheckBox_Click
'        .Object.OnAction = !RemoveCheckBox
'If ActiveSheet.OLEObjects("NewCheckBox").Object.Value = False Then
'Call RemoveCheckBox
'End If
  End With

Next c

' Developer: "Ayan Bishnu" 22/01/2022: End of the activeX checkbox creation

End Sub

