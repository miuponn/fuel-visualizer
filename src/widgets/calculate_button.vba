Sub CalculateAll()

Call RunMonthlyReport
Call RunFuelReport
Call RunGasReport
Call RunNuclearReport

Dim data As Worksheet
Dim DateYear As String

Set data = Sheets("Data")
DateYear = Right(data.Range("A3"), 4)

On Error Resume Next
Sheets(DateYear & " Output By Generator").Activate
Call CreateSheet

On Error Resume Next
Sheets(DateYear & " Output By Fuel Type").Activate
Call CreateFuelSheet

On Error Resume Next
Sheets(DateYear & " Gas Measurements").Activate
Call CreateGasSheet

On Error Resume Next
Sheets(DateYear & " Nuclear Measurements").Activate
Call CreateNuclearSheet


Application.DisplayAlerts = False 'switching off the alert button
Sheets("Monthly Output").Delete
Sheets("Monthly Output By Fuel").Delete
Sheets("Gas").Delete
Sheets("Nuclear").Delete
Application.DisplayAlerts = True 'switching off the alert button

Sheets("Data").Activate

End Sub
