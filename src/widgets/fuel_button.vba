Sub by_fuel()

Call RunFuelReport

Dim data As Worksheet
Dim DateYear As String

Set data = Sheets("Data")
DateYear = Right(data.Range("A3"), 4)

On Error Resume Next
Sheets(DateYear & " Output By Fuel Type").Activate
Call CreateFuelSheet

End Sub
