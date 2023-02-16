Sub gas_only()

Call RunGasReport

Dim data As Worksheet
Dim DateYear As String

Set data = Sheets("Data")
DateYear = Right(data.Range("A3"), 4)

On Error Resume Next
Sheets(DateYear & " Gas Measurements").Activate
Call CreateGasSheet

End Sub
