Sub nuclear_only()

Call RunNuclearReport

Dim data As Worksheet
Dim DateYear As String

Set data = Sheets("Data")
DateYear = Right(data.Range("A3"), 4)

On Error Resume Next
Sheets(DateYear & " Nuclear Measurements").Activate
Call CreateNuclearSheet

End Sub
