# Cleveland
Open-text Feedback into a useful format


## Program File -- Cleveland.py

- Inputs = Cleveland_Data_Nov.csv
- Output = ClevelandResults_Set1_Nov_1.csv

- The Workbook, containing multiple Worksheets, is separated using the following script:

```
Sub Splitbook()

Dim xPath As String
xPath = Application.ActiveWorkbook.Path
Application.ScreenUpdating = False
Application.DisplayAlerts = False
For Each xWs In ThisWorkbook.Sheets
    xWs.Copy
    Application.ActiveWorkbook.SaveAs Filename:=xPath & "\" & xWs.Name & ".xlsx"
    Application.ActiveWorkbook.Close False
Next
Application.DisplayAlerts = True
Application.ScreenUpdating = True

End Sub
```
