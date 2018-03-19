# VBAscripts
# macros 4 data handling
Sub dram()
# dram Makro / Makro am 14.04.00 von fitsilis aufgezeichnet
# Tastenkombination: Strg+d

Dim zeile, hilfszeile As Integer
Dim name As String
Dim diagramm As Integer
Dim repeat As Integer
Dim filenr As Integer

# ChDir "D:\Excel MOCVD\Excel-files\DRAM und RELAX\141B annealt\"
filenr = 2
hilfszeile = 10
Anfang:
Windows("DRAM_PULSE.xls").Activate
zeile = 10
diagramm = 1
Sheets("Makro").Select
Cells(filenr, 2).Select
If ActiveCell.FormulaR1C1 = "" Then GoTo Ende
name = ActiveCell.FormulaR1C1
filenr = filenr + 1

# name = InputBox("Choose file name: ", "Search")
Sheets.Add
ActiveSheet.name = name
Workbooks.OpenText filename:=name & ".dat", Origin:=xlWindows, StartRow:=1, DataType:=xlDelimited,
TextQualifier:= _
xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, Semicolon:=False, _
Comma:=False, Space:=True, Other:=False, FieldInfo:=Array(Array(1, 1), _
Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), _
Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 1), Array(15 _
, 1), Array(16, 1), Array(17, 1))
Columns("A:B").Select
Selection.NumberFormat = "General"
hier:
While ActiveCell.FormulaR1C1 <> ""
Cells(zeile, 3).Select
ActiveCell.FormulaR1C1 = "=RC[-2]*0.000001"
Cells(zeile, 4).Select
ActiveCell.FormulaR1C1 = "=RC[-2]*0.000001"
zeile = zeile + 1
Cells(zeile, 1).Select
Wend
Range("c" & hilfszeile, "d" & zeile - 1).Select

# --------------------Chart------------------------
Charts.Add
ActiveChart.ChartType = xlXYScatter
ActiveChart.SetSourceData Source:=Sheets(name).Range("c" & hilfszeile, "d" & zeile - 1)
ActiveChart.Location Where:=xlLocationAsObject, name:=name
ActiveChart.Axes(xlValue).Select
With ActiveChart.Axes(xlValue)
'.MinimumScale = 0
End With
ActiveChart.ChartArea.Select
With ActiveChart
.HasTitle = True
.ChartTitle.Characters.Text = "DRAM Pulse - " & name
.Axes(xlCategory, xlPrimary).HasTitle = True
.Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "t [s]"
.Axes(xlValue, xlPrimary).HasTitle = True
.Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "E [kV/cm]"
End With
ActiveChart.Legend.Select
Selection.Delete
ActiveChart.ChartArea.Select
ActiveChart.ChartArea.Copy
Windows("DRAM_PULSE.xls").Activate
ActiveSheet.Paste
If diagramm <> 1 Then
ActiveSheet.Shapes("Diagramm " & diagramm).IncrementTop 220#
End If
diagramm = diagramm + 1
Windows(name).Activate
zeile = zeile + hilfszeile
Range("A1").Select
Cells(zeile, 1).Select
If ActiveCell.FormulaR1C1 <> "" Then GoTo hier Else: GoTo Anfang
Ende:
MsgBox ("Fertig!!!")
End Sub
