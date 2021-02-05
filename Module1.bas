Attribute VB_Name = "Module1"
Option Explicit

Sub Run()

With UserForm1
      .StartUpPosition = 0
      .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
      .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
      .Show
End With

End Sub

Sub Graph()
' For histogram'
Dim datamin As Double, datamax As Double, datarange As Double
Dim lowbins As Integer, highbins As Integer, nbins As Double
Dim binrangeinit As Double, binrangefinal As Double
Dim bins() As Double, bincenters() As Double, j As Integer
Dim c As Integer, i As Integer, R() As Double
Dim bincounts() As Integer, ChartRange As String, nr As Integer
' end of histogram"

'   DO NOT MODIFY THE CODE BELOW!  ONCE A VECTOR R OF RESULTS HAS BEEN CREATED,
'   THE CODE BELOW WILL CREATE A HISTOGRAM.  MAKE SURE NOT TO DELETE OR CHANGE
'   THE NAME OF THE "HISTOGRAM DATA" WORKSHEET IN THIS FILE!

'   The code below creates a histogram of the vector R that you should create above
'   R is a vector of the end result of each simulation (profit per cookie in this case)

datamin = WorksheetFunction.Min(R)
datamax = WorksheetFunction.Max(R)
datarange = datamax - datamin
lowbins = Int(WorksheetFunction.Log(nsimulations, 2)) + 1
highbins = Int(Sqr(nsimulations))
nbins = (lowbins + highbins) / 2
binrangeinit = datarange / nbins
ReDim bins(1) As Double
If binrangeinit < 1 Then
    c = 1
    Do
        If 10 * binrangeinit > 1 Then
            binrangefinal = 10 * binrangeinit Mod 10
            Exit Do
        Else
            binrangeinit = 10 * binrangeinit
            c = c + 1
        End If
    Loop
    binrangefinal = binrangefinal / 10 ^ c
ElseIf binrangeinit < 10 Then
    binrangefinal = binrangeinit Mod 10
Else
    c = 1
    Do
        If binrangeinit / 10 < 10 Then
            binrangefinal = binrangeinit / 10 Mod 10
            Exit Do
        Else
            binrangeinit = binrangeinit / 10
            c = c + 1
        End If
    Loop
    binrangefinal = binrangefinal * 10 ^ c
End If
i = 1
bins(1) = (datamin - ((datamin) - (binrangefinal * Fix(datamin / binrangefinal))))
Do
    i = i + 1
    ReDim Preserve bins(i) As Double
    bins(i) = bins(i - 1) + binrangefinal
Loop Until bins(i) > datamax
nbins = i
ReDim Preserve bincounts(nbins - 1) As Integer
ReDim Preserve bincenters(nbins - 1) As Double
For j = 1 To nbins - 1
    c = 0
    For i = 1 To nsimulations
        If R(i) > bins(j) And R(i) <= bins(j + 1) Then
            c = c + 1
        End If
    Next i
    bincounts(j) = c
    bincenters(j) = (bins(j) + bins(j + 1)) / 2
Next j
Sheets("Histogram Data").Select
Cells.Clear
Range("A1").Select
Range("A1:A" & nbins - 1) = WorksheetFunction.Transpose(bincenters)
Range("B1:B" & nbins - 1) = WorksheetFunction.Transpose(bincounts)
MainForm.Hide
Application.ScreenUpdating = False
Charts("Histogram").Delete
ActiveCell.Range("A1:B1").Select
    Range(Selection, Selection.End(xlDown)).Select
    nr = Selection.Rows.Count
    ChartRange = Selection.Address
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range("'Histogram Data'!" & ChartRange)
    ActiveChart.ChartTitle.Select
    Selection.Delete
    ActiveChart.PlotArea.Select
    Application.CutCopyMode = False
    ActiveChart.FullSeriesCollection(1).Delete
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    ActiveChart.FullSeriesCollection(1).XValues = "='Histogram Data'!" & "$A$1:$A$" & nr
    ActiveChart.Legend.Select
    Selection.Delete
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
    ActiveChart.SetElement (msoElementPrimaryValueAxisTitleAdjacentToAxis)
    Selection.Caption = "Count"
    ActiveChart.Axes(xlCategory).AxisTitle.Select
    Selection.Caption = "Bin Center"
    ActiveChart.ChartArea.Select
    ActiveChart.Location Where:=xlLocationAsNewSheet, Name:="Histogram"
    
'------------------
'   FEEL FREE TO ADD CODE BELOW THIS POINT, E.G. TO OUTPUT A SUMMARY OF THE RESULTS IN MESSAGE BOX(ES)

End Sub
End Sub
