VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Mont Carlo Simulation"
   ClientHeight    =   10380
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   15324
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private Sub btnGo_Click()
' For histogram'
Dim datamin As Double, datamax As Double, datarange As Double
Dim lowbins As Integer, highbins As Integer, nbins As Double
Dim binrangeinit As Double, binrangefinal As Double
Dim bins() As Double, bincenters() As Double, j As Integer
Dim c As Integer, i As Integer, R() As Double
Dim bincounts() As Integer, ChartRange As String, nr As Integer
' end of histogram"


'Dim k As Integer
Dim NumberOfStimulation As Double
Dim P As Double, Cland As Double, Tax As Double
'Dim Arr() As Double

If Me.txtNumberOfSimulation = "" Then
MsgBox "Please, enter the number of simulation you would like to run"
Exit Sub
End If

NumberOfStimulation = Me.txtNumberOfSimulation

MsgBox "It will take some time. Please, be patient"
ReDim R(NumberOfStimulation)

For i = 1 To NumberOfStimulation
P = Rnd()

'LandCost_Discrete

    If P < 1 * Me.txtCostOfLandChance1 / 100 Then
        Cland = 1 * Me.txtCostOfLandCost1
        ThisWorkbook.Worksheets("Main").Range("Cland") = Cland

    ElseIf P < 1 * Me.txtCostOfLandChance2 / 100 Then
        Cland = 1 * Me.txtCostOfLandCost2
        ThisWorkbook.Worksheets("Main").Range("Cland") = Cland
    Else:
        Cland = 1 * Me.txtCostOfLandCost3
        ThisWorkbook.Worksheets("Main").Range("Cland") = Cland
    End If

'Royalties Cost_Beta-PERT
        
    Dim Alpha As Double
    Dim Beta As Double, N2 As Double
    Dim Max As Double, Low As Double, Mode As Double
 
    Max = 1 * Me.txtCostOfRoyaltiesHigh
    Mode = 1 * Me.txtCostOfRoyaltiesMode
    Low = 1 * Me.txtCostOfRoyaltiesLow
    
    If Max < 0 Then
        Max = -1 * Max
        End If
    If Mode < 0 Then
         Mode = -1 * Mode
        End If
    If Low < 0 Then
        Low = -1 * Low
    End If
    
    
      Alpha = (4 * Mode + Max - 5 * Low) / (Max - Low)
      Beta = (5 * Max - Low - 4 * Mode) / (Max - Low)
      N2 = WorksheetFunction.Beta_Inv(P, Alpha, Beta, Low, Max)
     
    
      If (1 * Me.txtCostOfRoyaltiesLow < 0) Then
        ThisWorkbook.Worksheets("Main").Range("CRoyal") = -1 * Round(N2, 2)
      Else
        ThisWorkbook.Worksheets("Main").Range("CRoyal") = Round(N2, 2)
      End If
'Total Depreciable Capital_Normal
    ThisWorkbook.Worksheets("Main").Range("CTDC") = Application.WorksheetFunction.Norm_Inv(P, 1 * Me.txtTotalDeprCapitalAvg, 1 * Me.txtTotalDeprCapitalStd)

'Working Capital_Uniform
    ThisWorkbook.Worksheets("Main").Range("WC") = 1 * Me.txtWorkingCapitalMin + (1 * Me.txtWorkingCapitalMax - 1 * Me.txtWorkingCapitalMin) * P

'Startup Cost_Normal
    ThisWorkbook.Worksheets("Main").Range("Cstart") = Application.WorksheetFunction.Norm_Inv(P, 1 * Me.txtStartupCostsAvg, 1 * Me.txtStartupCostsStd)

'Sales Revenue_Beta-PERT
    Max = 1 * Me.txtSalesRevenueHigh
    Mode = 1 * Me.txtSalesRevenueMode
    Low = 1 * Me.txtSalesRevenueLow
    

      Alpha = (4 * Mode + Max - 5 * Low) / (Max - Low)
      Beta = (5 * Max - Low - 4 * Mode) / (Max - Low)

    ThisWorkbook.Worksheets("Main").Range("S").Value = Application.WorksheetFunction.Beta_Inv(P, Alpha, Beta, Low, Max)


'ProductionCost _Triangular
    ThisWorkbook.Worksheets("Main").Range("COS").Value = TriangDistribution(P, 1 * Me.txtProductionCostsLow, 1 * Me.txtProductionCostsMode, 1 * Me.txtProductionCostsHigh)
    'ThisWorkbook.Worksheets("Main").Range("COS").Value = triangular_inverse(P, 1 * Me.txtProductionCostsLow, 1 * Me.txtProductionCostsMode, 1 * Me.txtProductionCostsHigh)
'Tax_Discrete
    If P < 1 * Me.txtTaxChance1 / 100 Then
        Tax = 1 * Me.txtTaxRate1
        ThisWorkbook.Worksheets("Main").Range("Tax") = Tax
    Else:
        Tax = 1 * Me.txtTaxRate2
        ThisWorkbook.Worksheets("Main").Range("Tax") = Tax
   
    End If

'Interest Rate_uniform
    ThisWorkbook.Worksheets("Main").Range("i") = 1 * Me.txtInterestRateMin + (1 * Me.txtInterestRateMax - 1 * Me.txtInterestRateMin) * P

    R(i) = ThisWorkbook.Worksheets("Main").Range("Final")
    
    Dim v As Integer
    If R(i) > 0 Then
        v = v + 1
    End If

Next i

Dim o As Double

o = (v / NumberOfStimulation) * 100
MsgBox (FormatNumber(o, 2) & "%" & " " & "percent of simulations were profitable")
'MsgBox "You profitability of the project is : " & Application.WorksheetFunction.Average(R())

'   The code below creates a histogram of the vector R that you should create above
'   R is a vector of the end result of each simulation (profit per cookie in this case)

datamin = WorksheetFunction.Min(R)
datamax = WorksheetFunction.Max(R)
datarange = datamax - datamin
lowbins = Int(WorksheetFunction.Log(NumberOfStimulation, 2)) + 1
highbins = Int(Sqr(NumberOfStimulation))
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
On Error Resume Next
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
    For i = 1 To NumberOfStimulation
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
UserForm1.Hide
Application.ScreenUpdating = False
Application.DisplayAlerts = False
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
Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub

Private Sub CommandButton1_Click()
Unload Me
End Sub



Function triangular_inverse(P As Double, L As Double, M As Double, U As Double) As Double
'Given a probability P and lower (L), upper (U), and most common (M) inputs, this
'function calculates the corresponding x value
Dim a As Double, b As Double, c As Double
If P < (M - L) / (U - L) Then
    a = 1
    b = -2 * L
    c = L ^ 2 - P * (M - L) * (U - L)
    triangular_inverse = (-b + Sqr(b ^ 2 - 4 * a * c)) / 2 / a
ElseIf P <= 1 Then
    a = 1
    b = -2 * U
    c = U ^ 2 - (1 - P) * (U - L) * (U - M)
    triangular_inverse = (-b - Sqr(b ^ 2 - 4 * a * c)) / 2 / a
End If
End Function

Function TriangDistribution(R, Low As Double, Mode As Double, Max As Double) As Double
    
    Dim M, U, Minv  As Double
    
    M = Mode - Low
    U = Max - Low
    Minv = Max - Mode
    
    If R <= (M / U) Then
    TriangDistribution = Low + (R * M * U) ^ 0.5
    
    Else:
    TriangDistribution = Max + ((1 - R) * U * Minv) ^ 0.5
    End If
  
End Function
