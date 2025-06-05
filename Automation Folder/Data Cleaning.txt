Sub DataCleanup()

Application.Calculation = xlCalculationManual


'Inserting column for Combination/Duplicate removal
Range("I1").EntireColumn.Insert
Application.Worksheets("Data").Range("I1") = "Combo"

'Add Combinations to combo colum with DOB
Application.Worksheets("Data").Range("I2").Formula = "=H2&G2&N2"
Application.Worksheets("Data").Range("I2:I5000").FillDown

'Check for Duplicates in Data
    Columns("I:I").Select
    Selection.FormatConditions.AddUniqueValues
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).DupeUnique = xlDuplicate
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
 'Replace all Not Availables
 Worksheets("Data").Columns("A:BV").Replace _
 What:="Not Available", Replacement:="", _
 SearchOrder:=xlByColumns, MatchCase:=False
 
  Worksheets("Data").Columns("A:BV").Replace "N/A", "", xlWhole
 
 'Replace Food or non fasting with NonFasting
 Worksheets("Data").Columns("AU").Replace "No", "NonFasting", xlWhole
 

 Worksheets("Data").Columns("AU").Replace _
 What:="Food 0-2 hrs ago", Replacement:="NonFasting", _
 SearchOrder:=xlByColumns, MatchCase:=False
 
 Worksheets("Data").Columns("AU").Replace _
 What:="", Replacement:="NonFasting", _
 SearchOrder:=xlByColumns, MatchCase:=False
 
 
 'Calculate the CVD RISK by SUMMING the rows
Application.Worksheets("Data").Range("BN2").Formula = "=BB2+AT2+AK2+AD2+S2+P2+AA2"
Application.Worksheets("Data").Range("BN2:BN5000").FillDown

'Delete blank rows at end of work book
Application.Worksheets("Data").Range("A1:A5000").Select
Selection.SpecialCells(xlCellTypeBlanks).EntireRow.Delete

'Refresh All pivots
ActiveWorkbook.RefreshAll


'Save
ActiveWorkbook.Save

Application.ScreenUpdating = True

Application.Calculation = xlCalculationAutomatic

 
End Sub
