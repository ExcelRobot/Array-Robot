Attribute VB_Name = "modRange"
Option Explicit

Function ListTargetColumnIndexesInSpillingRange(targetCells As Range) As String
    Dim indexCol As Collection
    Dim area As Range
    Dim col As Range
    Dim spillRange As Range
    Dim targetCellsInSpillRange As Range
    Dim firstSpillColumn As Integer
    
    If targetCells.Cells(1).SpillParent Is Nothing Then Exit Function
    
    Set spillRange = targetCells.Cells(1).SpillParent.SpillingToRange
    Set targetCellsInSpillRange = Intersect(spillRange, targetCells)
    
    firstSpillColumn = spillRange.Column
    
    Set indexCol = New Collection
    
    For Each area In targetCellsInSpillRange.Areas
        For Each col In targetCellsInSpillRange.Columns
            indexCol.Add col.Column - firstSpillColumn + 1
        Next
    Next
    
    Dim ctr As Integer
    Dim sortFormula As String
    
    sortFormula = "=TEXTJOIN("","",FALSE,SORT(UNIQUE({"
    For ctr = 1 To indexCol.Count
        sortFormula = sortFormula & indexCol.Item(ctr) & IIf(ctr < indexCol.Count, ";", "")
    Next
    sortFormula = sortFormula & "})))"
    
    ListTargetColumnIndexesInSpillingRange = "{" & Evaluate(sortFormula) & "}"

End Function

Function ListNonTargetColumnIndexesInSpillingRange(targetCells As Range) As String
    Dim indexCol As Collection
    Dim area As Range
    Dim col As Range
    Dim spillRange As Range
    Dim targetCellsInSpillRange As Range
    Dim firstSpillColumn As Integer
    
    If targetCells.Cells(1).SpillParent Is Nothing Then Exit Function
    
    Set spillRange = targetCells.Cells(1).SpillParent.SpillingToRange
    Set targetCellsInSpillRange = Intersect(spillRange, targetCells)
    
    firstSpillColumn = spillRange.Column
    
    Set indexCol = New Collection
    Dim ctr As Integer
    
    For ctr = 1 To spillRange.Columns.Count
        indexCol.Add ctr, CStr(ctr)
    Next
    
    On Error Resume Next
    For Each area In targetCellsInSpillRange.Areas
        For Each col In targetCellsInSpillRange.Columns
            indexCol.Remove CStr(col.Column - firstSpillColumn + 1)
        Next
    Next
    On Error GoTo 0
    
    Dim sortFormula As String
    
    sortFormula = "=TEXTJOIN("","",FALSE,SORT(UNIQUE({"
    For ctr = 1 To indexCol.Count
        sortFormula = sortFormula & indexCol.Item(ctr) & IIf(ctr < indexCol.Count, ";", "")
    Next
    sortFormula = sortFormula & "})))"
    
    ListNonTargetColumnIndexesInSpillingRange = "{" & Evaluate(sortFormula) & "}"

End Function

Function ListTargetRowIndexesInSpillingRange(targetCells As Range) As String
    Dim indexCol As Collection
    Dim area As Range
    Dim row As Range
    Dim spillRange As Range
    Dim targetCellsInSpillRange As Range
    Dim firstSpillRow As Integer
    
    If targetCells.Cells(1).SpillParent Is Nothing Then Exit Function
    
    Set spillRange = targetCells.Cells(1).SpillParent.SpillingToRange
    Set targetCellsInSpillRange = Intersect(spillRange, targetCells)
    
    firstSpillRow = spillRange.row
    
    Set indexCol = New Collection
    
    For Each area In targetCellsInSpillRange.Areas
        For Each row In targetCellsInSpillRange.Rows
            indexCol.Add row.row - firstSpillRow + 1
        Next
    Next
    
    Dim ctr As Integer
    Dim sortFormula As String
    
    sortFormula = "=TEXTJOIN("","",FALSE,SORT(UNIQUE({"
    For ctr = 1 To indexCol.Count
        sortFormula = sortFormula & indexCol.Item(ctr) & IIf(ctr < indexCol.Count, ";", "")
    Next
    sortFormula = sortFormula & "})))"
    
    ListTargetRowIndexesInSpillingRange = "{" & Evaluate(sortFormula) & "}"

End Function

Function ListNonTargetRowIndexesInSpillingRange(targetCells As Range) As String
    Dim indexCol As Collection
    Dim area As Range
    Dim row As Range
    Dim spillRange As Range
    Dim targetCellsInSpillRange As Range
    Dim firstSpillRow As Integer
    
    If targetCells.Cells(1).SpillParent Is Nothing Then Exit Function
    
    Set spillRange = targetCells.Cells(1).SpillParent.SpillingToRange
    Set targetCellsInSpillRange = Intersect(spillRange, targetCells)
    
    firstSpillRow = spillRange.row
    
    Set indexCol = New Collection
    Dim ctr As Integer
    
    For ctr = 1 To spillRange.Rows.Count
        indexCol.Add ctr, CStr(ctr)
    Next
    
    On Error Resume Next
    For Each area In targetCellsInSpillRange.Areas
        For Each row In targetCellsInSpillRange.Rows
            indexCol.Remove CStr(row.row - firstSpillRow + 1)
        Next
    Next
    On Error GoTo 0
    
    Dim sortFormula As String
    
    sortFormula = "=TEXTJOIN("","",FALSE,SORT(UNIQUE({"
    For ctr = 1 To indexCol.Count
        sortFormula = sortFormula & indexCol.Item(ctr) & IIf(ctr < indexCol.Count, ";", "")
    Next
    sortFormula = sortFormula & "})))"
    
    ListNonTargetRowIndexesInSpillingRange = "{" & Evaluate(sortFormula) & "}"

End Function

