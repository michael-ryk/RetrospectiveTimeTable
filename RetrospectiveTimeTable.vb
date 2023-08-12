Option Explicit

Sub RetroTimeTable(cellStyle As String)
    Debug.Print ("===== Start =====")
    Debug.Print ("Color: " & cellStyle)
    
    Const StartDatesRow As Integer = 4
    Const StartDatesCol As String = "F"
    Const DaysCol As String = "E"
    
    Dim dtToday             As Date
    Dim intCurrentRow       As Integer
    Dim intLastColUsed      As Integer
    Dim rngShiftRange       As Range
    Dim rngTodayCell        As Range
    Dim rngDaysCount        As Range
    
    intCurrentRow = ActiveCell.Row
    Set rngTodayCell = Range(StartDatesCol & intCurrentRow)
    Set rngDaysCount = Range(DaysCol & intCurrentRow)
    Debug.Print ("Range Today cell: " & rngTodayCell.Address)
        
    'Shift cells Right
    If Not IsEmpty(rngTodayCell.Value) Then
        intLastColUsed = rngTodayCell.Offset(0, -1).End(xlToRight).Column   'Offset must to handle case with only 1 cell filled
        Debug.Print ("Last used column: " & intLastColUsed)
        Set rngShiftRange = Range(rngTodayCell, Cells(intCurrentRow, intLastColUsed))
        Debug.Print ("Range to shift: " & rngShiftRange.Address)
        rngShiftRange.Cut rngShiftRange.Cells(1).Offset(0, 1)
    End If
    
    'Set today date
    Set rngTodayCell = Range(StartDatesCol & intCurrentRow)
    Debug.Print ("Range Today cell after shift: " & rngTodayCell.Address)
    If IsEmpty(rngTodayCell.Value) Then
        dtToday = Date
        Debug.Print ("First cell is empty - set today value")
        rngTodayCell.Value = dtToday
    End If
    
    'Color today cell based on button pressed
    With rngTodayCell
        .Style = cellStyle
    End With
    
    'Fix formula for days to last
    rngDaysCount.Formula = "=TODAY()-" & StartDatesCol & intCurrentRow
  
End Sub
