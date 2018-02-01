Attribute VB_Name = "Module1"
Sub myDeleteRows2()
  
  Const MyTarget = "ADD_DELETE_CRITERION" ' <-- change to suit
  
  Dim Rng As Range, DelCol As New Collection, x
  Dim i As Long, j As Long, k As Long
  
  ' Calc last row number
  j = Cells.SpecialCells(xlCellTypeLastCell).Row  'can be: j = Range("C" & Rows.Count).End(xlUp).Row
  
  ' Collect rows range with MyTarget
  For i = 1 To j
    If WorksheetFunction.CountIf(Rows(i), MyTarget) > 0 Then
      k = k + 1
      If k = 1 Then
        Set Rng = Rows(i)
      Else
        Set Rng = Union(Rng, Rows(i))
        If k >= 100 Then
          DelCol.Add Rng
          k = 0
        End If
      End If
    End If
  Next
  If k > 0 Then DelCol.Add Rng
  
  ' Turn off screen updating and events
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  
  ' Delete rows with MyTarget
  For Each x In DelCol
    x.Delete
  Next
  
  ' Update UsedRange
  With ActiveSheet.UsedRange: End With
  
  ' Restore screen updating and events
  Application.ScreenUpdating = True
  Application.EnableEvents = True
  

End Sub
