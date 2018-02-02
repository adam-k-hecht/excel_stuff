Attribute VB_Name = "Module3"
Sub CreateCalcFieldCPM()

    Dim PvtTbl As PivotTable
    Set PvtTbl = Worksheets("YOUR_SHEET").PivotTables("YOUR_PIVOT_TABLE")
    
    'Show 0 in empty cells, like an IFERROR but cooler because VBA
    PvtTbl.NullString = "0"
    PvtTbl.DisplayNullString = True
    
    PvtTbl.CalculatedFields.Add Name:="Calc. eCPM", Formula:="=(YOUR_FORMULA), UseStandardFormula:=True
    
    With PvtTbl.PivotFields("YOUR_CALC_FIELD_NAME")
        .Orientation = xlDataField
        .Function = xlSum
        .Position = 4
        .NumberFormat = "$#0.00"
        .Caption = "YOUR_NEW_FIELD_NAME"
    End With
        
End Sub
