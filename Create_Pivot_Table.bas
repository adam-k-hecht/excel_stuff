Attribute VB_Name = "Module2"

'https://excelchamps.com/blog/vba-to-create-pivot-table/
'huge thanks to Puneet Gogia at Excel Champs for the post!
'please support Puneet's work

Sub createPivotTable()

    'Declare Variables
    
    Dim PSheet As Worksheet
    Dim DSheet As Worksheet
    Dim PCache As PivotCache
    Dim PTable As PivotTable
    Dim PRange As Range
    Dim LastRow As Long
    Dim LastCol As Long
    
    'Insert a blank sheet and delete Excel-created Pivot sheet
    
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets(PivotTable).Delete
    Sheets.Add Before:=ActiveSheet
    ActiveSheet.Name = "NAME_YOUR_SHEET"
    Application.DisplayAlerts = True
    Set PSheet = Worksheets("NAME_YOUR_SHEET")
    Set DSheet = Worksheets("NAME_YOUR_SOURCE")
    
    'Define Data Range
    
    LastRow = DSheet.Cells(Rows.Count, 1).End(xlUp).Row
    LastCol = DSheet.Cells(1, Columns.Count).End(xlToLeft).Column
    Set PRange = DSheet.Cells(1, 1).Resize(LastRow, LastCol)
    
    'Define Pivot Cache because Excel is stupid
    
    Set PCache = ActiveWorkbook.PivotCaches.Create _
    (SourceType:=xlDatabase, SourceData:=PRange). _
    createPivotTable(TableDestination:=PSheet.Cells(2, 2), _
    TableName:="NAME_YOUR_TABLE")
    
    'Insert Blank Pivot
    
    Set PTable = PCache.createPivotTable _
    (TableDestination:=PSheet.Cells(1, 1), TableName:="NAME_YOUR_TABLE")
    
    'Insert Rows
    
    With ActiveSheet.PivotTables("YOUR_TABLE").PivotFields("YOUR_ROW")
    .Orientation = xlRowField
    .Position = 1
    End With
    
    Insert Columns
    With ActiveSheet.PivotTables("YOUR_TABLE").PivotFields("YOUR_COLUMN")
    .Orientation = xlColumnField
    .Position = 1
    End With
    
    'Insert Data
    
    With ActiveSheet.PivotTables("YOUR_TABLE").PivotFields("YOUR_DATA")
    .Orientation = xlDataField
    .Position = 1
    .Function = xlSum
    .NumberFormat = "#,##0"
    .Name = "YOUR_DATA_DISPLAY_NAME"
    End With
       
    'Format Pivot
    TableActiveSheet.PivotTables("YOUR_TABLE").ShowTableStyleRowStripes = TrueActiveSheet.PivotTables("YOUR_TABLE").TableStyle2 = "PivotStyleMedium9"      

End Sub
