VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Activate()
    Dim Worksheet As Excel.Worksheet
    
    'Disable cell drag and drop
    Application.CellDragAndDrop = False
    
    'Disable row and column headers
    For Each Worksheet In Application.Worksheets
        Worksheet.Activate
        ActiveWindow.DisplayHeadings = False
    Next Worksheet
    
    Me.Worksheets(1).Activate
    
    RemoveToolbars
End Sub
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    'Enable cell drag and drop
    Application.CellDragAndDrop = True
    RestoreToolbars
End Sub
Private Sub Workbook_Deactivate()
    Dim Worksheet As Excel.Worksheet
    
    'Enable cell drag and drop
    Application.CellDragAndDrop = True
    
    'Enable row and column headers
    For Each Worksheet In Application.Worksheets
        Worksheet.Activate
        ActiveWindow.DisplayHeadings = True
    Next Worksheet
    
    RestoreToolbars
End Sub
Private Sub Workbook_Open()
    'Disable cell drag and drop
    Application.CellDragAndDrop = False
    RemoveToolbars
End Sub







