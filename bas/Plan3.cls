VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Plan3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()
Dim ws As Worksheet
Dim pt As PivotTable

For Each ws In ActiveWorkbook.Worksheets
  For Each pt In ws.PivotTables
    pt.PivotCache.Refresh
  Next pt
Next ws

End Sub
