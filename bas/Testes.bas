Attribute VB_Name = "Testes"

Sub teste2()

    MsgBox Worksheets("APOIO").Range(BancoLocal).Value

End Sub

Sub teste()

Dim lRow As Long
Dim lPart As Long
Dim ws As Worksheet
Dim x As Long

'Set ws = Worksheets(ActiveSheet.Name)
Set ws = Worksheets("APOIO")

'find  first empty row in database
lRow = ws.Cells(Rows.count, 2).End(xlUp).Offset(1, 0).Row


For x = 1 To lRow - 1
    With ws
        MsgBox .Cells(x, 1).Value
    End With

Next x

'With ws
'    If Me.txtData.Value <> "" Then .Cells(lRow, 1).Value = Format(CDate(Me.txtData.Value), "mm/dd/yy")
'    .Cells(lRow, 2).Value = Me.cboTelefones.Value
'    .Cells(lRow, 3).Value = CCur(Me.txtValor.Value)
'End With

ThisWorkbook.Save

End Sub


Sub teste4()
Dim ws As Worksheet
Dim cCol As Long

Set ws = Worksheets("Plan1")
ws.Activate

MsgBox ws.Cells(1, Columns.count).End(xlToLeft).Column


End Sub
