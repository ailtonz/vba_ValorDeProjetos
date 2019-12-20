Attribute VB_Name = "Projetos"
'' BANCOS DE DADOS
Public Const BancoLocal As String = "A2"
Public Const SenhaBanco As String = "abc"
Public Const Vendedor As String = "D2"


Sub Selecao(ByVal Control As IRibbonControl)
    frmSelecao.Show
End Sub

Sub SelecionarBanco(ByVal strGuia As String)

Dim db As DAO.Database
Dim fd As Office.FileDialog
Dim ws As Worksheet
Dim lRow As Long

Dim strBanco As String
Dim strQry As String
Dim strSQL As String

Inicio:

Set ws = Worksheets(strGuia)
strBanco = ws.Range(BancoLocal).Value

'SELECIONAR O BANCO
If Not getFileStatus(strBanco) Then

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Filters.Clear
    fd.Filters.Add "BDs do Access", "*.MDB"
    fd.Title = "Por favor selecione a base de dados para uso da planilha."
    fd.AllowMultiSelect = False
    
    'ATUALIZAR CAMINHO DO BANCO
    If fd.Show = -1 Then
        ws.Range(BancoLocal).Value = fd.SelectedItems(1)
        ThisWorkbook.Save
        GoTo Inicio
    End If
    
'ATUALIZAR BANCO
Else
    'CARREGAR BANCO
    Set db = DBEngine.OpenDatabase(strBanco, False, False, "MS Access;PWD=" & SenhaBanco)
        
    'ENCONTRAR PRIMEIRA LINHA VAZIA NA GUIA
    lRow = ws.Cells(Rows.count, 2).End(xlUp).Offset(1, 0).Row
        
    'CARREGAR PARAMETROS DAS NOVAS CONSULTAS
    For x = 2 To lRow - 1
        With ws
            strQry = .Cells(x, 2).Value
            strSQL = .Cells(x, 3).Value
            
            'VERIFICAR A EXISTENCIA DA CONSULTA NO BANCO
            If Not qryExists(strQry, strBanco, SenhaBanco) Then
                'CRIAR CONSULTA NO BANCO DE DADOS
                db.CreateQueryDef strQry, strSQL
            Else
                'EXCLUSÃO DE CONSULTA
                db.QueryDefs.Delete strQry
                'CRIAR CONSULTA NO BANCO DE DADOS
                db.CreateQueryDef strQry, strSQL
            End If
            
        End With
    
    Next x
    
    db.Close
    
    Set db = Nothing

End If

End Sub

Sub CarregarDados(ByVal strNomeDaGuia As String, ByVal strBanco As String, ByVal strSenhaBanco As String, ByVal strSQL As String)
Dim db As DAO.Database
Dim rst As DAO.Recordset
Dim ws As Worksheet

Dim lRow As Long
Dim cCol As Long

Set db = DBEngine.OpenDatabase(strBanco, False, False, "MS Access;PWD=" & strSenhaBanco)
Set rst = db.OpenRecordset(strSQL)

Set ws = Worksheets(strNomeDaGuia)
ws.Activate
    
'ENCONTRAR PRIMEIRA LINHA VAZIA NA GUIA
lRow = ws.Cells(Rows.count, 2).End(xlUp).Offset(1, 0).Row

''ENCONTRAR ULTIMA COLUNA
'cCol = ws.Cells(1, Columns.Count).End(xlToLeft).Column

With ActiveSheet
'    'CARREGAR TITULOS DOS CAMPOS
'    For i = 0 To rst.Fields.count - 1
'        .Cells(1, i + 1) = rst.Fields(i).Name
'    Next i
    
    'FORMATAR
    ws.Range(ws.Cells(1, 1), ws.Cells(1, rst.Fields.count)).Font.Bold = True
        
    'LIMPAR
    ws.Range("A2:P" & lRow).Clear
        
    'COLAR
    ws.Range("A2").CopyFromRecordset rst

End With

rst.Close
db.Close

Set rst = Nothing
Set db = Nothing

End Sub

Sub LimparDados(ByVal strNomeDaGuia As String)
Dim ws As Worksheet
Set ws = Worksheets(strNomeDaGuia)

    ws.Activate

    'ENCONTRAR PRIMEIRA LINHA VAZIA NA GUIA
    lRow = ws.Cells(Rows.count, 2).End(xlUp).Offset(1, 0).Row
    
    'LIMPAR DADOS
    ws.Range("A2:P" & lRow).Clear
    
        
End Sub
