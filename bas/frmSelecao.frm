VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelecao 
   Caption         =   "SELEÇÃO DE DADOS"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8040
   OleObjectBlob   =   "frmSelecao.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSelecao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdApoio_Click()
    Sheets("APOIO").Visible = xlSheetVisible
End Sub

Private Sub cmdCARREGAR_Click()
Dim strBanco As String: strBanco = Worksheets("Apoio").Range(BancoLocal)
Dim strVendedor As String: strVendedor = Worksheets("Apoio").Range(Vendedor)
Dim tmpVendedor As String

Dim strSQL As String
Dim lst As Controls
Dim strAno As String

Dim strMSG As String
Dim strTitulo As String




'SELECIONAR DADOS
'...

    If ListBoxChecarSelecao(Me, Me.lstAno.Name) = False Then
        strMSG = "Ops!!! " & Chr(10) & Chr(13) & Chr(13)
        strMSG = strMSG & "Você esqueceu de selecionar um ANO(S) da lista. " & Chr(10) & Chr(13) & Chr(13)
        strTitulo = "SELEÇÃO DO ANO(S)!"
        
        MsgBox strMSG, vbInformation + vbOKOnly, strTitulo
'        Exit Sub
        
    ElseIf ListBoxChecarSelecao(Me, Me.lstVendedores.Name) = False And strVendedor = "" Then
        strMSG = "Ops!!! " & Chr(10) & Chr(13) & Chr(13)
        strMSG = strMSG & "Você esqueceu de selecionar um VENDEDOR(ES) da lista. " & Chr(10) & Chr(13) & Chr(13)
        strTitulo = "SELEÇÃO DE VENDEDOR(ES)!"
        
        MsgBox strMSG, vbInformation + vbOKOnly, strTitulo
'        Exit Sub
        
    Else
        
        '---------------------
        'SELECIONAR ANOS
        '---------------------
        
        For intCurrentRow = 0 To Me.lstAno.ListCount - 1
            DoEvents
            
            If Me.lstAno.Selected(intCurrentRow) Then
                
                strAno = strAno & "'" & Me.lstAno.List(intCurrentRow) & "',"
                
                ''' DESMARCAR ITEM SELECIONADO
                Me.lstAno.Selected(intCurrentRow) = False
            End If
            
        Next intCurrentRow
        
        strAno = Left(strAno, Len(strAno) - 1) & ""
        
'        MsgBox strAno
        
        
        '---------------------
        'SELECIONAR VENDEDORES
        '---------------------
        
        If strVendedor = "" Then
        
            For intCurrentRow = 0 To Me.lstVendedores.ListCount - 1
                DoEvents
                
                If Me.lstVendedores.Selected(intCurrentRow) Then
                    
                    strVendedor = strVendedor & "'" & Me.lstVendedores.List(intCurrentRow) & "',"
                    
                    ''' DESMARCAR ITEM SELECIONADO
                    Me.lstVendedores.Selected(intCurrentRow) = False
                End If
                
            Next intCurrentRow
            
            strVendedor = Left(strVendedor, Len(strVendedor) - 1) & ""
            
        Else
        
            tmpVendedor = "'" & strVendedor & "'"
            strVendedor = tmpVendedor
            
        End If
        
'        MsgBox strVendedor
    
        strSQL = "Select * from " & "qryOrcamentosValorMinimosDosProjetos" & " Where "
        strSQL = strSQL & " (((qryOrcamentosValorMinimosDosProjetos.strAno) In (" & strAno & ")) AND ((qryOrcamentosValorMinimosDosProjetos.VENDEDOR) In (" & strVendedor & ")))"
        
               
        ''CARREGAR DADOS
        CarregarDados "DADOS", strBanco, SenhaBanco, strSQL
        
        ''SALVAR PLANILHA
        ThisWorkbook.Save
        
        Me.Hide
        
    End If




End Sub

Private Sub UserForm_Initialize()
Dim strAnos As String:  strAnos = "Select Distinct strAno from qryOrcamentosValorMinimosDosProjetos Order by strAno"
Dim strVendedores As String: strVendedores = "SELECT DISTINCT UCase([VENDEDOR]) AS strVendedor FROM qryOrcamentosValorMinimosDosProjetos ORDER BY UCase([VENDEDOR])"

Dim strVendedor As String: strVendedor = Worksheets("Apoio").Range(Vendedor)

Sheets("APOIO").Visible = xlSheetVeryHidden
'Sheets("APOIO").Visible = xlSheetVisible

If strVendedor <> "" Then
    Me.lstVendedores.Enabled = False
    Me.cmdAPOIO.Enabled = False
Else
    Me.lstVendedores.Enabled = True
    Me.cmdAPOIO.Enabled = True
End If


'SELECIONAR BANCO
SelecionarBanco "APOIO"

Dim strBanco As String: strBanco = Worksheets("Apoio").Range(BancoLocal)

'LIMPAR GUIA DE DADOS
LimparDados "DADOS"

'CARREGAR LIST BOX
ListBoxCarregar strBanco, Me, Me.lstAno.Name, "strAno", strAnos
ListBoxCarregar strBanco, Me, Me.lstVendedores.Name, "strVendedor", strVendedores



End Sub

