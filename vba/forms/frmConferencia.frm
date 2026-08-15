VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmConferencia 
   Caption         =   "CONFERÊNCIA NOTAS DE ENTRADA"
   ClientHeight    =   8415.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14460
   ' OleObjectBlob removido na versao publica
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmConferencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TipoEntrada As String
Public NumeroNFE_XML As String
Public IdCompra As Long

Private Sub btnAdcItens_Click()

Dim item As ListItem
Dim rs As ADODB.Recordset
Dim rsProd As ADODB.Recordset
Dim dados As Variant
Dim QtdeSelecionados As Long
Dim sql As String
Dim MsgRetorno As String

If IdCompra = 0 Then
    MsgBox "Compra inválida.", vbExclamation
    Exit Sub
End If

'Validação
For Each item In lvProdutosEntrada.ListItems

    If item.Checked Then

        If item.Text = "SEM CADASTRO" Then
            MsgBox "Existem produtos sem cadastro selecionados.", vbExclamation
            Exit Sub
        End If

        QtdeSelecionados = QtdeSelecionados + 1

    End If

Next

If QtdeSelecionados = 0 Then
    MsgBox "Selecione ao menos um produto.", vbExclamation
    Exit Sub
End If

'Processa
For Each item In lvProdutosEntrada.ListItems

    If item.Checked Then

        dados = Split(item.Tag, "|")

        sql = "SELECT CODIGOBARRASCX,MEDIDACOMPRA,QUANTIDADECOMPRA,QUANTIDADEEMBALAGEM " & _
              "FROM TAB_PRODUTOS WHERE IDPRODUTO=" & SqlNumero(item.Text)

        Set rsProd = Conn.Execute(sql)

        If rsProd.EOF Then
            MsgBox "Produto ID " & item.Text & " não encontrado.", vbExclamation
            GoTo ProximoItem
        End If

        If NzDbl(rsProd!QUANTIDADECOMPRA) <= 0 Then
            MsgBox "Produto '" & item.SubItems(1) & "' possui Quantidade Compra inválida.", vbExclamation
            GoTo ProximoItem
        End If

        If NzDbl(rsProd!QUANTIDADEEMBALAGEM) <= 0 Then
            MsgBox "Produto '" & item.SubItems(1) & "' possui Quantidade Embalagem inválida.", vbExclamation
            GoTo ProximoItem
        End If

        sql = "CALL PROC_ENTRADAPRODUTOS(" & _
              SqlNumero(IdCompra) & "," & _
              SqlNumero(IDUsuarioLogado) & "," & _
              SqlNumero(item.Text) & "," & _
              SqlNumero(dados(0)) & "," & _
              SqlTexto(dados(1)) & "," & _
              SqlTexto(Nz(rsProd!CODIGOBARRASCX)) & "," & _
              SqlTexto(Nz(rsProd!MEDIDACOMPRA)) & "," & _
              SqlNumero(NzDbl(rsProd!QUANTIDADECOMPRA)) & "," & _
              SqlNumero(NzDbl(rsProd!QUANTIDADEEMBALAGEM)) & "," & _
              SqlNumero(item.SubItems(3)) & "," & _
              SqlTexto("") & "," & _
              SqlNumero(item.SubItems(5)) & ")"

        Set rs = Conn.Execute(sql)

        If Not rs Is Nothing Then
            If Not rs.EOF Then MsgRetorno = Nz(rs.Fields(0).Value)
        End If

ProximoItem:

        If Not rs Is Nothing Then
            If rs.State = adStateOpen Then rs.Close
        End If

        If Not rsProd Is Nothing Then
            If rsProd.State = adStateOpen Then rsProd.Close
        End If

        Set rs = Nothing
        Set rsProd = Nothing

    End If

Next

Dim ws As Worksheet
Dim tb As ListObject
Dim Linha As ListRow
Dim ColNFE As Long
Dim ColStatus As Long

Set ws = ThisWorkbook.Worksheets("ENTRADA CONSOLIDADO")
Set tb = ws.ListObjects("STATUSNF")

ColNFE = tb.ListColumns("NUMERO NFE").Index
ColStatus = tb.ListColumns("STATUS").Index

For Each Linha In tb.ListRows

    If CStr(Linha.Range.Cells(1, ColNFE).Value) = CStr(NumeroNFE_XML) Then

        Linha.Range.Cells(1, ColStatus).Value = "PROCESSADO"
        Exit For

    End If

Next Linha

If MsgRetorno <> "" Then MsgBox MsgRetorno, vbInformation

With frmItensCompra
    .TipoEntrada = "MANUAL"
    .txtIdCompra.Value = IdCompra
    .Show vbModal
End With

Unload Me

End Sub

Private Sub UserForm_initialize()

    With lvProdutosEntrada

        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .HideSelection = False
        .CheckBoxes = True

        .ColumnHeaders.Clear

        .ColumnHeaders.Add , , "ID", 90
        .ColumnHeaders.Add , , "PRODUTO", 220
        .ColumnHeaders.Add , , "MEDIDA", 60
        .ColumnHeaders.Add , , "QUANTIDADE", 80
        .ColumnHeaders.Add , , "CUSTO UNITARIO", 100
        .ColumnHeaders.Add , , "SUBTOTAL", 90

    End With

End Sub

Private Sub Userform_activate()

    Static Carregado As Boolean

    If Carregado Then Exit Sub

    CarregarProdutosEntrada NumeroNFE_XML

    Carregado = True

End Sub
Public Function ItemJaLancadoCompra(ByVal IdCompra As Long, ByVal IDProduto As Long) As Boolean

    Dim rs As ADODB.Recordset
    Dim sql As String

    sql = "SELECT COUNT(*) AS QTDE " & _
          "FROM tab_itenscompra " & _
          "WHERE IDCOMPRA = " & SqlNumero(IdCompra) & _
          " AND IDPRODUTO = " & SqlNumero(IDProduto)

    Set rs = Conn.Execute(sql)

    ItemJaLancadoCompra = NzDbl(rs!qtde) > 0

    rs.Close
    Set rs = Nothing

End Function

Public Function BuscarProdutoXML_Geral(ByVal CodigoBarras As String, ByVal CodigoFornecedor As String) As Variant

    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim whereSql As String

    If Trim(CodigoBarras & "") <> "" Then
        whereSql = "CODIGOBARRAS = " & SqlTexto(CodigoBarras)
    End If

    If Trim(CodigoFornecedor & "") <> "" And IsNumeric(CodigoFornecedor) Then
        If whereSql <> "" Then whereSql = whereSql & " OR "
        whereSql = whereSql & "CODIGOFORNECEDOR = " & SqlNumero(CodigoFornecedor)
    End If

    If whereSql = "" Then
        BuscarProdutoXML_Geral = Empty
        Exit Function
    End If

    sql = "SELECT IDPRODUTO FROM TAB_PRODUTOS WHERE " & whereSql & " LIMIT 1"

    Set rs = Conn.Execute(sql)

    If Not rs.EOF Then
        BuscarProdutoXML_Geral = CLng(rs!IDProduto)
    Else
        BuscarProdutoXML_Geral = Empty
    End If

    rs.Close
    Set rs = Nothing

End Function

Private Sub CarregarProdutosEntrada(ByVal NumeroNFE_XML As String)

    Dim ws As Worksheet
    Dim tb As ListObject
    Dim i As Long
    Dim item As ListItem
    Dim IDProduto As Variant
    Dim CodFornecedor As String
    Dim CodBarras As String

    Set ws = ThisWorkbook.Worksheets("ENTRADA PRODUTOS")
    Set tb = ws.ListObjects("ENTRADA_PRODUTOS")

    lvProdutosEntrada.ListItems.Clear

    If tb.DataBodyRange Is Nothing Then Exit Sub

    For i = 1 To tb.ListRows.Count

        If NzDbl(tb.DataBodyRange(i, tb.ListColumns("NUMERO NFE").Index).Value) = NzDbl(NumeroNFE_XML) Then

            CodFornecedor = Nz(tb.DataBodyRange(i, tb.ListColumns("CODIGO FORNECEDOR").Index).Value)
            CodBarras = Nz(tb.DataBodyRange(i, tb.ListColumns("CODIGO BARRAS").Index).Value)

            IDProduto = BuscarProdutoXML(CodBarras, CodFornecedor)

            If Not IsEmpty(IDProduto) Then
                If ItemJaLancadoCompra(IdCompra, CLng(IDProduto)) Then GoTo ProximoItem
            End If

            If IsEmpty(IDProduto) Then
                Set item = lvProdutosEntrada.ListItems.Add(, , "SEM CADASTRO")
            Else
                Set item = lvProdutosEntrada.ListItems.Add(, , CStr(IDProduto))
            End If

            item.SubItems(1) = Nz(tb.DataBodyRange(i, tb.ListColumns("NOME").Index).Value)
            item.SubItems(2) = Nz(tb.DataBodyRange(i, tb.ListColumns("MEDIDA COMPRA").Index).Value)
            item.SubItems(3) = NzDbl(tb.DataBodyRange(i, tb.ListColumns("QUANTIDADE").Index).Value)
            item.SubItems(4) = Format(NzDbl(tb.DataBodyRange(i, tb.ListColumns("VALOR UNITARIO").Index).Value), "0.00")
            item.SubItems(5) = Format(NzDbl(tb.DataBodyRange(i, tb.ListColumns("VALOR TOTAL").Index).Value), "0.00")

            item.Tag = CodFornecedor & "|" & CodBarras

        End If

ProximoItem:

    Next i

    If lvProdutosEntrada.ListItems.Count = 0 Then
        MsgBox "Todos os produtos desta compra já foram lançados.", vbInformation
        Unload Me
    End If

End Sub
Private Function BuscarProdutoXML(ByVal CodigoBarras As String, ByVal CodigoFornecedor As String) As Variant

    BuscarProdutoXML = BuscarProdutoXML_Geral(CodigoBarras, CodigoFornecedor)

End Function
Private Sub btnNovoProd_Click()

Dim item As ListItem
Dim dados As Variant

If lvProdutosEntrada.SelectedItem Is Nothing Then Exit Sub

Set item = lvProdutosEntrada.SelectedItem

If item.Text <> "SEM CADASTRO" Then
    MsgBox "Produto já cadastrado.", vbInformation
    Exit Sub
End If

dados = Split(item.Tag, "|")

With frmCadProduto

    .OrigemXML = True

    .txtCodFornecedor.Value = dados(0)
    .txtEAN.Value = dados(1)

    .txtNome.Value = item.SubItems(1)
    .cmbUMcompra.Value = item.SubItems(2)
    .txtQtdeCompra.Value = item.SubItems(3)
    .txtCusto.Value = item.SubItems(4)

End With

frmCadProduto.Show vbModal

End Sub

Public Sub AtualizarProdutoXML(ByVal IDProduto As Long)

Dim item As ListItem

If lvProdutosEntrada.SelectedItem Is Nothing Then Exit Sub

Set item = lvProdutosEntrada.SelectedItem

item.Text = IDProduto

End Sub
