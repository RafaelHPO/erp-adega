VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmItensCompra 
   Caption         =   "ENTRADA DE PRODUTOS"
   ClientHeight    =   9600.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15405
   OleObjectBlob   =   "frmItensCompra.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmItensCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TipoEntrada As String
Public NumeroNFE_XML As String

Private Sub AplicarPaleta()

    AplicarTema Me
    AplicarTemaLv lvProdutos
    AplicarBotaoPrincipal btnAdcItem
    AplicarBotaoPrincipal btnExcluirItem
    AplicarBotaoPrincipal btnCancelar
    AplicarBotaoPrincipal btnFecharEntrada
    
End Sub

Private Sub UserForm_initialize()

On Error GoTo TratarErro

    With lvProdutos

        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .HideSelection = False

        .ColumnHeaders.Clear

        .ColumnHeaders.Add , , "ID", 40
        .ColumnHeaders.Add , , "PRODUTO", 200
        .ColumnHeaders.Add , , "U.M.", 50
        .ColumnHeaders.Add , , "QTDE", 50
        .ColumnHeaders.Add , , "CUSTO", 50
        .ColumnHeaders.Add , , "SUBTOTAL", 80
        .ColumnHeaders.Add , , "LUCRO", 60
        .ColumnHeaders.Add , , "STATUS", 100

    End With

    CarregarProdutosCombo cmbProduto
    CalcularSubtotal
    AplicarPaleta
    
        Exit Sub

TratarErro:

    modSistema.tela = "frmItensCompra - initialize"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"
    
End Sub

Private Sub userform_activate()

On Error GoTo TratarErro

    If TipoEntrada = "XML" Then
        
        Dim rs As ADODB.Recordset
        
        Set rs = Conn.Execute( _
        "SELECT COUNT(*) AS QTD FROM TAB_ITENSCOMPRA WHERE IDCOMPRA = " & SqlNumero(txtIdCompra.Value))
        
        If rs("QTD").Value > 0 Then
            CarregarItensCompra
        End If
        
        rs.Close
        Set rs = Nothing
        
    Else
        
        CarregarItensCompra
        
    End If


    CarregarProdutosCombo cmbProduto
    CarregarUnidadesMedida cmbUM

    Exit Sub

TratarErro:

    modSistema.tela = "frmItensCompra - activate"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"

End Sub


Private Sub txtCusto_Change()

FormatarMoeda txtCusto

End Sub

Private Sub txtVencimento_Change()

    FormatarData txtVencimento

End Sub

'====================================================
' sub total da compra
'====================================================
Private Sub CalcularSubtotal()

On Error GoTo TratarErro

    Dim i As Long
    Dim ValorTotal As Double
    Dim rs As ADODB.Recordset
    Dim sql As String


    If frmCompra.TipoEntrada = "XML" Then

        ValorTotal = 0

        For i = 1 To lvProdutos.ListItems.Count
            
            ValorTotal = ValorTotal + CDbl(NzDbl(lvProdutos.ListItems(i).SubItems(6)))
            
        Next i

        txtSubtotal.Value = Format(ValorTotal, "0.00")

        Exit Sub

    End If


    If Trim(txtIdCompra.Value) = "" Then Exit Sub


    sql = "SELECT COALESCE(SUM(SUBTOTAL),0) AS TOTAL " & _
          "FROM TAB_ITENSCOMPRA " & _
          "WHERE IDCOMPRA = " & CLng(txtIdCompra.Value)


    Set rs = New ADODB.Recordset
    rs.Open sql, Conn, adOpenStatic, adLockReadOnly


    If Not rs.EOF Then
        txtSubtotal.Value = Format(CDbl(rs.Fields("TOTAL").Value), "0.00")
    Else
        txtSubtotal.Value = "0.00"
    End If


    rs.Close
    Set rs = Nothing

    Exit Sub

TratarErro:

    modSistema.tela = "frmItenscompra - calcularsubtotal"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"

End Sub

Private Sub CarregarItensXML(ByVal NumeroNF As String)

On Error GoTo TratarErro

    Dim ws As Worksheet
    Dim tb As ListObject
    Dim i As Long
    Dim ITEM As ListItem
    Dim PRODUTO As Variant
    
    Dim CodigoBarras As String
    Dim CodigoFornecedor As String


    Set ws = ThisWorkbook.Worksheets("XML_PRODUTOS")
    Set tb = ws.ListObjects("XML_PRODUTOS")

    lvProdutos.ListItems.Clear


    If tb.DataBodyRange Is Nothing Then Exit Sub


    For i = 1 To tb.ListRows.Count
        
        If NzDbl(tb.DataBodyRange(i, tb.ListColumns("NUMERO NFE").Index).Value) = NzDbl(NumeroNF) Then


            CodigoBarras = Nz(tb.DataBodyRange(i, tb.ListColumns("CODIGO BARRAS").Index).Value)
            CodigoFornecedor = Nz(tb.DataBodyRange(i, tb.ListColumns("CODIGO FORNECEDOR").Index).Value)


            PRODUTO = BuscarProdutoXML(CodigoBarras, CodigoFornecedor)


            If PRODUTO(0) <> "" Then
                
                Set ITEM = lvProdutos.ListItems.Add(, , PRODUTO(0))
                ITEM.SubItems(1) = PRODUTO(1)
                ITEM.SubItems(7) = "CADASTRADO"
                
            Else
                
                Set ITEM = lvProdutos.ListItems.Add(, , "")
                ITEM.SubItems(1) = Nz(tb.DataBodyRange(i, tb.ListColumns("NOME").Index).Value)
                ITEM.SubItems(7) = "SEM CADASTRO"
                
            End If


            ITEM.SubItems(2) = Nz(tb.DataBodyRange(i, tb.ListColumns("MEDIDA COMPRA").Index).Value)
            ITEM.SubItems(3) = NzDbl(tb.DataBodyRange(i, tb.ListColumns("QUANTIDADE").Index).Value)
            ITEM.SubItems(4) = Format(NzDbl(tb.DataBodyRange(i, tb.ListColumns("VALOR UNITARIO").Index).Value), "0.00")
            ITEM.SubItems(5) = Format(NzDbl(tb.DataBodyRange(i, tb.ListColumns("VALOR TOTAL").Index).Value), "0.00")
            ITEM.SubItems(6) = "0.00"


        End If

    Next i


    CalcularSubtotal

    Exit Sub

TratarErro:

    modSistema.tela = "frmItensCompra - Carregar Itens xml"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"


End Sub

'====================================================
' list view itens compra
'====================================================
Private Sub CarregarItensCompra()

On Error GoTo TratarErro

    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim ITEM As ListItem

    If Trim(txtIdCompra.Value) = "" Then Exit Sub

sql = "SELECT IC.IDPRODUTO, " & _
      "P.NOME, " & _
      "IC.QUANTIDADE, " & _
      "IC.CUSTOUNITARIO, " & _
      "IC.MEDIDACOMPRA, " & _
      "IC.SUBTOTAL, " & _
      "((P.PRECOVENDA - P.CUSTOMEDIO) * IC.QUANTIDADE) AS LUCROESTIMADO " & _
      "FROM TAB_ITENSCOMPRA IC " & _
      "INNER JOIN TAB_PRODUTOS P ON P.IDPRODUTO = IC.IDPRODUTO " & _
      "WHERE IC.IDCOMPRA = " & CLng(txtIdCompra.Value)

    Set rs = New ADODB.Recordset
    rs.Open sql, Conn, adOpenStatic, adLockReadOnly

    lvProdutos.ListItems.Clear

    Do While Not rs.EOF

        Set ITEM = lvProdutos.ListItems.Add(, , CStr(rs!idproduto))

        ITEM.SubItems(1) = IIf(IsNull(rs!NOME), "", rs!NOME)
        ITEM.SubItems(2) = IIf(IsNull(rs!MEDIDACOMPRA), 0, rs!MEDIDACOMPRA)
        ITEM.SubItems(3) = IIf(IsNull(rs!QUANTIDADE), 0, rs!QUANTIDADE)
        ITEM.SubItems(4) = Format(IIf(IsNull(rs!CUSTOUNITARIO), 0, rs!CUSTOUNITARIO), "0.00")
        ITEM.SubItems(5) = Format(IIf(IsNull(rs!subtotal), 0, rs!subtotal), "0.00")

        rs.MoveNext

    Loop

    rs.Close
    Set rs = Nothing

    DoEvents

    CalcularSubtotal
    
        Exit Sub

TratarErro:

    modSistema.tela = "frmItensCompra - Carregar lv "
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"
    
End Sub

Private Sub txtEAN5_AfterUpdate()

    SelecionarProdutoPorEAN cmbProduto, txtEAN5.Value
    cmbProduto.SetFocus
    
End Sub
Private Sub cmbProduto_afterupdate()

On Error GoTo TratarErro

    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim idproduto As Long

    If cmbProduto.ListIndex = -1 Then Exit Sub

    idproduto = CLng(cmbProduto.Value)

    sql = "SELECT CUSTOUNITARIO, codigofornecedor, codigobarrascx, medidacompra, quantidadecompra, codigobarras, quantidadeembalagem, tipo " & _
          "FROM TAB_PRODUTOS " & _
          "WHERE IDPRODUTO = " & idproduto

    Set rs = Conn.Execute(sql)

    If Not rs.EOF Then

        '====================================================
        ' VALIDA SE É COMBO
        '====================================================
        If UCase(Nz(rs!tipo, "")) = "COMBO" Then

            MsgBox "Este produto é um COMBO." & vbCrLf & _
                   "Selecione um produto individual para continuar.", _
                   vbExclamation, "Produto inválido"

            cmbProduto.Value = Null
            cmbProduto.SetFocus

            rs.Close
            Set rs = Nothing
            Exit Sub

        End If

        txtEAN.Value = Nz(rs!CodigoBarras)
        txtEANCX.Value = Nz(rs!CODIGOBARRASCX)
        txtCodFornecedor = Nz(rs!CodigoFornecedor)
        cmbUM = Nz(rs!MEDIDACOMPRA)
        txtQtdeCompra = Nz(rs!QUANTIDADECOMPRA)
        txtQtdeEmb = Nz(rs!QUANTIDADEEMBALAGEM)

    End If

    rs.Close
    Set rs = Nothing
    
        txtQuantidade.SetFocus
    
        Exit Sub

TratarErro:

    modSistema.tela = "frmItensCompra - cmbProduto change"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"
    
End Sub

Private Sub btnAdcItem_Click()

On Error GoTo TratarErro

    Dim sql As String
    Dim rs As ADODB.Recordset

    sql = "SELECT STATUS FROM TAB_COMPRAS WHERE IDCOMPRA = " & txtIdCompra

    Set rs = Conn.Execute(sql)

  If UCase(Trim(Nz(rs!STATUS, ""))) <> "ABERTO" Then

        MsgBox "ENTRADA CANCELADA OU CONCLUIDA"

        rs.Close
        Set rs = Nothing

        Exit Sub

    End If

    rs.Close
    Set rs = Nothing

    If Trim(txtIdCompra.Value & "") = "" Then
        MsgBox "Nenhuma compra aberta."
        Exit Sub
    End If

    If cmbProduto.ListIndex = -1 Then
        MsgBox "Selecione um produto."
        Exit Sub
    End If

  If Not IsNumeric(txtQuantidade.Value) Then
    MsgBox "Quantidade inválida."
    Exit Sub
End If

If CDbl(txtQuantidade.Value) <= 0 Then
    MsgBox "Quantidade inválida."
    Exit Sub
End If

    sql = "CALL PROC_ENTRADAPRODUTOS(" & _
          SqlNumero(txtIdCompra.Value) & "," & _
          SqlNumero(IDUsuarioLogado) & "," & _
          SqlNumero(cmbProduto.Column(0)) & "," & _
          SqlNumero(txtCodFornecedor.Value) & "," & _
          SqlTexto(txtEAN.Value) & "," & _
          SqlTexto(txtEANCX.Value) & "," & _
          SqlTexto(cmbUM.Value) & "," & _
          SqlNumero(txtQtdeCompra.Value) & "," & _
          SqlNumero(txtQtdeEmb.Value) & "," & _
          SqlNumero(txtQuantidade.Value) & "," & _
          SqlTexto(txtObs.Value) & "," & _
          SqlNumero(txtCusto.Value) & ")"

    Set rs = Conn.Execute(sql)

    If Not rs Is Nothing Then

        If Not rs.EOF Then

            MsgBox rs.Fields(0).Value, vbInformation

        End If

        rs.Close

    End If

    Set rs = Nothing

    CarregarItensCompra

    txtQuantidade.Value = ""
    txtQtdeEmb.Value = ""
    txtCusto.Value = ""
    txtEAN.Value = ""
    cmbUM.Value = ""
    txtQtdeCompra.Value = ""
    txtObs.Value = ""
    txtCodFornecedor.Value = ""
    cmbProduto.ListIndex = -1

    cmbProduto.SetFocus

    Exit Sub

TratarErro:

    modSistema.tela = "frmItensCompra - btnAdcItem"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"

End Sub

'==================================
'cadastra novo produto
'================================
Private Sub btnNovoProd_Click()
    
    frmCadProduto.Show vbModeless

End Sub
'==================================
'retira item da entrada e qtde do estoque
'================================
Private Sub btnExcluirItem_Click()

On Error GoTo TratarErro

    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim rsRetorno As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim StatusCompra As String

    If lvProdutos.SelectedItem Is Nothing Then
        MsgBox "Selecione um item para remover.", vbExclamation
        Exit Sub
    End If

    If frmEntradas.lvCompras.SelectedItem Is Nothing Then
        MsgBox "Selecione uma entrada.", vbExclamation
        Exit Sub
    End If

    sql = "SELECT STATUS FROM TAB_COMPRAS WHERE IDCOMPRA = " & SqlNumero(frmEntradas.lvCompras.SelectedItem.Text)

    Set rs = Conn.Execute(sql)

    If Not rs.EOF Then StatusCompra = Nz(rs!STATUS, "")

    rs.Close
    Set rs = Nothing

    If UCase(Trim(StatusCompra)) <> "ABERTO" Then
        MsgBox "ENTRADA CANCELADA OU CONCLUÍDA.", vbExclamation
        Exit Sub
    End If

    Set cmd = New ADODB.Command

    With cmd
        Set .ActiveConnection = Conn
        .CommandType = adCmdStoredProc
        .CommandText = "PROC_RETIRARITENSCOMPRA"

        .Parameters.Append .CreateParameter("P_IDCOMPRA", adInteger, adParamInput, , CLng(txtIdCompra.Value))
        .Parameters.Append .CreateParameter("P_IDUSUARIO", adInteger, adParamInput, , IDUsuarioLogado)
        .Parameters.Append .CreateParameter("P_IDPRODUTO", adInteger, adParamInput, , CLng(lvProdutos.SelectedItem.Text))
        .Parameters.Append .CreateParameter("P_QUANTIDADE", adCurrency, adParamInput, , NzDbl(lvProdutos.SelectedItem.SubItems(3)))

        Set rsRetorno = .Execute
    End With

    If Not rsRetorno Is Nothing Then

        If Not rsRetorno.EOF Then
            MsgBox rsRetorno.Fields(0).Value, vbInformation
        End If

        rsRetorno.Close
        Set rsRetorno = Nothing

    End If

    Set cmd = Nothing

    CarregarItensCompra

    Exit Sub

TratarErro:

    modSistema.tela = "frmItensCompra - btnExcluir item"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"

End Sub
'==================================
'cancelar todo processo de entrada
'==================================
Private Sub btnCancelar_Click()

On Error GoTo TratarErro

    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim sql As String
    
    'VALIDA SE ENTRADA JA FOI CONCLUIDA OU CANCELADA

        sql = "SELECT STATUS FROM TAB_COMPRAS WHERE IDCOMPRA = " & frmEntradas.lvCompras.SelectedItem
    
    Set rs = Conn.Execute(sql)
    
  If UCase(Trim(Nz(rs!STATUS, ""))) <> "ABERTO" Then
    MsgBox "ENTRADA CANCELADA OU CONCLUIDA"
    
    rs.Close
    Set rs = Nothing

    Exit Sub
    End If
    
    Set cmd = New ADODB.Command

    With cmd

        Set .ActiveConnection = Conn
        .CommandType = adCmdStoredProc
        .CommandText = "PROC_CANCELARENTRADA"

        .Parameters.Append .CreateParameter("P_IDCOMPRA", adInteger, adParamInput, , CLng(txtIdCompra.Value))

        .Parameters.Append .CreateParameter("P_IDUSUARIO", adInteger, adParamInput, , IDUsuarioLogado)

    End With

    Set rs = cmd.Execute

    If Not rs Is Nothing Then

        If rs.State = adStateOpen Then

            If Not rs.EOF Then

                MsgBox rs.Fields(0).Value

            End If

        End If

    End If

    Unload Me

    Exit Sub

TratarErro:

    modSistema.tela = "frmItensCompra - btnCancelar entrada"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"

End Sub
Private Sub btnFecharEntrada_Click()

On Error GoTo TratarErro

    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim rsRetorno As ADODB.Recordset

    Dim StatusPagamento As String
    Dim DataSQL As String

    '==================================
    ' VALIDA STATUS DA COMPRA
    '==================================
    sql = "SELECT STATUS FROM TAB_COMPRAS " & _
          "WHERE IDCOMPRA = " & Nz(frmEntradas.lvCompras.SelectedItem, 0)

    Set rs = Conn.Execute(sql)

    If rs.EOF Then

        MsgBox "Compra nao encontrada.", vbExclamation

        rs.Close
        Set rs = Nothing

        Exit Sub

    End If

  If UCase(Trim(Nz(rs!STATUS, ""))) <> "ABERTO" Then

        MsgBox "ENTRADA CANCELADA OU CONCLUIDA"

        rs.Close
        Set rs = Nothing

        Exit Sub

    End If

    rs.Close
    Set rs = Nothing

    '==================================
    ' VALIDA SE EXISTEM ITENS
    '==================================
    sql = "SELECT 1 " & _
          "FROM TAB_ITENSCOMPRA " & _
          "WHERE IDCOMPRA = " & Nz(txtIdCompra.Value, 0)

    Set rs = Conn.Execute(sql)

    If rs.EOF Then

        MsgBox "Entrada nao possui produtos. Adicione ao menos um item.", vbExclamation

        rs.Close
        Set rs = Nothing

        Exit Sub

    End If

    rs.Close
    Set rs = Nothing

    '==================================
    ' STATUS PAGAMENTO
    '==================================
    If optPago.Value = True Then

        StatusPagamento = "PAGO"

        DataSQL = Format(Date, "yyyy-mm-dd")

    ElseIf optAPagar.Value = True Then

        StatusPagamento = "PENDENTE"

       If Nz(txtVencimento.Value, "") = "" Then

    MsgBox "Informe a data de vencimento.", vbExclamation
    Exit Sub

End If

DataSQL = Format$(NzDate(txtVencimento.Value), "yyyy-mm-dd")

    Else

        MsgBox "Selecione o status do pagamento.", vbExclamation
        Exit Sub

    End If

    '==================================
    ' EXECUTA PROCEDURE
    '==================================
    Set cmd = New ADODB.Command

    With cmd

        Set .ActiveConnection = Conn

        .CommandType = adCmdStoredProc

        .CommandText = "PROC_FECHARENTRADA"

        .Parameters.Append .CreateParameter( _
            "P_IDCOMPRA", _
            adInteger, _
            adParamInput, _
            , CLng(Nz(txtIdCompra.Value, 0)))

        .Parameters.Append .CreateParameter( _
            "P_IDUSUARIO", _
            adInteger, _
            adParamInput, _
            , CLng(Nz(IDUsuarioLogado, 0)))

        .Parameters.Append .CreateParameter( _
            "P_VENCIMENTO", _
            adVarChar, _
            adParamInput, _
            12, _
            Nz(DataSQL, ""))

        .Parameters.Append .CreateParameter( _
            "P_STATUSPAGAMENTO", _
            adVarChar, _
            adParamInput, _
            20, _
            Nz(StatusPagamento, ""))

    End With

    Set rsRetorno = cmd.Execute

    If Not rsRetorno Is Nothing Then

        If rsRetorno.State = adStateOpen Then

            If Not rsRetorno.EOF Then

                MsgBox Nz(rsRetorno.Fields(0).Value, "Operacao concluida.")

            End If

            rsRetorno.Close

        End If

    End If

    Set rsRetorno = Nothing
    Set cmd = Nothing

    Unload Me

    Exit Sub

TratarErro:

    modSistema.tela = "frmItensCOmpra - btnFechar (finalizar)"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"

End Sub

