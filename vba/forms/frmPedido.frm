VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPedido 
   Caption         =   "Pedido"
   ClientHeight    =   8790.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15810
   OleObjectBlob   =   "frmPedido.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private abrirPagamento As Boolean
Private idVendaPagamento As Long
Private bloqueioUI As Boolean
Public editandoitem As Boolean

Private Function VendaAberta() As Boolean

    Dim rs As ADODB.Recordset
    Dim sql As String

    VendaAberta = False

    If Trim(txtIdVenda.Value) = "" Then Exit Function

    sql = "SELECT STATUS " & _
          "FROM TAB_VENDAS " & _
          "WHERE IDVENDA = " & CLng(txtIdVenda.Value)

    Set rs = Conn.Execute(sql)

    If Not rs.EOF Then

        If UCase(rs!STATUS) = "ABERTO" Then

            VendaAberta = True
        End If

    End If

    rs.Close
    Set rs = Nothing

End Function

Private Sub btnEditarItem_Click()

    On Error GoTo TratarErro

    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim idproduto As Long
    Dim idvenda As Long
    
    editandoitem = True
    
    idproduto = lvProdutos.SelectedItem.Text
    idvenda = txtIdVenda.Value
  
        cmbProduto = idproduto
        
            Call cmbProduto_afterupdate
    
    sql = " select sum(iv.quantidade) as qtde from tab_itensvenda iv " & _
          " where iv.idproduto = " & idproduto & _
          " and iv.idvenda = " & idvenda
          
    Set rs = Conn.Execute(sql)
    
        If Not rs.EOF Then
    
            txtQuantidade.Value = Nz(rs!qtde)
            
        rs.Close
        Set rs = Nothing
        
    End If

    sql = " call PROC_EDITARITEMVENDA(" & IDUsuarioLogado & ", " & idproduto & ", " & idvenda & ") "
    
    Conn.Execute sql
    
    CarregarItensVenda
    
    Exit Sub

TratarErro:

    modSistema.tela = "frmPedido - btnEditarItem"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number
    
    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"
    
End Sub

Private Sub lvprodutos_dblclick()

  Call btnEditarItem_Click
    
End Sub

Private Sub btnPagamento_Click()

On Error GoTo TratarErro

    If Not VendaAberta Then
        MsgBox "VENDA CANCELADA OU CONCLUIDA", vbExclamation
        Exit Sub
    End If

    Dim ID As Long
    ID = CLng(txtIdVenda.Value)

    Set frmPagamento = New frmPagamento

    frmPagamento.idvenda = ID
    frmPagamento.Show vbModal

    CarregarItensVenda
    
    Exit Sub

TratarErro:

    modSistema.tela = "frmPedido - btnPagamento"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"
                        
End Sub

Private Sub btnPendente_Click()

On Error GoTo TratarErro

    If Not VendaAberta Then
        MsgBox "VENDA CANCELADA OU CONCLUIDA", vbExclamation
        Exit Sub
    End If

   Dim sql As String
    Dim REF As String

REF = CStr(txtReferencia)

If Trim(REF) <> "" Then
    sql = "UPDATE TAB_VENDAS SET REFERENCIA = UPPER('" & Replace(REF, "'", "''") & _
          "') WHERE IDVENDA = " & txtIdVenda
    Conn.Execute sql
End If

  Unload Me

    Exit Sub

TratarErro:

    modSistema.tela = "frmPedido - btnPendente"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"

End Sub

Private Sub txtDesconto_Change()

FormatarMoeda txtDesconto

End Sub

Private Sub txtquantidade_afterupdate()

    btnAdcItem.SetFocus
    
End Sub

Private Sub txtsubtotal_activate()

 btnAdcItem.SetFocus

End Sub


Private Sub UserForm_Initialize()

On Error GoTo TratarErro

    With lvProdutos

        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .HideSelection = False
        .AllowColumnReorder = True

        .ColumnHeaders.Clear

        .ColumnHeaders.Add , , "ID", 60
        .ColumnHeaders.Add , , "PRODUTO", 220
        .ColumnHeaders.Add , , "QTDE", 60
        .ColumnHeaders.Add , , "PREÇO", 80
        .ColumnHeaders.Add , , "SUBTOTAL", 90

    End With

    CarregarProdutosCombo cmbProduto
    CarregarReferenciaVenda
    txtEAN5.SetFocus
    
        Exit Sub

TratarErro:

    modSistema.tela = "frmPedido - initialize"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"
    
End Sub
Private Sub userform_activate()

    On Error GoTo TratarErro

    If Trim(txtIdVenda.Value) <> "" Then

        CarregarItensVenda
CarregarReferenciaVenda
    End If
    
      If abrirPagamento Then

        abrirPagamento = False

        Set frmPagamento = New frmPagamento
        frmPagamento.idvenda = idVendaPagamento
        frmPagamento.Show vbModeless

    End If
 txtEAN5.SetFocus
 
     Exit Sub

TratarErro:

    modSistema.tela = "frmPedido - activate"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"
 
End Sub

Private Sub CarregarReferenciaVenda()

    Dim rs As ADODB.Recordset
    Dim sql As String

    If Trim(txtIdVenda.Value) = "" Then Exit Sub

    sql = "SELECT REFERENCIA FROM TAB_VENDAS WHERE IDVENDA = " & CLng(txtIdVenda.Value)

    Set rs = Conn.Execute(sql)

    If Not rs.EOF Then
        txtReferencia.Value = Nz(rs!REFERENCIA, "")
    End If

    rs.Close
    Set rs = Nothing

End Sub

Private Sub CalcularSubtotal()

    Dim rs As ADODB.Recordset
    Dim sql As String

    If Trim(txtIdVenda.Value) = "" Then Exit Sub

    sql = "SELECT IFNULL(SUM(SUBTOTAL),0) AS TOTAL " & _
          "FROM TAB_ITENSVENDA " & _
          "WHERE IDVENDA = " & CLng(txtIdVenda.Value)

    Set rs = Conn.Execute(sql)

    If Not rs.EOF Then
        txtSubtotal.Value = Format(NzDbl(rs!Total), "0.00")
    End If

    rs.Close
    Set rs = Nothing

End Sub
Private Sub CarregarItensVenda()

On Error GoTo TratarErro

    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim ITEM As ListItem

    If Trim(txtIdVenda.Value) = "" Then Exit Sub

 sql = "SELECT " & _
      "I.IDPRODUTO, " & _
      "P.NOME, " & _
      "SUM(I.QUANTIDADE) AS QUANTIDADE, " & _
      "I.PRECOUNITARIO, " & _
      "SUM(I.SUBTOTAL) AS SUBTOTAL " & _
      "FROM TAB_ITENSVENDA I " & _
      "INNER JOIN TAB_PRODUTOS P ON P.IDPRODUTO = I.IDPRODUTO " & _
      "INNER JOIN TAB_VENDAS V ON V.IDVENDA = I.IDVENDA " & _
      "WHERE I.IDVENDA = " & CLng(txtIdVenda.Value) & _
      " GROUP BY I.IDPRODUTO, P.NOME, I.PRECOUNITARIO " & _
      " ORDER BY P.NOME ASC"

    Set rs = New ADODB.Recordset
    rs.Open sql, Conn, adOpenStatic, adLockReadOnly

    lvProdutos.ListItems.Clear

If rs.EOF Then

    rs.Close
    Set rs = Nothing

    txtSubtotal.Value = "0.00"

    Exit Sub

End If

    Do While Not rs.EOF

        Set ITEM = lvProdutos.ListItems.Add(, , CStr(rs!idproduto))

        ITEM.SubItems(1) = IIf(IsNull(rs!NOME), "", rs!NOME)
        ITEM.SubItems(2) = IIf(IsNull(rs!QUANTIDADE), 0, rs!QUANTIDADE)
        ITEM.SubItems(3) = Format(IIf(IsNull(rs!PRECOUNITARIO), 0, rs!PRECOUNITARIO), "0.00")
        ITEM.SubItems(4) = Format(IIf(IsNull(rs!subtotal), 0, rs!subtotal), "0.00")

        rs.MoveNext

    Loop

    rs.Close
    Set rs = Nothing

    DoEvents

    CalcularSubtotal

    Exit Sub

TratarErro:

    modSistema.tela = "frmPedido - Carrega itens venda"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"

End Sub

Private Sub txtEAN5_AfterUpdate()

    SelecionarProdutoPorEAN cmbProduto, txtEAN5.Value
    Call cmbProduto_afterupdate
    txtQuantidade.SetFocus
End Sub

Private Sub cmbProduto_afterupdate()

On Error GoTo TratarErro

    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim idproduto As Long

    If cmbProduto.ListIndex = -1 Then Exit Sub

    idproduto = CLng(cmbProduto.List(cmbProduto.ListIndex, 0))

    sql = "SELECT PRECOVENDA, ESTOQUEATUAL, TIPO, CODIGOBARRAS " & _
          "FROM TAB_PRODUTOS " & _
          "WHERE IDPRODUTO = " & idproduto

    Set rs = Conn.Execute(sql)

    If Not rs.EOF Then

        txtPreco.Value = Format(IIf(IsNull(rs!PRECOVENDA), 0, rs!PRECOVENDA), "0.00")
        txtTipo.Value = Nz(rs!tipo)
        txtEstoque.Value = IIf(IsNull(rs!ESTOQUEATUAL), 0, rs!ESTOQUEATUAL)
        txtEAN5.Value = Nz(rs!CodigoBarras)
        
    End If

    rs.Close
    Set rs = Nothing
    
'txtQuantidade.SetFocus

    Exit Sub

TratarErro:

    modSistema.tela = "frmPedido - cmbProduto change"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"

End Sub

Private Sub btnAdcItem_Click()

    editandoitem = False

On Error GoTo TratarErro

    Dim sql As String
    Dim preco As String
    Dim rs As ADODB.Recordset
    Dim tipo As String

If Not VendaAberta Then

    MsgBox "VENDA CANCELADA OU CONCLUIDA", vbExclamation

    Exit Sub

End If

    If txtIdVenda.Value = "" Then
        MsgBox "Nenhuma venda aberta.", vbExclamation
        Exit Sub
    End If

    If cmbProduto.ListIndex = -1 Then
        MsgBox "Selecione um produto.", vbExclamation
        Exit Sub
    End If

    If Val(txtQuantidade.Value) <= 0 Then
        MsgBox "Quantidade inválida.", vbExclamation
        Exit Sub
    End If

    preco = Replace(txtPreco.Value, ",", ".")

   sql = "CALL PROC_ADCITENSVENDA(" & _
      NzDbl(txtIdVenda.Value) & "," & _
      NzDbl(IDUsuarioLogado) & "," & _
      NzDbl(cmbProduto.Column(0)) & "," & _
      NzDbl(txtQuantidade.Value) & "," & _
      SqlTexto(txtTipo.Value) & "," & _
      SqlNumero(txtPreco.Value) & ")"
      
    Set rs = Conn.Execute(sql)

    If Not rs Is Nothing Then

        If Not rs.EOF Then

            MsgBox rs.Fields(0).Value, vbInformation

        End If

        rs.Close

    End If

    Set rs = Nothing

    CarregarItensVenda
    txtQuantidade.Value = ""
    txtPreco.Value = ""
    txtEstoque.Value = ""
    txtTipo.Value = ""
    cmbProduto.ListIndex = -1

    cmbProduto.SetFocus

    Exit Sub

TratarErro:

    modSistema.tela = "frmPedido - btnAdcItem"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"

End Sub
Private Sub btnExcluirItem_Click()

On Error GoTo TratarErro

    Dim cmd As ADODB.Command

If Not VendaAberta Then

    MsgBox "VENDA CANCELADA OU CONCLUIDA", vbExclamation

    Exit Sub

End If

    If lvProdutos.SelectedItem Is Nothing Then

        MsgBox "Selecione um item."

        Exit Sub

    End If

    Set cmd = New ADODB.Command

    With cmd

        Set .ActiveConnection = Conn
        .CommandType = adCmdStoredProc
        .CommandText = "PROC_RETIRARITENSVENDA"

        .Parameters.Append .CreateParameter("P_IDVENDA", adInteger, adParamInput, , CLng(txtIdVenda.Value))
        .Parameters.Append .CreateParameter("P_IDUSUARIO", adInteger, adParamInput, , IDUsuarioLogado)
        .Parameters.Append .CreateParameter("P_IDPRODUTO", adInteger, adParamInput, , CLng(lvProdutos.SelectedItem.Text))
        .Parameters.Append .CreateParameter("P_QUANTIDADE", adInteger, adParamInput, , Val(lvProdutos.SelectedItem.SubItems(2)))

        .Execute

    End With

    DoEvents

    CarregarItensVenda
    
    Exit Sub

TratarErro:

    modSistema.tela = "frmPedido - btnExcluirItem"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"

End Sub
Private Sub btnCancelarVenda_Click()

On Error GoTo TratarErro

If Not SolicitarSenhaCaixa() Then Exit Sub

    Dim cmd As ADODB.Command

If Not VendaAberta Then

    MsgBox "VENDA CANCELADA OU CONCLUIDA", vbExclamation

    Exit Sub

End If

    Set cmd = New ADODB.Command

    With cmd

        Set .ActiveConnection = Conn
        .CommandType = adCmdStoredProc
        .CommandText = "PROC_CANCELARVENDA"

        .Parameters.Append .CreateParameter("P_IDVENDA", adInteger, adParamInput, , CLng(txtIdVenda.Value))
        .Parameters.Append .CreateParameter("P_IDUSUARIO", adInteger, adParamInput, , IDUsuarioLogado)

        .Execute

    End With

    MsgBox "Venda cancelada."

    frmVendas.CarregarVendas
    frmVendas.ResumoOperacional
    
    LimparTelaVenda

    Me.Hide

    Exit Sub

TratarErro:

    modSistema.tela = "frmPedido - btnCancelarVenda"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"

End Sub

Private Sub btnFecharVenda_Click()

On Error GoTo TratarErro

    Dim sql As String
    Dim rs As ADODB.Recordset

    Dim IdCaixa As Long
    Dim totalItens As Double
    Dim totalPago As Double
    Dim pendente As Double

    '====================================================
    ' BUSCA O CAIXA ABERTO
    '====================================================
    IdCaixa = ObterCaixaAberto()

    If IdCaixa = 0 Then
        MsgBox "Nenhum caixa aberto.", vbExclamation
        Exit Sub
    End If

    If Not VendaAberta Then
        MsgBox "VENDA CANCELADA OU CONCLUIDA", vbExclamation
        Exit Sub
    End If

    '========================
    ' VERIFICA ITENS
    '========================
    sql = "SELECT 1 FROM TAB_ITENSVENDA WHERE IDVENDA = " & CLng(txtIdVenda.Value)
    Set rs = Conn.Execute(sql)

    If rs.EOF Then
        MsgBox "Venda não possui produtos, adicionar", vbExclamation
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If

    rs.Close
    Set rs = Nothing

    '========================
    ' TOTAL ITENS
    '========================
    sql = "SELECT IFNULL(SUM(SUBTOTAL),0) AS TOTAL FROM TAB_ITENSVENDA WHERE IDVENDA=" & CLng(txtIdVenda.Value)
    Set rs = Conn.Execute(sql)

    If Not rs.EOF Then
        totalItens = CDbl(NzDbl(rs!Total))
    Else
        totalItens = 0
    End If

    rs.Close
    Set rs = Nothing

    '========================
    ' TOTAL PAGO
    '========================
    sql = "SELECT IFNULL(SUM(VALORPAGO),0) AS TOTAL FROM TAB_PAGAMENTOS WHERE IDVENDA=" & CLng(txtIdVenda.Value)
    Set rs = Conn.Execute(sql)

    If Not rs.EOF Then
        totalPago = CDbl(NzDbl(rs!Total))
    Else
        totalPago = 0
    End If

    rs.Close
    Set rs = Nothing

    '========================
    ' PENDENTE REAL
    '========================
    pendente = totalItens - totalPago

    If pendente > 0.01 Then
        MsgBox "Ainda existe valor pendente: R$ " & Format(pendente, "0.00"), vbExclamation
        Exit Sub
    End If

    '========================
    ' FECHAR VIA PROCEDURE
    '========================
    Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command

    With cmd
        Set .ActiveConnection = Conn
        .CommandType = adCmdStoredProc
        .CommandText = "PROC_FECHARVENDA"

        .Parameters.Append .CreateParameter("P_IDVENDA", adInteger, adParamInput, , CLng(txtIdVenda.Value))
        .Parameters.Append .CreateParameter("P_IDCAIXA", adInteger, adParamInput, , CLng(IdCaixa))
        .Parameters.Append .CreateParameter("P_IDUSUARIO", adInteger, adParamInput, , IDUsuarioLogado)
        .Parameters.Append .CreateParameter("P_DESCONTO", adCurrency, adParamInput, , CDbl(Val(txtDesconto.Value)))

        .Execute
    End With

    MsgBox "Venda finalizada."

    frmVendas.CarregarVendas
    frmVendas.ResumoOperacional

    LimparTelaVenda
    Me.Hide

    Exit Sub

TratarErro:

    modSistema.tela = "frmPedido - FecharVenda"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"

End Sub

Private Sub LimparTelaVenda()

   txtIdVenda.Value = ""
txtQuantidade.Value = ""
txtPreco.Value = ""
txtSubtotal.Value = ""
txtDesconto.Value = ""
txtEstoque.Value = ""

cmbProduto.ListIndex = -1

lvProdutos.ListItems.Clear

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If editandoitem = True Then
        
        MsgBox "Não é possível fechar " & vbCrLf & _
               "Enquanto edita um item!", vbInformation, "AVISO"
        
        Cancel = True
        Exit Sub
        
    End If

    If CloseMode = vbFormControlMenu Then
        Cancel = False
        Exit Sub
    End If

End Sub

