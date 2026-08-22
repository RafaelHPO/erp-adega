VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEntradas 
   Caption         =   "REGISTRO DE ENTRADAS"
   ClientHeight    =   9615.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15450
   OleObjectBlob   =   "frmEntradas.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEntradas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AplicarPaleta()

    AplicarTema Me
    AplicarBotaoPrincipal btnAbrirCompra
    AplicarBotaoPrincipal btnCancelar
    AplicarTemaLv lvCompras

End Sub

Private Sub UserForm_initialize()

On Error GoTo TratarErro

    With lvCompras

        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = False
        .HideSelection = False

        .ColumnHeaders.Clear

        .ColumnHeaders.Add , , "ID", 60
        .ColumnHeaders.Add , , "FORNECEDOR", 200
        .ColumnHeaders.Add , , "NUMERO NF", 100
        .ColumnHeaders.Add , , "DATA", 160
        .ColumnHeaders.Add , , "VALOR TOTAL", 120
        .ColumnHeaders.Add , , "STATUS", 100

    End With

    AplicarPaleta
    
    Exit Sub

TratarErro:

    modSistema.tela = "frmEntradas - initialize"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"

End Sub
Private Sub userform_activate()

    CarregarCompras

End Sub

Private Sub btnCancelar_Click()

On Error GoTo TratarErro

If Not SolicitarSenhaCaixa() Then Exit Sub

    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim IdCompra As Long

    If lvCompras.SelectedItem Is Nothing Then
        MsgBox "Selecione uma compra.", vbExclamation
        Exit Sub
    End If

    IdCompra = CLng(lvCompras.SelectedItem.Text)

    '====================================================
    ' valida se existe compra
    '====================================================
    sql = "SELECT STATUS FROM TAB_COMPRAS WHERE IDCOMPRA = " & IdCompra
    Set rs = Conn.Execute(sql)

    If rs.EOF Then
        MsgBox "Compra não encontrada.", vbExclamation
        Exit Sub
    End If

    rs.Close
    Set rs = Nothing

    '====================================================
    ' chama procedure (SEM RESTRIÇÃO DE STATUS)
    '====================================================
    Set cmd = New ADODB.Command

    With cmd
        Set .ActiveConnection = Conn
        .CommandType = adCmdStoredProc
        .CommandText = "PROC_CANCELARENTRADA"

        .Parameters.Append .CreateParameter("P_IDCOMPRA", adInteger, adParamInput, , IdCompra)
        .Parameters.Append .CreateParameter("P_IDUSUARIO", adInteger, adParamInput, , IDUsuarioLogado)

    End With

    Set rs = cmd.Execute

    If Not rs Is Nothing Then
        If rs.EOF = False Then
            MsgBox rs.Fields(0).Value
        End If
    End If

    Set rs = Nothing
    Set cmd = Nothing
    
    CarregarCompras

    Exit Sub

TratarErro:

    modSistema.tela = "frmEntradas - btnCancelar"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"

End Sub

Private Sub btnConferencia_Click()

On Error GoTo TratarErro

    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim IdCompraSel As Long
    Dim NumeroNF As String
    Dim StatusCompra As String

    If lvCompras.SelectedItem Is Nothing Then
        MsgBox "Selecione uma compra.", vbExclamation
        Exit Sub
    End If

    IdCompraSel = CLng(lvCompras.SelectedItem.Text)

    sql = "SELECT IDCOMPRA, NUMERONF, STATUS " & _
          "FROM tab_compras " & _
          "WHERE IDCOMPRA = " & SqlNumero(IdCompraSel)

    Set rs = Conn.Execute(sql)

    If rs.EOF Then
        MsgBox "Compra não encontrada.", vbExclamation
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If

    NumeroNF = Nz(rs!NumeroNF, "")
    StatusCompra = UCase(Trim(Nz(rs!STATUS, "")))

    rs.Close
    Set rs = Nothing

    If StatusCompra = "CONCLUIDO" Or StatusCompra = "CONCLUÍDO" Then
        MsgBox "Esta compra já está concluída.", vbInformation
        Exit Sub
    End If

    If StatusCompra = "CANCELADO" Then
        MsgBox "Esta compra está cancelada.", vbExclamation
        Exit Sub
    End If

    If StatusCompra <> "ABERTO" Then
        MsgBox "A conferência só pode ser aberta para compras com status ABERTO.", vbExclamation
        Exit Sub
    End If

    If Not CompraPossuiItensPendentes(IdCompraSel, NumeroNF) Then
        MsgBox "Todos os produtos desta compra já foram lançados.", vbInformation
        Exit Sub
    End If

    With frmConferencia
        .IdCompra = IdCompraSel
        .NumeroNFE_XML = NumeroNF
        .TipoEntrada = "XML"
        .Show vbModal
    End With

    Exit Sub

TratarErro:

    modSistema.tela = "frmEntradas - btnConferencia"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"

End Sub

Private Function CompraPossuiItensPendentes(ByVal IdCompra As Long, ByVal NumeroNF As String) As Boolean

    Dim ws As Worksheet
    Dim tb As ListObject
    Dim i As Long
    Dim CodFornecedor As String
    Dim CodBarras As String
    Dim idproduto As Variant

    Set ws = ThisWorkbook.Worksheets("ENTRADA PRODUTOS")
    Set tb = ws.ListObjects("ENTRADA_PRODUTOS")

    If tb.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To tb.ListRows.Count

        If NzDbl(tb.DataBodyRange(i, tb.ListColumns("NUMERO NFE").Index).Value) = NzDbl(NumeroNF) Then

            CodFornecedor = Nz(tb.DataBodyRange(i, tb.ListColumns("CODIGO FORNECEDOR").Index).Value)
            CodBarras = Nz(tb.DataBodyRange(i, tb.ListColumns("CODIGO BARRAS").Index).Value)

            idproduto = BuscarProdutoXML_Geral(CodBarras, CodFornecedor)

            If IsEmpty(idproduto) Then
                CompraPossuiItensPendentes = True
                Exit Function
            End If

            If Not ItemJaLancadoCompra(IdCompra, CLng(idproduto)) Then
                CompraPossuiItensPendentes = True
                Exit Function
            End If

        End If

    Next i

End Function

Public Function BuscarProdutoXML_Geral(ByVal CodigoBarras As String, ByVal CodigoFornecedor As String) As Variant

On Error GoTo TratarErro

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
        BuscarProdutoXML_Geral = CLng(rs!idproduto)
    Else
        BuscarProdutoXML_Geral = Empty
    End If

    rs.Close
    Set rs = Nothing

    Exit Function

TratarErro:

    modSistema.tela = "frmEntradas - Funcao buscarprodutoxml"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"

End Function

Public Function ItemJaLancadoCompra(ByVal IdCompra As Long, ByVal idproduto As Long) As Boolean

    Dim rs As ADODB.Recordset
    Dim sql As String

    sql = "SELECT COUNT(*) AS QTDE " & _
          "FROM tab_itenscompra " & _
          "WHERE IDCOMPRA = " & SqlNumero(IdCompra) & _
          " AND IDPRODUTO = " & SqlNumero(idproduto)

    Set rs = Conn.Execute(sql)

    ItemJaLancadoCompra = NzDbl(rs!qtde) > 0

    rs.Close
    Set rs = Nothing

End Function

Private Sub lvCompras_dblclick()

    btnEditarCompra_Click

End Sub


Private Sub CarregarCompras()

On Error GoTo TratarErro

    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim ITEM As ListItem
    
sql = "SELECT C.IDCOMPRA, F.NOME, " & _
      "C.NUMERONF, C.DATACOMPRA, C.VALORTOTAL, C.STATUS " & _
      "FROM TAB_COMPRAS C " & _
      "INNER JOIN TAB_FORNECEDORES F" & _
      " ON F.IDFORNECEDOR = C.IDFORNECEDOR " & _
      "ORDER BY IDCOMPRA DESC "

Set rs = New ADODB.Recordset
rs.Open sql, Conn, adOpenStatic, adLockReadOnly
    
lvCompras.ListItems.Clear
    
    Do While Not rs.EOF
    
        Set ITEM = lvCompras.ListItems.Add(, , Nz(rs!IdCompra))
        
        ITEM.SubItems(1) = Nz(rs!NOME)
        ITEM.SubItems(2) = Nz(rs!NumeroNF)
        ITEM.SubItems(3) = Format(Nz(rs!datacompra), "dd/mm/yyyy hh:mm")
        ITEM.SubItems(4) = Format(Nz(rs!ValorTotal, 0), "R$    #,##0.00")
        ITEM.SubItems(5) = Nz(rs!STATUS)
        
        rs.MoveNext
    Loop

    rs.Close
Set rs = Nothing
    
        Exit Sub

TratarErro:

    modSistema.tela = "frmEntradas - carregar lv"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"
    
End Sub

Private Sub btnAbrirCompra_Click()

    frmCompra.Show vbModeless

End Sub

Private Sub btnEditarCompra_Click()
    
    If lvCompras.SelectedItem Is Nothing Then
    MsgBox "selecione um Pedido de Compra para Editar", vbExclamation
    Exit Sub
    End If
    
frmItensCompra.txtIdCompra = lvCompras.SelectedItem
frmItensCompra.Show vbModeless

End Sub
