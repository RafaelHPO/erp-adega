VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmVendas 
   Caption         =   "Registro de Vendas"
   ClientHeight    =   11610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19755
   ' OleObjectBlob removido na versao publica
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmVendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private FiltroDataIni As Date
Private FiltroDataFim As Date
Private filtroStatus As String
Private FiltroUsuario As Long
Private UsarFiltroUsuario As Boolean

Private Sub btnAbrirCaixa_Click()

fraCaixa.Visible = False

If Not SolicitarSenhaCaixa() Then Exit Sub
    '====================================================
    ' VALIDA SE JÁ EXISTE CAIXA ABERTO
    '====================================================
    If ExisteCaixaAberto Then

        MsgBox "Já existe um caixa aberto.", vbExclamation
        Exit Sub

    End If

    '====================================================
    ' DECLARAÇÃO DAS VARIÁVEIS
    '====================================================
    Dim rs As ADODB.Recordset
    Dim sql As String

    Dim Fundo As Double
    Dim Retorno As String

    '====================================================
    ' SOLICITA O FUNDO INICIAL
    '====================================================
    Retorno = InputBox("Informe o fundo inicial do caixa:")

    If Len(Trim(Retorno)) = 0 Then Exit Sub

    If Not IsNumeric(Retorno) Then

        MsgBox "Informe um valor numérico válido.", vbExclamation
        Exit Sub

    End If

    Fundo = CDbl(Retorno)

    If Fundo < 0 Then

        MsgBox "O fundo inicial não pode ser negativo.", vbExclamation
        Exit Sub

    End If

    '====================================================
    ' EXECUTA PROCEDURE
    '====================================================
    sql = "CALL PROC_ABRIRCAIXA(" & _
          IDUsuarioLogado & "," & _
          Replace(Fundo, ",", ".") & ")"

    Set rs = Conn.Execute(sql)

    If Not rs.EOF Then

        If CampoExiste(rs, "ERRO") Then

            MsgBox rs("ERRO"), vbExclamation

        Else

            frmPrincipal.txtStatusCaixa.Value = "ABERTO"
            frmPrincipal.txtIdCaixa.Value = rs("IDCAIXA")

            MsgBox "Caixa Nº " & rs("IDCAIXA") & _
                   " aberto com fundo de R$ " & _
                   Format(rs("FUNDO"), "0.00"), _
                   vbInformation

        End If

    End If

    rs.Close
    Set rs = Nothing

Call ResumoOperacional

End Sub

Private Sub btnConsulta_Click()

    frmConsulta.Show vbModeless

End Sub

Private Sub btnExibir_Click()
btnOcultar.Visible = True
btnExibir.Visible = False
txtVendido.PasswordChar = ""
txtQtdeVendas.PasswordChar = ""
txtTm.PasswordChar = ""
txtCaixa.PasswordChar = ""
txtQtdAberto.PasswordChar = ""
txtAberto.PasswordChar = ""

End Sub

Private Sub btnOcultar_Click()
btnExibir.Visible = True
btnOcultar.Visible = False
txtVendido.PasswordChar = "*"
txtQtdeVendas.PasswordChar = "*"
txtTm.PasswordChar = "*"
txtCaixa.PasswordChar = "*"
txtQtdAberto.PasswordChar = "*"
txtAberto.PasswordChar = "*"

End Sub

Private Sub btnSangria_Click()
fraCaixa.Visible = False
If Not SolicitarSenhaCaixa() Then Exit Sub

    If Not MostrarResumoCaixa("realizar a sangria") Then Exit Sub

    Dim rs As ADODB.Recordset
    Dim sql As String

    Dim IdCaixa As Long
    Dim valor As Double

    Dim RetornoValor As String

    IdCaixa = ObterCaixaAberto()

    If IdCaixa = 0 Then

        MsgBox "Nenhum caixa aberto.", vbExclamation
        Exit Sub

    End If

    RetornoValor = InputBox("Informe o valor da sangria:")

    If Len(Trim(RetornoValor)) = 0 Then Exit Sub

    If Not IsNumeric(RetornoValor) Then

        MsgBox "Informe um valor numérico válido.", vbExclamation
        Exit Sub

    End If

    valor = CDbl(RetornoValor)

    sql = "CALL PROC_SANGRIACAIXA(" & _
          IdCaixa & "," & _
          Replace(valor, ",", ".") & "," & _
          IDUsuarioLogado & ")"

    Set rs = Conn.Execute(sql)

    If Not rs.EOF Then

    If CampoExiste(rs, "MSG") Then

        If CampoExiste(rs, "VALORRETIRADO") Then

            MsgBox rs!MSG & vbCrLf & _
                   "Valor retirado: R$ " & _
                   Format(NzDbl(rs!VALORRETIRADO), "0.00"), vbInformation

        Else

            MsgBox rs!MSG, vbInformation

        End If

    ElseIf CampoExiste(rs, "RETORNO") Then

        MsgBox rs!Retorno, vbExclamation

    ElseIf CampoExiste(rs, "ERRO") Then

        MsgBox rs!Erro, vbExclamation

    Else

        MsgBox "A operação foi concluída, porém a procedure não retornou uma mensagem.", vbInformation

    End If

End If

    rs.Close
    Set rs = Nothing

Call ResumoOperacional

End Sub

Private Sub btnFecharCaixa_Click()
fraCaixa.Visible = False
If Not SolicitarSenhaCaixa() Then Exit Sub

    If Not MostrarResumoCaixa("fechar o caixa") Then Exit Sub

    '====================================================
    ' DECLARAÇÃO DAS VARIÁVEIS
    '====================================================
    Dim rs As ADODB.Recordset
    Dim sql As String

    Dim IdCaixa As Long

    '====================================================
    ' BUSCA O CAIXA ABERTO
    '====================================================
    IdCaixa = ObterCaixaAberto()

    If IdCaixa = 0 Then

        MsgBox "Nenhum caixa aberto.", vbExclamation
        Exit Sub

    End If

    '====================================================
    ' EXECUTA PROCEDURE
    '====================================================
    sql = "CALL PROC_FECHARCAIXA(" & _
          IdCaixa & "," & _
          IDUsuarioLogado & ")"

    Set rs = Conn.Execute(sql)

    If Not rs.EOF Then

        If CampoExiste(rs, "ERRO") Then

            MsgBox rs("ERRO"), vbExclamation

        Else

            frmPrincipal.txtStatusCaixa.Value = "FECHADO"
            frmPrincipal.txtIdCaixa.Value = ""

            MsgBox rs("MSG") & vbCrLf & _
                   "Total Sangrias: R$ " & _
                   Format(rs("VALOR TOTAL"), "0.00") & vbCrLf & _
                   "Saldo Caixa: R$ " & _
                   Format(rs("SOBRA"), "0.00"), vbInformation

        End If

    End If

    rs.Close
    Set rs = Nothing

Call ResumoOperacional

End Sub

 Private Sub btnCaixa_Click()

 ' Alterna entre mostrar e esconder o frame
    fraCaixa.Visible = Not fraCaixa.Visible
    
End Sub

Sub lvVendas_dblClick()

    Call btnEditarVenda_Click

End Sub

Private Sub btnEditarVenda_Click()
    
    If lvVendas.SelectedItem Is Nothing Then
    MsgBox "selecione uma venda", vbExclamation
    Exit Sub
    End If
    
    frmPedido.txtIdVenda = lvVendas.SelectedItem
    frmPedido.Show vbModeless

End Sub


Private Sub Userform_activate()

fraCaixa.Visible = False

    Call ResumoOperacional
    Call CarregarVendas
    ValidarCaixaOperacional
    
End Sub


Private Sub UserForm_initialize()

    With lvVendas

        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .HideSelection = False

        .ColumnHeaders.Clear

        .ColumnHeaders.Add , , "ID", 50
        .ColumnHeaders.Add , , "DATA", 100
        .ColumnHeaders.Add , , "REFERENCIA", 140
        .ColumnHeaders.Add , , "VALOR TOTAL", 110
        .ColumnHeaders.Add , , "DESCONTO", 110
        .ColumnHeaders.Add , , "VALOR FINAL", 110
        .ColumnHeaders.Add , , "STATUS", 90
        .ColumnHeaders.Add , , "USUÁRIO", 120
        .ColumnHeaders.Add , , "DIAS PENDENTE", 110

    End With

  CarregarCmbData
    CarregarCmbUsuario
    CarregarCmbStatus

    CarregarVendas
    ResumoOperacional

btnExibir.Visible = True
btnOcultar.Visible = False
lblAvisoCaixa.Visible = False

txtQtdAberto.PasswordChar = "*"
txtAberto.PasswordChar = "*"
txtQtdeVendas.PasswordChar = "*"
txtVendido.PasswordChar = "*"
txtTm.PasswordChar = "*"
txtCaixa.PasswordChar = "*"

End Sub
Public Sub CarregarVendas()

    Dim rs As ADODB.Recordset
    Dim sql As String

   sql = "SELECT " & _
      "V.IDVENDA, " & _
      "V.DATAVENDA, " & _
      "UPPER(V.REFERENCIA) AS REFERENCIA, " & _
      "IFNULL(SUM(IV.SUBTOTAL),0) AS SUBTOTAL, " & _
      "V.VALORTOTAL, " & _
      "V.DESCONTO, " & _
      "V.VALORFINAL, " & _
      "V.STATUS, " & _
      "CASE " & _
      "WHEN V.STATUS = 'ABERTO' THEN DATEDIFF(CURDATE(), V.DATAVENDA) " & _
      "ELSE NULL " & _
      "END AS DIASPENDENTE, " & _
      "U.USUARIO " & _
      "FROM TAB_VENDAS V " & _
      "LEFT JOIN TAB_ITENSVENDA IV ON IV.IDVENDA = V.IDVENDA " & _
      "INNER JOIN TAB_USUARIOS U ON U.IDUSUARIO = V.IDUSUARIO " & _
      "WHERE 1=1 "

'========================================
' STATUS + PERÍODO
'========================================

If cmbStatus.Value <> "" And cmbStatus.Value <> "TODOS" Then

    sql = sql & _
        " AND V.STATUS = " & SqlTexto(cmbStatus.Value)

    If cmbStatus.Value <> "ABERTO" Then

        If cmbDataInicial.ListIndex > 0 Then
            sql = sql & _
                " AND V.DATAVENDA >= '" & _
                Format$(NzDate(cmbDataInicial.Value), "yyyy-mm-dd") & "'"
        End If

        If cmbDataFinal.ListIndex > 0 Then
            sql = sql & _
                " AND V.DATAVENDA <= '" & _
                Format$(NzDate(cmbDataFinal.Value), "yyyy-mm-dd") & "'"
        End If

    End If

Else

    sql = sql & " AND ("

    sql = sql & "V.STATUS='ABERTO'"

    sql = sql & " OR (V.STATUS<>'ABERTO'"

    If cmbDataInicial.ListIndex > 0 Then
        sql = sql & _
            " AND V.DATAVENDA >= '" & _
            Format$(NzDate(cmbDataInicial.Value), "yyyy-mm-dd") & "'"
    End If

    If cmbDataFinal.ListIndex > 0 Then
        sql = sql & _
            " AND V.DATAVENDA <= '" & _
            Format$(NzDate(cmbDataFinal.Value), "yyyy-mm-dd") & "'"
    End If

    sql = sql & " AND V.STATUS <> 'CANCELADO'"

    sql = sql & "))"

End If
    '========================================
    ' USUÁRIO
    '========================================
    If cmbUsuario.Value <> "" And cmbUsuario.Value <> "TODOS" Then

        sql = sql & _
        " AND U.USUARIO = " & SqlTexto(cmbUsuario.Value)

    End If

    '========================================
    ' STATUS
    '========================================
    If cmbStatus.Value <> "" And cmbStatus.Value <> "TODOS" Then

        sql = sql & _
        " AND V.STATUS = " & SqlTexto(cmbStatus.Value)

    Else

        sql = sql & _
        " AND V.STATUS <> 'CANCELADO' "

    End If

    '========================================
    ' GROUP BY
    '========================================
    sql = sql & _
          " GROUP BY " & _
          "V.IDVENDA, " & _
          "V.DATAVENDA, " & _
          "V.REFERENCIA, " & _
          "V.VALORTOTAL, " & _
          "V.DESCONTO, " & _
          "V.VALORFINAL, " & _
          "V.STATUS, " & _
          "U.USUARIO "

    '========================================
    ' ORDER BY
    '========================================
    sql = sql & _
          " ORDER BY " & _
          "CASE WHEN V.STATUS = 'ABERTO' THEN 0 ELSE 1 END, " & _
          "V.IDVENDA DESC"

    Set rs = Conn.Execute(sql)

    lvVendas.ListItems.Clear

    Do While Not rs.EOF

        With lvVendas.ListItems.Add(, , Nz(rs!idvenda))

    .SubItems(1) = Format(rs!DATAVENDA, "dd/mm/yyyy")
    .SubItems(2) = Nz(rs!REFERENCIA)
    .SubItems(3) = Format(Nz(rs!subtotal, 0), "R$       #,##0.00")
    .SubItems(4) = Format(Nz(rs!Desconto, 0), "R$       #,##0.00")
    .SubItems(5) = Format(Nz(rs!VALORFINAL, 0), "R$       #,##0.00")
    .SubItems(6) = Nz(rs!STATUS)
    .SubItems(7) = Nz(rs!Usuario)
    .SubItems(8) = IIf(IsNull(rs!DIASPENDENTE), "-", _
                        rs!DIASPENDENTE & " Dias Pendente")

    If Nz(rs!DIASPENDENTE, 0) >= 7 Then

        .ListSubItems(8).ForeColor = vbRed

    ElseIf Nz(rs!DIASPENDENTE, 0) >= 3 Then

        .ListSubItems(8).ForeColor = RGB(255, 140, 0)

    End If

End With

        rs.MoveNext

    Loop

    rs.Close
    Set rs = Nothing

    ResumoOperacional

End Sub
Private Sub btnCancelarVenda_Click()

If Not SolicitarSenhaCaixa() Then Exit Sub

    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim idvenda As Long
    Dim resp As VbMsgBoxResult

    If lvVendas.SelectedItem Is Nothing Then
        MsgBox "Selecione uma venda.", vbExclamation
        Exit Sub
    End If

    idvenda = CLng(lvVendas.SelectedItem.Text)

    resp = MsgBox("Deseja cancelar a venda " & idvenda & "?", _
                  vbYesNo + vbQuestion)

    If resp = vbNo Then Exit Sub

    Set cmd = New ADODB.Command

    With cmd

        Set .ActiveConnection = Conn
        .CommandType = adCmdStoredProc
        .CommandText = "PROC_CANCELARVENDA"

        .Parameters.Append .CreateParameter("P_IDVENDA", adInteger, adParamInput, , idvenda)
        .Parameters.Append .CreateParameter("P_IDUSUARIO", adInteger, adParamInput, , IDUsuarioLogado)

    End With

    Set rs = cmd.Execute

    If Not rs Is Nothing Then
        If Not rs.EOF Then
            MsgBox rs.Fields(0).Value
        End If
    End If

    CarregarVendas

    Set rs = Nothing
    Set cmd = Nothing

ResumoOperacional

End Sub
Private Sub btnAbrirVenda_Click()

    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset

    If UCase(Trim(frmPrincipal.txtStatusCaixa.Value)) <> "ABERTO" Then
        MsgBox "Não existe caixa aberto.", vbExclamation
        Exit Sub
    End If

    Set cmd = New ADODB.Command

    With cmd
        Set .ActiveConnection = Conn
        .CommandType = adCmdStoredProc
        .CommandText = "PROC_ABRIRVENDA"

        .Parameters.Append .CreateParameter("P_IDUSUARIO", _
                                            adInteger, _
                                            adParamInput, _
                                            , IDUsuarioLogado)
    End With

    Set rs = cmd.Execute

    If Not rs.EOF Then
        frmPedido.txtIdVenda.Value = rs.Fields(0).Value

    End If

    rs.Close
    Set rs = Nothing
    
  
    
    frmPedido.Show

End Sub
Public Sub ResumoOperacional()

    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim qtde As Double
    Dim Total As Double
    Dim cond As String
    Dim aberto As String
    Dim qtdAberto As String
    
    cond = FiltroVendas()

    '========================
    ' TOTAL VENDIDO
    '========================
    sql = "SELECT IFNULL(SUM(VALORFINAL),0) AS TOTAL " & _
          "FROM TAB_VENDAS V " & _
          "INNER JOIN TAB_USUARIOS U ON U.IDUSUARIO = V.IDUSUARIO " & _
          cond

    Set rs = Conn.Execute(sql)

    Total = Nz(rs!Total, 0)
    txtVendido.Value = Format(Total, "0.00")

    rs.Close

    '========================
    ' QTDE VENDAS
    '========================
    sql = "SELECT COUNT(*) AS QTDE " & _
          "FROM TAB_VENDAS V " & _
          "INNER JOIN TAB_USUARIOS U ON U.IDUSUARIO = V.IDUSUARIO " & _
          cond

    Set rs = Conn.Execute(sql)

    qtde = Nz(rs!qtde, 0)
    txtQtdeVendas.Value = qtde

    rs.Close

    '========================
    ' TICKET MÉDIO
    '========================
    If qtde > 0 Then
        txtTm.Value = Format(Total / qtde, "0.00")
    Else
        txtTm.Value = "0.00"
    End If

    '========================
    ' CAIXA
    '========================
    sql = "SELECT IFNULL(SUM(VALORCAIXA),0) AS CAIXA " & _
          "FROM TAB_CAIXA " & _
          "WHERE STATUS = 'ABERTO'"

    Set rs = Conn.Execute(sql)

    If Not rs.EOF Then
        txtCaixa.Value = Format(Nz(rs!CAIXA, 0), "0.00")
    Else
        txtCaixa.Value = "0.00"
    End If

    rs.Close
    Set rs = Nothing

    sql = "SELECT SUM(V.TOTALVENDA - V.PAGO) AS TOTAL, " & _
          "COUNT(*) AS QTD " & _
          "FROM (SELECT V.IDVENDA, COALESCE(SUM(IV.SUBTOTAL), 0) AS TOTALVENDA, " & _
          "COALESCE((SELECT SUM(P.VALORPAGO) FROM tab_pagamentos P WHERE P.IDVENDA = V.IDVENDA), 0) AS PAGO " & _
          "FROM tab_vendas V JOIN TAB_ITENSVENDA IV ON IV.IDVENDA = V.IDVENDA " & _
          "WHERE V.STATUS = 'ABERTO' GROUP BY V.IDVENDA ) V "
          
    Set rs = Conn.Execute(sql)

If Not rs.EOF Then
    
    aberto = NzDbl(rs!Total)
    txtAberto = Format(aberto, "R$  #,##0.00")
    
        If aberto >= 1000 Then
        txtAberto.ForeColor = vbRed
        
        ElseIf aberto >= 500 Then
        txtAberto.ForeColor = RGB(255, 140, 0)
                
        Else: txtAberto.ForeColor = vbBlack
        
        End If
        
    qtdAberto = Nz(rs!qtd)
    txtQtdAberto = CLng(qtdAberto)

        If qtdAberto >= 15 Then
        txtQtdAberto.ForeColor = vbRed
        
        ElseIf qtdAberto >= 5 Then
        txtQtdAberto.ForeColor = RGB(255, 140, 0)
        
        Else: txtQtdAberto.ForeColor = vbBlack
        
        End If
        
    rs.Close
    Set rs = Nothing
    
    End If
    
End Sub
Private Function FiltroVendas() As String

    Dim cond As String

    cond = " WHERE 1=1 "

    '========================
    ' USUÁRIO
    '========================
    If cmbUsuario.Value <> "" And cmbUsuario.Value <> "TODOS" Then

        cond = cond & _
            " AND U.USUARIO = " & SqlTexto(cmbUsuario.Value)

    End If

    '========================
    ' STATUS + PERÍODO
    '========================
    If cmbStatus.Value <> "" And cmbStatus.Value <> "TODOS" Then

        cond = cond & _
            " AND V.STATUS = " & SqlTexto(cmbStatus.Value)

        If cmbStatus.Value <> "ABERTO" Then

            If cmbDataInicial.ListIndex > 0 Then
                cond = cond & _
                    " AND V.DATAVENDA >= '" & _
                    Format$(NzDate(cmbDataInicial.Value), "yyyy-mm-dd") & "'"
            End If

            If cmbDataFinal.ListIndex > 0 Then
                cond = cond & _
                    " AND V.DATAVENDA <= '" & _
                    Format$(NzDate(cmbDataFinal.Value), "yyyy-mm-dd") & "'"
            End If

        End If

    Else

        cond = cond & _
            " AND (V.STATUS='ABERTO'"

        cond = cond & _
            " OR (V.STATUS<>'ABERTO'"

        If cmbDataInicial.ListIndex > 0 Then
            cond = cond & _
                " AND V.DATAVENDA >= '" & _
                Format$(NzDate(cmbDataInicial.Value), "yyyy-mm-dd") & "'"
        End If

        If cmbDataFinal.ListIndex > 0 Then
            cond = cond & _
                " AND V.DATAVENDA <= '" & _
                Format$(NzDate(cmbDataFinal.Value), "yyyy-mm-dd") & "'"
        End If

        cond = cond & _
            " AND V.STATUS <> 'CANCELADO'))"

    End If

    FiltroVendas = cond

End Function
Private Sub CarregarCmbData()

    Dim rs As ADODB.Recordset
    Dim sql As String

    cmbDataInicial.Clear
    cmbDataFinal.Clear

    cmbDataInicial.AddItem "TODOS"
    cmbDataFinal.AddItem "TODOS"

    sql = "SELECT DISTINCT DATE(DATAVENDA) AS DATA " & _
          "FROM TAB_VENDAS " & _
          "ORDER BY DATA DESC"

    Set rs = Conn.Execute(sql)

    Do While Not rs.EOF

        cmbDataInicial.AddItem Format(rs!Data, "dd/mm/yyyy")
        cmbDataFinal.AddItem Format(rs!Data, "dd/mm/yyyy")

        rs.MoveNext

    Loop

    rs.Close
    Set rs = Nothing

   cmbDataInicial.ListIndex = 1
cmbDataFinal.ListIndex = 1

End Sub

Private Sub CarregarCmbUsuario()

    Dim rs As ADODB.Recordset
    Dim sql As String

    cmbUsuario.Clear
    cmbUsuario.AddItem "TODOS"

    sql = "SELECT DISTINCT USUARIO FROM TAB_USUARIOS ORDER BY USUARIO"

    Set rs = Conn.Execute(sql)

    Do While Not rs.EOF

        cmbUsuario.AddItem rs!Usuario

        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing

    cmbUsuario.Value = "TODOS"

End Sub

Private Sub CarregarCmbStatus()

    cmbStatus.Clear

    cmbStatus.AddItem "TODOS"
    cmbStatus.AddItem "ABERTO"
    cmbStatus.AddItem "CONCLUIDO"
    cmbStatus.AddItem "PENDENTE"
    cmbStatus.AddItem "CANCELADO"

    cmbStatus.Value = "TODOS"

End Sub

Private Sub cmbDataInicial_Change()
FormatarDatacmb cmbDataInicial
    ValidarPeriodo
    CarregarVendas
    ResumoOperacional

End Sub

Private Sub cmbDataFinal_Change()
FormatarDatacmb cmbDataFinal
    ValidarPeriodo
    CarregarVendas
    ResumoOperacional

End Sub

Private Sub cmbUsuario_Change()
    CarregarVendas
    ResumoOperacional
End Sub

Private Sub cmbStatus_Change()
    CarregarVendas
    ResumoOperacional
End Sub

Private Sub ValidarPeriodo()

    If cmbDataInicial.ListIndex > 0 _
    And cmbDataFinal.ListIndex > 0 Then

        If NzDate(cmbDataInicial.Value) > NzDate(cmbDataFinal.Value) Then

            MsgBox "A data inicial não pode ser maior que a data final.", vbExclamation

            cmbDataFinal.ListIndex = 0

        End If

    End If

End Sub


