VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCadProduto 
   Caption         =   "CADASTRO DE PRODUTOS"
   ClientHeight    =   8415.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13755
   ' OleObjectBlob removido na versao publica
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCadProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public OrigemXML As Boolean
Public IdProdutoCriado As Long

Private Sub txtPrecoVenda_Change()

FormatarMoeda txtPrecoVenda

End Sub

Private Sub userform_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me
    End If

End Sub

Private Sub UserForm_initialize()

    ckAtivo.Value = True

CarregarUnidadesMedida cmbUM
CarregarUnidadesMedida cmbUMcompra
CarregarSetor cmbSetor
        
        cmbTipo.List = Array( _
        "UNIT", _
        "COMBO" _
        )

 CarregarCategorias cmbCategoria
carregarsubgrupos cmbSubgrupo


lblIdProduto.Visible = False
txtidproduto.Visible = False
btnSalvarEdit.Visible = False
btnCancelarEdit.Visible = False
lblCadastro.Visible = False
txtCadastro.Visible = False
cmbTipo.Visible = False
lblTipo.Visible = False
Label14.Visible = False
txtCustoMedio.Visible = False
' bloco de subgrupo
Label17.Visible = False
Label18.Visible = False
cmbSubgrupo.Visible = False
txtQtdeSub.Visible = False
btnAdd.Visible = False
btnEx.Visible = False
txtMarkup.Locked = True
txtSugerido.Locked = True
        
End Sub

Private Sub Userform_activate()
    
    If Trim(txtidproduto) <> "" Then
    CarregarProduto
      ckAtivo.Value = True
      CarregarGrupos
      
      Label17.Visible = True
                Label18.Visible = True
                cmbSubgrupo.Visible = True
                txtQtdeSub.Visible = True
                btnAdd.Visible = True
                btnEx.Visible = True
      
End If

With lvsubgrupos

    .View = lvwReport
    .Gridlines = True
    .HideSelection = False
    .FullRowSelect = True
    
    .ColumnHeaders.Clear
    
    .ColumnHeaders.Add , , "ID", 30
    .ColumnHeaders.Add , , "SUBGRUPO", 100
    
    End With

End Sub

Private Sub txtCusto_Change()

FormatarMoeda txtCusto

    CalcularPrecoSugerido

End Sub

'Private Sub txtCusto_AfterUpdate()
'
 '   Dim Resposta As VbMsgBoxResult
  '  Dim Custo As Double
'
'    Custo = NzDbl(txtCusto.Value)
'
 '   Resposta = MsgBox( _
  '      "Confirmar custo de R$ " & _
   '     Format(Custo, "#,##0.00") & " ?", _
    '    vbYesNo + vbQuestion, _
     '   "Confirmação de Custo")
'
 '   If Resposta = vbNo Then
'
 '       txtCusto.SetFocus
  '      txtCusto.SelStart = 0
   '     txtCusto.SelLength = Len(txtCusto.Text)
'
 '       Exit Sub
'
 '   End If
'
 '   If Custo >= 2000 Then
''
  '      MsgBox "ATENÇÃO CRÍTICA: custo de R$ " & _
   '            Format(Custo, "#,##0.00") & _
    '           ". Valor de custo muito elevado. Revise o cadastro.", _
     '          vbCritical
'
 '   ElseIf Custo >= 1000 Then
'
 '       MsgBox "ATENÇÃO: custo de R$ " & _
  '             Format(Custo, "#,##0.00") & _
   '            ". Verifique se o valor foi digitado corretamente.", _
    '           vbExclamation
 '   End If
'
'End Sub
Private Sub btnSalvarEdit_Click()

Dim Custo As Double

Custo = NzDbl(txtCusto.Value)

If Custo >= 2000 Then

    MsgBox "Custo acima de R$ 2.000,00. Revise o valor informado.", vbCritical
    Exit Sub

ElseIf Custo >= 1000 Then

    If MsgBox( _
        "ATENÇÃO!" & vbCrLf & vbCrLf & _
        "Custo informado: R$ " & Format(Custo, "#,##0.00") & vbCrLf & _
        "Deseja realmente continuar?", _
        vbYesNo + vbExclamation) = vbNo Then

        Exit Sub

    End If

Else

    If MsgBox( _
        "Confirmar custo de R$ " & Format(Custo, "#,##0.00") & " ?", _
        vbYesNo + vbQuestion) = vbNo Then

        Exit Sub

    End If

End If

    Dim sql As String
    Dim StatusProduto As String

    If ckAtivo.Value Then
        StatusProduto = "ATIVO"
    Else
        StatusProduto = "INATIVO"
    End If

    sql = "UPDATE TAB_PRODUTOS SET " & _
          "CODIGOFORNECEDOR = " & SqlTexto(txtCodFornecedor.Value) & ", " & _
          "NOME = " & SqlTexto(txtNome.Value) & ", " & _
          "CODIGOBARRASCX = " & SqlTexto(txtEANCX.Value) & ", " & _
          "CODIGOBARRAS = " & SqlTexto(txtEAN.Value) & ", " & _
          "MEDIDACOMPRA = " & SqlTexto(cmbUMcompra.Value) & ", " & _
          "QUANTIDADECOMPRA = " & SqlNumero(NzDbl(txtQtdeCompra.Value)) & ", " & _
          "QUANTIDADEEMBALAGEM = " & SqlNumero(NzDbl(txtQtdeEmb.Value)) & ", " & _
          "MEDIDAVENDA = " & SqlTexto(cmbUM.Value) & ", " & _
          "IDCATEGORIA = " & SqlNumero(cmbCategoria.Column(0)) & ", " & _
          "SETOR = " & SqlTexto(cmbSetor.Value) & ", " & _
          "TIPO = " & SqlTexto(cmbTipo.Value) & ", " & _
          "MARCA = " & SqlTexto(txtMarca.Value) & ", " & _
          "CUSTOUNITARIO = " & SqlNumero(NzDbl(txtCusto.Value)) & ", " & _
          "CUSTOMEDIO = " & SqlNumero(NzDbl(txtCustoMedio.Value)) & ", " & _
          "PRECOVENDA = " & SqlNumero(NzDbl(txtPrecoVenda.Value)) & ", " & _
          "STATUS = " & SqlTexto(StatusProduto) & " " & _
          "WHERE IDPRODUTO = " & CLng(txtidproduto.Value)

    Conn.Execute sql

    MsgBox "Produto atualizado com sucesso!", vbInformation

    frmProdutos.CarregarProdutos

    Unload Me

End Sub
Private Sub btnSalvar_Click()

Dim Custo As Double

Custo = NzDbl(txtCusto.Value)

If Custo >= 2000 Then

    MsgBox "Custo acima de R$ 2.000,00. Revise o valor informado.", vbCritical
    Exit Sub

ElseIf Custo >= 1000 Then

    If MsgBox( _
        "ATENÇÃO!" & vbCrLf & vbCrLf & _
        "Custo informado: R$ " & Format(Custo, "#,##0.00") & vbCrLf & _
        "Deseja realmente continuar?", _
        vbYesNo + vbExclamation) = vbNo Then

        Exit Sub

    End If

Else

    If MsgBox( _
        "Confirmar custo de R$ " & Format(Custo, "#,##0.00") & " ?", _
        vbYesNo + vbQuestion) = vbNo Then

        Exit Sub

    End If

End If

    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim StatusProduto As String

    If ckAtivo.Value Then
        StatusProduto = "ATIVO"
    Else
        StatusProduto = "INATIVO"
    End If

    Set cmd = New ADODB.Command

    With cmd

        Set .ActiveConnection = Conn

        .CommandType = adCmdStoredProc
        .CommandText = "PROC_CADASTRARPRODUTO"

        .Parameters.Append .CreateParameter("P_CODFORNECEDOR", adInteger, adParamInput, , NzDB(txtCodFornecedor.Value))

        .Parameters.Append .CreateParameter("P_NOME", adVarChar, adParamInput, 80, NzDB(txtNome.Value))

        .Parameters.Append .CreateParameter("P_CODIGOBARRASCX", adVarChar, adParamInput, 15, NzDB(txtEANCX.Value))

        .Parameters.Append .CreateParameter("P_CODIGOBARRAS", adVarChar, adParamInput, 14, NzDB(txtEAN.Value))

        .Parameters.Append .CreateParameter("P_MEDIDACOMPRA", adVarChar, adParamInput, 5, NzDB(cmbUMcompra.Value))

        .Parameters.Append .CreateParameter("P_QUANTIDADECOMPRA", adInteger, adParamInput, , NzDbl(txtQtdeCompra.Value))

        .Parameters.Append .CreateParameter("P_QUANTIDADEEMBALAGEM", adInteger, adParamInput, , NzDbl(txtQtdeEmb.Value))

        .Parameters.Append .CreateParameter("P_MEDIDAVENDA", adVarChar, adParamInput, 5, NzDB(cmbUM.Value))

        .Parameters.Append .CreateParameter("P_IDCATEGORIA", adInteger, adParamInput, , NzDB(cmbCategoria.Column(0)))

        .Parameters.Append .CreateParameter("P_SETOR", adVarChar, adParamInput, 30, NzDB(cmbSetor.Value))

        .Parameters.Append .CreateParameter("P_MARCA", adVarChar, adParamInput, 40, NzDB(txtMarca.Value))

        .Parameters.Append .CreateParameter("P_CUSTO", adCurrency, adParamInput, , NzDbl(txtCusto.Value))

        .Parameters.Append .CreateParameter("P_PRECO", adCurrency, adParamInput, , NzDbl(txtPrecoVenda.Value))

        .Parameters.Append .CreateParameter("P_STATUS", adVarChar, adParamInput, 10, StatusProduto)

        .Parameters.Append .CreateParameter("P_IDUSUARIO", adInteger, adParamInput, , IDUsuarioLogado)

    End With
Dim IdProdutoCriado As Long

Set rs = cmd.Execute

If Not rs.EOF Then
    IdProdutoCriado = CLng(rs("IDPRODUTO").Value)
End If

rs.Close
Set rs = Nothing
Set cmd = Nothing


If OrigemXML Then

    frmConferencia.AtualizarProdutoXML IdProdutoCriado

Else

    frmProdutos.CarregarProdutos
    MsgBox "Produto cadastrado com sucesso!", vbInformation

End If


Unload Me

End Sub
Public Sub CarregarProduto()

    Dim rs As ADODB.Recordset
    Dim sql As String

    sql = "SELECT CODIGOFORNECEDOR, NOME, CODIGOBARRASCX, CODIGOBARRAS, MEDIDACOMPRA, " & _
          "QUANTIDADECOMPRA,QUANTIDADEEMBALAGEM, MEDIDAVENDA, IDCATEGORIA, SETOR, TIPO, MARCA, CUSTOUNITARIO, CUSTOMEDIO, PRECOVENDA, STATUS, DATACADASTRO " & _
          "FROM TAB_PRODUTOS WHERE IDPRODUTO = " & CLng(txtidproduto)

    Set rs = New ADODB.Recordset
    rs.Open sql, Conn, adOpenStatic, adLockReadOnly

    If Not rs.EOF Then

        txtCodFornecedor.Value = Nz(rs!CodigoFornecedor)
        txtNome.Value = Nz(rs!NOME)
        txtEANCX.Value = Nz(rs!CODIGOBARRASCX)
        txtEAN.Value = Nz(rs!CodigoBarras)
        cmbUM.Value = Nz(rs!MEDIDAVENDA)
        txtQtdeCompra.Value = Nz(rs!QUANTIDADECOMPRA)
        txtQtdeEmb.Value = Nz(rs!QUANTIDADEEMBALAGEM)
        txtCustoMedio.Value = Nz(rs!customedio)
        cmbUMcompra.Value = Nz(rs!MEDIDACOMPRA)
        cmbCategoria.Value = Nz(rs!IdCategoria)
        cmbSetor.Value = Nz(rs!Setor)
        cmbTipo.Value = Nz(rs!tipo)
        txtMarca.Value = Nz(rs!MARCA)
        txtCusto.Value = Nz(rs!CUSTOUNITARIO)
        txtPrecoVenda.Value = Nz(rs!PRECOVENDA)
        txtCadastro.Value = Nz(rs!datacadastro)
        
        cmbCategoria_Change
        ckAtivo.Value = (Nz(rs!STATUS) = "ATIVO")
        
    End If

    rs.Close
    Set rs = Nothing

End Sub
Public Sub carregarsubgrupos(cmb As ComboBox)

    Dim rs As ADODB.Recordset

    Set rs = Conn.Execute("SELECT IDGRUPO, DESCRICAO FROM TAB_GRUPOS ORDER BY DESCRICAO")

    With cmb

        .Clear
        .ColumnCount = 2
        .ColumnWidths = "0 pt;100 pt"

        Do While Not rs.EOF

            .AddItem rs!DESCRICAO
            .List(.ListCount - 1, 0) = rs!IdGrupo

            rs.MoveNext

        Loop

    End With

    rs.Close
    Set rs = Nothing

End Sub

Public Sub CarregarUsuarios()

    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim item As ListItem

    sql = "SELECT IDUSUARIO, NOME, USUARIO FROM TAB_USUARIOS"

    Set rs = New ADODB.Recordset
    rs.Open sql, Conn, adOpenStatic, adLockReadOnly

    lvUsuarios.ListItems.Clear

    Do While Not rs.EOF

        Set item = lvUsuarios.ListItems.Add(, , rs!idUsuario)
        item.SubItems(1) = rs!NOME
        item.SubItems(2) = rs!Usuario

        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing

End Sub
Private Sub CarregarGrupos()

    If Trim(txtidproduto.Value) = "" Then Exit Sub
    
    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim item As ListItem
    
    sql = "SELECT g.idgrupo AS ID, g.descricao AS SUBGRUPO " & _
          "FROM tab_produtosgrupo pg " & _
          "LEFT JOIN tab_grupos g ON pg.idgrupo = g.idgrupo " & _
          "WHERE pg.idProduto = " & txtidproduto.Value
    
    Set rs = New ADODB.Recordset
    rs.Open sql, Conn, adOpenStatic, adLockReadOnly

    lvsubgrupos.ListItems.Clear
    
    Do While Not rs.EOF
    
        Set item = lvsubgrupos.ListItems.Add(, , rs!id)
        item.SubItems(1) = rs!SUBGRUPO
    
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing

End Sub
Private Sub cmbCategoria_Change()

    Dim rs As ADODB.Recordset

    If cmbCategoria.ListIndex = -1 Then Exit Sub

    Set rs = Conn.Execute( _
        "SELECT MARKUP FROM TAB_CATEGORIAS WHERE IDCATEGORIA = " & cmbCategoria.Column(0))

    If Not rs.EOF Then
        txtMarkup.Value = rs!Markup
    End If

    rs.Close
    Set rs = Nothing

    CalcularPrecoSugerido

End Sub
Private Sub CalcularPrecoSugerido()

    Dim Custo As Double
    Dim Markup As Double

    If Trim(txtCusto.Value) = "" Then
        txtSugerido.Value = ""
        Exit Sub
    End If

    If Trim(txtMarkup.Value) = "" Then
        txtSugerido.Value = ""
        Exit Sub
    End If

    If Not IsNumeric(txtCusto.Value) Then Exit Sub
    If Not IsNumeric(txtMarkup.Value) Then Exit Sub


    Custo = CDbl(txtCusto.Value)
    Markup = CDbl(txtMarkup.Value)

    txtSugerido.Value = Format( _
        Custo * (1 + Markup / 100), _
        "0.00")

End Sub

Private Sub btnCancelarEdit_Click()

    '====================================================
    ' FECHA O FORMULÁRIO
    '====================================================
    Unload Me

End Sub


Private Sub btnFechar_Click()

    '====================================================
    ' FECHA O FORMULÁRIO
    '====================================================
    Unload Me

End Sub


