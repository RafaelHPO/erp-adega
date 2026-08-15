VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProdutos 
   Caption         =   "PRODUTOS"
   ClientHeight    =   11610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19755
   ' OleObjectBlob removido na versao publica
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAcerto_Click()

    fraCombo.Visible = False
    frmAcerto.Show vdmodeless

End Sub

Private Sub btnCadCombo_Click()

    fraCombo.Visible = False
    frmCadCombo.Show vbModal

End Sub

Private Sub btnCalculaPrecovenda_Click()

    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim MSG As String

    sql = "CALL PROC_CALCULAPRECOVENDA()"
    Set rs = Conn.Execute(sql)

    MsgBox rs.Fields(0).Value, vbInformation

    rs.Close
    Set rs = Nothing

    CarregarProdutos

End Sub

Private Carregando As Boolean

Private Sub btnCombos_Click()

    fraCombo.Visible = Not fraCombo.Visible

End Sub

Private Sub btnEditarCombo_Click()

    fraCombo.Visible = False

    Dim rs As ADODB.Recordset
    Dim sql As String

    If lvProdutosCadastro.SelectedItem Is Nothing Then
        MsgBox "Selecione um produto", vbExclamation
        Exit Sub
    End If

    sql = "SELECT TIPO " & _
          "FROM TAB_PRODUTOS " & _
          "WHERE IDPRODUTO = " & CLng(lvProdutosCadastro.SelectedItem.Text)

    Set rs = Conn.Execute(sql)

    If rs.EOF Then
        MsgBox "Produto não encontrado."
        Exit Sub
    End If

    If UCase(Nz(rs!tipo, "")) <> "COMBO" Then
        MsgBox "O produto selecionado não é um combo.", vbExclamation
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If

    rs.Close
    Set rs = Nothing

    frmCadCombo.txtIdCombo.Value = CLng(lvProdutosCadastro.SelectedItem.Text)
    frmCadCombo.Show vbModal

    fraCombo.Visible = False

End Sub


Private Sub UserForm_initialize()

    Carregando = True

    With lvProdutosCadastro
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .HideSelection = False
        .ColumnHeaders.Clear
    End With

    With lvProdutosCadastro.ColumnHeaders
        .Add , , "ID", 60
        .Add , , "NOME", 250
        .Add , , "ESTOQUE", 80
        .Add , , "CUSTO", 90
        .Add , , "PREÇO VENDA", 100
        .Add , , "SALDO ESTOQUE", 100
        .Add , , "LUCRO ESTIMADO", 120
        .Add , , "TIPO", 80
        .Add , , "STATUS", 80
    End With

    CarregarIndicadores
    CarregarProdutos

    CarregarProdutosCombo cmbProduto
With cmbCategoria

    .Clear
    .ColumnCount = 2
    .ColumnWidths = "0 pt;120 pt"

    .AddItem 0
    .List(.ListCount - 1, 1) = "TODOS"

End With

CarregarCategorias cmbCategoria
    CarregarSetor cmbSetor

    cmbTipo.List = Array("TODOS", "UNIT", "COMBO")
    cmbStatus.List = Array("TODOS", "ATIVO", "INATIVO")

    cmbCategoria.AddItem "TODOS", 0
    cmbSetor.AddItem "TODOS", 0

    cmbCategoria.ListIndex = 1
    cmbSetor.ListIndex = 0
    cmbTipo.ListIndex = 0
    cmbStatus.ListIndex = 1

    Carregando = False
    fraCombo.Visible = False

End Sub

Private Sub Userform_activate()

    CarregarIndicadores

End Sub

Private Sub btnNovoProd_Click()

    fraCombo.Visible = False
    frmCadProduto.Show vbModeless

End Sub

Private Sub cmbCategoria_Change()

    If Carregando Then Exit Sub
    AtualizarTela

End Sub

Private Sub cmbSetor_Change()

    If Carregando Then Exit Sub
    AtualizarTela

End Sub

Private Sub cmbTipo_Change()

    If Carregando Then Exit Sub
    AtualizarTela

End Sub

Private Sub cmbStatus_Change()

    If Carregando Then Exit Sub
    AtualizarTela

End Sub

Private Sub cmbProduto_change()

    If Carregando Then Exit Sub
    AtualizarTela

End Sub

Private Function WhereFiltros() As String

    Dim W As String

    W = " WHERE 1=1 "

If cmbCategoria.ListIndex > -1 Then

    If Not IsNull(cmbCategoria.Column(0)) Then

        If Val(cmbCategoria.Column(0)) > 0 Then

            W = W & " AND IDCATEGORIA = " & Val(cmbCategoria.Column(0))

        End If

    End If

End If
    If cmbSetor.Value <> "TODOS" Then
        W = W & " AND SETOR = " & SqlTexto(cmbSetor.Value)
    End If

    If cmbTipo.Value <> "TODOS" Then
        W = W & " AND TIPO = " & SqlTexto(cmbTipo.Value)
    End If

    If cmbStatus.Value <> "TODOS" Then
        W = W & " AND STATUS = " & SqlTexto(cmbStatus.Value)
    End If

    If Trim(cmbProduto.Value) <> "" Then
        W = W & " AND IDPRODUTO = " & CLng(cmbProduto.Value)
    End If

    WhereFiltros = W

End Function

Public Sub AtualizarTela()

    CarregarProdutos
    CarregarIndicadores

End Sub

Public Sub CarregarProdutos()

    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim item As ListItem

    sql = "SELECT IDPRODUTO, UPPER(NOME) NOME, TIPO, " & _
          "ESTOQUEATUAL, CUSTOUNITARIO, PRECOVENDA, UPPER(STATUS) AS STATUS, " & _
          "(ESTOQUEATUAL * CUSTO) AS SALDOESTOQUE, " & _
          "((PRECOVENDA - CUSTO) * ESTOQUEATUAL) AS LUCROESTIMADO " & _
          "FROM TAB_PRODUTOS "

    sql = sql & WhereFiltros
    sql = sql & " ORDER BY NOME"

    Set rs = New ADODB.Recordset
    rs.Open sql, Conn, adOpenStatic, adLockReadOnly

    lvProdutosCadastro.ListItems.Clear

    Do While Not rs.EOF

        Set item = lvProdutosCadastro.ListItems.Add(, , Nz(rs!IDProduto))

        item.SubItems(1) = Nz(rs!NOME)
        item.SubItems(2) = Nz(rs!ESTOQUEATUAL)
        item.SubItems(3) = Format(NzDbl(rs!CUSTOUNITARIO), "R$   #,##0.00")
        item.SubItems(4) = Format(NzDbl(rs!PRECOVENDA), "R$   #,##0.00")
        item.SubItems(5) = Format(NzDbl(rs!SALDOESTOQUE), "R$   #,##0.00")
        item.SubItems(6) = Format(NzDbl(rs!LUCROESTIMADO), "R$   #,##0.00")
        item.SubItems(7) = Nz(rs!tipo)
        item.SubItems(8) = Nz(rs!STATUS)

        rs.MoveNext

    Loop

    rs.Close
    Set rs = Nothing

End Sub

Public Sub CarregarIndicadores()

    Dim rs As ADODB.Recordset
    Dim sql As String

    sql = "SELECT " & _
          "COUNT(*) TOTALPRODUTOS, " & _
          "SUM(ESTOQUEATUAL*CUSTO) VALORESTOQUE, " & _
          "SUM((PRECOVENDA-CUSTO)*ESTOQUEATUAL) LUCROESTIMADO " & _
          "FROM TAB_PRODUTOS"

    sql = sql & WhereFiltros

    Set rs = New ADODB.Recordset
    rs.Open sql, Conn, adOpenStatic, adLockReadOnly

    If Not rs.EOF Then

        txtProdutos.Value = Nz(rs!TOTALPRODUTOS)
        txtEstTotal.Value = Format(Nz(rs!VALORESTOQUE), "R$ #,##0.00")
        txtEstimado.Value = Format(Nz(rs!LUCROESTIMADO), "R$ #,##0.00")

    End If

    rs.Close
    Set rs = Nothing

End Sub

Private Sub btnEditarProduto_Click()

    fraCombo.Visible = False

    If lvProdutosCadastro.SelectedItem Is Nothing Then
        MsgBox "Selecione um produto", vbExclamation
        Exit Sub
    End If

    frmCadProduto.lblIdProduto.Visible = True
    frmCadProduto.txtidproduto.Visible = True
    frmCadProduto.btnSalvar.Visible = False
    frmCadProduto.btnFechar.Visible = False
    frmCadProduto.btnSalvarEdit.Visible = True
    frmCadProduto.btnCancelarEdit.Visible = True
    frmCadProduto.lblTipo.Visible = True
    frmCadProduto.cmbTipo.Visible = True
    frmCadProduto.lblCadastro.Visible = True
    frmCadProduto.txtCadastro.Visible = True
    frmCadProduto.Label14.Visible = True
    frmCadProduto.txtCustoMedio.Visible = True

    frmCadProduto.txtidproduto = lvProdutosCadastro.SelectedItem.Text
    frmCadProduto.Show vbModeless

End Sub

Private Sub btnExcluirProduto_Click()

    fraCombo.Visible = False

    Dim sql As String
    Dim resp As VbMsgBoxResult
    Dim IDProduto As Long

    If lvProdutosCadastro.SelectedItem Is Nothing Then
        MsgBox "Selecione um produto", vbExclamation
        Exit Sub
    End If

    IDProduto = CLng(lvProdutosCadastro.SelectedItem.Text)

    resp = MsgBox("Tem certeza que deseja excluir este produto?", _
                  vbYesNo + vbQuestion, _
                  "Confirmação")

    If resp = vbNo Then Exit Sub

    sql = "DELETE FROM TAB_PRODUTOS WHERE IDPRODUTO = " & IDProduto

    Conn.Execute sql

    MsgBox "Produto excluído com sucesso!", vbInformation

    Call CarregarProdutos

End Sub

