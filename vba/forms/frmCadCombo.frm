VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCadCombo 
   Caption         =   "CADASTRO DE COMBO"
   ClientHeight    =   7590
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11145
   ' OleObjectBlob removido na versao publica
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCadCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub userform_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me
    End If

End Sub

Private Sub Label8_Click()

End Sub

Private Sub UserForm_initialize()

CarregarCategorias cmbCategoria
   CarregarUnidadesMedida cmbUM
   
With lvItensCombo

    .View = lvwReport
    .Gridlines = True
    .FullRowSelect = True
    .HideSelection = False

    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "ID", 40
    .ColumnHeaders.Add , , "PRODUTO", 180
    .ColumnHeaders.Add , , "QTDE", 40
    .ColumnHeaders.Add , , "CUSTO", 70
    .ColumnHeaders.Add , , "ESTOQUE", 70

End With

AjustarLayout

CarregarProdutosCombo cmbProduto

txtEAN5.Visible = False
cmbProduto.Visible = False
txtQuantidade.Visible = False
lvItensCombo.Visible = False
btnAdicionar.Visible = False
btnRemover.Visible = False
btnSalvar.Visible = False
btnCancelar.Visible = False
Label1.Visible = False
Label7.Visible = False
Label2.Visible = False
Label3.Visible = False
txtIdCombo.Visible = False

       
End Sub

Private Sub AjustarLayout()

    With lvItensCombo
        .Left = 18
        .Top = 234
        .Width = 408
        .Height = 132
    End With

End Sub

Private Sub btnCancelar_Click()
    
  Conn.Execute "ROLLBACK"
    
    Unload Me
        
End Sub

Private Sub btnCriar_Click()

    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim IdCombo As Long

If Trim(txtNome.Value) = "" Then
    MsgBox "Informe o nome do combo."
    Exit Sub
End If

If cmbCategoria.ListIndex = -1 Then
    MsgBox "Selecione a categoria."
    Exit Sub
End If

If cmbUM.ListIndex = -1 Then
    MsgBox "Selecione a unidade."
    Exit Sub
End If

sql = "CALL PROC_CADASTRARCOMBO(" & _
      SqlTexto(txtNome.Value) & "," & _
      SqlTexto(cmbUM.Value) & "," & _
      SqlTexto(cmbCategoria.Value) & "," & _
      SqlTexto(txtMarca.Value) & "," & _
      SqlNumero(txtPrecoVenda.Value) & ")"

    Set rs = Conn.Execute(sql)

    If Not rs.EOF Then
        IdCombo = CLng(rs!IdCombo)
    End If

    MsgBox "Combo cadastrado com sucesso!", vbInformation

    rs.Close
    Set rs = Nothing

btnCriar.Visible = False

txtNome.Enabled = False
cmbCategoria.Enabled = False
cmbUM.Enabled = False
txtMarca.Enabled = False

txtIdCombo.Value = IdCombo
txtEAN5.Visible = True
cmbProduto.Visible = True
txtQuantidade.Visible = True
lvItensCombo.Visible = True
btnAdicionar.Visible = True
btnRemover.Visible = True
btnSalvar.Visible = True
btnCancelar.Visible = True
Label1.Visible = True
Label7.Visible = True
Label2.Visible = True
Label3.Visible = True
txtIdCombo.Visible = True

End Sub
Private Sub Userform_activate()

    CarregarProdutosCombo cmbProduto
    
  lvItensCombo.Visible = False
    DoEvents
    AjustarLayout
    lvItensCombo.Visible = True

 If Trim(txtIdCombo.Value) <> "" Then

    CarregarCabecalhoCombo

    btnCriar.Visible = False

    txtNome.Enabled = True
    cmbCategoria.Enabled = True
    cmbUM.Enabled = True
    txtMarca.Enabled = True
    txtPrecoVenda.Enabled = True

    txtEAN5.Visible = True
    cmbProduto.Visible = True
    txtQuantidade.Visible = True
    lvItensCombo.Visible = True
    btnAdicionar.Visible = True
    btnRemover.Visible = True
    btnSalvar.Visible = True
    btnCancelar.Visible = True
    Label1.Visible = True
Label7.Visible = True
Label2.Visible = True
Label3.Visible = True
txtIdCombo.Visible = True

End If

End Sub

Private Sub txtEAN5_AfterUpdate()

   SelecionarProdutoPorEAN cmbProduto, txtEAN5.Value

End Sub

Private Sub cmbProduto_afterupdate()

    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim IDProduto As Long

    If cmbProduto.ListIndex = -1 Then Exit Sub

    IDProduto = CLng(cmbProduto.Value)

    sql = "SELECT TIPO " & _
          "FROM TAB_PRODUTOS " & _
          "WHERE IDPRODUTO = " & IDProduto

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
    End If
End Sub

Private Sub CarregarCabecalhoCombo()

    Dim rs As ADODB.Recordset
    Dim sql As String

    If Trim(txtIdCombo.Value) = "" Then Exit Sub

    sql = "SELECT P.IDPRODUTO, " & _
          "P.NOME, " & _
          "P.IDCATEGORIA, " & _
          "P.MEDIDACOMPRA, " & _
          "P.PRECOVENDA, " & _
          "P.MARCA " & _
          "FROM TAB_PRODUTOS P " & _
          "WHERE P.IDPRODUTO = " & SqlNumero(txtIdCombo.Value)

    Set rs = Conn.Execute(sql)

    If Not rs.EOF Then

        txtIdCombo.Value = rs!IDProduto
        txtNome.Value = Nz(rs!NOME)
        
               Dim i As Long

For i = 0 To cmbCategoria.ListCount - 1
    If CLng(cmbCategoria.List(i, 0)) = CLng(rs!IdCategoria) Then
        cmbCategoria.ListIndex = i
        Exit For
    End If
Next i
        cmbCategoria.Value = Nz(rs!IdCategoria)
        cmbUM.Value = Nz(rs!MEDIDACOMPRA)
        txtPrecoVenda.Value = Format(Nz(rs!PRECOVENDA, 0), "0.00")
        txtMarca.Value = Nz(rs!MARCA)

    End If

    rs.Close
    Set rs = Nothing

    CarregarItensCombo

End Sub

Private Sub CarregarItensCombo()

    Dim rs As New ADODB.Recordset
    Dim item As ListItem
    Dim sql As String

    lvItensCombo.ListItems.Clear

    sql = "SELECT IC.IDITEMCOMBO, P.NOME, IC.QUANTIDADE, IC.CUSTO, P.ESTOQUEATUAL " & _
          "FROM TAB_ITENSCOMBO IC " & _
          "INNER JOIN TAB_PRODUTOS P ON P.IDPRODUTO = IC.IDPRODUTO " & _
          "WHERE IC.IDCOMBO = " & SqlNumero(txtIdCombo.Value)

    rs.Open sql, Conn, adOpenStatic, adLockReadOnly

    Do While Not rs.EOF

        Set item = lvItensCombo.ListItems.Add(, , rs!IDITEMCOMBO)

        item.SubItems(1) = rs!NOME
        item.SubItems(2) = rs!QUANTIDADE
        item.SubItems(3) = Format(rs!Custo, "0.00")
        item.SubItems(4) = Nz(rs!ESTOQUEATUAL)

        rs.MoveNext

    Loop

    rs.Close

End Sub

Private Sub btnAdicionar_Click()

    Dim sql As String

If cmbProduto.ListIndex = -1 Then
    MsgBox "Selecione um produto."
    Exit Sub
End If

If NzDbl(txtQuantidade.Value) <= 0 Then
    MsgBox "Quantidade inválida."
    Exit Sub
End If

sql = "CALL PROC_ADCITENSCOMBO(" & _
      SqlNumero(txtIdCombo.Value) & "," & _
      SqlNumero(cmbProduto.Column(0)) & "," & _
      SqlNumero(txtQuantidade.Value) & ")"

    Conn.Execute sql

CarregarItensCombo

txtQuantidade.Value = ""
cmbProduto.ListIndex = -1
txtEAN5.Value = ""

txtEAN5.SetFocus

MsgBox "Item adicionado.", vbInformation

End Sub

Private Sub btnRemover_Click()

    Dim sql As String
    Dim rs As ADODB.Recordset

    If lvItensCombo.SelectedItem Is Nothing Then
        MsgBox "Selecione um item para remover.", vbExclamation
        Exit Sub
    End If

sql = "CALL PROC_REMOVEITEMCOMBO(" & _
      SqlNumero(lvItensCombo.SelectedItem.Text) & ")"

    Set rs = Conn.Execute(sql)

    If Not rs Is Nothing Then

        If Not rs.EOF Then
            MsgBox rs.Fields(0).Value, vbInformation
        End If

        rs.Close

    End If

    Set rs = Nothing

CarregarItensCombo

End Sub
Private Sub btnSalvar_Click()

    Dim sql As String
    Dim rs As ADODB.Recordset

    '===========================
    ' Atualiza cabeçalho
    '===========================
    sql = "CALL PROC_EDITARCOMBO(" & _
          SqlNumero(txtIdCombo.Value) & "," & _
          SqlTexto(txtNome.Value) & "," & _
          SqlTexto(cmbUM.Value) & "," & _
          SqlTexto(cmbCategoria.Value) & "," & _
          SqlTexto(txtMarca.Value) & "," & _
          SqlNumero(txtPrecoVenda.Value) & ")"

    Set rs = Conn.Execute(sql)

    If Not rs Is Nothing Then
        If Not rs.EOF Then
            MsgBox rs.Fields(0).Value, vbInformation
        End If
        rs.Close
    End If

    '===========================
    ' Recalcula custo do combo
    '===========================
    sql = "CALL PROC_SALVARCOMBO(" & SqlNumero(txtIdCombo.Value) & ")"

    Set rs = Conn.Execute(sql)

    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
    End If

    frmProdutos.CarregarProdutos

    MsgBox "Combo salvo com sucesso.", vbInformation

    Unload Me

End Sub


