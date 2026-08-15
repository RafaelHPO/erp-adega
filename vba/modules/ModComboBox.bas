Attribute VB_Name = "ModComboBox"
Public Sub CarregarProdutosCombo(cmb As MSForms.ComboBox, Optional SomenteUnitario As Boolean = False)

    Dim rs As ADODB.Recordset
    Dim sql As String

    sql = "SELECT IDPRODUTO, NOME " & _
          "FROM TAB_PRODUTOS " & _
          "WHERE STATUS = 'ATIVO' "

    If SomenteUnitario = True Then
        sql = sql & "AND TIPO <> 'COMBO' "
    End If

    sql = sql & "ORDER BY NOME"

    Set rs = Conn.Execute(sql)

    cmb.Clear
    cmb.ColumnCount = 2
    cmb.ColumnWidths = "0 pt;200 pt"

    Do While Not rs.EOF

        cmb.AddItem rs!IDProduto & ""
        cmb.List(cmb.ListCount - 1, 1) = rs!NOME

        rs.MoveNext

    Loop

    rs.Close
    Set rs = Nothing

End Sub
Public Sub SelecionarProdutoPorEAN(cmb As MSForms.ComboBox, CodBarras As String)

    If Trim(CodBarras) = "" Then Exit Sub

    Dim rs As ADODB.Recordset
    Dim sql As String

    sql = "SELECT IDPRODUTO " & _
          "FROM TAB_PRODUTOS " & _
          "WHERE CODIGOBARRAS = '" & _
          Replace(Trim(CodBarras), "'", "''") & "'"

    Set rs = Conn.Execute(sql)

    If Not rs.EOF Then
        cmb.Value = rs!IDProduto
    Else
        cmb.ListIndex = -1
    End If

    rs.Close
    Set rs = Nothing

End Sub

Public Sub CarregarCategorias(cmb As ComboBox)

    Dim rs As ADODB.Recordset

    Set rs = Conn.Execute( _
        "SELECT IDCATEGORIA, DESCRICAO FROM TAB_CATEGORIAS ORDER BY DESCRICAO")

    With cmb

        .Clear
        .ColumnCount = 2
        .ColumnWidths = "0 pt;120 pt"

        .AddItem 0
        .List(.ListCount - 1, 1) = "TODOS"

        Do While Not rs.EOF

            .AddItem rs!IdCategoria
            .List(.ListCount - 1, 1) = rs!DESCRICAO

            rs.MoveNext

        Loop

    End With

    rs.Close
    Set rs = Nothing

End Sub
Public Sub CarregarUnidadesMedida(cmb As ComboBox)

    With cmb

        .Clear

        '====================================================
        ' UNIDADES PADRÃO (SIGLAS ATÉ 3 LETRAS)
        '====================================================
        .AddItem "UN"   ' unidade
        .AddItem "CX"   ' caixa
        .AddItem "FD"   ' fardo
        .AddItem "PC"   ' pacote
        .AddItem "KT"   ' kit

        '====================================================
        ' BEBIDAS / VAREJO
        '====================================================
        .AddItem "LT"   ' litro
        .AddItem "ML"   ' mililitro
        .AddItem "LQ"   ' long neck (lógica de embalagem)
        .AddItem "LTN"  ' latão
        .AddItem "LTA"  ' lata
        .AddItem "GRF"  ' garrafa

        '====================================================
        ' PESO
        '====================================================
        .AddItem "KG"   ' quilo
        .AddItem "G"    ' grama

        '====================================================
        ' OUTROS
        '====================================================
        .AddItem "UN"   ' redundante ok se quiser padrão simples

    End With

End Sub

Public Sub CarregarSetor(cmb As ComboBox)

    With cmb

        .Clear

        '====================================================
        ' SETORES DO SISTEMA
        '====================================================
        .AddItem "BEBIDAS"
        .AddItem "TABACARIA"
        .AddItem "CONVENIENCIA"

    End With

End Sub

Public Sub CarregarComboFornecedor(cmb As ComboBox)

    Dim rs As ADODB.Recordset

    Set rs = Conn.Execute("SELECT IDFORNECEDOR,NOME FROM TAB_FORNECEDORES WHERE STATUS='ATIVO' ORDER BY NOME")

    With cmb

        .Clear
        .ColumnCount = 2
        .ColumnWidths = "0 pt;120 pt"
        Do While Not rs.EOF

            .AddItem rs!idFornecedor
            .List(.ListCount - 1, 1) = rs!NOME

            rs.MoveNext

        Loop

    End With

    rs.Close
    Set rs = Nothing

End Sub

Public Sub CarregarComboStatus()

    cmbStatus.Clear

    cmbStatus.AddItem "ATIVO"
    cmbStatus.AddItem "TODOS"

    cmbStatus.ListIndex = 0

End Sub
