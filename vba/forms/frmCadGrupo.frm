VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCadGrupo 
   Caption         =   "CADASTRO DE GRUPO"
   ClientHeight    =   7590
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11145
   OleObjectBlob   =   "frmCadGrupo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCadGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private IDGrupoAtual As Long

Private Sub userform_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me
    End If

End Sub

Private Sub UserForm_Initialize()

    CarregarProdutosCombo cmbProduto, True

    With lvItensGrupo
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .HideSelection = False

        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "ID", 40
        .ColumnHeaders.Add , , "PRODUTO", 240
        .ColumnHeaders.Add , , "QTDE", 60
    End With

    AjustarLayout

    'segunda etapa escondida
    cmbProduto.Visible = False
    txtQuantidadeItem.Visible = False
    lvItensGrupo.Visible = False
    btnAdicionar.Visible = False
    btnRemover.Visible = False
    btnSalvar.Visible = False
    btnCancelar.Visible = False
Label1.Visible = False
Label7.Visible = False
txtIdgrupo.Visible = False
txtEAN5.Visible = False
Label2.Visible = False
cmbProduto.Visible = False
Label3.Visible = False

End Sub

Private Sub AjustarLayout()

    With lvItensGrupo

        .Left = 18
        .Top = 234
        .Width = 408
        .Height = 132

    End With

End Sub

Private Sub userform_activate()

  lvItensGrupo.Visible = False
    DoEvents
    AjustarLayout
    lvItensGrupo.Visible = True

    CarregarProdutosCombo cmbProduto, True

If Me.Tag <> "" Then

    txtIdgrupo.Value = Me.Tag

    CarregarCabecalhoGrupo

    btnCriar.Visible = False

    txtDescricao.Enabled = False

    cmbProduto.Visible = True
    txtQuantidadeItem.Visible = True
    lvItensGrupo.Visible = True
    btnAdicionar.Visible = True
    btnRemover.Visible = True
    btnSalvar.Visible = True
    btnCancelar.Visible = True
    Label1.Visible = True
Label7.Visible = True
txtIdgrupo.Visible = True
txtEAN5.Visible = True
Label2.Visible = True
cmbProduto.Visible = True
Label3.Visible = True
    
CarregarItensGrupo

End If

Me.Tag = ""

End Sub
Private Sub CarregarCabecalhoGrupo()

    Dim rs As ADODB.Recordset
    Dim sql As String


    If Trim(txtIdgrupo.Value) = "" Then Exit Sub


sql = "SELECT IDGRUPO, DESCRICAO " & _
      "FROM TAB_GRUPOS " & _
      "WHERE IDGRUPO = " & SqlNumero(txtIdgrupo.Value)


    Set rs = Conn.Execute(sql)


    If Not rs.EOF Then

        txtIdgrupo.Value = rs!IdGrupo
        txtDescricao.Value = Nz(rs!DESCRICAO)

    End If


    rs.Close
    Set rs = Nothing

End Sub
Private Sub btnCriar_Click()

    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim IdGrupo As Long

    If Trim(txtDescricao.Value) = "" Then
        MsgBox "Informe a descrição do grupo."
        Exit Sub
    End If

sql = "CALL PROC_CADASTRARGRUPO(" & _
      SqlTexto(txtDescricao.Value) & ")"

    Set rs = Conn.Execute(sql)

    If Not rs.EOF Then
        IdGrupo = CLng(rs!IdGrupo)
    End If

    rs.Close
    Set rs = Nothing

    txtIdgrupo.Value = IdGrupo

    MsgBox "Grupo criado com sucesso!", vbInformation

    btnCriar.Visible = False

    txtDescricao.Enabled = False

    cmbProduto.Visible = True
    txtQuantidadeItem.Visible = True
    lvItensGrupo.Visible = True
    btnAdicionar.Visible = True
    btnRemover.Visible = True
    btnSalvar.Visible = True
btnCancelar.Visible = True
Label1.Visible = True
Label7.Visible = True
txtIdgrupo.Visible = True
txtEAN5.Visible = True
Label2.Visible = True
cmbProduto.Visible = True
Label3.Visible = True

lvItensGrupo.Visible = False
DoEvents

AjustarLayout

lvItensGrupo.Visible = True
DoEvents

End Sub

Private Sub btnAdicionar_Click()

    Dim rs As ADODB.Recordset
    Dim sql As String

    If txtIdgrupo.Value = "" Then
        MsgBox "Crie o grupo primeiro.", vbExclamation
        Exit Sub
    End If

    If cmbProduto.ListIndex = -1 Then
        MsgBox "Selecione um produto.", vbExclamation
        Exit Sub
    End If

    If Val(txtQuantidadeItem.Value) < 0 Then
        MsgBox "Informe uma quantidade válida.", vbExclamation
        Exit Sub
    End If


    sql = "CALL PROC_ADCITEMGRUPO(" & _
          txtIdgrupo.Value & "," & _
          cmbProduto.Column(0) & "," & _
          Replace(txtQuantidadeItem.Value, ",", ".") & ")"


    Set rs = Conn.Execute(sql)


    If Not rs.EOF Then

        MsgBox rs!Retorno, vbInformation

        If rs!Retorno = "ITEM ADICIONADO AO GRUPO" Then
            CarregarItensGrupo
        End If

    End If


    rs.Close
    Set rs = Nothing

cmbProduto.ListIndex = -1
txtQuantidadeItem.Value = ""

End Sub

Private Sub btnRemover_Click()

    Dim rs As ADODB.Recordset
    Dim sql As String

    If lvItensGrupo.SelectedItem Is Nothing Then
        MsgBox "Selecione um item para remover.", vbExclamation
        Exit Sub
    End If
    
sql = "CALL PROC_REMOVEITEMGRUPO(" & _
      SqlNumero(txtIdgrupo.Value) & "," & _
      SqlNumero(lvItensGrupo.SelectedItem.Text) & ")"

    Set rs = Conn.Execute(sql)

    If Not rs.EOF Then
        MsgBox rs!Retorno, vbInformation
    End If

    rs.Close
    Set rs = Nothing

    CarregarItensGrupo

End Sub
Private Sub btnSalvar_Click()

    MsgBox "Grupo salvo com sucesso.", vbInformation

    Unload Me

End Sub
Private Sub btnCancelar_Click()

    Unload Me

End Sub
Private Sub CarregarItensGrupo()

    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim ITEM As ListItem

    lvItensGrupo.ListItems.Clear

    If Trim(txtIdgrupo.Value) = "" Then Exit Sub


    sql = "SELECT PG.IDPRODUTO, P.NOME, PG.QUANTIDADE " & _
          "FROM TAB_PRODUTOSGRUPO PG " & _
          "INNER JOIN TAB_PRODUTOS P ON P.IDPRODUTO = PG.IDPRODUTO " & _
          "WHERE PG.IDGRUPO = " & SqlNumero(txtIdgrupo.Value) & _
          " ORDER BY P.NOME"


    Set rs = Conn.Execute(sql)


    Do While Not rs.EOF

        Set ITEM = lvItensGrupo.ListItems.Add(, , rs!idproduto)

        ITEM.SubItems(1) = Nz(rs!NOME)
        ITEM.SubItems(2) = Nz(rs!QUANTIDADE)

        rs.MoveNext

    Loop


    rs.Close
    Set rs = Nothing

End Sub
