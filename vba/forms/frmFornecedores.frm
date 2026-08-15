VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFornecedores 
   Caption         =   "FORNECEDORES"
   ClientHeight    =   7170
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12060
   ' OleObjectBlob removido na versao publica
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFornecedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_initialize()

    With lvFornecedores
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .HideSelection = False
        .ColumnHeaders.Clear
    End With

    With lvFornecedores.ColumnHeaders
        .Add , , "ID", 60
        .Add , , "Nome", 200
        .Add , , "CNPJ", 130
        .Add , , "Contato", 100
        .Add , , "Status", 80
    End With

CarregarComboFornecedor cmbFornecedor
CarregarFornecedores

cmbStatus.List = Array("TODOS", "ATIVO", "INATIVO")
cmbStatus.ListIndex = 1
End Sub
Private Sub btnNovoFornecedor_Click()

    frmCadFornecedor.Show vbModeless

End Sub
Private Sub cmbFornecedor_Change()

    CarregarFornecedores

End Sub


Private Sub cmbStatus_Change()

    CarregarFornecedores

End Sub
Public Sub CarregarFornecedores()

    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim item As ListItem
    Dim filtroFornecedor As String
    Dim filtroStatus As String


    sql = "SELECT IDFORNECEDOR,NOME,CNPJ,CONTATO,STATUS FROM TAB_FORNECEDORES WHERE 1=1 "


    If cmbFornecedor.ListIndex <> -1 Then
        filtroFornecedor = cmbFornecedor.Column(0)
        sql = sql & " AND IDFORNECEDOR = " & filtroFornecedor
    End If


    If cmbStatus.Value <> "TODOS" And cmbStatus.Value <> "" Then
        filtroStatus = cmbStatus.Value
        sql = sql & " AND STATUS = '" & filtroStatus & "'"
    Else
        sql = sql & " AND STATUS = 'ATIVO'"
    End If


    sql = sql & " ORDER BY NOME"


    Set rs = New ADODB.Recordset
    rs.Open sql, Conn, adOpenStatic, adLockReadOnly


    lvFornecedores.ListItems.Clear


    Do While Not rs.EOF

        Set item = lvFornecedores.ListItems.Add(, , Nz(rs!idFornecedor))

        item.SubItems(1) = Nz(rs!NOME)
        item.SubItems(2) = Nz(rs!CNPJ)
        item.SubItems(3) = Nz(rs!CONTATO)
        item.SubItems(4) = Nz(rs!STATUS)

        rs.MoveNext

    Loop


    rs.Close
    Set rs = Nothing

End Sub
Private Sub btnEditarFornecedor_Click()

    If lvFornecedores.SelectedItem Is Nothing Then
        MsgBox "Selecione um fornecedor", vbExclamation
        Exit Sub
    End If

    frmCadFornecedor.Tag = lvFornecedores.SelectedItem.Text
    frmCadFornecedor.Show
    

End Sub

Private Sub btnExcluirFornecedor_Click()

    Dim sql As String
    Dim resp As VbMsgBoxResult
    Dim IDProduto As Long

    If lvFornecedores.SelectedItem Is Nothing Then
        MsgBox "Selecione um fornecedor", vbExclamation
        Exit Sub
    End If

    idFornecedor = CLng(lvFornecedores.SelectedItem.Text)

    resp = MsgBox("Tem certeza que deseja excluir este fornecedor?", _
                  vbYesNo + vbQuestion, _
                  "Confirmação")

    If resp = vbNo Then Exit Sub

    sql = "DELETE FROM TAB_FORNECEDORES WHERE IDFORNECEDOR = " & idFornecedor

    Conn.Execute sql

    MsgBox "Fornecedor excluído com sucesso!", vbInformation

    Call CarregarFornecedores   ' recarrega o ListView

End Sub

