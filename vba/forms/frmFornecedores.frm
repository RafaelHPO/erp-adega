VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFornecedores 
   Caption         =   "FORNECEDORES"
   ClientHeight    =   7170
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12060
   OleObjectBlob   =   "frmFornecedores.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFornecedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AplicarPaleta()

    AplicarTema Me
    AplicarBotaoPrincipal btnNovoFornecedor
    AplicarBotaoPrincipal btnExcluirFornecedor
    AplicarTemaLv lvFornecedores
    
End Sub

Private Sub UserForm_initialize()

On Error GoTo TratarErro

    With lvFornecedores
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = False
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

AplicarPaleta

    Exit Sub

TratarErro:

    modSistema.tela = "Fornecedores - initialize"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"

End Sub
Private Sub btnNovoFornecedor_Click()

    frmCadFornecedor.Show vbModeless

End Sub
Private Sub cmbFornecedor_Change()

    CarregarFornecedores

End Sub


Private Sub cmbStatus_change()

    CarregarFornecedores

End Sub
Public Sub CarregarFornecedores()

On Error GoTo TratarErro

    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim ITEM As ListItem
    Dim filtroFornecedor As String
    Dim FiltroStatus As String


    sql = "SELECT IDFORNECEDOR,NOME,CNPJ,CONTATO,STATUS FROM TAB_FORNECEDORES WHERE 1=1 "


    If cmbFornecedor.ListIndex <> -1 Then
        filtroFornecedor = cmbFornecedor.Column(0)
        sql = sql & " AND IDFORNECEDOR = " & filtroFornecedor
    End If


    If cmbStatus.Value <> "TODOS" And cmbStatus.Value <> "" Then
        FiltroStatus = cmbStatus.Value
        sql = sql & " AND STATUS = '" & FiltroStatus & "'"
    Else
        sql = sql & " AND STATUS = 'ATIVO'"
    End If


    sql = sql & " ORDER BY NOME"


    Set rs = New ADODB.Recordset
    rs.Open sql, Conn, adOpenStatic, adLockReadOnly


    lvFornecedores.ListItems.Clear


    Do While Not rs.EOF

        Set ITEM = lvFornecedores.ListItems.Add(, , Nz(rs!idFornecedor))

        ITEM.SubItems(1) = Nz(rs!NOME)
        ITEM.SubItems(2) = Nz(rs!CNPJ)
        ITEM.SubItems(3) = Nz(rs!CONTATO)
        ITEM.SubItems(4) = Nz(rs!STATUS)

        rs.MoveNext

    Loop


    rs.Close
    Set rs = Nothing

    Exit Sub

TratarErro:

    modSistema.tela = "Fornecedores - carregar lv"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"

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

On Error GoTo TratarErro

    Dim sql As String
    Dim resp As VbMsgBoxResult
    Dim idproduto As Long

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

    Exit Sub

TratarErro:

    modSistema.tela = "Fornecedores - btnExcluir"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"

End Sub

