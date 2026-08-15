VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCompra 
   Caption         =   "ENTRADA DE MERCADORIA"
   ClientHeight    =   7155
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12150
   ' OleObjectBlob removido na versao publica
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TipoEntrada As String
Public NumeroNFE_XML As String

Private Sub btnAtualizar_Click()
        
ThisWorkbook.RefreshAll
    
End Sub

Private Sub btnImportar_Click()

    Dim arquivo As String

    arquivo = SelecionarArquivo()

    If arquivo <> "" Then

        ProcessarNFCe arquivo

    End If

CarregarNotasXML
    CarregarFornecedores

End Sub

Function SelecionarArquivo() As String

    With Application.FileDialog(msoFileDialogFilePicker)

        .Title = "Selecionar NFC-e"
        .Filters.Clear
        .Filters.Add "Arquivo Excel", "*.xlsx"

        If .Show = -1 Then

            SelecionarArquivo = .SelectedItems(1)

        End If

    End With

End Function



Private Sub UserForm_initialize()

    TipoEntrada = "MANUAL"

    With lvNotas
    
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .AllowColumnReorder = True
        .HideSelection = False
        .Font = Tahoma
        .Font.Size = 11
        
        .ColumnHeaders.Clear
        
        .ColumnHeaders.Add , , "DATA", 80
        .ColumnHeaders.Add , , "CNPJ", 100
        .ColumnHeaders.Add , , "FORNECEDOR", 200
        .ColumnHeaders.Add , , "NUMERO NFE", 90
        .ColumnHeaders.Add , , "VALOR TOTAL", 90
        
    End With

    CarregarNotasXML
    CarregarComboFornecedor cmbFornecedor

End Sub

Private Sub Userform_activate()

    CarregarComboFornecedor cmbFornecedor
    CarregarNotasXML

End Sub

Public Sub CarregarNotasXML()

    Dim ws As Worksheet
    Dim tb As ListObject
    Dim i As Long
    Dim item As ListItem

    Set ws = ThisWorkbook.Worksheets("ENTRADA CONSOLIDADO")
    Set tb = ws.ListObjects("ENTRADA_CONSOLIDADO")

    lvNotas.ListItems.Clear

    If tb.DataBodyRange Is Nothing Then Exit Sub

    For i = 1 To tb.ListRows.Count

        If UCase(Trim(tb.DataBodyRange(i, tb.ListColumns("STATUS").Index).Value)) <> "PROCESSADO" Then

            Set item = lvNotas.ListItems.Add(, , _
                tb.DataBodyRange(i, tb.ListColumns("DATA").Index).Value)

            item.SubItems(1) = tb.DataBodyRange(i, tb.ListColumns("CNPJ").Index).Value
            item.SubItems(2) = tb.DataBodyRange(i, tb.ListColumns("FORNECEDOR").Index).Value
            item.SubItems(3) = tb.DataBodyRange(i, tb.ListColumns("NUMERO NFE").Index).Value
            item.SubItems(4) = tb.DataBodyRange(i, tb.ListColumns("VALOR TOTAL").Index).Value

            'Guarda a chave da NFe para uso posterior
            item.Tag = tb.DataBodyRange(i, tb.ListColumns("CHAVE NFE").Index).Value

        End If

    Next i

End Sub
Private Sub lvNotas_DblClick()

    Dim item As ListItem
    Dim i As Long
    Dim encontrou As Boolean
    Dim ws As Worksheet
    Dim tb As ListObject
    
    If lvNotas.SelectedItem Is Nothing Then Exit Sub
    
    Set item = lvNotas.SelectedItem
    
    Set ws = ThisWorkbook.Worksheets("ENTRADA CONSOLIDADO")
    Set tb = ws.ListObjects("ENTRADA_CONSOLIDADO")
    
    
    '====================================
    ' DADOS DA NOTA XML
    '====================================
    
    txtNf.Value = item.SubItems(3)
    txtValorTotal.Value = item.SubItems(4)
    
    txtChaveNf.Value = item.Tag
    
    NumeroNFE_XML = item.SubItems(3)
    TipoEntrada = "XML"
    
    
    '====================================
    ' PROCURA FORNECEDOR CADASTRADO
    '====================================
    
    For i = 0 To cmbFornecedor.ListCount - 1
        
        If cmbFornecedor.List(i, 1) = item.SubItems(2) Then
            
            cmbFornecedor.ListIndex = i
            txtId.Value = cmbFornecedor.List(i, 0)
            encontrou = True
            
            Exit For
            
        End If
        
    Next i
    
    
    '====================================
    ' FORNECEDOR NÃO ENCONTRADO
    '====================================
    
    If encontrou = False Then
        
        If MsgBox("Fornecedor não cadastrado." & vbCrLf & vbCrLf & _
                  "Deseja cadastrar agora?", _
                  vbQuestion + vbYesNo, _
                  "Fornecedor não encontrado") = vbYes Then
            
            Call AbrirCadastroFornecedorXML(item)
            
        Else
            
            txtId.Value = ""
            
        End If
        
    End If

End Sub
Private Sub AbrirCadastroFornecedorXML(ByVal item As ListItem)

    Dim ws As Worksheet
    Dim tb As ListObject
    Dim Linha As ListRow
    
    Set ws = ThisWorkbook.Worksheets("ENTRADA CONSOLIDADO")
    Set tb = ws.ListObjects("ENTRADA_CONSOLIDADO")
    
    
    For Each Linha In tb.ListRows
        
        If Linha.Range(tb.ListColumns("CNPJ").Index).Value = item.SubItems(1) Then
            
            With frmCadFornecedor
                
                .txtNome.Value = Linha.Range(tb.ListColumns("FORNECEDOR").Index).Value
                .txtCNPJ.Value = Linha.Range(tb.ListColumns("CNPJ").Index).Value
                
                .txtEndereco.Value = Linha.Range(tb.ListColumns("ENDERECO").Index).Value
                .txtNumero.Value = Linha.Range(tb.ListColumns("NUMERO").Index).Value
                .txtBairro.Value = Linha.Range(tb.ListColumns("BAIRRO").Index).Value
                
                .txtContato.Value = Linha.Range(tb.ListColumns("CONTATO").Index).Value
                .ckAtivo.Value = True
                
                .Show
                
            End With
            
            Exit For
            
        End If
        
    Next Linha

End Sub

Private Sub cmbFornecedor_Change()
    
   Dim idFornecedor As Long

    If cmbFornecedor.ListIndex = -1 Then Exit Sub

    idFornecedor = CLng(cmbFornecedor.Column(0))
    
        txtId.Value = (idFornecedor)

    
End Sub
Private Sub btnSalvar_Click()

On Error GoTo Erro

Dim cmd As ADODB.Command
Dim rs As ADODB.Recordset
Dim IdCompraGerada As Long


If txtId.Value = "" Then

    MsgBox "Selecione o fornecedor.", vbExclamation
    Exit Sub

End If


Set cmd = New ADODB.Command


With cmd

    Set .ActiveConnection = Conn
    .CommandType = adCmdStoredProc
    .CommandText = "PROC_REGISTRARCOMPRA"

    .Parameters.Append .CreateParameter("P_IDFORNECEDOR", adInteger, adParamInput, , CLng(txtId.Value))
    .Parameters.Append .CreateParameter("P_CHAVENF", adVarChar, adParamInput, 44, txtChaveNf.Value)
    .Parameters.Append .CreateParameter("P_NNF", adInteger, adParamInput, , CLng(txtNf.Value))
    .Parameters.Append .CreateParameter("P_IDUSUARIO", adInteger, adParamInput, , IDUsuarioLogado)
  .Parameters.Append .CreateParameter("P_VALORTOTAL", adCurrency, adParamInput, , NzDbl(txtValorTotal.Value))

End With


Set rs = cmd.Execute


If rs.EOF Then

    MsgBox "Não foi possível registrar a compra.", vbExclamation
    GoTo Limpar

End If


IdCompraGerada = rs("IDCOMPRA").Value


'================================================
' ENVIA PARA CONFERENCIA XML
'================================================
MsgBox TipoEntrada
If TipoEntrada = "XML" Then

    frmConferencia.TipoEntrada = TipoEntrada
    frmConferencia.NumeroNFE_XML = NumeroNFE_XML
    frmConferencia.IdCompra = IdCompraGerada

    frmConferencia.Show vbModeless


'================================================
' ENTRADA MANUAL
'================================================

Else

    frmItensCompra.TipoEntrada = TipoEntrada
    frmItensCompra.txtIdCompra.Value = IdCompraGerada

    frmItensCompra.Show vbModeless

End If


Unload Me


Limpar:

If Not rs Is Nothing Then rs.Close

Set rs = Nothing
Set cmd = Nothing

Exit Sub


Erro:

MsgBox "Erro ao registrar compra:" & vbCrLf & Err.Description, vbCritical

Resume Limpar


End Sub
Private Sub btnFechar_Click()

    '====================================================
    ' FECHA O FORMULÁRIO
    '====================================================
    Unload Me

End Sub

 '====================================================
' CADASTRAR NOVO FORNECEDOR
'====================================================


Private Sub btnNovoForn_Click()

    frmCadFornecedor.Show vbModeless

End Sub
