VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPagamentos 
   Caption         =   "ENTRADAS DE CAIXA"
   ClientHeight    =   5595
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7020
   ' OleObjectBlob removido na versao publica
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPagamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public idvenda As Long

Private Sub Userform_activate()

    Static Carregado As Boolean

    If Not Carregado Then
        CarregarPagamentos idvenda
        Carregado = True
    End If

End Sub

Private Sub UserForm_initialize()

 With lvPagamentos

        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .HideSelection = False

        .ColumnHeaders.Clear

        .ColumnHeaders.Add , , "ID", 40
        .ColumnHeaders.Add , , "ID VENDA", 50
        .ColumnHeaders.Add , , "FORMA PAGAMENTO", 100
        .ColumnHeaders.Add , , "VALOR PAGO", 60
        .ColumnHeaders.Add , , "DATA", 110

    End With

End Sub

Public Sub CarregarPagamentos(ByVal idvenda As Long)

    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim item As ListItem

    sql = "SELECT " & _
          "IDPAGAMENTO," & _
          "IDVENDA," & _
          "FORMAPAGAMENTO," & _
          "VALORPAGO," & _
          "DATAPAGAMENTO " & _
          "FROM TAB_PAGAMENTOS " & _
          "WHERE IDVENDA = " & idvenda & " " & _
          "AND VALORPAGO > 0 " & _
          "ORDER BY IDPAGAMENTO;"

    Set rs = New ADODB.Recordset
    rs.Open sql, Conn, adOpenForwardOnly, adLockReadOnly

    lvPagamentos.ListItems.Clear

    Do Until rs.EOF

        Set item = lvPagamentos.ListItems.Add(, , rs!IdPagamento)

        item.SubItems(1) = rs!idvenda
        item.SubItems(2) = rs!FORMAPAGAMENTO
        item.SubItems(3) = Format(rs!ValorPago, "R$ #,##0.00")
        item.SubItems(4) = Format(rs!DATAPAGAMENTO, "dd/mm/yyyy")

        rs.MoveNext

    Loop

    rs.Close
    Set rs = Nothing

End Sub
Private Sub btnEstornar_Click()

    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim Retorno As String

    Dim IdPagamento As Long
    Dim valor As Variant
    Dim Motivo As String

    If lvPagamentos.SelectedItem Is Nothing Then
        MsgBox "Selecione um pagamento.", vbExclamation
        Exit Sub
    End If

    IdPagamento = CLng(lvPagamentos.SelectedItem.Text)

    valor = InputBox("Informe o valor do estorno:", "Estorno")
    If Trim(valor & "") = "" Then Exit Sub

    If Not IsNumeric(valor) Then
        MsgBox "Valor inválido.", vbExclamation
        Exit Sub
    End If

    Motivo = InputBox("Informe o motivo do estorno:", "Motivo")
    If Trim(Motivo) = "" Then Exit Sub

    If Not SolicitarSenhaCaixa() Then Exit Sub

    sql = "CALL PROC_ESTORNAR(" & _
            IDUsuarioLogado & "," & _
            IdPagamento & "," & _
            SqlNumero(valor) & "," & _
            SqlTexto(Motivo) & ");"
Debug.Print sql

    Set rs = New ADODB.Recordset
    rs.Open sql, Conn, adOpenForwardOnly, adLockReadOnly

    If Not rs.EOF Then
        MsgBox Nz(rs!Retorno), vbInformation
    End If

    rs.Close
    Set rs = Nothing

frmPagamento.AtualizarFinanceiro

    Unload Me

End Sub

Private Sub btnCancelar_Click()

    Unload Me

End Sub
