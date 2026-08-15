VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAcerto 
   Caption         =   "ACERTO DE ESTOQUE"
   ClientHeight    =   5655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8355.001
   ' OleObjectBlob removido na versao publica
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAcerto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub userform_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me
    End If

End Sub

Private Sub btnCancelar_Click()

    Unload Me

End Sub

Private Sub btnSalvar_Click()

    Dim sql As String
    Dim rs As ADODB.Recordset

    If cmbProduto.ListIndex = -1 Then
        MsgBox "Selecione um produto.", vbExclamation
        Exit Sub
    End If

    If Val(txtQuantidade.Value) <= 0 Then
        MsgBox "Quantidade inválida.", vbExclamation
        Exit Sub
    End If
    
    If txtMotivo.Value = "" Then
    MsgBox "Descreva o motivo do acerto!", vbExclamation
    Exit Sub
    End If

   sql = "CALL PROC_ACERTOESTOQUE(" & _
      NzDbl(IDUsuarioLogado) & "," & _
      NzDbl(cmbProduto.Column(0)) & "," & _
      NzDbl(txtQuantidade.Value) & "," & _
      SqlTexto(cmbAcerto.Value) & "," & _
      SqlTexto(txtMotivo.Value) & ")"
      
    Set rs = Conn.Execute(sql)

    If Not rs Is Nothing Then

        If Not rs.EOF Then

            MsgBox rs.Fields(0).Value, vbInformation

        End If

        rs.Close

    End If

    Set rs = Nothing

    txtQuantidade.Value = ""
    txtMotivo.Value = ""
    cmbAcerto.ListIndex = -1
    cmbProduto.ListIndex = -1
    txtAtual.Value = ""
            txtPosterior.Value = ""

    cmbProduto.SetFocus
frmProdutos.CarregarProdutos


End Sub

Private Sub UserForm_initialize()

    CarregarProdutosCombo cmbProduto

    cmbAcerto.List = Array( _
        "ENTRADA", _
        "SAIDA" _
    )

End Sub

Private Sub txtEAN5_AfterUpdate()

    SelecionarProdutoPorEAN5 cmbProduto, txtEAN5.Value

cmbProduto.SetFocus

End Sub

Private Sub cmbProduto_afterupdate()

 Dim rs As ADODB.Recordset
    Dim sql As String
    Dim IDProduto As Long

    If cmbProduto.ListIndex = -1 Then Exit Sub

    IDProduto = CLng(cmbProduto.List(cmbProduto.ListIndex, 0))

    sql = "SELECT ESTOQUEATUAL,TIPO " & _
          "FROM TAB_PRODUTOS " & _
          "WHERE IDPRODUTO = " & IDProduto

    Set rs = Conn.Execute(sql)

    If Not rs.EOF Then

 If UCase(Nz(rs!tipo, "")) = "COMBO" Then

            MsgBox "Este produto é um COMBO." & vbCrLf & _
                   "Selecione um produto individual para continuar.", _
                   vbExclamation, "Produto inválido"

            cmbProduto.Value = Null
            txtAtual.Value = ""
            txtPosterior.Value = ""
            txtQuantidade.Value = ""
            txtEAN5.Value = ""
            cmbProduto.SetFocus

            rs.Close
            Set rs = Nothing
            Exit Sub

        End If

        txtAtual.Value = IIf(IsNull(rs!ESTOQUEATUAL), 0, rs!ESTOQUEATUAL)

    End If

    rs.Close
    Set rs = Nothing
    
    txtQuantidade.SetFocus

End Sub

Private Sub cmbAcerto_afterupdate()

    Dim posterior As Currency
    
    If cmbAcerto.Value = "ENTRADA" Then
    
posterior = NzDbl(txtAtual.Value) + NzDbl(txtQuantidade.Value)

Else

posterior = NzDbl(txtAtual.Value) - NzDbl(txtQuantidade.Value)

End If

    txtPosterior.Value = posterior
    
 txtMotivo.SetFocus

End Sub

Private Sub txtquantidade_afterupdate()

    Dim posterior As Currency
    
    If cmbAcerto.Value = "ENTRADA" Then
    
posterior = NzDbl(txtAtual.Value) + NzDbl(txtQuantidade.Value)

Else

posterior = NzDbl(txtAtual.Value) - NzDbl(txtQuantidade.Value)

End If

    txtPosterior.Value = posterior
    
 txtMotivo.SetFocus

End Sub
