VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCadFornecedor 
   Caption         =   "CADASTRO DE FORNECEDOR"
   ClientHeight    =   4830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8010
   ' OleObjectBlob removido na versao publica
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCadFornecedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub userform_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me
    End If

End Sub

Private Sub UserForm_initialize()

btnSalvarEdit.Visible = False
btnCancelarEdit.Visible = False

End Sub

Private Sub Userform_activate()

ckAtivo.Value = True

    If Trim(Me.Tag) <> "" Then

        CarregarFornecedor

        btnSalvar.Visible = False
        btnSalvarEdit.Visible = True

    Else

        btnSalvar.Visible = True
        btnSalvarEdit.Visible = False

    End If

End Sub

Private Sub btnCancelar_Click()

    ' Limpa os campos
    Call LimparCampos

    ' Fecha o formulário
    Unload Me

End Sub
Private Sub LimparCampos()

   txtNome.Value = ""
txtCNPJ.Value = ""
txtContato.Value = ""
txtEndereco.Value = ""
txtNumero.Value = ""
txtBairro.Value = ""

End Sub


Private Sub btnSalvar_Click()

    '--------------------------------------------------
    ' Variáveis
    '--------------------------------------------------
    Dim cmd As ADODB.Command
    
    '====================================================
    ' DEFINE O STATUS PELO CHECKBOX
    '====================================================
    If ckAtivo.Value = True Then
        StatusFornecedor = "ATIVO"
    Else
        StatusFornecedor = "INATIVO"
    End If

    '--------------------------------------------------
    ' Cria objeto de comando
    '--------------------------------------------------
   
Dim rs As ADODB.Recordset

Set cmd = New ADODB.Command

With cmd

    Set .ActiveConnection = Conn
    .CommandType = adCmdStoredProc
    .CommandText = "PROC_CADASTRARFORNECEDOR"

    .Parameters.Append .CreateParameter("P_NOME", adVarChar, adParamInput, 80, NzDB(txtNome.Value))
    .Parameters.Append .CreateParameter("P_CNPJ", adVarChar, adParamInput, 24, NzDB(txtCNPJ.Value))
    .Parameters.Append .CreateParameter("P_CONTATO", adVarChar, adParamInput, 13, NzDB(txtContato.Value))
    .Parameters.Append .CreateParameter("P_ENDERECO", adVarChar, adParamInput, 100, NzDB(txtEndereco.Value))
    .Parameters.Append .CreateParameter("P_NUMERO", adVarChar, adParamInput, 10, NzDB(txtNumero.Value))
    .Parameters.Append .CreateParameter("P_BAIRRO", adVarChar, adParamInput, 50, NzDB(txtBairro.Value))
    .Parameters.Append .CreateParameter("P_STATUS", adVarChar, adParamInput, 15, NzDB(StatusFornecedor))
    .Parameters.Append .CreateParameter("P_IDUSUARIO", adInteger, adParamInput, , NzDB(IDUsuarioLogado))

    Set rs = .Execute

End With

If Not rs Is Nothing Then
    If Not rs.EOF Then
      
        MsgBox rs.Fields(0).Value
    End If
    rs.Close
End If

Set rs = Nothing
Set cmd = Nothing
frmFornecedores.CarregarFornecedores

Unload Me

End Sub


Private Sub CarregarFornecedor()

    Dim rs As ADODB.Recordset
    Dim sql As String

    sql = "SELECT IDFORNECEDOR, NOME, CNPJ, CONTATO, ENDERECO, " & _
          "NUMERO, BAIRRO, STATUS " & _
          "FROM TAB_FORNECEDORES " & _
          "WHERE IDFORNECEDOR = " & CLng(Me.Tag)

    Set rs = New ADODB.Recordset
    rs.Open sql, Conn, adOpenStatic, adLockReadOnly

    If Not rs.EOF Then

        txtNome.Value = Nz(rs!NOME)
        txtCNPJ.Value = Nz(rs!CNPJ)
        txtContato.Value = Nz(rs!CONTATO)
        txtEndereco.Value = Nz(rs!ENDERECO)
        txtNumero.Value = Nz(rs!NUMERO)
        txtBairro.Value = Nz(rs!BAIRRO)

        ckAtivo.Value = (Nz(rs!STATUS) = "ATIVO")

    End If

    rs.Close
    Set rs = Nothing

End Sub

Private Sub btnSalvarEdit_Click()

    Dim sql As String
    Dim StatusFornecedor As String

    If ckAtivo.Value Then
        StatusFornecedor = "ATIVO"
    Else
        StatusFornecedor = "INATIVO"
    End If

    sql = "UPDATE TAB_FORNECEDORES SET " & _
          "NOME = '" & Replace(txtNome.Value, "'", "''") & "', " & _
          "CNPJ = '" & txtCNPJ.Value & "', " & _
          "CONTATO = '" & txtContato.Value & "', " & _
          "ENDERECO = '" & Replace(txtEndereco.Value, "'", "''") & "', " & _
          "NUMERO = '" & txtNumero.Value & "', " & _
          "BAIRRO = '" & Replace(txtBairro.Value, "'", "''") & "', " & _
          "STATUS = '" & StatusFornecedor & "' " & _
          "WHERE IDFORNECEDOR = " & CLng(Me.Tag)

    Conn.Execute sql

    MsgBox "Fornecedor atualizado com sucesso!", vbInformation

Conn.Execute sql

frmFornecedores.CarregarFornecedores

Unload Me

End Sub

Private Sub btnCancelarEdit_Click()

    ' Fecha o formulário
    Unload Me

End Sub

