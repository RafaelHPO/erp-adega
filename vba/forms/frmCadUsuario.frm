VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCadUsuario 
   Caption         =   "CADASTRO DE USÚARIO"
   ClientHeight    =   4830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8010
   ' OleObjectBlob removido na versao publica
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCadUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub userform_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me
    End If

End Sub

Private Sub Userform_activate()

    If Trim(Me.Tag) <> "" Then

        CarregarUsuario

        btnSalvar.Visible = False
        btnSalvarEdit.Visible = True

    Else

        btnSalvar.Visible = True
        btnSalvarEdit.Visible = False

    End If

End Sub

Private Sub btnSalvar_Click()

    '--------------------------------------------------
    ' Variáveis
    '--------------------------------------------------
    Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command

    With cmd

        Set .ActiveConnection = Conn

        .CommandType = adCmdStoredProc

        .CommandText = "PROC_CADASTRARUSUARIO"

        ' Nome completo
        .Parameters.Append .CreateParameter("P_NOME", adVarChar, adParamInput, 40, txtNome.Value)

        ' Login
        .Parameters.Append .CreateParameter("P_USUARIO", adVarChar, adParamInput, 20, txtUsuario.Value)

        ' Senha
        .Parameters.Append .CreateParameter("P_SENHA", adVarChar, adParamInput, 30, txtSenha.Value)

        .Execute

    End With

    MsgBox "Usuário cadastrado com sucesso!", vbInformation

    Set cmd = Nothing

frmUsuarios.CarregarUsuarios

Me.Hide

End Sub

Private Sub btnCancelar_Click()

    ' Limpa os campos
    Call LimparCampos

    ' Fecha o formulário
   Unload Me

End Sub

Private Sub UserForm_initialize()
    
    btnSalvarEdit.Visible = False

End Sub
Public Sub LimparCampos()

    txtNome.Value = ""
    txtUsuario.Value = ""
    txtSenha.Value = ""

End Sub

Private Sub CarregarUsuario()

    Dim rs As ADODB.Recordset
    Dim sql As String

    sql = "SELECT NOME, USUARIO FROM TAB_USUARIOS WHERE IDUSUARIO = " & Me.Tag

    Set rs = New ADODB.Recordset
    rs.Open sql, Conn, adOpenStatic, adLockReadOnly

    If Not rs.EOF Then
        txtNome = rs!NOME
        txtUsuario = rs!Usuario
    End If

    rs.Close
    Set rs = Nothing

End Sub


Private Sub btnSalvarEdit_Click()

    Dim sql As String

    sql = "UPDATE TAB_USUARIOS SET " & _
          "NOME = '" & txtNome & "', " & _
          "USUARIO = '" & txtUsuario & "', " & _
          "SENHA = SHA2('" & txtSenha & "', 256) " & _
          "WHERE IDUSUARIO = " & Me.Tag

    Conn.Execute sql

    MsgBox "Usuário atualizado com sucesso!", vbInformation
    
Conn.Execute sql

frmUsuarios.CarregarUsuarios

Unload Me

End Sub


