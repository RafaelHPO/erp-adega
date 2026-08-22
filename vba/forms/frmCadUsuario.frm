VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCadUsuario 
   Caption         =   "CADASTRO DE USÚARIO"
   ClientHeight    =   4095
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6405
   OleObjectBlob   =   "frmCadUsuario.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCadUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AplicarPaleta()
    
    AplicarTema Me
    
    AplicarBotaoPrincipal btnSalvar

End Sub

Private Sub UserForm_initialize()
    
    btnSalvarEdit.Visible = False
    AplicarPaleta

End Sub

Private Sub userform_activate()

    On Error GoTo TratarErro

    If Trim(Me.Tag) <> "" Then

        CarregarUsuario

        btnSalvar.Visible = False
        btnSalvarEdit.Visible = True

    Else

        btnSalvar.Visible = True
        btnSalvarEdit.Visible = False

    End If

    Exit Sub

TratarErro:

    modSistema.tela = "CadUsuario - activate"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"

End Sub

Private Sub btnSalvar_Click()

    On Error GoTo TratarErro

    Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command

    With cmd

        Set .ActiveConnection = Conn

        .CommandType = adCmdStoredProc

        .CommandText = "PROC_CADASTRARUSUARIO"
        
        .Parameters.Append .CreateParameter("P_NOME", adVarChar, adParamInput, 40, txtNome.Value)
        .Parameters.Append .CreateParameter("P_USUARIO", adVarChar, adParamInput, 20, txtUsuario.Value)
        .Parameters.Append .CreateParameter("P_SENHA", adVarChar, adParamInput, 30, txtSenha.Value)
        
        .Execute

    End With

    MsgBox "Usuário cadastrado com sucesso!", vbInformation

    Set cmd = Nothing

frmUsuarios.CarregarUsuarios

Unload Me

    Exit Sub

TratarErro:

    modSistema.tela = "CadUsuario - btnSalvar"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"

End Sub

Private Sub btnCancelar_Click()

    Call LimparCampos
   Unload Me

End Sub

Public Sub LimparCampos()

    txtNome.Value = ""
    txtUsuario.Value = ""
    txtSenha.Value = ""

End Sub

Private Sub CarregarUsuario()

On Error GoTo TratarErro

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

    Exit Sub

TratarErro:

    modSistema.tela = "CadUsuario - CarregarUsuario"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"

End Sub


Private Sub btnSalvarEdit_Click()

    On Error GoTo TratarErro

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

        Exit Sub

TratarErro:

    modSistema.tela = "CadUsuario - btnSalvarEdit"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"

End Sub


