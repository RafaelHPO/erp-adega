VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLogin 
   Caption         =   "Login - ERP Comercial"
   ClientHeight    =   7245
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6825
   ' OleObjectBlob removido na versao publica
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_initialize()

    lblSistema.Caption = "SISTEMA DEMO"

    If Conn.State = 1 Then

        lblBanco.Caption = "Banco conectado"

    Else

        lblBanco.Caption = "Sem conexão"

    End If

Application.Visible = False

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

  If CloseMode = vbFormControlMenu Then
       Cancel = True
       Call FecharSistema
   End If

End Sub

Private Sub btnEntrar_Click()

    Dim rs As ADODB.Recordset
    Dim sql As String

    sql = "CALL PROC_LOGIN('" & _
          txtUsuario.Value & "','" & _
          txtSenha.Value & "')"

    Set rs = Conn.Execute(sql)

    If rs("STATUS") = 1 Then

    UsuarioLogado = rs("USUARIO")

    IDUsuarioLogado = rs("IDUSUARIO")


    MsgBox rs("MSG") & vbCrLf & vbCrLf & _
    "Bem vindo! " & Nz(rs!NOME), vbInformation, "Login"

    Me.Hide

    frmPrincipal.Show vbModeless

    Unload Me

Else

    MsgBox rs("MSG"), vbExclamation, "Acesso negado"

End If

    rs.Close
    Set rs = Nothing

End Sub

Private Sub imgGear_Click()

Call manutencao

End Sub
