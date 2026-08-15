Attribute VB_Name = "modSistema"
Public gAbrirPagamento As Boolean
Public gIdVenda As Long

Private Sub Workbook_Open()
    VerificarFila
End Sub
Public Function BuscarConfig(Chave As String) As String

    Dim rs As ADODB.Recordset
    Dim sql As String

    sql = "SELECT VALOR FROM TAB_CONFIG WHERE CHAVE = " & SqlTexto(Chave)

    Set rs = Conn.Execute(sql)

    If Not rs.EOF Then
        BuscarConfig = rs!valor
    Else
        BuscarConfig = ""
    End If

    rs.Close
    Set rs = Nothing

End Function

Public Sub VerificarFila()

    On Error Resume Next

    If gAbrirPagamento = True Then

        gAbrirPagamento = False

        Dim f As frmPagamento
        Set f = New frmPagamento

        f.idvenda = gIdVenda
        f.Show vbModeless

    End If

    Application.OnTime Now + TimeValue("00:00:01"), "VerificarFila"

End Sub

Public Sub FormatarMoeda(Txt As MSForms.TextBox)

    Static travar As Boolean
    If travar Then Exit Sub

    travar = True

    Dim v As String

    v = Replace(Txt.Text, ".", "")
    v = Replace(v, ",", "")

    If IsNumeric(v) And v <> "" Then
        Txt.Text = Format(CDbl(v) / 100, "#,##0.00")
        Txt.SelStart = Len(Txt.Text)
    End If

    travar = False

End Sub

Public Sub FormatarData(Txt As MSForms.TextBox)

    Static travar As Boolean
    If travar Then Exit Sub
    travar = True

    Dim s As String

    s = Replace(Txt.Text, "/", "")

    If Len(s) > 2 Then s = Left(s, 2) & "/" & Mid(s, 3)
    If Len(s) > 5 Then s = Left(s, 5) & "/" & Mid(s, 6)

    Txt.Text = Left(s, 10)
    Txt.SelStart = Len(Txt.Text)

    travar = False

End Sub

Public Sub FormatarDatacmb(cmb As MSForms.ComboBox)

    Static travar As Boolean
    If travar Then Exit Sub
    travar = True

    Dim s As String

    s = Replace(cmb.Text, "/", "")

    If Len(s) > 2 Then s = Left(s, 2) & "/" & Mid(s, 3)
    If Len(s) > 5 Then s = Left(s, 5) & "/" & Mid(s, 6)

    cmb.Text = Left(s, 10)
    cmb.SelStart = Len(cmb.Text)

    travar = False

End Sub

Public Sub FormatarCNPJ(Txt As MSForms.TextBox)

    Static travar As Boolean
    If travar Then Exit Sub
    travar = True

    Dim s As String

    s = Replace(Txt.Text, ".", "")
    s = Replace(s, "/", "")
    s = Replace(s, "-", "")

    If Len(s) > 2 Then s = Left(s, 2) & "." & Mid(s, 3)
    If Len(s) > 5 Then s = Left(s, 6) & "." & Mid(s, 6)
    If Len(s) > 8 Then s = Left(s, 10) & "/" & Mid(s, 9)
    If Len(s) > 12 Then s = Left(s, 15) & "-" & Mid(s, 13)

    Txt.Text = Left(s, 18)
    Txt.SelStart = Len(Txt.Text)

    travar = False

End Sub

Public Sub FecharSistema()

    On Error Resume Next

    Application.DisplayAlerts = False

    ' Fecha conexão com banco se existir
    If Not Conn Is Nothing Then
        
        If Conn.State = 1 Then
            Conn.Close
        End If
        
        Set Conn = Nothing
        
    End If

    ' Mostra Excel antes de encerrar
    Application.Visible = True

    ' Evita perguntar para salvar
    ThisWorkbook.Saved = True

    Application.Quit

End Sub
 
Public Sub LiberarModoAdmin()

    Dim uf As Object

    'Fecha todos os UserForms abertos
    For Each uf In VBA.UserForms
        Unload uf
    Next uf

    Application.Visible = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub
Public Sub VoltarModoSistema()

    Dim uf As Object

    'Fecha todos os UserForms abertos
    For Each uf In VBA.UserForms
        Unload uf
    Next uf

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    'Fecha o editor VBA se estiver aberto
    On Error Resume Next
    Application.VBE.MainWindow.Close
    On Error GoTo 0

    'Valida usuário logado
    If IDUsuarioLogado = 0 Or IsNull(IDUsuarioLogado) Then
        
        Application.Visible = False
        frmLogin.Show
        
    Else
        
        Application.Visible = False
        frmPrincipal.Show
        
    End If

End Sub

Public Function SolicitarSenhaCaixa() As Boolean

    frmSenha.SenhaValida = False
    frmSenha.txtSenha.Value = ""

    frmSenha.Show vbModal

    SolicitarSenhaCaixa = frmSenha.SenhaValida

End Function

Public Function ValidarSenha(Senha As String) As Boolean

    Dim rs As ADODB.Recordset
    Dim sql As String

    sql = "SELECT COUNT(*) AS TOTAL " & _
          "FROM tab_usuarios " & _
          "WHERE IDUSUARIO = " & IDUsuarioLogado & _
          " AND SENHA = SHA2('" & Replace(Senha, "'", "''") & "',256)"

    Set rs = Conn.Execute(sql)

    ValidarSenha = (Nz(rs!Total, 0) > 0)

    rs.Close
    Set rs = Nothing

End Function

Public Function ConfirmarSenha() As Boolean

    Dim Senha As String

    Senha = InputBox("Digite sua senha para continuar:", "Confirmação")

    If Len(Trim(Senha)) = 0 Then Exit Function

    ConfirmarSenha = ValidarSenha(Senha)

    If Not ConfirmarSenha Then
        MsgBox "Senha inválida.", vbExclamation
    End If

End Function

Public Function MostrarResumoCaixa(Optional Operacao As String = "") As Boolean

    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim Resumo As String
    Dim Resposta As VbMsgBoxResult

    sql = "CALL PROC_RESUMOCAIXA()"

    Set rs = Conn.Execute(sql)

    If rs.EOF Then Exit Function

    If CampoExiste(rs, "ERRO") Then

        MsgBox rs("ERRO"), vbExclamation

        rs.Close
        Set rs = Nothing

        Exit Function

    End If

    Resumo = "========== RESUMO DO CAIXA ==========" & vbCrLf & vbCrLf

    Resumo = Resumo & _
             "PIX.................. " & Format(rs("PIX"), "R$ #,##0.00") & vbCrLf & _
             "CARTÃO............... " & Format(rs("CARTAO"), "R$ #,##0.00") & vbCrLf & _
             "DINHEIRO............. " & Format(rs("DINHEIRO"), "R$ #,##0.00") & vbCrLf & vbCrLf & _
             "FUNDO INICIAL........ " & Format(rs("FUNDO"), "R$ #,##0.00") & vbCrLf & _
             "SANGRIAS............ " & Format(rs("SANGRIA"), "R$ #,##0.00") & vbCrLf & _
             "CAIXA ATUAL......... " & Format(rs("CAIXA"), "R$ #,##0.00") & vbCrLf & _
             "DINHEIRO ESPERADO... " & Format(rs("DINHEIRO ESPERADO"), "R$ #,##0.00")

    rs.Close
    Set rs = Nothing

    If Operacao = "" Then Operacao = "continuar"

    Resposta = MsgBox(Resumo & vbCrLf & vbCrLf & _
                      "Deseja " & Operacao & "?", _
                      vbQuestion + vbYesNo, _
                      "Resumo do Caixa")

    MostrarResumoCaixa = (Resposta = vbYes)

End Function

Public Sub manutencao()

    On Error GoTo TrataErro

    Dim Senha As String
    Dim sql As String
    Dim rs As ADODB.Recordset

    Senha = InputBox("Senha de administrador:")

    If Trim(Senha) = "" Then
        MsgBox "DIGITE UMA SENHA", vbInformation
        Exit Sub
    End If

    sql = "SELECT COUNT(*) AS ACESSO " & _
          "FROM TAB_CONFIG " & _
          "WHERE CHAVE = 'SENHA_ADMIN' " & _
          "AND VALOR = SHA2('" & Replace(Senha, "'", "''") & "', 256)"

    Set rs = Conn.Execute(sql)

    If Not rs.EOF Then

        If Nz(rs!ACESSO, 0) < 1 Then

            MsgBox "Senha inválida!", vbExclamation

        Else

            LiberarModoAdmin

        End If

    End If

    rs.Close
    Set rs = Nothing

    Exit Sub

TrataErro:

    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
    End If

    Set rs = Nothing

    MsgBox "Erro: " & Err.Description, vbCritical

End Sub

Public Sub ResetarSistema()

    On Error Resume Next
    
    
    'Fecha formulários abertos
    Dim frm As Object
    
    For Each frm In VBA.UserForms
        Unload frm
    Next frm
    
    
    'Limpa objetos globais se existirem
    Set rs = Nothing
    
    
    'Limpa erro atual
    Err.Clear
    
    
    'Volta tratamento normal
    On Error GoTo 0
    
    
    'Reabre tela principal

End Sub

