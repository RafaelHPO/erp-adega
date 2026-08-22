Attribute VB_Name = "modSistema"
Public gAbrirPagamento As Boolean
Public gIdVenda As Long
Public tela As String
Public nErro As Long
Public DescErro As String

Private Sub Workbook_Open()
    VerificarFila
End Sub

Public Sub AguardarSegundos(ByVal Segundos As Double)

    Dim Inicio As Double
    Inicio = Timer

    Do While Timer - Inicio < Segundos
        DoEvents
    Loop

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

Public Function SolicitarSenhaCaixa() As Boolean

    'frmSenha.supervisor = True
    frmSenha.admin = False

    frmSenha.SenhaValida = False
    frmSenha.txtSenha.Value = ""

    frmSenha.Show vbModal

    SolicitarSenhaCaixa = frmSenha.SenhaValida

End Function

Public Function SolicitarSenhaAdmin() As Boolean

    frmSenha.admin = True
    frmSenha.supervisor = False

    frmSenha.SenhaValida = False
    frmSenha.txtSenha.Value = ""

    frmSenha.Show vbModal

    SolicitarSenhaAdmin = frmSenha.SenhaValida

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
    Dim resposta As VbMsgBoxResult

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

    resposta = MsgBox(Resumo & vbCrLf & vbCrLf & _
                      "Deseja " & Operacao & "?", _
                      vbQuestion + vbYesNo, _
                      "Resumo do Caixa")

    MostrarResumoCaixa = (resposta = vbYes)

End Function

Public Sub manutencao()

    On Error GoTo TrataErro

    If Not SolicitarSenhaAdmin() Then Exit Sub

    LiberarModoAdmin

    Exit Sub

TrataErro:

    MsgBox "Erro: " & Err.Description, vbCritical, "SISTEMA"

End Sub

Public Sub ReportarErro()

    Dim sql As String

    tela = UCase(Trim(tela))
    etapa = UCase(Trim(etapa))
    DescErro = Trim(DescErro)

    sql = "INSERT INTO TAB_OCORRENCIAS " & _
          "(OCORRENCIA, TELA, DESCRICAO, STATUS, IDUSUARIO, VERSAO_SISTEMA) " & _
          "VALUES (" & _
          SqlNumero(nErro) & ", " & _
          SqlTexto(tela) & ", " & _
          SqlTexto(DescErro) & ", " & _
          "'PENDENTE', " & _
          IDUsuarioLogado & ", " & _
          SqlTexto(BuscarConfig("VERSAO_SISTEMA")) & ")"

    Conn.Execute sql

End Sub

Public Sub ReiniciarSistema()

    On Error GoTo TratarErro

    Dim caminho As String

    caminho = ThisWorkbook.FullName

    Shell """" & Application.Path & "\EXCEL.EXE"" """ & caminho & """", vbNormalFocus

    Application.Quit

    Exit Sub

TratarErro:

    MsgBox "Não foi possível reiniciar o sistema." & vbCrLf & _
           Err.Description, vbExclamation, "SISTEMA"

End Sub

