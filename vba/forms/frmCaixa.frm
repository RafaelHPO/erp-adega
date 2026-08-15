VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCaixa 
   Caption         =   "HISTÓRICO DE CAIXA"
   ClientHeight    =   8415.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13755
   ' OleObjectBlob removido na versao publica
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lvCaixa_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub userform_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me
    End If

End Sub

Private Sub btnAbrirCaixa_Click()

 If Not SolicitarSenhaCaixa() Then Exit Sub

    '====================================================
    ' VALIDA SE JÁ EXISTE CAIXA ABERTO
    '====================================================
    If ExisteCaixaAberto Then

        MsgBox "Já existe um caixa aberto.", vbExclamation
        Exit Sub

    End If

    '====================================================
    ' DECLARAÇÃO DAS VARIÁVEIS
    '====================================================
    Dim rs As ADODB.Recordset
    Dim sql As String

    Dim Fundo As Double
    Dim Retorno As String

    '====================================================
    ' SOLICITA O FUNDO INICIAL
    '====================================================
    Retorno = InputBox("Informe o fundo inicial do caixa:")

    If Len(Trim(Retorno)) = 0 Then Exit Sub

    If Not IsNumeric(Retorno) Then

        MsgBox "Informe um valor numérico válido.", vbExclamation
        Exit Sub

    End If

    Fundo = CDbl(Retorno)

    If Fundo < 0 Then

        MsgBox "O fundo inicial não pode ser negativo.", vbExclamation
        Exit Sub

    End If

    '====================================================
    ' EXECUTA PROCEDURE
    '====================================================
    sql = "CALL PROC_ABRIRCAIXA(" & _
          IDUsuarioLogado & "," & _
          Replace(Fundo, ",", ".") & ")"

    Set rs = Conn.Execute(sql)

    If Not rs.EOF Then

        If CampoExiste(rs, "ERRO") Then

            MsgBox rs("ERRO"), vbExclamation

        Else

            frmPrincipal.txtStatusCaixa.Value = "ABERTO"
            frmPrincipal.txtIdCaixa.Value = rs("IDCAIXA")

            MsgBox "Caixa Nº " & rs("ID CAIXA") & _
                   " aberto com fundo de R$ " & _
                   Format(rs("FUNDO"), "0.00"), _
                   vbInformation

        End If

    End If

    rs.Close
    Set rs = Nothing

End Sub

Private Sub btnSangria_Click()

 If Not SolicitarSenhaCaixa() Then Exit Sub

    If Not MostrarResumoCaixa("realizar a sangria") Then Exit Sub

    Dim rs As ADODB.Recordset
    Dim sql As String

    Dim IdCaixa As Long
    Dim valor As Double

    Dim RetornoValor As String

    IdCaixa = ObterCaixaAberto()

    If IdCaixa = 0 Then

        MsgBox "Nenhum caixa aberto.", vbExclamation
        Exit Sub

    End If

    RetornoValor = InputBox("Informe o valor da sangria:")

    If Len(Trim(RetornoValor)) = 0 Then Exit Sub

    If Not IsNumeric(RetornoValor) Then

        MsgBox "Informe um valor numérico válido.", vbExclamation
        Exit Sub

    End If

    valor = CDbl(RetornoValor)

    sql = "CALL PROC_SANGRIACAIXA(" & _
          IdCaixa & "," & _
          Replace(valor, ",", ".") & "," & _
          IDUsuarioLogado & ")"

    Set rs = Conn.Execute(sql)

    If Not rs.EOF Then

        If CampoExiste(rs, "ERRO") Then

            MsgBox rs("ERRO"), vbExclamation

        Else

            MsgBox rs("MSG") & vbCrLf & _
                   "Valor retirado: R$ " & _
                   Format(rs("VALOR RETIRADO"), "0.00"), vbInformation

        End If

    End If

    rs.Close
    Set rs = Nothing

End Sub

Private Sub btnFecharCaixa_Click()

If Not SolicitarSenhaCaixa() Then Exit Sub

    If Not MostrarResumoCaixa("fechar o caixa") Then Exit Sub

    '====================================================
    ' DECLARAÇÃO DAS VARIÁVEIS
    '====================================================
    Dim rs As ADODB.Recordset
    Dim sql As String

    Dim IdCaixa As Long

    '====================================================
    ' BUSCA O CAIXA ABERTO
    '====================================================
    IdCaixa = ObterCaixaAberto()

    If IdCaixa = 0 Then

        MsgBox "Nenhum caixa aberto.", vbExclamation
        Exit Sub

    End If

    '====================================================
    ' EXECUTA PROCEDURE
    '====================================================
    sql = "CALL PROC_FECHARCAIXA(" & _
          IdCaixa & "," & _
          IDUsuarioLogado & ")"

    Set rs = Conn.Execute(sql)

    If Not rs.EOF Then

        If CampoExiste(rs, "ERRO") Then

            MsgBox rs("ERRO"), vbExclamation

        Else

            frmPrincipal.txtStatusCaixa.Value = "FECHADO"
            frmPrincipal.txtIdCaixa.Value = ""

            MsgBox rs("MSG") & vbCrLf & _
                   "Total Sangrias: R$ " & _
                   Format(rs("VALOR TOTAL"), "0.00") & vbCrLf & _
                   "Saldo Caixa: R$ " & _
                   Format(rs("SOBRA"), "0.00"), vbInformation

        End If

    End If

    rs.Close
    Set rs = Nothing

End Sub

Private Sub CarregarHistoricoCaixa()

    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim item As ListItem

    lvCaixa.ListItems.Clear

    sql = "SELECT " & _
          "C.IDCAIXA, " & _
          "C.DATAABERTURA, " & _
          "U.USUARIO AS USUARIO, " & _
          "C.FUNDO, " & _
          "C.SANGRIA, " & _
          "C.VALORCAIXA, " & _
          "(C.FUNDO + C.VALORCAIXA) AS DINHEIROESPERADO, " & _
          "C.STATUS " & _
          "FROM TAB_CAIXA C " & _
          "INNER JOIN TAB_USUARIOS U " & _
          "ON U.IDUSUARIO = C.IDUSUARIO " & _
          "ORDER BY C.IDCAIXA DESC"

    Set rs = Conn.Execute(sql)

    Do While Not rs.EOF

        Set item = lvCaixa.ListItems.Add(, , rs("IDCAIXA"))

        item.SubItems(1) = Format(rs("DATAABERTURA"), "dd/mm/yyyy")
        item.SubItems(2) = rs("USUARIO")

        item.SubItems(3) = Format(rs("FUNDO"), "0.00")
        item.SubItems(4) = Format(rs("SANGRIA"), "0.00")
        item.SubItems(5) = Format(rs("VALORCAIXA"), "0.00")
        item.SubItems(6) = Format(rs("DINHEIROESPERADO"), "0.00")
        item.SubItems(7) = rs("STATUS")

        rs.MoveNext

    Loop

    rs.Close
    Set rs = Nothing

End Sub
Private Sub Userform_activate()

    '====================================================
    ' CONFIGURAÇÃO DO LISTVIEW
    '====================================================
    With lvCaixa

        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .HideColumnHeaders = False

        .ListItems.Clear
        .ColumnHeaders.Clear

        '================================================
        ' CABEÇALHOS
        '================================================
        .ColumnHeaders.Add , , "ID", 40
        .ColumnHeaders.Add , , "DATA ABERTURA", 70
        .ColumnHeaders.Add , , "USUÁRIO", 110
        .ColumnHeaders.Add , , "FUNDO", 70
        .ColumnHeaders.Add , , "SANGRIA", 70
        .ColumnHeaders.Add , , "VALOR CAIXA", 80
        .ColumnHeaders.Add , , "ESPERADO", 70
        .ColumnHeaders.Add , , "STATUS", 70

    End With

    '====================================================
    ' CARREGA DADOS
    '====================================================
    CarregarHistoricoCaixa

End Sub
