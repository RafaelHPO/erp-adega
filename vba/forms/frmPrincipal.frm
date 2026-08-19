VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPrincipal 
   Caption         =   "ERP Adega"
   ClientHeight    =   12015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20955
   OleObjectBlob   =   "frmPrincipal.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public ProximoBackup As Date
Public xlDashboard As Object
Public wbDashboard As Object

Public Sub ExecutarBackup(Optional ByVal ExibirMensagem As Boolean = False)

    On Error Resume Next

    Dim Shell As Object

    Set Shell = CreateObject("WScript.Shell")

    '0 = oculto
    Shell.Run """C:\SISTEMA\backup.bat""", 0, True

    Set Shell = Nothing

    If ExibirMensagem Then
        MsgBox "Backup concluído com sucesso!", vbInformation
    End If

End Sub

Public Sub AgendarBackups()

    Dim Hoje As Date
    Dim Horario As Date

    Hoje = Date

    If Time < TimeValue("15:00:00") Then

        Horario = Hoje + TimeValue("15:00:00")

    ElseIf Time < TimeValue("22:00:00") Then

        Horario = Hoje + TimeValue("22:00:00")

    Else

        Horario = Hoje + 1 + TimeValue("15:00:00")

    End If

    ProximoBackup = Horario

    Application.OnTime ProximoBackup, "BackupPeriodico"

End Sub

Public Sub BackupPeriodico()

    ExecutarBackup False

    AgendarBackups

End Sub
Private Sub btnBackup_Click()

    fraConfig.Visible = False

    ExecutarBackup True

End Sub

Private Sub btnCategorias_Click()

fraEstoque.Visible = False
frmCategorias.Show vbModeless

End Sub

Private Sub btnDashboard_Click()

fraConfig.Visible = False
fraCadastro.Visible = False
fraOperacao.Visible = False
fraEstoque.Visible = False

    Dim caminho As String

    caminho = "C:\Sistema\DASHBOARD.xlsm"

    If Dir(caminho) = "" Then
        MsgBox "Dashboard não encontrado em:" & vbCrLf & caminho, vbExclamation
        Exit Sub
    End If

    If Not xlDashboard Is Nothing Then
        On Error Resume Next
        xlDashboard.Visible = True
        xlDashboard.WindowState = -4137 'xlMaximized
        xlDashboard.Activate
        If Err.Number = 0 Then Exit Sub
        On Error GoTo 0
    End If

    Set xlDashboard = CreateObject("Excel.Application")

    xlDashboard.Visible = True
    xlDashboard.DisplayAlerts = False
    xlDashboard.EnableEvents = True

    Set wbDashboard = xlDashboard.Workbooks.Open(caminho)

End Sub

Private Sub btnEstoque_Click()

    fraEstoque.Visible = Not fraEstoque.Visible

End Sub

Private Sub btnLog_Click()

     fraConfig.Visible = False
     
     frmLogEventos.Show vbModeless

End Sub

Private Sub btnManutencao_Click()

    fraConfig.Visible = False
    Call manutencao
    
End Sub

Private Sub btnGrupos_Click()
fraEstoque.Visible = False
frmGrupos.Show vbModeless

End Sub

Private Sub btnMovimento_Click()

fraEstoque.Visible = False
   fraCadastro.Visible = False
   frmMovimentoEstoque.Show vbModeless

End Sub

Private Sub btnReport_Click()

    frmOcorrencias.Show vbModeless
    
End Sub

Private Sub imgGear_Click()

    fraConfig.Visible = Not fraConfig.Visible
    fraCadastro.Visible = False
fraOperacao.Visible = False
fraEstoque.Visible = False

End Sub
Private Sub btnCompras_Click()
    
    fraOperacao.Visible = False
    
    frmEntradas.Show vbModeless
    
End Sub

Private Sub btnFinanceiro_Click()
 fraConfig.Visible = False
fraCadastro.Visible = False
fraOperacao.Visible = False
fraEstoque.Visible = False
frmFluxoCaixa.Show vbModeless

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

  If CloseMode = vbFormControlMenu Then
       Cancel = True
       Call FecharSistema
       On Error Resume Next
Application.OnTime ProximoBackup, "BackupPeriodico", , False

ExecutarBackup False
   End If

End Sub

Private Sub userform_activate()

    CarregarKpiPrincipal
    
    txtAtt.Value = " ---------Versão: 1.2.0--------- " & vbCrLf & _
                   " ADICIONADO BOTAO EDITAR EM PEDIDO DE VENDA, AJUSTE NA BORDA DAS TELAS, OTIMIZAÇÃO NA TELA DE VENDAS E LEGENDA DE AVISO. " & vbCrLf & vbCrLf & _
                   " ---------Versão: 1.1.0--------- " & vbCrLf & _
                   " FRMPRODUTOS INICIANDO EM BRANCO, CORRIGIDO. LOG DE DEPURAÇÃO DE ERROS NO BANCO, IMPLANTADO CALCULO DE ESTOQUE MÍNIMO. " & vbCrLf & vbCrLf & _
                   " ---------Versão: 1.0.4--------- " & vbCrLf & _
                   " FORMATAÇÃO DE DATA E VALORES EM TEXTBOX E LISTVIEW, BTNEDITAR EM CONTAS "

End Sub

Public Sub UserForm_Initialize()

ExecutarBackup False
AgendarBackups

  lblVersao.Caption = "Versão " & BuscarConfig("VERSAO_SISTEMA")

    lblAtualizacao.Caption = _
        "Atualizado em " & BuscarConfig("DATA_VERSAO")

    ' só UI aqui
    fraEstoque.Visible = False
    fraOperacao.Visible = False
    fraCadastro.Visible = False
 
    txtUsuario.Value = IDUsuarioLogado & " - " & UsuarioLogado
    fraConfig.Visible = False

Dim rs As ADODB.Recordset
    Dim sql As String

    sql = "SELECT IDCAIXA " & _
          "FROM TAB_CAIXA " & _
          "WHERE STATUS = 'ABERTO' " & _
          "LIMIT 1"

    Set rs = Conn.Execute(sql)

    If rs.EOF Then
       
        txtStatusCaixa.Value = "FECHADO"

    Else
 txtIdCaixa.Value = (rs!IdCaixa)
        txtStatusCaixa.Value = "ABERTO"

    End If

    rs.Close
    Set rs = Nothing

End Sub

Public Sub CarregarKpiPrincipal()

    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim ttMes As Double
    Dim qtdMes As Long
    Dim tmMes As Double
    Dim ttHoje As Double
    Dim qtdHoje As Long
    Dim tmHoje As Double
    Dim proxVenc As Date
    Dim forn As String
    Dim desc As String
    Dim titulos As Long
    Dim totaltitulos As Double
    Dim valorprox As Double
    
sql = "SELECT " & _
    "COALESCE(SUM(VALORFINAL),0) AS TOTALMES, " & _
    "COALESCE(COUNT(*),0) AS QTDEMES, " & _
    "COALESCE(ROUND(SUM(VALORFINAL) / IFNULL(COUNT(*),0),2),0) AS TMMES " & _
"FROM tab_vendas " & _
"WHERE MONTH(DATAVENDA) = MONTH(CURRENT_DATE) AND YEAR(DATAVENDA) = YEAR(CURRENT_DATE) "

Set rs = Conn.Execute(sql)

If Not rs.EOF Then

ttMes = NzDbl(rs!TOTALMES)
qtdMes = NzDbl(rs!QTDEMES)
tmMes = NzDbl(rs!tmMes)

End If

rs.Close
Set rs = Nothing

sql = "SELECT " & _
    "COALESCE(SUM(VALORFINAL),0) AS TOTALHOJE, " & _
    "COALESCE(COUNT(*),0) AS QTDEHOJE, " & _
    "COALESCE(ROUND(SUM(VALORFINAL) / IFNULL(COUNT(*),0),2),0) AS TMHOJE " & _
"FROM tab_vendas " & _
"WHERE DATE(DATAVENDA) = CURRENT_DATE "

Set rs = Conn.Execute(sql)

If Not rs.EOF Then

ttHoje = NzDbl(rs!TOTALHOJE)
qtdHoje = NzDbl(rs!QTDEHOJE)
tmHoje = NzDbl(rs!tmHoje)

End If

rs.Close
Set rs = Nothing

sql = "SELECT MIN(VENCIMENTO) AS PROX, " & _
      "(SELECT VALOR FROM tab_contasapagar WHERE VENCIMENTO = " & _
      "(SELECT MIN(VENCIMENTO) FROM TAB_CONTASAPAGAR WHERE STATUS = 'PENDENTE')) AS VALORPROX, " & _
      "(SELECT NOME FROM tab_contasapagar CP JOIN tab_fornecedores F " & _
      "ON F.IDFORNECEDOR = CP.IDFORNECEDOR WHERE VENCIMENTO = " & _
      "(SELECT MIN(VENCIMENTO) FROM tab_contasapagar WHERE STATUS = 'PENDENTE')) AS FORNECEDOR, " & _
      "(SELECT DESCRICAO FROM TAB_CONTASAPAGAR WHERE VENCIMENTO = " & _
      "(SELECT MIN(VENCIMENTO) FROM tab_contasapagar WHERE STATUS = 'PENDENTE')) AS DESCRICAO, " & _
      "COUNT(*) AS TITULOS, " & _
      "COALESCE(Sum(valor), 0) As VLTOTAL " & _
      "FROM tab_contasapagar WHERE STATUS = 'PENDENTE' "
         
Set rs = Conn.Execute(sql)

If Not rs.EOF Then

    proxVenc = NzDate(rs!PROX)
    forn = Nz(rs!fornecedor)
    desc = Nz(rs!DESCRICAO)
    valorprox = NzDbl(rs!valorprox)
    titulos = NzDbl(rs!titulos)
    totaltitulos = NzDbl(rs!vltotal)

End If

    txtTtMes.Value = Format(ttMes, "R$ #,##0.00")
    txtQtdMes.Value = qtdMes
    txtTmMes.Value = Format(tmMes, "R$ #,##0.00")
    txtTtHoje.Value = Format(ttHoje, "R$ #,##0.00")
    txtQtdHoje.Value = qtdHoje
    txtTmHoje.Value = Format(tmHoje, "R$ #,##0.00")
txtProxVenc.Value = _
    proxVenc & vbCrLf & vbCrLf & _
    forn & vbCrLf & vbCrLf & _
    "VALOR: " & Format(valorprox, "R$ #,##0.00") & vbCrLf & vbCrLf & _
    "DESCRIÇÃO: " & desc
    txtTitulos.Value = titulos
    txtTotalTitulos.Value = Format(totaltitulos, "R$ #,##0.00")
    

End Sub

Private Sub btnCadastro_click()
 fraEstoque.Visible = False
 fraConfig.Visible = False
   fraOperacao.Visible = False
  
    fraCadastro.Visible = Not fraCadastro.Visible
 
End Sub

Private Sub btnCaixa_Click()

If Not SolicitarSenhaCaixa() Then Exit Sub

fraOperacao.Visible = Not fraOperacao.Visible
    frmCaixa.Show vbModeless
    
End Sub


'=====================================================
' MENU OPERAÇÕES
'=====================================================

Private Sub btnOperacao_Click()
 fraConfig.Visible = False
   fraCadastro.Visible = False
  fraEstoque.Visible = False
    ' Alterna entre mostrar e esconder o frame
    fraOperacao.Visible = Not fraOperacao.Visible
 
End Sub


'=====================================================
' CADASTRO DE PRODUTOS
'=====================================================

Private Sub btnCadProduto_Click()
fraEstoque.Visible = False
   fraCadastro.Visible = False
    ' Abre o formulário de cadastro de produtos
    frmProdutos.Show vbModeless
  
End Sub


'=====================================================
' CADASTRO DE FORNECEDORES
'=====================================================

Private Sub btnCadFornecedor_Click()
fraEstoque.Visible = False
    fraCadastro.Visible = False
    ' Abre o formulário de fornecedores
    frmFornecedores.Show vbModeless
 
End Sub


'=====================================================
' CADASTRO DE USUÁRIOS
'=====================================================

Private Sub btnCadUsuario_Click()
fraEstoque.Visible = False
  fraCadastro.Visible = False
    ' Abre o formulário de usuários
    frmUsuarios.Show vbModeless
  
End Sub


'=====================================================
' VENDAS
'=====================================================

Private Sub btnVendas_Click()

fraOperacao.Visible = Not fraOperacao.Visible
    frmVendas.Show vbModeless

End Sub

'====================================================
' BUSCA O CAIXA ABERTO
'====================================================
Public Function ObterCaixaAberto() As Long

    Dim rs As ADODB.Recordset
    Dim sql As String

    sql = "SELECT IDCAIXA " & _
          "FROM TAB_CAIXA " & _
          "WHERE STATUS = 'ABERTO' " & _
          "LIMIT 1"

    Set rs = Conn.Execute(sql)

    If rs.EOF Then

        ObterCaixaAberto = 0

    Else

        ObterCaixaAberto = rs("IDCAIXA")

    End If

    rs.Close
    Set rs = Nothing

End Function

Private Sub btnAbrirCaixa_Click()

If Not ConfirmarSenha Then Exit Sub

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

            '============================================
            ' ATUALIZA CAMPOS DA TELA
            '============================================
            txtStatusCaixa.Value = "ABERTO"
            txtIdCaixa.Value = rs("ID CAIXA")

            MsgBox "Caixa Nº " & rs("ID CAIXA") & _
                   " aberto com fundo de R$ " & _
                   Format(rs("FUNDO"), "0.00"), vbInformation

        End If

    End If

    rs.Close
    Set rs = Nothing

End Sub

Private Sub btnSangria_Click()

If Not ConfirmarSenha Then Exit Sub

    '====================================================
    ' DECLARAÇÃO DAS VARIÁVEIS
    '====================================================
    Dim rs As ADODB.Recordset
    Dim sql As String

    Dim IdCaixa As Long
    Dim valor As Double

    Dim RetornoValor As String

    '====================================================
    ' BUSCA O CAIXA ABERTO
    '====================================================
    IdCaixa = ObterCaixaAberto()

    If IdCaixa = 0 Then

        MsgBox "Nenhum caixa aberto.", vbExclamation
        Exit Sub

    End If

    '====================================================
    ' SOLICITA VALOR DA SANGRIA
    '====================================================
    RetornoValor = InputBox("Informe o valor da sangria:")

    If Len(Trim(RetornoValor)) = 0 Then Exit Sub

    If Not IsNumeric(RetornoValor) Then

        MsgBox "Informe um valor numérico válido.", vbExclamation
        Exit Sub

    End If

    valor = CDbl(RetornoValor)

    '====================================================
    ' EXECUTA PROCEDURE
    '====================================================
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

If Not ConfirmarSenha Then Exit Sub

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

            '============================================
            ' ATUALIZA CAMPOS DA TELA
            '============================================
            txtStatusCaixa.Value = "FECHADO"
            txtIdCaixa.Value = ""

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
