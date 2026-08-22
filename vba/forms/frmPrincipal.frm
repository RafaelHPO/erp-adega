VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPrincipal 
   Caption         =   "E R P - FLAME SYSTEM"
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
Private mAjustandoLayout As Boolean
Public mostrarvalores As Boolean

Private Const BASE_INSIDE_WIDTH As Single = 1046.25
Private Const BASE_INSIDE_HEIGHT As Single = 599.25

Private Const BASE_IMAGE_WIDTH As Single = 816
Private Const MARGEM_DIREITA_IMAGEM As Single = 26.25
Private Const RESERVA_INFERIOR As Single = 53.25

Private Sub AplicarPaleta()

    'Aplica o tema genérico ao formulário inteiro
    AplicarTema Me

    '==================================================
    ' FRAMES ESPECÍFICOS
    '==================================================

    fraMenu.BackColor = COR_FUNDO_SECUNDARIO

    fraCadastro.BackColor = COR_FUNDO_SECUNDARIO
    fraOperacao.BackColor = COR_FUNDO_SECUNDARIO
    fraEstoque.BackColor = COR_FUNDO_SECUNDARIO

    fraConfig.BackColor = COR_PAINEL

    '==================================================
    ' NAVEGAÇÃO PRINCIPAL
    '==================================================

    AplicarBotaoPrincipal btnCadastro
    AplicarBotaoPrincipal btnOperacao
    AplicarBotaoPrincipal btnFinanceiro
    AplicarBotaoPrincipal btnDashboard
    AplicarBotaoPrincipal btnMostrarOcultar

    '==================================================
    ' TÍTULO E IDENTIDADE
    '==================================================

    lblTitulo.ForeColor = COR_VERMELHO
    Label16.ForeColor = COR_VERMELHO

    lblVersao.ForeColor = COR_TEXTO_SECUNDARIO
    lblAtualizacao.ForeColor = COR_TEXTO_SECUNDARIO

    '==================================================
    ' CABEÇALHOS DOS PAINÉIS
    '==================================================

    AplicarCabecalhoPainel Label10
    AplicarCabecalhoPainel Label11

    'Fechar atualizações
    Label12.BackColor = COR_VERMELHO_ESCURO
    Label12.ForeColor = COR_TEXTO

    'Exibir atualizações
    Label13.BackColor = COR_VERMELHO
    Label13.ForeColor = COR_TEXTO

End Sub

Private Sub UserForm_Resize()

    AjustarLayout

End Sub

Private Sub AjustarLayout()

    Dim novaLarguraImagem As Single
    Dim novaAlturaPrincipal As Single

    If mAjustandoLayout Then Exit Sub
    If Me.InsideWidth <= 0 Or Me.InsideHeight <= 0 Then Exit Sub

    On Error GoTo TratarErro

    mAjustandoLayout = True

    '==================================================
    ' ÁREA PRINCIPAL
    '==================================================

    novaLarguraImagem = _
        Me.InsideWidth - 204 - MARGEM_DIREITA_IMAGEM

    novaAlturaPrincipal = _
        Me.InsideHeight - 6 - RESERVA_INFERIOR

    If novaLarguraImagem < 1 Then novaLarguraImagem = 1
    If novaAlturaPrincipal < 1 Then novaAlturaPrincipal = 1

    '==================================================
    ' MENU LATERAL
    '==================================================

    With fraMenu
        .Left = 12
        .Top = 6
        .Width = 168
        .Height = novaAlturaPrincipal
    End With

    '==================================================
    ' IMAGEM DA ÁREA PRINCIPAL
    '==================================================

    With Image1
        .Left = 204
        .Top = 6
        .Width = novaLarguraImagem
        .Height = novaAlturaPrincipal
    End With

    '==================================================
    ' SUBMENUS
    ' Apenas os Frames são movidos.
    ' Os controles internos permanecem intactos.
    '==================================================

    fraCadastro.Left = fraMenu.Left + fraMenu.Width - 6
    fraCadastro.Top = fraMenu.Top + btnCadastro.Top + 6

    fraOperacao.Left = fraMenu.Left + fraMenu.Width - 6
    fraOperacao.Top = fraMenu.Top + btnOperacao.Top + 6

    fraEstoque.Left = _
        fraCadastro.Left + btnEstoque.Left + 18

    fraEstoque.Top = _
        fraCadastro.Top + fraCadastro.Height

    '==================================================
    ' INDICADORES SUPERIORES
    '==================================================

    PosicionarKPI txtTtMes, Label2, 72, 18
    PosicionarKPI txtQtdMes, Label3, 204, 12.05
    PosicionarKPI txtTmMes, Label4, 336, 12

    PosicionarKPI txtTtHoje, Label5, 480, 18
    PosicionarKPI txtQtdHoje, Label6, 612, 12
    PosicionarKPI txtTmHoje, Label7, 744, 12.05

    '==================================================
    ' PAINÉIS INFERIORES DA IMAGEM
    '==================================================

    AjustarPainelAtualizacoes
    AjustarPainelContas

    '==================================================
    ' RODAPÉ
    '==================================================

    AjustarRodape

Saida:

    mAjustandoLayout = False
    Exit Sub

TratarErro:

    Debug.Print "Erro em AjustarLayout: " & _
                Err.Number & " - " & Err.Description

    Resume Saida

End Sub

Private Sub PosicionarKPI( _
    ByVal caixa As Object, _
    ByVal rotulo As Object, _
    ByVal centroBase As Single, _
    ByVal deslocamentoRotulo As Single)

    Dim centroAtual As Single

    centroAtual = Image1.Left + _
        (centroBase / BASE_IMAGE_WIDTH) * Image1.Width

    caixa.Left = centroAtual - caixa.Width / 2
    caixa.Top = Image1.Top + 12

    rotulo.Left = caixa.Left + deslocamentoRotulo
    rotulo.Top = Image1.Top + 42

End Sub

Private Sub AjustarPainelAtualizacoes()

    Dim limiteInferior As Single
    Dim topoCabecalho As Single

    limiteInferior = Image1.Top + Image1.Height
    topoCabecalho = limiteInferior - 120

    txtAtt.Left = Image1.Left + 12
    txtAtt.Top = limiteInferior - 106.1

    Label11.Left = txtAtt.Left

    Label12.Left = txtAtt.Left + 210
    Label12.Top = topoCabecalho

    Label13.Left = txtAtt.Left
    Label13.Top = limiteInferior - 24

    If txtAtt.Visible Then
        Label11.Top = topoCabecalho
    Else
        Label11.Top = limiteInferior - 24
    End If

End Sub

Private Sub AjustarPainelContas()

    Dim limiteDireito As Single
    Dim limiteInferior As Single

    limiteDireito = _
        Image1.Left + Image1.Width - 12

    limiteInferior = _
        Image1.Top + Image1.Height

    'Próxima conta a vencer
    txtProxVenc.Left = limiteDireito - txtProxVenc.Width
    txtProxVenc.Top = limiteInferior - 106.1

    Label10.Left = txtProxVenc.Left
    Label10.Top = limiteInferior - 120

    'Quantidade de títulos
    txtTitulos.Left = _
        txtProxVenc.Left - 12 - txtTitulos.Width

    txtTitulos.Top = limiteInferior - 120

    Label8.Left = txtTitulos.Left + 12
    Label8.Top = limiteInferior - 90

    'Valor total
    txtTotalTitulos.Left = txtTitulos.Left
    txtTotalTitulos.Top = limiteInferior - 60

    Label9.Left = txtTotalTitulos.Left + 18
    Label9.Top = limiteInferior - 30

End Sub

Private Sub AjustarRodape()

    Dim topoCampos As Single
    Dim topoRotulos As Single
    Dim topoVersao As Single
    Dim inicioGrupoDireito As Single

    topoCampos = Me.InsideHeight - 30.25
    topoRotulos = Me.InsideHeight - 30.2
    topoVersao = Me.InsideHeight - 30.25

    '==================================================
    ' GRUPO ESQUERDO
    '==================================================

    lblTitulo.Left = 6
    lblTitulo.Top = topoCampos

    lblUsuario.Left = 198
    lblUsuario.Top = topoRotulos

    txtUsuario.Left = 252
    txtUsuario.Top = topoCampos

    lblIdCaixa.Left = 360
    lblIdCaixa.Top = topoRotulos

    txtIdCaixa.Left = 414
    txtIdCaixa.Top = topoCampos

    lblCaixa.Left = 468
    lblCaixa.Top = topoRotulos

    txtStatusCaixa.Left = 552
    txtStatusCaixa.Top = topoCampos

    '==================================================
    ' GRUPO DIREITO
    '==================================================

    inicioGrupoDireito = Me.InsideWidth - 374.25

    Label16.Left = inicioGrupoDireito
    Label16.Top = topoVersao

    lblVersao.Left = inicioGrupoDireito + 90
    lblVersao.Top = topoVersao

    lblAtualizacao.Left = inicioGrupoDireito + 186
    lblAtualizacao.Top = topoVersao

End Sub
Private Sub userform_activate()

    CarregarKpiPrincipal
    
    lblTitulo.FontSize = 18
    lblUsuario.FontSize = 12
    lblIdCaixa.FontSize = 12
    lblCaixa.FontSize = 12
    Label16.FontSize = 12
    lblVersao.FontSize = 14
    lblAtualizacao.FontSize = 12
    
    
Label12.BackColor = COR_TEXTO_SECUNDARIO
Label13.BackColor = COR_TEXTO_SECUNDARIO
Label12.BackStyle = fmBackStyleOpaque
Label13.BackStyle = fmBackStyleOpaque

lblVersao.ForeColor = COR_TEXTO
lblAtualizacao.ForeColor = COR_TEXTO
CarregarAlteracoes
    VerificarAtualizacao
    
End Sub

Public Sub UserForm_initialize()

Call attBanco

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

    txtAtt.SetFocus
    txtAtt.SelStart = Len(txtAtt.Value)
    txtAtt.SelLength = 0

Me.Left = 0
Me.Top = 0
Me.Width = Application.Width
Me.Height = Application.Height

AjustarLayout
AplicarPaleta
mostrarvalores = False
AtualizarExibicaoValores

End Sub

Private Sub VerificarAtualizacao()

AguardarSegundos 3

On Error GoTo TratarErro

    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim status As String
    
        sql = " SELECT VALOR FROM TAB_CONFIG WHERE CHAVE = 'STATUS' "
        
            Set rs = Conn.Execute(sql)
                
                If Not rs.EOF Then
                    
                    status = rs!valor
                    
                   If Trim(status) = "" Then
                
                    Dim resposta As VbMsgBoxResult
                
                    resposta = MsgBox( _
                        "Há uma nova atualização disponível." & vbCrLf & _
                        "Deseja atualizar o sistema agora?", _
                        vbYesNo + vbInformation, _
                        "SISTEMA")
                
                        If resposta = vbYes Then
                            ReiniciarSistema
                        End If
                
                    End If
                
                End If
                            
        Exit Sub

TratarErro:

    modSistema.tela = "frmPrincipal - verificar att"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"

End Sub

Public Sub CarregarAlteracoes()
    
    On Error GoTo TrataErro

    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim Texto As String

    sql = "SELECT VERSAO, ALTERACAO " & _
          "FROM TAB_ALTERACOES " & _
          "ORDER BY DATA_ALTERACAO DESC , IDALTERACAO DESC "

    Set rs = Conn.Execute(sql)

    Texto = ""

    Do While Not rs.EOF

        Texto = Texto & _
                "--------- Versão: " & rs!versao & " ---------" & vbCrLf & _
                vbCrLf & _
                rs!ALTERACAO & vbCrLf & _
                vbCrLf

        rs.MoveNext

    Loop

    txtAtt.TextAlign = fmTextAlignCenter
    txtAtt.Value = Texto

    rs.Close
    Set rs = Nothing

    Exit Sub

TrataErro:

    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
    End If

    Set rs = Nothing

    MsgBox "Erro ao carregar alterações: " & Err.Description, vbCritical

End Sub

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
    frmSenha.admin = True
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
Private Sub Label12_Click()

    txtAtt.Visible = Not txtAtt.Visible

    Label13.Visible = True
    Label12.Visible = False

    AjustarPainelAtualizacoes

End Sub

Private Sub Label13_Click()

    txtAtt.Visible = Not txtAtt.Visible

    Label12.Visible = True
    Label13.Visible = False

    AjustarPainelAtualizacoes

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

Private Sub AtualizarExibicaoValores()

    Dim mascara As String

    If mostrarvalores Then
        mascara = vbNullString
    Else
        mascara = "*"
    End If

    txtTtMes.PasswordChar = mascara
    txtQtdMes.PasswordChar = mascara
    txtTmMes.PasswordChar = mascara

    txtTtHoje.PasswordChar = mascara
    txtQtdHoje.PasswordChar = mascara
    txtTmHoje.PasswordChar = mascara

    txtProxVenc.PasswordChar = mascara
    txtTitulos.PasswordChar = mascara
    txtTotalTitulos.PasswordChar = mascara

End Sub

Private Sub btnMostrarOcultar_Click()

    mostrarvalores = Not mostrarvalores

    AtualizarExibicaoValores

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

Private Sub btnOperacao_Click()
 fraConfig.Visible = False
   fraCadastro.Visible = False
  fraEstoque.Visible = False
    ' Alterna entre mostrar e esconder o frame
    fraOperacao.Visible = Not fraOperacao.Visible
 
End Sub

Private Sub btnCadProduto_Click()
fraEstoque.Visible = False
   fraCadastro.Visible = False
    ' Abre o formulário de cadastro de produtos
    frmProdutos.Show vbModeless
  
End Sub

Private Sub btnCadFornecedor_Click()
fraEstoque.Visible = False
    fraCadastro.Visible = False
    ' Abre o formulário de fornecedores
    frmFornecedores.Show vbModeless
 
End Sub


Private Sub btnCadUsuario_Click()
fraEstoque.Visible = False
  fraCadastro.Visible = False
    ' Abre o formulário de usuários
    frmUsuarios.Show vbModeless
  
End Sub

Private Sub btnVendas_Click()

fraOperacao.Visible = Not fraOperacao.Visible
    frmVendas.Show vbModeless

End Sub

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

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

  If CloseMode = vbFormControlMenu Then
       Cancel = True
       Call FecharSistema
       On Error Resume Next
Application.OnTime ProximoBackup, "BackupPeriodico", , False

ExecutarBackup False
   End If

End Sub
