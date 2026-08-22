VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmContas 
   Caption         =   "CONTAS A PAGAR / RECEBER"
   ClientHeight    =   5370
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7035
   OleObjectBlob   =   "frmContas.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmContas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public IdConta As Long
Public TipoConta As String
Public ModoNovo As Boolean

Private Sub AplicarPaleta()

    AplicarTema Me
    AplicarBotaoPrincipal btnBaixarConta
    AplicarBotaoPrincipal btnexcluir
    
End Sub

Private Sub UserForm_initialize()

    CarregarComboFornecedor cmbFornecedor
    AplicarPaleta
    
End Sub

Private Sub txtDataLan_Change()

FormatarData txtDataLan

End Sub

Private Sub txtDataPag_Change()

FormatarData txtDataPag

End Sub

Private Sub txtDataVen_Change()

  FormatarData txtDataVen

End Sub


Public Sub PrepararNovaConta()

    cmbFornecedor.ListIndex = -1
    txtDataLan = Format(Date, "dd/mm/yyyy")
    txtDescricao = ""
    txtStatus = "PENDENTE"
    txtDataVen = ""
    txtDataPag = ""
    txtValor = ""

    btnSalvar.Visible = True
    btnBaixarConta.Visible = False
    btnexcluir.Visible = False

End Sub

Public Sub CarregarConta()

On Error GoTo TratarErro

    Dim rs As ADODB.Recordset
    Dim sql As String

    If TipoConta = "PAGAR" Then

        sql = "SELECT C.IDFORNECEDOR, C.DESCRICAO, C.VALOR, C.VENCIMENTO, C.STATUS, " & _
              "DATE(C.DATALANCAMENTO) AS DATALANCAMENTO, C.DATAPAGAMENTO " & _
              "FROM tab_contasapagar C " & _
              "WHERE IDAPAGAR = " & IdConta

    ElseIf TipoConta = "RECEBER" Then

        sql = "SELECT DESCRICAO, VALOR, VENCIMENTO, STATUS, DATARECEBIMENTO " & _
              "FROM tab_contasareceber " & _
              "WHERE IDARECEBER = " & IdConta

    Else
        MsgBox "Tipo de conta inválido.", vbExclamation
        Exit Sub
    End If

    Set rs = Conn.Execute(sql)

    If rs.EOF Then
        MsgBox "Conta não encontrada.", vbExclamation
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If


    If TipoConta = "PAGAR" Then

       Dim i As Long

For i = 0 To cmbFornecedor.ListCount - 1
    If CLng(cmbFornecedor.List(i, 0)) = CLng(rs!idFornecedor) Then
        cmbFornecedor.ListIndex = i
        Exit For
    End If
Next i
        txtDataLan.Value = FormatarDataCampo(rs!DATALANCAMENTO)
        txtDescricao.Value = Nz(rs!DESCRICAO)
        txtStatus.Value = Nz(rs!STATUS)
        txtDataVen.Value = FormatarDataCampo(rs!VENCIMENTO)
        txtDataPag.Value = FormatarDataCampo(rs!DATAPAGAMENTO)
        txtValor.Value = Format(NzDbl(rs!valor), "#,##0.00")
        
    Else
        cmbFornecedor.ListIndex = -1
        txtDataLan.Value = ""
        txtDescricao.Value = Nz(rs!DESCRICAO)
        txtStatus.Value = Nz(rs!STATUS)
        txtDataVen.Value = FormatarDataCampo(rs!VENCIMENTO)
        txtDataPag.Value = FormatarDataCampo(rs!DATARECEBIMENTO)
        txtValor.Value = Format(NzDbl(rs!valor), "#,##0.00")
    End If

    btnSalvar.Visible = False
    btnBaixarConta.Visible = True
    btnexcluir.Visible = True

    rs.Close
    Set rs = Nothing

    Exit Sub

TratarErro:

    modSistema.tela = "Contas - CarregarConta"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"

End Sub

Private Sub btnBaixarConta_Click()

On Error GoTo TratarErro

If Not SolicitarSenhaCaixa() Then Exit Sub

    Dim sql As String

    If IdConta = 0 Then Exit Sub

    If TipoConta = "PAGAR" Then

        sql = "UPDATE tab_contasapagar SET " & _
              "STATUS = 'PAGO', " & _
              "DATAPAGAMENTO = CURDATE() " & _
              "WHERE IDAPAGAR = " & IdConta

    ElseIf TipoConta = "RECEBER" Then

        sql = "UPDATE tab_contasareceber SET " & _
              "STATUS = 'RECEBIDO', " & _
              "DATARECEBIMENTO = CURDATE() " & _
              "WHERE IDARECEBER = " & IdConta

    End If

    Conn.Execute sql

    RecarregarFluxoCaixa
    MsgBox "Baixado!", vbInformation
    Unload Me

    Exit Sub

TratarErro:

    modSistema.tela = "Contas - Baixar Conta"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"

End Sub

Private Sub btnExcluir_Click()

On Error GoTo TratarErro

If Not SolicitarSenhaCaixa() Then Exit Sub

    Dim sql As String

    If IdConta = 0 Then Exit Sub

    If MsgBox("Deseja cancelar esta conta?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

    If TipoConta = "PAGAR" Then

        sql = "UPDATE tab_contasapagar SET " & _
              "STATUS = 'CANCELADO' " & _
              "WHERE IDAPAGAR = " & IdConta

    ElseIf TipoConta = "RECEBER" Then

        sql = "UPDATE tab_contasareceber SET " & _
              "STATUS = 'CANCELADO' " & _
              "WHERE IDARECEBER = " & IdConta

    End If

    Conn.Execute sql

    RecarregarFluxoCaixa
    Unload Me

    Exit Sub

TratarErro:

    modSistema.tela = "Contas - Excluir"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"

End Sub

Private Sub btnSalvar_Click()

On Error GoTo TratarErro

If Not SolicitarSenhaCaixa() Then Exit Sub

    Dim sql As String

    If Not ModoNovo Then Exit Sub

    sql = "INSERT INTO tab_contasapagar " & _
          "(IDFORNECEDOR, DESCRICAO, VALOR, VENCIMENTO, STATUS, DATALANCAMENTO, DATAPAGAMENTO) VALUES (" & _
          SqlNumero(cmbFornecedor.Column(0)) & ", " & _
          SqlTexto(txtDescricao.Value) & ", " & _
          SqlNumero(NzDbl(Replace(txtValor.Value, ".", ""))) & ", " & _
          SqlData(txtDataVen.Value) & ", " & _
          SqlTexto(txtStatus.Value) & ", " & _
          "NOW(), " & _
          SqlData(txtDataPag.Value) & ")"

    Conn.Execute sql

    RecarregarFluxoCaixa
    MsgBox "Conta Lançada!", vbInformation
    Unload Me

    Exit Sub

TratarErro:

    modSistema.tela = "Contas - btnSalvar"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"

End Sub

Private Sub btnEditar_Click()

On Error GoTo TratarErro

    If Not SolicitarSenhaCaixa() Then Exit Sub
    
    If IdConta = 0 Then Exit Sub
    
    Dim sql As String
    
    sql = "update tab_contasapagar set " & _
          "idfornecedor = " & SqlNumero(cmbFornecedor.Column(0)) & ", " & _
          "descricao = " & SqlTexto(txtDescricao.Value) & ", " & _
          "valor = " & SqlNumero(NzDbl(Replace(txtValor.Value, ".", ""))) & ", " & _
          "vencimento = " & SqlData(txtDataVen.Value) & ", " & _
          "status = " & SqlTexto(txtStatus.Value) & ", " & _
          "DATAPAGAMENTO = " & SqlData(txtDataPag.Value) & " " & _
          "WHERE IDAPAGAR = " & IdConta
    
    Conn.Execute sql
    
    RecarregarFluxoCaixa
    MsgBox "Conta Alterada!", vbInformation
    Unload Me
    
Exit Sub

    Exit Sub

TratarErro:

    modSistema.tela = "Contas - Editar Conta"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"

End Sub

Public Sub RecarregarFluxoCaixa()

    On Error Resume Next
    frmFluxoCaixa.PreencherLvPagar
    frmFluxoCaixa.PreencherLvReceber
    On Error GoTo 0

End Sub

Private Sub txtValor_Change()

FormatarMoeda txtValor

End Sub
