VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPagamento 
   Caption         =   "PAGAMENTO"
   ClientHeight    =   5955
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5325
   OleObjectBlob   =   "frmPagamento.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPagamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private totalVenda As Currency
Private ValorPago As Currency

Option Explicit

Public idvenda As Long

Private Sub btnEstorno_Click()

    Set frmPagamentos = New frmPagamentos

    frmPagamentos.idvenda = idvenda
    frmPagamentos.Show vbModal

End Sub

Private Sub txtCartao_Change()

FormatarMoeda txtCartao

End Sub

Private Sub txtDinheiro_Change()

FormatarMoeda txtDinheiro

End Sub

Private Sub txtPix_Change()

FormatarMoeda txtPix

End Sub

Private Sub txtVr_Change()

FormatarMoeda txtVr

End Sub

Private Sub userform_activate()

    AtualizarFinanceiro

End Sub

Private Sub UserForm_Initialize()

On Error GoTo TratarErro

    txtDinheiro.Visible = False
    txtPix.Visible = False
    txtCartao.Visible = False

    txtDinheiro.Value = ""
    txtPix.Value = ""
    txtCartao.Value = ""
    txtPendente.Value = ""

    ckDinheiro.Value = False
    ckPix.Value = False
    ckCartao.Value = False
    ckMulti.Value = False
    
    txtVr.Visible = False
txtVr.Value = ""
ckVR.Value = False
    
        Exit Sub

TratarErro:

    modSistema.tela = "Pagamento - initialize"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"
    
End Sub
Private Function GetPendenteVR() As Double

    GetPendenteVR = GetPendente() * IIf(ckVR.Value, 1.15, 1)

End Function
Private Function CalcularVR() As Double

    Dim base As Double

    base = GetPendente()

    If ckVR.Value Then
        CalcularVR = base * 1.15
    Else
        CalcularVR = base
    End If

End Function

Private Sub ckVR_Click()

    If ckVR.Value = True Then
    txtVr.Visible = True
    Else
    txtVr.Visible = False
        txtAcres.Value = ""
        Exit Sub
    End If

    txtAcres.Value = "Total com VR: " & Format(GetPendente() * 1.15, "0.00")

End Sub
Private Function AplicarVR(valor As Double) As Double

    AplicarVR = valor * 1.15

End Function

Private Sub ckDinheiro_Click()

    If ckDinheiro.Value = True Then
        txtDinheiro.Visible = True
    Else
        txtDinheiro.Visible = False
    End If

End Sub
Private Sub ckPix_Click()

    If ckPix.Value = True Then
        txtPix.Visible = True
    Else
        txtPix.Visible = False
    End If

End Sub
Private Sub ckCartao_Click()

    If ckCartao.Value = True Then
        txtCartao.Visible = True
    Else
        txtCartao.Visible = False
    End If

End Sub
Private Sub ckMulti_Click()

    If ckMulti.Value = True Then
        txtDinheiro.Visible = True
        txtPix.Visible = True
        txtCartao.Visible = True
        txtVr.Visible = True
    Else
        txtDinheiro.Visible = False
        txtPix.Visible = False
        txtCartao.Visible = False
         txtVr.Visible = False
    End If

End Sub
Private Function GetPendente() As Double

    Dim rs As ADODB.Recordset
    Dim totalVenda As Double
    Dim totalPago As Double

    Set rs = Conn.Execute( _
        "SELECT IFNULL(SUM(SUBTOTAL),0) AS TOTAL " & _
        "FROM TAB_ITENSVENDA WHERE IDVENDA=" & idvenda)

    totalVenda = rs!Total
    rs.Close

    Set rs = Conn.Execute( _
        "SELECT IFNULL(SUM(VALORPAGO),0) AS TOTAL " & _
        "FROM TAB_PAGAMENTOS WHERE IDVENDA=" & idvenda)

    totalPago = rs!Total
    rs.Close

    GetPendente = totalVenda - totalPago

End Function
Private Function GetTotalPago() As Double

    Dim rs As ADODB.Recordset

    Set rs = Conn.Execute( _
        "SELECT IFNULL(SUM(VALORPAGO),0) AS TOTAL " & _
        "FROM TAB_PAGAMENTOS WHERE IDVENDA=" & idvenda)

    GetTotalPago = rs!Total

    rs.Close

End Function
Public Sub AtualizarFinanceiro()

    Dim base As Double

    base = GetPendente()

    txtPago.Value = Format(GetTotalPago(), "0.00")
    txtPendente.Value = Format(base, "0.00")

    If ckVR.Value Then
        txtAcres.Value = "Com VR: " & Format(base * 1.15, "0.00")
    Else
        txtAcres.Value = ""
    End If

End Sub
Private Sub btnPagar_Click()

On Error GoTo TratarErro

    Dim recebido As Double
    Dim pago As Double
    Dim troco As Double
    Dim pendente As Double
    Dim multi As Boolean
    Dim rs As ADODB.Recordset

    pendente = GetPendente()
    multi = ckMulti.Value

    '================ DINHEIRO
    If ckDinheiro.Value Then

        recebido = NzDbl(txtDinheiro.Value)

        If recebido > 0 Then

            If multi = False Then

                If recebido < pendente Then
                    MsgBox "Tem que pagar o valor total: " & Format(pendente, "0.00")
                    Exit Sub
                End If

                pago = pendente

            Else

                pago = recebido

            End If

            troco = recebido - pago

            If troco > 0 Then
                MsgBox "Troco: " & Format(troco, "0.00")
            End If

            Set rs = Conn.Execute("CALL PROC_BAIXARPAGAMENTOS(" & _
                SqlNumero(IDUsuarioLogado) & "," & _
                SqlNumero(idvenda) & "," & _
                SqlTexto("DINHEIRO") & "," & _
                SqlNumero(recebido) & "," & _
                SqlNumero(pago) & ")")

            If Not rs Is Nothing Then
                If Not rs.EOF Then MsgBox rs.Fields(0).Value, vbInformation
                rs.Close
            End If

            Set rs = Nothing

            pendente = GetPendente()

        End If

    End If

    '================ PIX
    If ckPix.Value Then

        recebido = NzDbl(txtPix.Value)

        If recebido > 0 Then

            If multi = False Then

                If recebido < pendente Then
                    MsgBox "Tem que pagar o valor total: " & Format(pendente, "0.00")
                    Exit Sub
                End If

                pago = pendente

            Else

                pago = recebido

            End If

            Set rs = Conn.Execute("CALL PROC_BAIXARPAGAMENTOS(" & _
                SqlNumero(IDUsuarioLogado) & "," & _
                SqlNumero(idvenda) & "," & _
                SqlTexto("PIX") & "," & _
                SqlNumero(recebido) & "," & _
                SqlNumero(pago) & ")")

            If Not rs Is Nothing Then
                If Not rs.EOF Then MsgBox rs.Fields(0).Value, vbInformation
                rs.Close
            End If

            Set rs = Nothing

            pendente = GetPendente()

        End If

    End If

    '================ CARTAO
    If ckCartao.Value Then

        recebido = NzDbl(txtCartao.Value)

        If recebido > 0 Then

            If multi = False Then

                If recebido < pendente Then
                    MsgBox "Tem que pagar o valor total: " & Format(pendente, "0.00")
                    Exit Sub
                End If

                pago = pendente

            Else

                pago = recebido

            End If

            Set rs = Conn.Execute("CALL PROC_BAIXARPAGAMENTOS(" & _
                SqlNumero(IDUsuarioLogado) & "," & _
                SqlNumero(idvenda) & "," & _
                SqlTexto("CARTAO") & "," & _
                SqlNumero(recebido) & "," & _
                SqlNumero(pago) & ")")

            If Not rs Is Nothing Then
                If Not rs.EOF Then MsgBox rs.Fields(0).Value, vbInformation
                rs.Close
            End If

            Set rs = Nothing

            pendente = GetPendente()

        End If

    End If

    '================ VR
    If ckVR.Value Then

        recebido = NzDbl(txtVr.Value)

        If recebido > 0 Then

            pago = recebido

            Set rs = Conn.Execute("CALL PROC_BAIXARPAGAMENTOS(" & _
                SqlNumero(IDUsuarioLogado) & "," & _
                SqlNumero(idvenda) & "," & _
                SqlTexto("VR") & "," & _
                SqlNumero(recebido) & "," & _
                SqlNumero(pago) & ")")

            If Not rs Is Nothing Then
                If Not rs.EOF Then MsgBox rs.Fields(0).Value, vbInformation
                rs.Close
            End If

            Set rs = Nothing

            pendente = GetPendente()

        End If

    End If

    AtualizarFinanceiro

    pendente = GetPendente()

    If pendente <= 0 Then

        MsgBox "Pagamento concluído!"
        Unload Me

    Else

        MsgBox "Ainda falta: R$ " & Format(pendente, "0.00")

    End If

    Exit Sub

TratarErro:

    modSistema.tela = "Pagamento - btnPagar"
    modSistema.DescErro = Err.Description
    modSistema.nErro = Err.Number

    Call modSistema.ReportarErro
    
    MsgBox "Erro: " & Err.Number & vbCrLf & _
                        Err.Description, vbInformation, "SISTEMA"

End Sub
