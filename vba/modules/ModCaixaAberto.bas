Attribute VB_Name = "ModCaixaAberto"
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

Public Function ExisteCaixaAberto() As Boolean

    ExisteCaixaAberto = (ObterCaixaAberto() > 0)

End Function

Public Sub ValidarCaixaOperacional()

    Dim rs As ADODB.Recordset
    Dim sql As String

    Dim DataCaixa As Date
    Dim ValorCaixa As Double

    frmVendas.lblAvisoCaixa.Visible = False
    frmVendas.lblAvisoCaixa.Caption = ""

    If Not ExisteCaixaAberto Then Exit Sub

    sql = "SELECT DATAABERTURA, VALORCAIXA " & _
          "FROM TAB_CAIXA " & _
          "WHERE STATUS = 'ABERTO' " & _
          "LIMIT 1"

    Set rs = Conn.Execute(sql)

    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If

    DataCaixa = rs!DATAABERTURA
    ValorCaixa = Nz(rs!ValorCaixa, 0)

    rs.Close
    Set rs = Nothing

    '=========================================
    ' AVISOS
    '=========================================
    If DataCaixa < Date Then

  With frmVendas.lblAvisoCaixa
    .Caption = "[!] Existe um caixa aberto do dia " & _
               Format(DataCaixa, "dd/mm/yyyy")
    .ForeColor = RGB(192, 0, 0)
    .Font.Bold = True
    .Visible = True
End With

    ElseIf ValorCaixa >= 300 Then

        With frmVendas.lblAvisoCaixa
            .Caption = "[!] Caixa com R$ " & _
                       Format(ValorCaixa, "0.00") & _
                       ". Recomenda-se realizar uma sangria."
        .ForeColor = RGB(192, 0, 0)
    .Font.Bold = True
            .Visible = True
        End With

    End If

End Sub
