Attribute VB_Name = "ModETLitens"
Option Explicit

Function LinhaSegura(pos As Long) As String
    If pos >= LBound(dados) And pos <= UBound(dados) Then
        LinhaSegura = CStr(dados(pos))
    Else
        LinhaSegura = ""
    End If
End Function

Function TextoJanela(pos As Long, Optional qtdLinhas As Long = 4) As String
    Dim j As Long
    Dim fim As Long
    Dim Txt As String

    fim = pos + qtdLinhas
    If fim > UBound(dados) Then fim = UBound(dados)

    For j = pos To fim
        Txt = Txt & " " & LinhaSegura(j)
    Next j

    TextoJanela = Application.WorksheetFunction.Trim(Txt)
End Function

Function NormalizarTexto(Txt As String) As String
    Txt = Replace(Txt, vbTab, " ")
    Txt = Replace(Txt, Chr(160), " ")
    Txt = Replace(Txt, "Vl.Unit.", "Vl. Unit.")
    Txt = Replace(Txt, "Vl Unit", "Vl. Unit")
    Txt = Replace(Txt, "Qtde:", "Qtde.:")
    Txt = Replace(Txt, "Qtd.:", "Qtde.:")
    NormalizarTexto = Application.WorksheetFunction.Trim(Txt)
End Function
Sub ExtrairItens()

    Dim ws As Worksheet
    Dim i As Long
    Dim Linha As Long
    Dim bloco As String

    Set ws = wbOrigem.Worksheets("ETL_ITEM")
    Linha = 2

    For i = 1 To UBound(dados)

        If InStr(1, dados(i), "Código", vbTextCompare) > 0 Then

            bloco = NormalizarTexto(TextoJanela(i, 6))

            ws.Cells(Linha, 1) = BuscarNumeroNota
            ws.Cells(Linha, 2) = ExtrairEAN(bloco)
            ws.Cells(Linha, 3) = ExtrairCodigo(dados(i))
            ws.Cells(Linha, 4) = ExtrairDescricao(dados(i))
            ws.Cells(Linha, 5) = ExtrairQuantidade(bloco)
            ws.Cells(Linha, 6) = ExtrairUnidade(bloco)
            ws.Cells(Linha, 7) = ExtrairValorUnitario(bloco)
            ws.Cells(Linha, 8) = BuscarValorTotalItem(i)
            ws.Cells(Linha, 9) = ExtrairCNPJ

            Linha = Linha + 1

        End If

        If InStr(1, dados(i), "Qtd. total de itens", vbTextCompare) > 0 Then Exit For

    Next i

End Sub

Function ExtrairQuantidade(Txt As String) As String
    ExtrairQuantidade = ExtrairEntre(Txt, "Qtde.:", "UN:")
End Function

Function ExtrairUnidade(Txt As String) As String
    Dim un As String

    un = ExtrairEntre(Txt, "UN:", "Vl. Unit.:")
    un = Trim(un)

    If un = "" Then
        un = ExtrairEntre(Txt, "Un:", "Vl. Unit.:")
    End If

    ExtrairUnidade = un
End Function

Function ExtrairValorUnitario(Txt As String) As String
    Dim valor As String

    valor = ExtrairEntre(Txt, "Vl. Unit.:", "Vl. Total")
    If valor = "" Then valor = ExtrairDepois(Txt, "Vl. Unit.:")

    ExtrairValorUnitario = LimparValor(valor)
End Function

Function ExtrairEntre(Txt As String, inicio As String, fim As String) As String
    Dim a As Long
    Dim b As Long

    a = InStr(1, Txt, inicio, vbTextCompare)
    If a = 0 Then Exit Function

    b = InStr(a + Len(inicio), Txt, fim, vbTextCompare)
    If b = 0 Then Exit Function

    ExtrairEntre = Trim(Mid(Txt, a + Len(inicio), b - a - Len(inicio)))
End Function

Function ExtrairDepois(Txt As String, campo As String) As String
    Dim p As Long

    p = InStr(1, Txt, campo, vbTextCompare)
    If p = 0 Then Exit Function

    ExtrairDepois = Trim(Mid(Txt, p + Len(campo)))
End Function

Function LimparValor(Txt As String) As String
    Dim i As Long
    Dim C As String
    Dim r As String

    Txt = Trim(Txt)

    For i = 1 To Len(Txt)
        C = Mid(Txt, i, 1)

        If C Like "[0-9]" Or C = "," Or C = "." Then
            r = r & C
        ElseIf r <> "" Then
            Exit For
        End If
    Next i

    LimparValor = r
End Function

Function ExtrairCodigo(Txt As String) As String

    Dim p As Long
    Dim resultado As String

    p = InStr(1, Txt, "Código", vbTextCompare)

    If p = 0 Then Exit Function

    resultado = Mid(Txt, p + 7)

    resultado = Replace(resultado, ":", "")
    resultado = Replace(resultado, ")", "")
    
    ExtrairCodigo = Trim(resultado)

End Function
Function ExtrairDescricao(Txt As String) As String

    ExtrairDescricao = Trim(Split(Txt, "(Código:")(0))

End Function
Function ExtrairCampo(Txt As String, inicio As String, fim As String) As String

    Dim a As Long
    Dim b As Long

    Txt = Replace(Txt, ".", "")

    a = InStr(Txt, Replace(inicio, ".", ""))
    b = InStr(Txt, Replace(fim, ".", ""))

    If a = 0 Or b = 0 Then Exit Function

    ExtrairCampo = Trim(Mid(Txt, _
    a + Len(Replace(inicio, ".", "")), _
    b - a - Len(Replace(inicio, ".", ""))))

End Function
Function ExtrairValorCampo(Txt As String, campo As String) As String

    Dim p As Long
    Dim resultado As String

    p = InStr(1, Txt, campo, vbTextCompare)

    If p = 0 Then Exit Function

    resultado = Mid(Txt, p + Len(campo))

    resultado = Trim(resultado)

    ExtrairValorCampo = resultado

End Function

Function BuscarNumeroNota() As String

    Dim i As Long

    For i = 1 To UBound(dados)

        If InStr(dados(i), "Número:") > 0 Then

            BuscarNumeroNota = Trim(Split(Split(dados(i), "Número:")(1), "Série")(0))

            Exit Function

        End If

    Next

End Function

Function BuscarValorTotalItem(pos As Long) As String

    Dim i As Long

    For i = pos + 1 To pos + 8

        If InStr(1, dados(i), "Vl. Total", vbTextCompare) > 0 Then
            
            If i + 1 <= UBound(dados) Then
                BuscarValorTotalItem = dados(i + 1)
            End If

            Exit Function

        End If

    Next i

End Function

Function BuscarValorUnitario(Txt As String) As String

    Dim p As Long
    Dim valor As String

    p = InStr(1, Txt, "Vl. Unit.:", vbTextCompare)

    If p = 0 Then Exit Function

    valor = Mid(Txt, p + Len("Vl. Unit.:"))

    valor = Trim(valor)

    BuscarValorUnitario = valor

End Function
