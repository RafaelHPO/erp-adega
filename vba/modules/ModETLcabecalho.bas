Attribute VB_Name = "ModETLcabecalho"
Option Explicit

Sub ExtrairCabecalho()

    Dim ws As Worksheet
    Dim Linha As Long

    Set ws = wbOrigem.Worksheets("ETL_CAB")

    Linha = 2

    ws.Cells(Linha, 1) = ExtrairNumero
    ws.Cells(Linha, 2) = dados(1)
    ws.Cells(Linha, 3) = ExtrairCNPJ
    ws.Cells(Linha, 4) = ExtrairNF
    ws.Cells(Linha, 5) = ExtrairSerie
    ws.Cells(Linha, 6) = ExtrairEmissao
    ws.Cells(Linha, 7) = ExtrairValor("Valor total R$")
    ws.Cells(Linha, 8) = ExtrairValor("Descontos R$")
    ws.Cells(Linha, 9) = ExtrairValor("Valor pago R$")
    ws.Cells(Linha, 10) = ExtrairPagamento

End Sub

Function ExtrairCNPJ() As String

    Dim i As Long

    For i = 1 To UBound(dados)

        If InStr(dados(i), "CNPJ:") > 0 Then

            ExtrairCNPJ = Replace(dados(i), "CNPJ:", "")
            Exit Function

        End If

    Next

End Function

Function ExtrairNF() As String

    Dim i As Long
    Dim Txt As String

    For i = 1 To UBound(dados)

        If InStr(dados(i), "Número:") > 0 Then

            Txt = dados(i)

            ExtrairNF = Trim(Split(Split(Txt, "Número:")(1), "Série")(0))

            Exit Function

        End If

    Next

End Function

Function ExtrairSerie() As String

    Dim i As Long
    Dim Txt As String

    For i = 1 To UBound(dados)

        If InStr(dados(i), "Série:") > 0 Then

            Txt = dados(i)

            ExtrairSerie = Trim(Split(Split(Txt, "Série:")(1), "Emissão")(0))

            Exit Function

        End If

    Next

End Function

Function ExtrairEmissao() As String

    Dim i As Long

    For i = 1 To UBound(dados)

        If InStr(dados(i), "Emissão:") > 0 Then

            ExtrairEmissao = Trim(Split(Split(dados(i), "Emissão:")(1), "-")(0))
            Exit Function

        End If

    Next

End Function

Function ExtrairValor(Tag As String) As String

    Dim i As Long

    For i = 1 To UBound(dados)

        If InStr(1, dados(i), Tag, vbTextCompare) > 0 Then

            If IsNumeric(Replace(dados(i + 1), ",", ".")) Then
                ExtrairValor = dados(i + 1)
                Exit Function
            End If

        End If

    Next

End Function

Function ExtrairPagamento() As String

    Dim i As Long

    For i = 1 To UBound(dados)

        If InStr(dados(i), "Forma de pagamento") > 0 Then

            ExtrairPagamento = dados(i + 1)
            Exit Function

        End If

    Next

End Function

Function ExtrairNumero() As String

    ExtrairNumero = Format(Now, "yyyymmddhhmmss")

End Function

Function ExtrairEAN(Txt As String) As String

    Dim p As Long
    Dim codigo As String

    If InStr(Txt, "EAN") > 0 Then

        p = InStr(Txt, "EAN")

        codigo = Mid(Txt, p + 3)

        ExtrairEAN = Trim(codigo)

    End If

End Function
