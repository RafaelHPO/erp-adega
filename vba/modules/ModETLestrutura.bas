Attribute VB_Name = "ModETLestrutura"
Option Explicit

Sub CriarAbaETL()

    Dim ws As Worksheet

    On Error Resume Next

    Application.DisplayAlerts = False

wbOrigem.Worksheets("ETL_CAB").Delete
wbOrigem.Worksheets("ETL_ITEM").Delete

    Application.DisplayAlerts = True

    On Error GoTo 0


    Set ws = wbOrigem.Worksheets.Add
    ws.Name = "ETL_CAB"

    ws.Range("A1:J1") = Array( _
        "ID_NOTA", _
        "EMITENTE", _
        "CNPJ", _
        "NUMERO", _
        "SERIE", _
        "EMISSAO", _
        "VALOR_TOTAL", _
        "DESCONTO", _
        "VALOR_PAGO", _
        "FORMA_PAGAMENTO")


    Set ws = wbOrigem.Worksheets.Add
    ws.Name = "ETL_ITEM"

    ws.Range("A1:H1") = Array( _
        "ID_NOTA", _
        "EAN", _
        "CODIGO FORNECEDOR", _
        "DESCRICAO", _
        "QTDE", _
        "UN", _
        "VL_UNIT", _
        "VL_TOTAL", _
        "CNPJ")


End Sub

