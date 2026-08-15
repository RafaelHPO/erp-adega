Attribute VB_Name = "ModETLprocess"
Option Explicit

'==========================================
' PROCESSADOR PRINCIPAL
'==========================================

Public Sub ProcessarNFCe(ByVal caminho As String)

    On Error GoTo Erro

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False


    '==================================
    ' ABRE NFC-e
    '==================================

    If AbrirArquivo(caminho) = False Then

        MsgBox "Erro ao abrir arquivo."

        GoTo Finalizar

    End If


    '==================================
    ' LÊ PLANILHA
    '==================================

    CarregarDados


    '==================================
    ' CRIA ABA ETL
    '==================================

    CriarAbaETL


    '==================================
    ' EXTRAÇÕES
    '==================================

    ExtrairCabecalho

    ExtrairItens


    '==================================
    ' SALVA E FECHA
    '==================================

    FecharArquivo


Finalizar:

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    Exit Sub


Erro:

    MsgBox "Erro: " & Err.Description

    If Not wbOrigem Is Nothing Then

        wbOrigem.Close False

    End If

    Resume Finalizar


End Sub
