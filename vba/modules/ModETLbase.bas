Attribute VB_Name = "ModETLbase"
Option Explicit

'==========================================
' VARIÁVEIS GLOBAIS
'==========================================

Public wbOrigem As Workbook
Public wsOrigem As Worksheet

Public CaminhoArquivo As String

Public dados() As String

'==========================================
' ABRE O XLSX
'==========================================

Public Function AbrirArquivo(ByVal arquivo As String) As Boolean

    On Error GoTo Erro

    Set wbOrigem = Application.Workbooks.Open( _
    Filename:=arquivo, _
    ReadOnly:=False, _
    UpdateLinks:=False)

    Set wsOrigem = wbOrigem.Worksheets(1)

    'wbOrigem.Windows(1).Visible = False

    CaminhoArquivo = arquivo

    AbrirArquivo = True

    Exit Function

Erro:

    AbrirArquivo = False

End Function

'==========================================
' FECHA O XLSX
'==========================================

Public Sub FecharArquivo()

    If wbOrigem Is Nothing Then Exit Sub

    wbOrigem.Close SaveChanges:=True

    Set wsOrigem = Nothing
    Set wbOrigem = Nothing

Application.Visible = False

ThisWorkbook.RefreshAll

End Sub

'==========================================
' LÊ TODA A PLANILHA
'==========================================

Public Sub CarregarDados()

    Dim cel As Range

    Dim Texto As String

    Dim Partes() As String

    Dim i As Long
    Dim j As Long

    ReDim dados(1 To 5000)

    i = 0

    For Each cel In wsOrigem.UsedRange

        Texto = Trim(CStr(cel.Value))

        If Texto <> "" Then

            Texto = Replace(Texto, vbCrLf, vbLf)
            Texto = Replace(Texto, vbCr, vbLf)

            Partes = Split(Texto, vbLf)

            For j = LBound(Partes) To UBound(Partes)

                If Trim(Partes(j)) <> "" Then

                    i = i + 1

                    If i > UBound(dados) Then

                        ReDim Preserve dados(1 To UBound(dados) + 5000)

                    End If

                    dados(i) = Trim(Partes(j))

                End If

            Next j

        End If

    Next cel

    ReDim Preserve dados(1 To i)

End Sub

'==========================================
' DEBUG
'==========================================

Public Sub MostrarDados()

    Dim i As Long

    For i = LBound(dados) To UBound(dados)

        Debug.Print i & " -> " & dados(i)

    Next i

End Sub
