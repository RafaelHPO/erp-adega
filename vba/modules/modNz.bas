Attribute VB_Name = "modNz"
Public Function Nz(valor As Variant, Optional valorPadrao As Variant = "") As Variant

    If IsNull(valor) Then
        Nz = valorPadrao
    Else
        Nz = valor
    End If

End Function

Public Function NzDB(valor As Variant) As Variant

    If IsNull(valor) Then
        NzDB = Null
    ElseIf Trim(valor & "") = "" Then
        NzDB = Null
    Else
        NzDB = valor
    End If

End Function


Public Function SqlTexto(valor As Variant) As String

    If Trim(valor & "") = "" Then
        SqlTexto = "NULL"
    Else
        SqlTexto = "'" & Replace(valor, "'", "''") & "'"
    End If

End Function
Public Function SqlNumero(valor As Variant) As String

    If Trim(valor & "") = "" Then
        SqlNumero = "NULL"
    ElseIf Not IsNumeric(valor) Then
        Err.Raise vbObjectError + 1, , "Valor numérico inválido: " & valor
    Else
        SqlNumero = Replace(CStr(CDbl(valor)), ",", ".")
    End If

End Function
Public Function NzDbl(valor As Variant) As Double

    If IsNull(valor) Then
        NzDbl = 0
        Exit Function
    End If

    valor = Trim(valor & "")

    If valor = "" Then
        NzDbl = 0
        Exit Function
    End If

    valor = Replace(valor, ".", ",")

    If IsNumeric(valor) Then
        NzDbl = CDbl(valor)
    Else
        NzDbl = 0
    End If

End Function

Public Function SqlData(valor As Variant) As String

    If Trim(valor & "") = "" Then
        SqlData = "NULL"
    Else
        SqlData = "'" & Format(CDate(valor), "yyyy-mm-dd") & "'"
    End If

End Function

Public Function FormatarDataCampo(valor As Variant) As String

    If IsNull(valor) Or IsEmpty(valor) Or Trim(valor & "") = "" Then
        FormatarDataCampo = ""
    Else
        FormatarDataCampo = Format(CDate(valor), "dd/mm/yyyy")
    End If

End Function

Public Function NzDate(valor As Variant) As Date

    If IsNull(valor) Then
        NzDate = 0
        Exit Function
    End If

    If Trim(valor & "") = "" Then
        NzDate = 0
        Exit Function
    End If

    If IsDate(valor) Then
        NzDate = CDate(valor)
    Else
        NzDate = 0
    End If

End Function
