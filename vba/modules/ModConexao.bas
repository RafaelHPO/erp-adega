Attribute VB_Name = "ModConexao"
Public Conn As ADODB.Connection
Public MotivoLicenca As String
Public DataValidadeLicenca As Date
Public MensagemLicenca As String



Public Function ConectarBanco() As Boolean

    On Error GoTo Erro

    Set Conn = New ADODB.Connection
' <ip-removido>
  Conn.ConnectionString = _
        "Driver={MySQL ODBC 9.7 Unicode Driver};" & _
        "Server=<host-do-banco>;" & _
        "Port=<porta>;" & _
        "Database=<nome-do-banco>;" & _
        "User=<usuario>;" & _
        "Password=<senha>;" & _
        "Option=3;"

    Conn.Open

    ConectarBanco = True

    Exit Function

Erro:

    MsgBox "Erro ao conectar: " & Err.Description

    ConectarBanco = False

End Function

Public Function VerificarLicenca() As Boolean

    Dim rs As ADODB.Recordset

    Set rs = Conn.Execute("CALL PROC_VERIFICARLICENCA();")


    MotivoLicenca = rs!Retorno

    MensagemLicenca = rs!MSG


    If IsNull(rs!VALIDADE) Then

        DataValidadeLicenca = 0

    Else

        DataValidadeLicenca = rs!VALIDADE

    End If


    VerificarLicenca = (rs!SUCESSO = 1)


    rs.Close
    Set rs = Nothing

End Function

Sub TestarConexao()

    If ConectarBanco Then

        MsgBox "Banco conectado!"

        Conn.Close

    Else

        MsgBox "Falha conexão"

    End If

End Sub

