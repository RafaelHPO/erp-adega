Attribute VB_Name = "tailscale"
Public Sub AbrirTailscale()

    On Error GoTo TrataErro

    Dim caminho As String

    caminho = "C:\Program Files\Tailscale\tailscale-ipn.exe"

    If Dir(caminho) = "" Then Exit Sub

    If Not ProcessoRodando("tailscale-ipn.exe") Then
        Shell """" & caminho & """", vbNormalFocus
    End If

    Exit Sub

TrataErro:

End Sub

Public Function ProcessoRodando(ByVal NomeProcesso As String) As Boolean

    Dim objWMI As Object
    Dim processos As Object
    Dim processo As Object

    ProcessoRodando = False

    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")

    Set processos = objWMI.ExecQuery( _
        "SELECT Name FROM Win32_Process WHERE Name = '" & NomeProcesso & "'" _
    )

    For Each processo In processos

        ProcessoRodando = True
        Exit For

    Next processo

    Set processo = Nothing
    Set processos = Nothing
    Set objWMI = Nothing

End Function
