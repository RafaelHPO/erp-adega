Attribute VB_Name = "modGlob"
Public Function CampoExiste(rs As ADODB.Recordset, NomeCampo As String) As Boolean

    Dim campo As ADODB.Field

    CampoExiste = False

    For Each campo In rs.Fields

        If UCase(campo.Name) = UCase(NomeCampo) Then

            CampoExiste = True
            Exit Function

        End If

    Next campo

End Function
