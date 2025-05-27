Module Module1
    Public cn As New ADODB.Connection
    Public rs As New ADODB.Recordset

    Sub con()
        With cn
            If .State = 1 Then .Close()
            .ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:register.accdb;Persist Security Info=False;"
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .Open()
        End With
    End Sub
End Module
