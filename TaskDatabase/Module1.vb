Imports System.Data.SqlClient
Module Module1
    Public Conn As SqlConnection
    Public da As SqlDataAdapter
    Public ds As DataSet
    Public cmd As SqlCommand
    Public rd As SqlDataReader
    Public Str As String
    Public Sub Koneksi()
        Str = "data source=DESKTOP-N3KO3RS; initial catalog=JualBuku; integrated security=true"
        Conn = New SqlConnection(Str)
        If Conn.State = ConnectionState.Closed Then
            Conn.Open()
        End If
    End Sub
End Module
