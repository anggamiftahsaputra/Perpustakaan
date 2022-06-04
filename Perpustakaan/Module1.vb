Imports System.Data.SqlClient
Module Module1
    Public Conn As SqlConnection
    Public Da As SqlDataAdapter
    Public Ds As DataSet
    Public Rd As SqlDataReader
    Public Cmd As SqlCommand
    Public MyDB As String
    Public Sub koneksi()
        MyDB = "data source=DESKTOP-88CF28Q\ANGGAIT; initial catalog=DbPerpus; integrated security=true"
        Conn = New SqlConnection(MyDB)
        If Conn.State = ConnectionState.Closed Then Conn.Open()
    End Sub
End Module
