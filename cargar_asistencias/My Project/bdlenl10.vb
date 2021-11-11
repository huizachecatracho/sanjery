Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Public Class bdlenl10
    Public constr As String = ""
    Public Conn As IDbConnection = New SqlConnection
    Public Cmd As IDbCommand
    Public Consulta As String
    Public reader As IDataReader
    Public mensaje_error As String = ""

    Public Sub conectar()
        Conn.ConnectionString = "Data Source=10.34.67.16;Initial Catalog=ACCESSCONTROL;User ID=apilac;PASSWORD=apilac;timeout=600000;"
        constr = "Data Source=10.34.67.16;Initial Catalog=ACCESSCONTROL;User ID=apilac;PASSWORD=apilac"
        Try
            Conn.Open()

        Catch ex As Exception
            mensaje_error = "No se encuentra la Base de Datos"
        End Try
    End Sub
    Public Sub consultar(ByVal consulta As String)
        Cmd = Nothing
        Cmd = Conn.CreateCommand
        Cmd.CommandText = consulta
        Cmd.CommandTimeout = 600000
        reader = Cmd.ExecuteReader()
    End Sub

    Public Sub cerrar()
        Conn.Close()
    End Sub
    Public Sub ejecuta_sentencia(ByVal sentencia As String)
        Cmd = Nothing
        Cmd = Conn.CreateCommand
        Cmd.CommandTimeout = 600000
        Cmd.CommandText = sentencia
        Cmd.ExecuteNonQuery()
    End Sub

End Class
