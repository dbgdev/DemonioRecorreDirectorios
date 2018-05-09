Imports System.Data.SqlClient

Public Class QCError
    'INSERT INTO [dbo].[QCerror]
    '       ([itemCode]
    '       ,[tcin_QCerror]
    '       ,[timeCodeIn_QCerror]
    '       ,[detallesXml_QCerror])
    ' VALUES
    Private _cadenaConexion As String = "server=192.168.108.233;user id=admateriales;password=Mariquita10;database=materiales"

    Private itemCode As String
    Private tcin As String
    Private timeCodeIn As String
    Private detallesXml As String

    'Shared log2 As ObjectLog

    Public Sub New()
        Me.itemCode = String.Empty
        Me.tcin = String.Empty
        Me.timeCodeIn = String.Empty
        Me.detallesXml = String.Empty
    End Sub

    Public Sub New(itemCode As String, tcin As String, timeCodeIn As String, detallesXml As String)
        Me.itemCode = itemCode
        Me.tcin = tcin
        Me.timeCodeIn = timeCodeIn
        Me.detallesXml = detallesXml
    End Sub

    Friend Function InsertBD(cadenaConexion As String) As Boolean
        Dim resultOK As Boolean = False
        Using conn As New SqlConnection(cadenaConexion)
            conn.Open()
            Dim command As SqlCommand = conn.CreateCommand()
            command.CommandText = "INSERT INTO [dbo].[QCerror]
           ([itemCode]
           ,[tcin_QCerror]
           ,[timeCodeIn_QCerror]
           ,[detallesXml_QCerror])

            VALUES (  @itemCode, @tcin_QCerror, @timeCodeIn_QCerror, @detallesXml_QCerror )"

            command.Parameters.AddWithValue("@itemCode", Me.itemCode)
            command.Parameters.AddWithValue("@tcin_QCerror", Me.tcin)
            command.Parameters.AddWithValue("@timeCodeIn_QCerror", Me.timeCodeIn)
            command.Parameters.AddWithValue("@detallesXml_QCerror", Me.detallesXml)

            Try
                'Compruebo que haya insertado almenos 1 fila en la BD
                resultOK = command.ExecuteNonQuery() > 0
            Catch ex As Exception
                'log2.WriteLog(ex.Message)
            End Try

            conn.Close()

        End Using
        Return resultOK
    End Function
End Class
