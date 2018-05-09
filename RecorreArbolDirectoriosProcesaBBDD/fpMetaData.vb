Imports System.Data.SqlClient
Imports System.IO

Public Class FpMetaData
    Public log2 As Specialized.StringCollection = New Specialized.StringCollection()
#Region "Campos de la BD"
    'id_fpMetaData int PRIMARY KEY IDENTITY (1,1) Not NULL ,
    'fileName_fpMetaData varchar(50) Default NULL,
    'date_fpMetaData datetime Default NULL,
    'path_fpMetaData varchar(100) Default NULL,
    'error_cvmc varchar(250) Default NULL,
    'estadoDalet varchar(20) Default NULL,
    'festadoDalet datetime Default NULL,
    'urlproxydalet varchar(250) Default NULL,
    'categoryDalet varchar(250) Default NULL,
    'qcDalet int Default NULL,
    'itemCode varchar(45) Default NULL,
    'ObjectIdDalet varchar(50) Default NULL,
    'estadoProvys int Default NULL,
    'festadoProvys datetime Default NULL,
    'qcVisionados int Default NULL,
    'fqcVisionados datetime Default NULL,
    'UsuarioVisionados varchar(45) Default NULL,
    'qcTecnico int Default NULL,
    'fqctecnico datetime Default NULL,
    'usuarioTecnico varchar(45) Default NULL,
#End Region
    Private _id_BD As Integer
    'Private _fileName As String
    '??? Aun no tengo claro si es un String o un Datetime
    'Private _fileDate As String
    'Private _filePath As String
    'Private _error_cvmc As String
    Private _estadoDalet As String
    Private _festadoDalet As String
    Private _urlProxyDalet As String
    Private _categoryOriginal As String
    'Private _categoryDalet As String
    Private _qcDalet As Integer
    'Private _itemCode As String
    Private _ObjectIdDalet As String
    Private _estadoProvys As String
    Private _festadoProvys As String
    Private _qcVisionados As Integer
    Private _fqcVisionados As String
    Private _usuarioVisionados As String
    Private _qcTecnico As Integer
    Private _fqcTecnico As String
    Private _usuariotecnico As String
    'Private _categoryOriginal As String

    Public Sub New()
        _id_BD = -1
        Me.FileName = String.Empty
        _FileDate = DateAndTime.Today
        _FilePath = String.Empty
        _Error_CVMC = String.Empty
        _estadoDalet = String.Empty
        _festadoDalet = String.Empty
        _urlProxyDalet = String.Empty
        _CategoryDalet = String.Empty
        _qcDalet = -1
        _ItemCode = String.Empty
        _ObjectIdDalet = String.Empty
        _estadoProvys = String.Empty
        _festadoProvys = String.Empty
        _qcVisionados = -1
        _fqcVisionados = String.Empty
        _usuarioVisionados = String.Empty
        _qcTecnico = -1
        _fqcTecnico = String.Empty
        _usuariotecnico = String.Empty
    End Sub

    Public Sub New(fileName As String, fileDate As String, filePath As String, error_cvmc As String, estadoDalet As String, festadoDalet As String, urlProxyDalet As String, categoryDalet As String, qcDalet As Integer, itemCode As String, ObjectIdDalet As String, estadoProvys As String, festadoProvys As String, qcVisionados As Integer, fqcVisionados As String, usuarioVisionados As String, qcTecnico As String, fqcTecnico As String, usuariotecnico As String)
        _id_BD = -1
        _FileName = fileName
        _FileDate = fileDate
        _FilePath = filePath
        _Error_CVMC = error_cvmc
        _estadoDalet = estadoDalet
        _festadoDalet = festadoDalet
        _urlProxyDalet = urlProxyDalet
        _CategoryDalet = categoryDalet
        _qcDalet = qcDalet
        _ItemCode = itemCode
        _ObjectIdDalet = ObjectIdDalet
        _estadoProvys = estadoProvys
        _festadoProvys = festadoProvys
        _qcVisionados = qcVisionados
        _fqcVisionados = fqcVisionados
        _usuarioVisionados = usuarioVisionados
        _qcTecnico = qcTecnico
        _fqcTecnico = fqcTecnico
        _usuariotecnico = usuariotecnico
    End Sub

    Friend Function InsertBD(cadenaConexion As String) As Boolean
        Dim resultOK As Boolean = False
        Using conn As New SqlConnection(cadenaConexion)
            conn.Open()
            Dim command As SqlCommand = conn.CreateCommand()
            command.CommandText = "INSERT INTO [dbo].[fileprocessmetadata]
           ([fileName_fpMetaData]
           ,[date_fpMetaData]
           ,[path_fpMetaData]
           ,[error_cvmc]
           ,[itemCode]
           ,[categoryOriginal]
           )
     VALUES ( @fileName , @fileDateCreation , @filePath , @error_cvmc, @itemCode, @categoryOriginal )"

            command.Parameters.AddWithValue("@fileName", FileName)
            command.Parameters.AddWithValue("@fileDateCreation", Me.FileDate)
            command.Parameters.AddWithValue("@filePath", Me.FilePath)
            command.Parameters.AddWithValue("@error_cvmc", """OK""")
            command.Parameters.AddWithValue("@itemCode", Me.ItemCode)
            command.Parameters.AddWithValue("@categoryOriginal", Me.CategoryOriginal)

            Try
                'Compruebo que haya insertado almenos 1 fila en la BD
                resultOK = command.ExecuteNonQuery() > 0
            Catch ex As Exception
                log2.Add(ex.Message + command.CommandText)
            End Try



            conn.Close()

        End Using
        Return resultOK
    End Function

    Public Property FileName As String
    Public Property FileDate As DateTime
    Public Property FilePath As String
    Public Property Error_CVMC As String
    Public Property ItemCode As String
    Public Property CategoryDalet As String

    Public Property CategoryOriginal As String
        Get
            Return CategoryOriginal1
        End Get
        Set(value As String)
            CategoryOriginal1 = value
        End Set
    End Property

    Public Property CategoryOriginal1 As String
        Get
            Return _categoryOriginal
        End Get
        Set(value As String)
            _categoryOriginal = value
        End Set
    End Property

    Public Sub GetAssetByItemCode(cadenaConexion As String, ByVal ItemCode As String)
        Dim resultOK As Boolean = False
        'Lanzar aqui un select y devolver 1 fila solo

        Using conn As New SqlConnection(cadenaConexion)
            conn.Open()
            Dim command As SqlCommand = conn.CreateCommand()
            Dim reader As SqlDataReader

            command.CommandText = " SELECT fp.id_fpMetaData, fp.fileName_fpMetaData, fp.date_fpMetaData, fp.path_fpMetaData,
                    fp.error_cvmc,
                    fp.estadoDalet, fp.festadoDalet, fp.urlproxydalet, fp.categoryOriginal, fp.categoryDalet, fp.qcDalet, fp.itemCode, fp.ObjectIdDalet,
                    fp.estadoProvys, fp.festadoProvys,
                    fp.qcVisionados, fp.fqcVisionados, fp.UsuarioVisionados,
                    fp.qcTecnico, fp.fqctecnico, fp.usuarioTecnico " &
                                  " FROM fileprocessmetadata as fp " &
                                  " WHERE fp.itemCode = @ItemCode " &
                                  " ORDER BY  fp.date_fpMetaData DESC "
            command.Parameters.AddWithValue("@ItemCode", ItemCode)
            Try
                'Compruebo que haya insertado almenos 1 fila en la BD
                reader = command.ExecuteReader()
                If reader.HasRows Then
                    'añado el resultado a las filas
                    reader.Read()
                    Me._id_BD = reader.Item("id_fpMetaData")
                    Me.FileName = IIf(IsDBNull(reader.Item("fileName_fpMetaData")), String.Empty, reader.Item("fileName_fpMetaData"))
                    Me.FileDate = IIf(IsDBNull(reader.Item("date_fpMetaData")), Date.MinValue, reader.Item("date_fpMetaData"))
                    Me.FilePath = IIf(IsDBNull(reader.Item("error_cvmc")), String.Empty, reader.Item("error_cvmc"))
                    Me._estadoDalet = IIf(IsDBNull(reader.Item("estadoDalet")), String.Empty, reader.Item("estadoDalet"))
                    Me._festadoDalet = IIf(IsDBNull(reader.Item("festadoDalet")), String.Empty, reader.Item("festadoDalet"))
                    Me._urlProxyDalet = IIf(IsDBNull(reader.Item("urlproxydalet")), String.Empty, reader.Item("urlproxydalet"))
                    Me.CategoryOriginal = IIf(IsDBNull(reader.Item("categoryOriginal")), String.Empty, reader.Item("categoryOriginal"))
                    Me._qcDalet = IIf(IsDBNull(reader.Item("qcDalet")), -1, reader.Item("qcDalet"))
                    Me.ItemCode = IIf(IsDBNull(reader.Item("itemCode")), String.Empty, reader.Item("itemCode"))
                    Me._ObjectIdDalet = IIf(IsDBNull(reader.Item("ObjectIdDalet")), String.Empty, reader.Item("ObjectIdDalet"))
                    Me._estadoProvys = IIf(IsDBNull(reader.Item("estadoProvys")), String.Empty, reader.Item("estadoProvys"))
                    Me._festadoProvys = IIf(IsDBNull(reader.Item("festadoProvys")), String.Empty, reader.Item("festadoProvys"))
                    Me._qcVisionados = IIf(IsDBNull(reader.Item("qcVisionados")), -1, reader.Item("qcVisionados"))
                    Me._fqcVisionados = IIf(IsDBNull(reader.Item("fqcVisionados")), String.Empty, reader.Item("fqcVisionados"))
                    Me._usuarioVisionados = IIf(IsDBNull(reader.Item("UsuarioVisionados")), String.Empty, reader.Item("UsuarioVisionados"))
                    Me._qcTecnico = IIf(IsDBNull(reader.Item("qcTecnico")), -1, reader.Item("qcTecnico"))
                    Me._fqcTecnico = IIf(IsDBNull(reader.Item("fqctecnico")), String.Empty, reader.Item("fqctecnico"))
                    Me._usuariotecnico = IIf(IsDBNull(reader.Item("usuarioTecnico")), String.Empty, reader.Item("usuarioTecnico"))

                End If
                reader.Close()
            Catch ex As Exception
                log2.Add(ex.Message)
            End Try
        End Using

    End Sub
    Public Sub SetAssetByItemCode(cadenaConexion As String, ByVal ItemCode As String)
        Dim resultOK As Boolean = False
        'Lanzar aqui un select y devolver 1 fila solo

        Using conn As New SqlConnection(cadenaConexion)
            conn.Open()
            Dim command As SqlCommand = conn.CreateCommand()

            command.CommandText = " UPDATE [dbo].[fileprocessmetadata] " &
                                  " SET [fileName_fpMetaData] = @fileName ,
		                                [date_fpMetaData] = @fileDateCreation ,
                                        [path_fpMetaData] = @filePath ,
                                        [error_cvmc] = @error_cvmc ,
                                        [estadoDalet] = @estadoDalet ,
                                        [festadoDalet] = @festadoDalet ,
                                        [urlproxydalet] = @urlProxyDalet ,
                                        [categoryOriginal] = @categoryOriginal ,
                                        [categoryDalet] = @categoryDalet ,
                                        [qcDalet] = @qcDalet ,
                                        [itemCode] = @ItemCode ,
                                        [ObjectIdDalet] = @ObjectIdDalet,
                                        [estadoProvys] = @estadoProvys,
                                        [festadoProvys] = @festadoProvys,
                                        [qcVisionados] = @qcVisionados,
                                        [fqcVisionados] = @fqcVisionados,
                                        [UsuarioVisionados] = @UsuarioVisionados,
                                        [qcTecnico] = @qcTecnico,
                                        [fqctecnico] = @fqctecnico,
                                        [usuarioTecnico] = @usuarioTecnico " &
                                  " WHERE id_fpMetaData = @id_fpMetaData AND itemCode = @ItemCode "
            'Add Parameters here, first filter parameters, next ones updating values
            command.Parameters.AddWithValue("@ItemCode", ItemCode)
            If Me._id_BD <> -1 Then
                command.Parameters.AddWithValue("@id_fpMetaData", Me._id_BD)
            End If
            'File Values
            command.Parameters.AddWithValue("@fileName", Me.FileName)
            command.Parameters.AddWithValue("@fileDateCreation", Me.FileDate)
            command.Parameters.AddWithValue("@filePath", Me.FilePath)
            command.Parameters.AddWithValue("@error_cvmc", Me.Error_CVMC)
            'Dalet values BD
            command.Parameters.AddWithValue("@estadoDalet", Me._estadoDalet)
            command.Parameters.AddWithValue("@festadoDalet", Me._festadoDalet)
            command.Parameters.AddWithValue("@urlProxyDalet", Me._urlProxyDalet)
            command.Parameters.AddWithValue("@categoryOriginal", Me.CategoryOriginal)
            command.Parameters.AddWithValue("@categoryDalet", Me.CategoryDalet)
            command.Parameters.AddWithValue("@qcDalet", Me._qcDalet)
            command.Parameters.AddWithValue("@ObjectIdDalet", Me._ObjectIdDalet)
            'Provys values BD
            command.Parameters.AddWithValue("@estadoProvys", Me._estadoProvys)
            command.Parameters.AddWithValue("@festadoProvys", Me._festadoProvys)

            command.Parameters.AddWithValue("@qcVisionados", Me._qcVisionados)
            command.Parameters.AddWithValue("@fqcVisionados", Me._fqcVisionados)
            command.Parameters.AddWithValue("@UsuarioVisionados", Me._usuarioVisionados)

            command.Parameters.AddWithValue("@qcTecnico", Me._qcTecnico)
            command.Parameters.AddWithValue("@fqcTecnico", Me._fqcTecnico)
            command.Parameters.AddWithValue("@usuarioTecnico", Me._usuariotecnico)

            Try
                'Compruebo que haya se haya actualizado 1 fila en la solo BD
                resultOK = command.ExecuteNonQuery() = 1
            Catch ex As Exception
                log2.Add(ex.Message)
            End Try

        End Using
    End Sub
    Public Function GetStatusDaletByNameStatus(cadenaConexion As String, ByVal NameStatus As String)
        Dim status As String = String.Empty
        Using conn As New SqlConnection(cadenaConexion)
            conn.Open()
            Dim command As SqlCommand = conn.CreateCommand()
            Dim reader As SqlDataReader

            command.CommandText = " SELECT ed.* " &
                                  " FROM [dbo].[estadoDalet] as ed " &
                                  " WHERE ed.nombre_estadoDalet = @NameStatus "

            command.Parameters.AddWithValue("@NameStatus", NameStatus)
            Try
                'Compruebo que haya insertado almenos 1 fila en la BD
                reader = command.ExecuteReader()
                If reader.HasRows Then
                    'añado el resultado a las filas
                    reader.Read()
                    status = IIf(IsDBNull(reader.Item("estadoAsset_estadoDalet")), String.Empty, reader.Item("estadoAsset_estadoDalet"))
                End If
            Catch ex As Exception
                log2.Add(ex.Message)
            End Try

        End Using

        Return status
    End Function
    Public Shared Sub CreateFileLCK(pathOrigin As String, nameFile As String)
        Dim LogCreate As ObjectLog = New ObjectLog(nameFile, "C:\test")
        Try
            If Not Directory.Exists(pathOrigin) Then
                Directory.CreateDirectory(pathOrigin)
            End If

            If Not File.Exists(nameFile) Then
                File.Create(pathOrigin, 0).Close()
            End If
        Catch ex As Exception
            LogCreate.WriteLog(ex.Message)
        End Try
        LogCreate.WriteLog(nameFile & "Processed OK")
    End Sub
End Class
