Imports System.Data.SqlClient
Imports System.IO
Imports System.Configuration
Imports System.Threading
Imports System.Web

Public Class RecursiveFileSearch
    Private _pathFolderOrigin As String = String.Empty
    Private _pathFolderDestiny As String = String.Empty
    Private listaFicheros As List(Of String) = New List(Of String)
    Shared log As Specialized.StringCollection = New Specialized.StringCollection()
    Shared log2 As ObjectLog


    Public Sub New()
        _pathFolderOrigin = "C:\test\Carpeta_1"
        _pathFolderDestiny = "\\Nassau\cvmc\CVMC_S&C\Desarrollo\pruebas\destino"
        Initialize(_pathFolderOrigin, _pathFolderDestiny)
    End Sub

    Public Sub New(originPath As String, destinyPath As String)
        Me._pathFolderOrigin = originPath
        If Not IsDBNull(originPath) And originPath <> "" Then
            If Not IsDBNull(destinyPath) And destinyPath <> "" Then
                Initialize(originPath, destinyPath)
            End If
        End If
    End Sub

    Private Sub Initialize(pathFolderOrigin As String, pathFolderDestiny As String)
        'server produccion de App Config
        '<add key = "pathOrigen" value="\\nasapunt\FTPContinguts" />
        '<add key = "pathDestino" value="\\netapp.cvmc.es\dalet\Import\Productoras"/>
        '<add key = "ConnectionString" value="server=192.168.108.233;user id=admateriales;password=Mariquita10;database=materiales" />

        'test en mi maquina de App Config con carpeta de red de estino
        '<add key = "pathOrigen" value="C:\test\Carpeta_1" />
        '<add key = "pathDestino" value="\\Nassau\cvmc\CVMC_S%26C\Desarrollo\pruebas\destino"/>
        '<add key = "ConnectionString" value="server=192.168.108.233;user id=admateriales;password=Mariquita10;database=MATERIALES_DEV" />

        'test en mi maquina de App Config
        '<add key = "pathOrigen" value="C:\test\Carpeta_1" />
        '<add key = "pathDestino" value="C:\desconeguda"/>
        '<add key = "ConnectionString" value="server=192.168.108.233;user id=admateriales;password=Mariquita10;database=MATERIALES_DEV" />

        log2 = New ObjectLog("inicializando", "C:\test")
        Dim rootDir As System.IO.DirectoryInfo = New IO.DirectoryInfo(HttpUtility.UrlDecode(pathFolderOrigin))
        Dim destDir As System.IO.DirectoryInfo = New IO.DirectoryInfo(HttpUtility.UrlDecode(pathFolderDestiny))

        If rootDir.Exists Then
            'log.Add("Comienzo a procesar: " & rootDir.FullName)
            WalkDirectoryTree(rootDir, destDir)
        Else
            'log.Add("ERROR: Invalid root directory or unaccesible")
            log2.WriteLog("ERROR: Invalid root directory or unaccesible")
        End If

    End Sub

    Private Sub WalkDirectoryTree(rootDir As DirectoryInfo, destDir As DirectoryInfo)
        Dim files As System.IO.FileInfo() = Nothing
        Dim subDirs As System.IO.DirectoryInfo() = Nothing

        Try
            'Filtro por tipo de archivos  {".avi", ".mpeg", ".mpg", ".qt", ".wmv", ".asf", ".rm", ".ram", ".flv", ".mxf", ".xml", ".mov", ".mp4"}
            'Mi filtro de pruebas ' {".avi", ".mpeg", ".mpg", ".qt", ".wmv", ".asf", ".rm", ".ram", ".flv", ".mxf", ".xml", ".mov", ".mp4"}
            Dim readExtensions As String = HttpUtility.UrlEncode(ConfigurationManager.AppSettings.Get("allowedExtensionFiles"))

            Dim extensions As String() = HttpUtility.UrlDecode(readExtensions).Replace(" ", "").Split(",")
            files = rootDir.GetFiles().Where(Function(f) extensions.Contains(f.Extension.ToLower())).ToArray()

        Catch e As UnauthorizedAccessException
            log2.WriteLog(e.Message)
        Catch e As System.IO.DirectoryNotFoundException
            log2.WriteLog(e.Message)
        End Try

        If files IsNot Nothing Then
            For Each fi As System.IO.FileInfo In files
                'Dim fileInfo As FileInfo = New FileInfo(e.FullPath)
                While IsFileLocked(fi)
                    Thread.Sleep(5000)
                End While

                'Procesar fichero aquí

                log2.WriteLog("Processing " & fi.FullName)


                Try
                    log2.WriteLog("Antes de crear fpMeta")

                    Dim fpMeta As FpMetaData = New FpMetaData With {
                    .FileName = fi.Name,
                    .FileDate = DateAndTime.Now,
                    .FilePath = rootDir.FullName,
                    .Error_CVMC = """ OK """,
                    .ItemCode = fi.Name.Split(".")(0),
                    .CategoryOriginal = rootDir.FullName.Replace("\\", "\").Replace("C:", "")
                    }
                    log2.WriteLog("Antes del insert")
                    'DB: Cadena de conexion de pruebas en local con MySQL--->"server=192.168.170.20;database=davidbarreiro_schema;user id=user;password=user;port=3306;"
                    'DB: Cadena de conexion para la BD de Produccion "server=192.168.108.233;user id=admateriales;password=Mariquita10;database=materiales"
                    'DB: Cadena de conexion para la BD de Desarrollo "server=192.168.108.233;user id=admateriales;password=Mariquita10;database=MATERIALES_DEV"
                    Dim cadenaConexion As String = ConfigurationManager.AppSettings.Get("ConnectionString")
                    If Not listaFicheros.Contains(fi.Name) Then
                        listaFicheros.Add(fi.Name)
                        Dim estaInsertado As Boolean = fpMeta.InsertBD(cadenaConexion)
                        log2.WriteLog("Despues del insert de " + fi.Name)
                    End If


                    For Each elem As String In fpMeta.log2
                        log2.WriteLog(elem)
                    Next

                    'DB: Demonio 2.0 ahora debe consultar en una tabla de la BD (routingPath) 
#Region "Demonio 2.0"
                    '   una serie de patrones basados segun su path de Origen (rootDir.FullName)???
                    '   Si encuentro el patron de carpetas en la tabla lo envio al destino que me indique la tabla si no
                    '   va al directorio por defecto pasado en Appconfig

                    '_pathFolderDestiny = SearchDestinyDirectoryByPatern(rootDir.FullName, cadenaConexion)

                    'If Not IsNothing(_pathFolderDestiny) And _pathFolderDestiny <> "" Then
                    '    log2.WriteLog("Moviendo---> " & fi.FullName & " to " & _pathFolderDestiny + "\" + fi.Name)
                    '    File.Move(fi.FullName, _pathFolderDestiny + "\" + fi.Name)
                    'Else
                    '    log2.WriteLog("Moviendo---> " & fi.FullName & " to " & HttpUtility.UrlDecode(destDir.FullName) + "\" + fi.Name)
                    '    File.Move(fi.FullName, HttpUtility.UrlDecode(destDir.FullName) + "\" + fi.Name)
                    'End If


                    'File.Move(fi.FullName, HttpUtility.UrlDecode(destDir.FullName) + "\" + fi.Name)
#End Region

                    'TODO: Intentar refactorizar este cacho codigo
                    _pathFolderDestiny = SearchDestinyDirectoryByPatern(rootDir.FullName, cadenaConexion)

                    If Not IsNothing(_pathFolderDestiny) And _pathFolderDestiny <> "" Then
                        log2.WriteLog("Inicio el copiado---> " & fi.FullName & " to " & HttpUtility.UrlDecode(_pathFolderDestiny) + "\" + fi.Name)

                        Dim archivodestino As System.IO.FileInfo = New FileInfo(HttpUtility.UrlDecode(_pathFolderDestiny) + "\" + fi.Name)
                        While IsFileLocked(fi)
                            Thread.Sleep(5001)
                        End While
                        If Not archivodestino.Exists Then
                            log2.WriteLog("Copiando Destino BD---> " & fi.FullName & " to " & HttpUtility.UrlDecode(_pathFolderDestiny) + "\" + fi.Name)
                            File.Copy(fi.FullName, HttpUtility.UrlDecode(_pathFolderDestiny) + "\" + fi.Name)
                        End If

                        While IsFileLocked(fi)
                            Thread.Sleep(5015)
                        End While

                        log2.WriteLog("Borrando fichero original: " + fi.FullName)
                        File.Delete(fi.FullName)
                    Else
                        log2.WriteLog("Inicio el copiado---> " & fi.FullName & " to " & HttpUtility.UrlDecode(destDir.FullName) + "\" + fi.Name)

                        Dim archivodestino As System.IO.FileInfo = New FileInfo(HttpUtility.UrlDecode(destDir.FullName) + "\" + fi.Name)
                        While IsFileLocked(fi)
                            Thread.Sleep(5001)
                        End While
                        If Not archivodestino.Exists Then
                            log2.WriteLog("Copiando Destino por Defecto---> " & fi.FullName & " to " & HttpUtility.UrlDecode(destDir.FullName) + "\" + fi.Name)
                            File.Copy(fi.FullName, HttpUtility.UrlDecode(destDir.FullName) + "\" + fi.Name)
                        End If

                        While IsFileLocked(fi)
                            Thread.Sleep(5015)
                        End While

                        log2.WriteLog("Borrando fichero original: " + fi.FullName)
                        File.Delete(fi.FullName)
                    End If
                    'TODO: Intentar refactorizar este cacho codigo



                Catch ex As Exception
                    log2.WriteLog(ex.Message)
                End Try
                'If Not rootDir.GetFiles().Contains(fi) Then
                '    listaFicheros.Remove(fi.Name)
                'End If
            Next

            If rootDir.GetFiles().Count > 0 Then
                log2.WriteLog("Borrando lista ficheros no permitidos")
                For Each fileItem As FileInfo In rootDir.GetFiles()
                    listaFicheros.Remove(fileItem.FullName)
                Next
            End If

            subDirs = rootDir.GetDirectories()
            For Each dirInfo As System.IO.DirectoryInfo In subDirs
                'Imprimo nombre directorio , proceso directorio, sigo recorriendo aquí
                'DB: Por regla de negocio me salto cualquier directorio que contenga en su ruta una #
                If Not dirInfo.FullName.Contains("#") Then
                    WalkDirectoryTree(dirInfo, destDir)
                End If

            Next
        End If
    End Sub

    Private Function SearchDestinyDirectoryByPatern(fullName As String, conectionString As String) As String

        Dim path As String = String.Empty
        'Conectar con la BD
        'Dado un path de origen busca el path de destion en la tabla routing path
        Using conn As New SqlConnection(conectionString)
            conn.Open()
            Dim command As SqlCommand = conn.CreateCommand()
            Dim reader As SqlDataReader

            command.CommandText = " SELECT rp.*
                                    FROM dbo.routingPath rp
                                    WHERE rp.patternType_routingPath= @type AND  pathOrigin_routingPath = @pathOrigin "

            command.Parameters.AddWithValue("@pathOrigin", fullName)
            command.Parameters.AddWithValue("@type", "folder")
            Try
                'Compruebo que haya insertado almenos 1 fila en la BD
                reader = command.ExecuteReader()
                If reader.HasRows Then
                    'añado el resultado a las filas
                    reader.Read()

                    path = IIf(IsDBNull(reader.Item("pathDestiny_routingPath")), String.Empty, reader.Item("pathDestiny_routingPath"))
                End If
                reader.Close()
            Catch ex As Exception
                log2.WriteLog(ex.Message)
            End Try
        End Using

        Return path
    End Function

    'Comprueba si el archivo esta abierto si es así da un delay de 5 seg
    Private Shared Function IsFileLocked(ByVal file As FileInfo) As Boolean
        Dim stream As FileStream = Nothing
        Try
            stream = file.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None)
        Catch __unusedIOException1__ As IOException
            Return True
        Finally
            If stream IsNot Nothing Then stream.Close()
        End Try

        Return False
    End Function
#Region "Old code"
    'Private Shared Sub Initialize_old(path As String)
    '    Dim rootDir As System.IO.DirectoryInfo = New IO.DirectoryInfo(path)

    '    If rootDir.Exists Then
    '        Console.WriteLine(rootDir.FullName)
    '        WalkDirectoryTree(rootDir)
    '    Else
    '        Console.WriteLine("Invalid root directory or unaccesible")
    '    End If

    '    Console.WriteLine("Files with restricted access:")
    '    For Each s As String In log
    '        Console.WriteLine(s)
    '    Next

    '    Console.WriteLine("Press any key")
    '    Console.ReadKey()
    'End Sub
    'Private Shared Sub Initialize(path As String)
    '    Dim rootDir As System.IO.DirectoryInfo = New IO.DirectoryInfo(path)

    '    'Dim filterFile As String = "*.*"
    '    'Dim rootDirListener As FileSystemWatcher = New FileSystemWatcher(path, filterFile)

    '    'rootDirListener.IncludeSubdirectories = True

    '    If rootDir.Exists Then
    '        Console.WriteLine(rootDir.FullName)
    '        WalkDirectoryTree(rootDir)
    '    Else
    '        Console.WriteLine("Invalid root directory or unaccesible")
    '    End If

    '    Console.WriteLine("Files with restricted access:")
    '    For Each s As String In log
    '        Console.WriteLine(s)
    '    Next

    '    Console.WriteLine("Press any key")
    '    Console.ReadKey()
    'End Sub
    'Private Shared Sub WalkDirectoryTree(ByVal root As IO.DirectoryInfo)
    '    Dim files As IO.FileInfo() = Nothing
    '    Dim subDirs As IO.DirectoryInfo() = Nothing
    '    Try
    '        files = root.GetFiles("*.*")
    '    Catch e As UnauthorizedAccessException
    '        log.Add(e.Message)
    '    Catch e As IO.DirectoryNotFoundException
    '        Console.WriteLine(e.Message)
    '    End Try

    '    If files IsNot Nothing Then
    '        For Each fi As IO.FileInfo In files
    '            Console.WriteLine(fi.FullName)
    '            'TODO: Lanzar aquí  el insert a la BD
    '            Dim cadenaConexion As String = "server=192.168.170.20;database=davidbarreiro_schema;user id=user;password=user;port=3306;"

    '            Using conn As New SqlConnection(cadenaConexion)
    '                conn.Open()
    '                Dim command As SqlCommand = conn.CreateCommand()
    '                'command.Parameters
    '                command.CommandText = "INSERT INTO [dbo].[fileprocessmetadata]
    '       ([fileName_fpMetaData]
    '       ,[date_fpMetaData]
    '       ,[path_fpMetaData]
    '       ,[error_cvmc]
    '       )
    ' VALUES
    '       ( " & fi.Name & " , " &
    '          fi.CreationTime & ", " &
    '          fi.FullName & ", " &
    '          "OK" & " )"
    '                Dim result As Integer = command.ExecuteNonQuery()
    '                If result > 0 Then
    '                    MsgBox("exito!")
    '                End If
    '                conn.Close()
    '            End Using
    '        Next

    '        subDirs = root.GetDirectories()
    '        For Each dirInfo As System.IO.DirectoryInfo In subDirs
    '            WalkDirectoryTree(dirInfo)
    '        Next
    '    End If

    'End Sub
#End Region

End Class
