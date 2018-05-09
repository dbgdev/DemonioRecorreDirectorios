Imports System.Configuration
Imports System.IO
Imports System.Threading
Imports System.Timers
Imports System.Web

Public Class Service1
    Private recorreDirectorio As RecursiveFileSearch
    Private eventId As Integer = 1
    Private _pathOriginService As String
    Private _pathDestinyservice As String
    Private escribeLog As ObjectLog = New ObjectLog("traza servicio", "C:\test")

    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Agregue el código aquí para iniciar el servicio. Este método debería poner
        ' en movimiento los elementos para que el servicio pueda funcionar.
        ' Set up a timer to trigger every minute.
        escribeLog.WriteLog("On start")

        _pathOriginService = HttpUtility.UrlEncode(ConfigurationManager.AppSettings.Get("pathOrigen")) '"C:\root"
        escribeLog.WriteLog("origenvariable")
        _pathDestinyservice = HttpUtility.UrlEncode(ConfigurationManager.AppSettings.Get("pathDestino")) '"C:\movidos" ' 
        escribeLog.WriteLog("destinovariable")

        escribeLog.WriteLog("On argumentos 3--> " + _pathOriginService + " , " + _pathDestinyservice)

        Dim timer As Timers.Timer = New Timers.Timer()
        timer.Interval = 60001 ' 60 seconds  
        escribeLog.WriteLog("before timer")
        AddHandler timer.Elapsed, AddressOf Me.OnTimer
        timer.Start()
        escribeLog.WriteLog("after timer")

    End Sub

    Protected Overrides Sub OnStop()
        ' Agregue el código aquí para realizar cualquier anulación necesaria para detener el servicio.

    End Sub
    Private Sub OnTimer(sender As Object, e As ElapsedEventArgs)
        ' TODO: makes things here  
        recorreDirectorio = New RecursiveFileSearch(_pathOriginService, _pathDestinyservice)  'New RecursiveFileSearch() '

        eventId = eventId + 1
    End Sub
End Class
