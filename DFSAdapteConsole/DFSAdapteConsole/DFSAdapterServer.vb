Imports System.Runtime.InteropServices
Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Diagnostics
Imports System.Threading
Imports System.Net.Sockets
Imports System.ComponentModel
Imports DfsDataObject = Emc.Documentum.FS.DataModel.Core.DataObject
Imports Emc.Documentum.FS.DataModel.Core.Content
Imports Emc.Documentum.FS.DataModel.Core.Properties
Imports Emc.Documentum.FS.DataModel.Core.Context
Imports Emc.Documentum.FS.Runtime.Context
Imports Emc.Documentum.FS.DataModel.Core.Profiles
Imports Emc.Documentum.FS.Services.Core
Imports System.IO
Imports Emc.Documentum.FS.DataModel.Core
Imports Emc.Documentum.FS.Runtime.Ucf
Imports System.Xml

Module DFSAdapterServer

#Region "Переменные"

    Private sServer As String 'Сервер
    Private sPort As String  'Порт
    Private sRepositoryName As String  'Имя репозитория
    Private sUserName As String  'Имя пользователя
    Private sPassword As String  'Пароль
    Private clientOrchestrated As Boolean = False 'при client-версии работа UCF выполняется сборкой, которую приносит UcfInstaller; если DFS-версия, то всю эту работу возмет на себя DFS
    Private strFilePath As String  'Файл
    Private strFileID As String  'ID файла
    Private strAction As String  'ID файла

    Public service As IObjectService

#End Region

#Region "Внутренние функции"
    Public Function Raise1CException(ByVal strMessage As String, Optional ByVal strOwner As String = "", Optional ByVal strDetails As String = "")
        Dim StrErrMessage As String
        StrErrMessage = "ERR:DFSAdapteConsole:" & strOwner.Trim & ":" & strMessage
        Trace.TraceError(StrErrMessage & vbCrLf & strDetails)
        Raise1CException = StrErrMessage
    End Function

    Private Function initializationDFS() As Boolean
        Try
            Dim strURL As String = "http://" & sServer & ":" & sPort & "/services"
            Dim contextFactory As ContextFactory = contextFactory.Instance
            Dim context As IServiceContext = contextFactory.NewContext()
            Dim repoId As New RepositoryIdentity()
            repoId.RepositoryName = sRepositoryName
            repoId.UserName = sUserName
            repoId.Password = sPassword
            context.AddIdentity(repoId)
            Dim profile As New ContentTransferProfile()
            profile.TransferMode = ContentTransferMode.UCF
            If clientOrchestrated Then
                Dim activityInfo As New ActivityInfo(True)
                MsgBox(activityInfo.ToString())
                Dim connection As New UcfConnection(New Uri(strURL & "/core"))
                activityInfo.ActivityId = connection.GetUcfId()
                activityInfo.SessionId = connection.GetJsessionId()
                activityInfo.TargetDeploymentId = connection.GetDeploymentId()
                activityInfo.AddCookies(connection.GetCookies())
                profile.ActivityInfo = activityInfo
            End If
            context.SetProfile(profile)
            Dim propProfile As New PropertyProfile(PropertyFilterMode.ALL)
            context.SetProfile(propProfile)
            service = ServiceFactory.Instance.GetRemoteService(Of IObjectService)(context, "core", strURL)
            Return True
        Catch ex As Exception
            Raise1CException(ex.Message, "initializationDFS()", ex.StackTrace)
            Return False
        End Try
    End Function

    Private Function DownloadFileFromRepository(ByVal DocId As String, ByVal FilePath As String) As String
        Try
            Dim docIdentity As New ObjectIdentity(New ObjectId(DocId), sRepositoryName)
            Dim cp As New ContentProfile(FormatFilter.ANY, "", PageFilter.ANY, 0, PageModifierFilter.ANY, "")
            Dim options As New OperationOptions(Nothing, cp)
            Dim result As DataPackage = service.[Get](New ObjectIdentitySet(docIdentity), options)
            Dim properties As PropertySet = result.DataObjects(0).Properties
            Dim content As Content = result.DataObjects(0).Contents(0)
            Dim object_name As String = properties.[Get]("object_name").GetValueAsString
            FilePath = FilePath & IIf(Right(FilePath.Trim(), 1) = "\", object_name, "\" & object_name)
            FileCopy(content.GetAsFile().FullName, FilePath)
            'Kill(content.GetAsFile().FullName)
            Return FilePath 'Возвращаем путь к файлу
        Catch ex As Exception
            Raise1CException(ex.Message, "DownloadFileFromRepository()", ex.StackTrace)
            Return ""
        End Try
    End Function

    Private Function UploadFileToRepository(ByVal FilePath As String) As String
        Try
            If Dir$(FilePath) <> vbNullString Then
                Dim objectName As String = New FileInfo(FilePath).Name
                Dim dataObject As New DataObject(New ObjectIdentity(sRepositoryName), "dm_document")
                dataObject.Properties.[Set]("object_name", objectName)
                dataObject.Contents.Add(New FileContent(FilePath, "emcmf"))
                Dim result As DataPackage = service.Create(New DataPackage(dataObject), Nothing)
                Dim properties As PropertySet = result.DataObjects(0).Properties
                Return properties.[Get]("r_object_id").ToString() 'Возвращаем ID файла
            Else
                Raise1CException("Файл " & FilePath & " не существует!", "UploadFileToRepository()")
                Return ""
            End If
        Catch ex As Exception
            Raise1CException(ex.Message, "UploadFileToRepository()", ex.StackTrace)
            Return ""
        End Try
    End Function

    Private Function UpdateFileInRepository(ByVal FilePath As String, ByVal DocId As String) As String
        Try
            If Dir$(FilePath) <> vbNullString Then
                Dim objectName As String = New FileInfo(FilePath).Name
                Dim ObjectId As New ObjectId(DocId)
                Dim dataObject As New DataObject(New ObjectIdentity(ObjectId, sRepositoryName), "dm_document")
                dataObject.Properties.[Set]("object_name", objectName)
                dataObject.Contents.Add(New FileContent(FilePath, "emcmf"))
                Dim result As DataPackage = service.Update(New DataPackage(dataObject), Nothing)
                Dim properties As PropertySet = result.DataObjects(0).Properties
                Return properties.[Get]("r_object_id").GetValueAsString 'Возвращаем ID файла
            Else
                Raise1CException("Файл " & FilePath & " не существует!", "UploadFileToRepository()")
                Return ""
            End If
        Catch ex As Exception
            Raise1CException(ex.Message, "UploadFileToRepository()", ex.StackTrace)
            Return ""
        End Try
    End Function

#End Region

    Private Sub ReadParam(ByVal strLineArgs As String)
        Dim myXml As New XmlDataDocument()
        Dim fs As New FileStream(strLineArgs, FileMode.Open, FileAccess.Read)
        myXml.Load(fs)

        sServer = myXml.CreateNavigator.SelectSingleNode("/Parametrs/@Server").Value
        sPort = myXml.CreateNavigator.SelectSingleNode("/Parametrs/@Port").Value
        sRepositoryName = myXml.CreateNavigator.SelectSingleNode("/Parametrs/@RepositoryName").Value
        sUserName = myXml.CreateNavigator.SelectSingleNode("/Parametrs/@UserName").Value
        sPassword = clsCryptData.DecodingData(myXml.CreateNavigator.SelectSingleNode("/Parametrs/@Password").Value)
        strFilePath = myXml.CreateNavigator.SelectSingleNode("/Parametrs/@FilePath").Value
        strFileID = myXml.CreateNavigator.SelectSingleNode("/Parametrs/@FileID").Value
        strAction = myXml.CreateNavigator.SelectSingleNode("/Parametrs/@Action").Value
    End Sub

    Private Function DoActionMain()
        Try
            If sServer = "" Or sPort = "" Or sRepositoryName = "" Or sUserName = "" Or sPassword = "" Then
                DoActionMain = Raise1CException("Не задан один или несколько параметров!!!", "Main()", _
                                 "Server = " & sServer & vbCrLf & _
                                 "Port = " & sPort & vbCrLf & _
                                 "RepositoryName = " & sRepositoryName & vbCrLf & _
                                 "UserName = " & sUserName & vbCrLf & _
                                 "Password = " & sPassword)
            Else
                initializationDFS()
                Select Case strAction 'Порядковый номер метода
                    '//////////////////////////////////////////////////////////
                    Case "send"
                        If strFilePath = "" Then
                            DoActionMain = Raise1CException("Не указан файл!", "UploadFileToRepository()")
                        Else
                            DoActionMain = UploadFileToRepository(strFilePath)
                        End If
                    Case "get"
                        If strFileID = "" Or strFilePath = "" Then
                            DoActionMain = Raise1CException("Не задан путь или ID файла!", "DownloadFileFromRepository()")
                        Else
                            DoActionMain = DownloadFileFromRepository(strFileID, strFilePath)
                        End If
                    Case "update"
                        If strFileID = "" Or strFilePath = "" Then
                            DoActionMain = Raise1CException("Не задан путь или ID файла!", "UpdateFileInRepository()")
                        Else
                            DoActionMain = UpdateFileInRepository(strFilePath, strFileID)
                        End If
                    Case Else
                        DoActionMain = Raise1CException("Не верный параметр <Action>!", "DoActionMain()")
                End Select
            End If
        Catch ex As Exception 'Обрабатываем исключение (ошибку)
            DoActionMain = Raise1CException(ex.Message, "Main()", ex.ToString)
        End Try
    End Function

    Sub Main()
        Trace.AutoFlush = True
        Trace.Listeners.Add(New TextWriterTraceListener("DFSAdapterServer.log"))
        Trace.Listeners.Add(New EventLogTraceListener("DFSAdapterServer"))
        'Trace.TraceInformation("Начало работы DFSAdapterServer (C) ООО ППВТИ Михаил Немчинов (mnemchinov@mail.ru).")
        Try
            Dim OnPort As Integer = CInt(Environment.GetCommandLineArgs().GetValue(1).ToString())
            DoListen(OnPort)
        Catch ex As Exception
            Raise1CException(ex.Message, "Main()", ex.ToString)
        End Try
    End Sub

    Private Sub DoListen(ByVal OnPort As Integer)
        Try
            'Accept the pending client connection and return a TcpClient initialized for communication. 
            Dim tcpListener As New TcpListener(OnPort)
            Dim tcpClient As TcpClient

            tcpListener.Start()
            tcpClient = tcpListener.AcceptTcpClient()
            ' Get the stream
            Dim networkStream As NetworkStream = tcpClient.GetStream()
            ' Read the stream into a byte array
            Dim bytes(tcpClient.ReceiveBufferSize) As Byte
            networkStream.Read(bytes, 0, CInt(tcpClient.ReceiveBufferSize))
            ' Return the data received from the client to the console.
            Dim clientdata As String = Encoding.ASCII.GetString(bytes)
            clientdata = Replace(clientdata, Chr(0), "")
            If clientdata.Trim() = "end" Then
                tcpClient.Close()
                tcpListener.Stop()
                'Trace.TraceInformation("Завршение работы DFSAdapterServer (C) ООО ППВТИ Михаил Немчинов (mnemchinov@mail.ru).")
                Exit Sub
            End If
            ReadParam(clientdata)
            Dim responseString As String = DoActionMain()
            Dim sendBytes As [Byte]() = Encoding.ASCII.GetBytes(responseString)
            networkStream.Write(sendBytes, 0, sendBytes.Length)
            tcpClient.Close()
            tcpListener.Stop()
        Catch ex As Exception
            Raise1CException(ex.Message, "DoAction()", ex.ToString)
        End Try
    End Sub
End Module

