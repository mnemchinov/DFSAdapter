Imports System.Runtime.InteropServices
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Diagnostics
Imports System.Threading
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
'Сгенерируйте уникальный идентификатор компоненты (меню Tools - Create GUID)  
'Уникальный идентификатор в пределах Вселенной за время ее существования. :-) 
'Укажите ProgID компоненты (по этому имени ее будет находить 1С).
'Пример регистрации компоненты в системном реестре, чтобы ее смогла найти 1С: 
'regasm.exe DFSAdapter.dll /codebase 

<ComVisible(True), Guid("ac6b9bb3-8dbb-4f62-aa41-79cb6a1e4a1a"), ProgId("AddIn.DFSAdapter")> Public Class DFSAdapter

    Implements IInitDone
    Implements ILanguageExtender
    Const c_AddinName As String = "DFSAdapter"

#Region "IInitDone implementation"
    '/////////////////////////////////////////////////////////////////////////////////
    Public Sub New() ' Обязательно для COM инициализации
        'Вызывается при начале работы внешней компоненты
    End Sub

    '/////////////////////////////////////////////////////////////////////////////////
    Private Sub Init(<MarshalAs(UnmanagedType.IDispatch)> ByVal pConnection As Object) Implements IInitDone.Init
        Trace.AutoFlush = True
        Trace.Listeners.Add(New TextWriterTraceListener("DFSAdapter.log"))
        Trace.Listeners.Add(New EventLogTraceListener("DFSAdapter"))
        Trace.TraceInformation("Начало работы " & c_AddinName & " (C) ООО ППВТИ Михаил Немчинов (mnemchinov@mail.ru).")
        V7Data.V7Object = pConnection
    End Sub

    '/////////////////////////////////////////////////////////////////////////////////
    Private Sub Done() Implements IInitDone.Done
        'Вызывается при завершении работы внешней компоненты
        V7Data.V7Object = Nothing
        GC.Collect()
        GC.WaitForPendingFinalizers()
        Trace.TraceInformation("Завершение работы " & c_AddinName & " (C) ООО ППВТИ Михаил Немчинов (mnemchinov@mail.ru).")
    End Sub

    '/////////////////////////////////////////////////////////////////////////////////
    Private Sub GetInfo(ByRef pInfo() As Object) Implements IInitDone.GetInfo
        pInfo.SetValue("2000", 0)
    End Sub

    '/////////////////////////////////////////////////////////////////////////////////
    Sub RegisterExtensionAs(ByRef bstrExtensionName As String) Implements ILanguageExtender.RegisterExtensionAs
        bstrExtensionName = c_AddinName
    End Sub

#End Region

#Region "Переменные"

    Dim sServer As String 'Сервер
    Dim sPort As String  'Порт
    Dim sRepositoryName As String  'Имя репозитория
    Dim sUserName As String  'Имя пользователя
    Dim sPassword As String  'Пароль
    Dim clientOrchestrated As Boolean = False 'при client-версии работа UCF выполняется сборкой, которую приносит UcfInstaller; если DFS-версия, то всю эту работу возмет на себя DFS
    Dim IsInitialized As Boolean = False

    Dim service As IObjectService

#End Region

#Region "Внутренние функции"

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
                Dim connection As New UcfConnection(New Uri(strURL & "/core"))
                Dim activityInfo As New ActivityInfo(True)
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
            Return (True)
        Catch ex As Exception
            Raise1CException(ex.Message, "initializationDFS()", ex.StackTrace)
            Return False
        End Try
    End Function

    Private Function DownloadFileFromRepository(ByVal strFileID As String) As String
        Try
            Dim docIdentity As New ObjectIdentity(New ObjectId(strFileID), sRepositoryName)
            Dim cp As New ContentProfile(FormatFilter.ANY, "", PageFilter.ANY, 0, PageModifierFilter.ANY, "")
            Dim options As New OperationOptions(Nothing, cp)
            Dim result As DataPackage = service.[Get](New ObjectIdentitySet(docIdentity), options)
            Dim properties As PropertySet = result.DataObjects(0).Properties
            Dim content As Content = result.DataObjects(0).Contents(0)
            Return content.GetAsFile().FullName 'Возвращаем путь к файлу
        Catch ex As Exception
            Raise1CException(ex.Message, "DownloadFileFromRepository()", ex.StackTrace)
            Return ""
        End Try
    End Function

    Private Function UploadFileToRepository(ByVal strFilePath As String) As String
        Try
            If Dir$(strFilePath) <> vbNullString Then
                Dim objectName As String = New FileInfo(strFilePath).Name
                Dim dataObject As New DataObject(New ObjectIdentity(sRepositoryName), "dm_document")
                dataObject.Properties.[Set]("object_name", objectName)
                dataObject.Contents.Add(New FileContent(strFilePath, "emcmf"))
                Dim result As DataPackage = service.Create(New DataPackage(dataObject), Nothing)
                Dim properties As PropertySet = result.DataObjects(0).Properties
                Return properties.[Get]("r_object_id").ToString() 'Возвращаем ID файла
            Else
                Raise1CException("Файл " & strFilePath & " не существует!", "UploadFileToRepository()")
                Return ""
            End If
        Catch ex As Exception
            Raise1CException(ex.Message, "UploadFileToRepository()", ex.StackTrace)
            Return ""
        End Try
    End Function

#End Region

#Region "Свойства"
    '/////////////////////////////////////////////////////////////////////////////////
    Enum Props
        'Числовые идентификаторы свойств внешней компоненты
        propServer = 0 'Сервер
        propPort = 1  'Порт
        propRepositoryName = 2  'Имя репозитория
        propUserName = 3 'Имя пользователя
        propPassword = 4  'Пароль
        propClientOrchestrated = 5  'при client-версии работа UCF выполняется сборкой, которую приносит UcfInstaller; если DFS-версия, то всю эту работу возмет на себя DFS
        propIsInitialized = 6

        LastProp = 7
    End Enum

    '/////////////////////////////////////////////////////////////////////////////////
    Sub GetNProps(ByRef plProps As Integer) Implements ILanguageExtender.GetNProps
        'Здесь 1С получает количество доступных из ВК свойств
        plProps = Props.LastProp
    End Sub

    '/////////////////////////////////////////////////////////////////////////////////
    Sub FindProp(ByVal bstrPropName As String, ByRef plPropNum As Integer) Implements ILanguageExtender.FindProp
        'Здесь 1С ищет числовой идентификатор свойства по его текстовому имени
        Select Case bstrPropName
            Case "Server", "Сервер"
                plPropNum = Props.propServer
            Case "Port", "Порт"
                plPropNum = Props.propPort
            Case "RepositoryName", "ИмяРепозитория"
                plPropNum = Props.propRepositoryName
            Case "UserName", "ИмяПользователя"
                plPropNum = Props.propUserName
            Case "Password", "Пароль"
                plPropNum = Props.propPassword
            Case "ClientOrchestrated", "ОбработкаНаКлиенте"
                plPropNum = Props.propClientOrchestrated
            Case "IsInitialized", "СервисИнициализирован"
                plPropNum = Props.propIsInitialized
            Case Else
                plPropNum = -1
        End Select
    End Sub

    '/////////////////////////////////////////////////////////////////////////////////
    Sub GetPropName(ByVal lPropNum As Integer, ByVal lPropAlias As Integer, ByRef pbstrPropName As String) Implements ILanguageExtender.GetPropName
        'Здесь 1С (теоретически) узнает имя свойства по его идентификатору. lPropAlias - номер псевдонима
        pbstrPropName = ""
    End Sub

    '/////////////////////////////////////////////////////////////////////////////////
    'Функция генерирует исключение в 1С
    Public Sub Raise1CException(ByVal strMessage As String, Optional ByVal strOwner As String = "", Optional ByVal strDetails As String = "")
        Dim ei As EXCEPINFO
        Dim StrErrMessage As String
        ei.wCode = 1006 'Вид пиктограммы
        '1000 - нет значка
        '1001 - обычный значок
        '1002 - красный значок !
        '1003 - красный значок !!
        '1004 - красный значок !!!
        '1005 - зеленый значок i
        '1006 - красный значок err
        '1007 - Окно предупреждения "Внимание"
        '1008 - Окно предупреждения "Информация"
        '1009 - Окно предупреждения "Ошибка"

        StrErrMessage = "ERR:" & c_AddinName & ":" & strOwner.Trim & ":" & strMessage
        Trace.TraceError(StrErrMessage & vbCrLf & strDetails)
        ei.scode = 1 'Генерируем ошибку времени исполнения
        ei.bstrDescription = StrErrMessage 'Сообщение
        ei.bstrSource = c_AddinName

        V7Data.ErrorLog.AddError(c_AddinName, ei)

    End Sub

    '/////////////////////////////////////////////////////////////////////////////////
    Sub GetPropVal(ByVal lPropNum As Integer, ByRef pvarPropVal As Object) Implements ILanguageExtender.GetPropVal
        'Здесь 1С узнает значения свойств 
        Try
            pvarPropVal = Nothing
            Select Case lPropNum
                Case Props.propServer
                    pvarPropVal = sServer
                Case Props.propPort
                    pvarPropVal = sPort
                Case Props.propRepositoryName
                    pvarPropVal = sRepositoryName
                Case Props.propUserName
                    pvarPropVal = sUserName
                Case Props.propPassword
                    pvarPropVal = sPassword
                Case Props.propClientOrchestrated
                    pvarPropVal = clientOrchestrated
                Case Props.propIsInitialized
                    pvarPropVal = IsInitialized
            End Select
        Catch ex As Exception 'Обработчик исключительных ситуаций (ошибок)
            Raise1CException(ex.Message, "GetPropVal()", ex.ToString)
        End Try
    End Sub

    '/////////////////////////////////////////////////////////////////////////////////
    Sub SetPropVal(ByVal lPropNum As Integer, ByRef varPropVal As Object) Implements ILanguageExtender.SetPropVal
        'Здесь 1С изменяет значения свойств 
        Select Case lPropNum
            Case Props.propServer
                sServer = CStr(varPropVal)
            Case Props.propPort
                sPort = CStr(varPropVal)
            Case Props.propRepositoryName
                sRepositoryName = CStr(varPropVal)
            Case Props.propUserName
                sUserName = CStr(varPropVal)
            Case Props.propPassword
                sPassword = CStr(varPropVal)
            Case Props.propClientOrchestrated
                clientOrchestrated = varPropVal
            Case Props.propIsInitialized
                IsInitialized = varPropVal
        End Select
    End Sub

    '/////////////////////////////////////////////////////////////////////////////////
    Sub IsPropReadable(ByVal lPropNum As Integer, ByRef pboolPropRead As Boolean) Implements ILanguageExtender.IsPropReadable
        'Здесь 1С узнает, какие свойства доступны для чтения

        pboolPropRead = True ' Все свойства доступны для чтения
    End Sub

    '/////////////////////////////////////////////////////////////////////////////////
    Sub IsPropWritable(ByVal lPropNum As Integer, ByRef pboolPropWrite As Boolean) Implements ILanguageExtender.IsPropWritable
        'Здесь 1С узнает, какие свойства доступны для записи
        Select Case lPropNum
            Case Props.propIsInitialized
                pboolPropWrite = False
            Case Else
                pboolPropWrite = True ' Все свойства доступны для записи
        End Select
    End Sub

#End Region

#Region "Методы"

    '/////////////////////////////////////////////////////////////////////////////////
    Enum Methods
        'Числовые идентификаторы методов (процедур или функций) внешней компоненты
        methDownloadFileFromRepository = 0 'Получить файл из репозитория
        methUploadFileToRepository = 1 'Отправить файл в репозиторий
        methInitializationDFS = 2 'Инициализировать сервис DFS

        LastMethod = 3
    End Enum

    '/////////////////////////////////////////////////////////////////////////////////
    Sub GetNMethods(ByRef plMethods As Integer) Implements ILanguageExtender.GetNMethods
        plMethods = Methods.LastMethod
    End Sub

    '/////////////////////////////////////////////////////////////////////////////////
    Sub FindMethod(ByVal bstrMethodName As String, ByRef plMethodNum As Integer) Implements ILanguageExtender.FindMethod
        'Здесь 1С получает числовой идентификатор метода (процедуры или функции) по имени (названию) процедуры или функции

        plMethodNum = -1
        Select Case bstrMethodName
            Case "DownloadFileFromRepository", "СкачатьФайлИзРепозитория"
                plMethodNum = Methods.methDownloadFileFromRepository
            Case "UploadFileToRepository", "ЗагрузитьФайлВРепозиторий"
                plMethodNum = Methods.methUploadFileToRepository
            Case "InitializationDFS", "Инициализация"
                plMethodNum = Methods.methInitializationDFS
        End Select
    End Sub

    '/////////////////////////////////////////////////////////////////////////////////
    Sub GetMethodName(ByVal lMethodNum As Integer, ByVal lMethodAlias As Integer, ByRef pbstrMethodName As String) Implements ILanguageExtender.GetMethodName
        'Здесь 1С (теоретически) получает имя метода по его идентификатору. lMethodAlias - номер синонима.
        pbstrMethodName = ""
    End Sub

    '/////////////////////////////////////////////////////////////////////////////////
    Sub GetNParams(ByVal lMethodNum As Integer, ByRef plParams As Integer) Implements ILanguageExtender.GetNParams
        'Здесь 1С получает количество параметров у метода (процедуры или функции)

        Select Case lMethodNum
            Case Methods.methDownloadFileFromRepository
                plParams = 1
            Case Methods.methUploadFileToRepository
                plParams = 1
            Case Methods.methInitializationDFS
                plParams = 0
        End Select
    End Sub

    '/////////////////////////////////////////////////////////////////////////////////
    Sub GetParamDefValue(ByVal lMethodNum As Integer, ByVal lParamNum As Integer, ByRef pvarParamDefValue As Object) Implements ILanguageExtender.GetParamDefValue
        'Здесь 1С получает значения параметров процедуры или функции по умолчанию

        pvarParamDefValue = Nothing 'Нет значений по умолчанию
    End Sub

    '/////////////////////////////////////////////////////////////////////////////////
    Sub HasRetVal(ByVal lMethodNum As Integer, ByRef pboolRetValue As Boolean) Implements ILanguageExtender.HasRetVal
        'Здесь 1С узнает, возвращает ли метод значение (т.е. является процедурой или функцией)

        pboolRetValue = True  'Все методы у нас будут функциями (т.е. будут возвращать значение). 
    End Sub

    '/////////////////////////////////////////////////////////////////////////////////
    Sub CallAsProc(ByVal lMethodNum As Integer, ByRef paParams As System.Array) Implements ILanguageExtender.CallAsProc
        'Здесь внешняя компонента выполняет код процедур. А процедур у нас нет.
    End Sub

    '/////////////////////////////////////////////////////////////////////////////////
    Sub CallAsFunc(ByVal lMethodNum As Integer, ByRef pvarRetValue As Object, ByRef paParams As System.Array) Implements ILanguageExtender.CallAsFunc

        'Здесь внешняя компонента выполняет код функций.

        Try
            pvarRetValue = 0 'Возвращаемое значение метода для 1С
            If sServer.Length = 0 Or sPort.Length = 0 Or sRepositoryName.Length = 0 Or sUserName.Length = 0 Or sPassword.Length = 0 Then
                Raise1CException("Не задан один или несколько параметров!!!", "CallAsFunc()", _
                                 "Server = " & sServer & vbCrLf & _
                                 "Port = " & sPort & vbCrLf & _
                                 "RepositoryName = " & sRepositoryName & vbCrLf & _
                                 "UserName = " & sUserName & vbCrLf & _
                                 "Password = " & sPassword)
            Else
                Select Case lMethodNum 'Порядковый номер метода
                    '//////////////////////////////////////////////////////////
                    Case Methods.methDownloadFileFromRepository
                        If Not IsInitialized Then
                            Raise1CException("Сервис не инициализирован!!!", "")
                        Else
                            If paParams.Length = 0 Then
                                pvarRetValue = False
                                Raise1CException("Не задан ID файла!", "DownloadFileFromRepository()")
                            Else
                                pvarRetValue = DownloadFileFromRepository(paParams.GetValue(0).ToString)
                            End If
                        End If
                    Case Methods.methUploadFileToRepository
                        If Not IsInitialized Then
                            Raise1CException("Сервис не инициализирован!!!", "")
                        Else
                            If paParams.Length = 0 Then
                                pvarRetValue = False
                                Raise1CException("Не указан файл!", "UploadFileToRepository()")
                            Else
                                pvarRetValue = UploadFileToRepository(paParams.GetValue(0).ToString)
                            End If
                        End If
                    Case Methods.methInitializationDFS
                        pvarRetValue = initializationDFS()
                        IsInitialized = pvarRetValue
                End Select
            End If
        Catch ex As Exception 'Обрабатываем исключение (ошибку)
            Raise1CException(ex.Message, "CallAsFunc()", ex.ToString)
        End Try
    End Sub
#End Region
End Class
