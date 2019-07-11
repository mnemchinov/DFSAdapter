Imports System.Runtime.InteropServices
Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Diagnostics
Imports System.Threading
Imports System.ComponentModel
Imports System.IO
Imports System.Net.Sockets
Imports System.Reflection

'Сгенерируйте уникальный идентификатор компоненты (меню Tools - Create GUID)  
'Уникальный идентификатор в пределах Вселенной за время ее существования. :-) 
'Укажите ProgID компоненты (по этому имени ее будет находить 1С).
'Пример регистрации компоненты в системном реестре, чтобы ее смогла найти 1С: 
'regasm.exe DFSAdapter.dll /codebase 

<ComVisible(True), Guid("5286E467-5B87-402C-9B3E-5D75F78BF8C9"), ProgId("AddIn.DFSAdapter")> Public Class DFSAdapter

    Implements IInitDone
    Implements ILanguageExtender
    Implements ISpecifyPropertyPages

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

        CreateRes("DFSAdapterServer_exe", My.Resources.DFSAdapterServer_exe)
        CreateRes("Emc_Documentum_FS_DataModel_Core_dll", My.Resources.Emc_Documentum_FS_DataModel_Core_dll)
        CreateRes("Emc_Documentum_FS_DataModel_Shared_dll", My.Resources.Emc_Documentum_FS_DataModel_Shared_dll)
        CreateRes("Emc_Documentum_FS_Runtime_dll", My.Resources.Emc_Documentum_FS_Runtime_dll)
        CreateRes("Emc_Documentum_FS_Runtime_Ucf_dll", My.Resources.Emc_Documentum_FS_Runtime_Ucf_dll)
        CreateRes("Emc_Documentum_FS_Services_Core_dll", My.Resources.Emc_Documentum_FS_Services_Core_dll)

        ReadConfigSettings()

        V7Data.V7Object = pConnection
    End Sub

    '/////////////////////////////////////////////////////////////////////////////////
    Private Sub Done() Implements IInitDone.Done
        'Вызывается при завершении работы внешней компоненты
        WriteConfigSettings()
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

#Region "ISpecifyPropertyPage implementation"
    Public Sub GetPages(ByRef pPages As CAUUID) Implements ISpecifyPropertyPages.GetPages
        pPages.cElems = 1
        pPages.pElems = Marshal.AllocCoTaskMem(Marshal.SizeOf(GetType(Guid)))
        ' GUID совпадает с ClientPage!
        Marshal.StructureToPtr(New Guid("CF0C6DFE-F97C-4DB6-8298-0F86DF4507B6"), pPages.pElems, False)
    End Sub
#End Region

#Region "Переменные"

    Public Shared sServer As String 'Сервер
    Public Shared sPort As String  'Порт
    Public Shared sRepositoryName As String  'Имя репозитория
    Public Shared sUserName As String  'Имя пользователя
    Public Shared sPassword As String  'Пароль
    Public Shared clientOrchestrated As Boolean = False 'при client-версии работа UCF выполняется сборкой, которую приносит UcfInstaller; если DFS-версия, то всю эту работу возмет на себя DFS

    Private Shared strPortForDFSAdapterServer As Integer

#End Region

#Region "Внутренние функции"

    '/////////////////////////////////////////////////////////////////////////////////
    'Функция генерирует исключение в 1С
    Public Sub Raise1CException(ByVal strMessage As String, Optional ByVal strOwner As String = "", Optional ByVal strDetails As String = "")
        Dim ei As ExcepInfo
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

    Shared Function GetFreeNetPort(ByVal StartPort As Integer) As Integer
        If StartPort = 10000 Then
            Return 0
        End If
        If StartPort = 0 Then
            StartPort = 1
        End If
        Dim hostadd As System.Net.IPAddress = System.Net.Dns.GetHostEntry("127.0.0.1").AddressList(0)
        Dim EPhost As New System.Net.IPEndPoint(hostadd, StartPort)
        Dim s As New System.Net.Sockets.Socket(System.Net.Sockets.AddressFamily.InterNetwork, System.Net.Sockets.SocketType.Stream, System.Net.Sockets.ProtocolType.Tcp)
        Try
            s.Connect(EPhost)
        Catch
        End Try
        If Not s.Connected Then
            Return StartPort
        Else
            GetFreeNetPort(StartPort + 1)
        End If
        Return 0
    End Function

    Public Shared Sub ReadConfigSettings()
        sServer = ConfigSettings.ReadSetting("Server")
        sPort = ConfigSettings.ReadSetting("Port")
        sRepositoryName = ConfigSettings.ReadSetting("RepositoryName")
        sUserName = ConfigSettings.ReadSetting("UserName")
        sPassword = clsCryptData.DecodingData(ConfigSettings.ReadSetting("Password"))
        strPortForDFSAdapterServer = CInt(ConfigSettings.ReadSetting("PortForDFSAdapterServer"))
    End Sub

    Public Shared Sub WriteConfigSettings()
        ConfigSettings.WriteSetting("Server", sServer)
        ConfigSettings.WriteSetting("Port", sPort)
        ConfigSettings.WriteSetting("RepositoryName", sRepositoryName)
        ConfigSettings.WriteSetting("UserName", sUserName)
        ConfigSettings.WriteSetting("Password", clsCryptData.EncodingData(sPassword))
        ConfigSettings.WriteSetting("PortForDFSAdapterServer", strPortForDFSAdapterServer.ToString())
    End Sub

    Private Sub CreateRes(ByVal strResName As String, ByVal strResValue() As Byte)
        Dim strFileName = Replace(strResName, "_", ".")
        Dim tempFile As String = New FileInfo(Assembly.GetExecutingAssembly().Location).DirectoryName & "\" & strFileName
        If Dir$(tempFile) = vbNullString Then
            Using fs As New FileStream(tempFile, FileMode.Create)
                fs.Write(strResValue, 0, strResValue.Length)
                fs.Close()
            End Using
        End If
    End Sub

    Private Function ParamToXML(ByVal strAction As String, ByVal FilePath As String, ByVal FileID As String) As String

        Dim strXmlFile As String = Path.GetTempPath() & "dfsprm.xml"
        Dim Writer As New Xml.XmlTextWriter(strXmlFile, System.Text.Encoding.UTF8)

        ' Записываем объявление версии XML
        Writer.WriteStartDocument(True)
        ' Указываем, что XML-документ должен быть отформатирован
        Writer.Formatting = Xml.Formatting.Indented
        ' Задаем 2 пробела для выделения вложенных данных
        Writer.Indentation = 2
        ' Записываем открывающий тег
        Writer.WriteStartElement("Parametrs")
        Writer.WriteAttributeString("Server", sServer)
        Writer.WriteAttributeString("Port", sPort)
        Writer.WriteAttributeString("UserName", sUserName)
        Writer.WriteAttributeString("Password", clsCryptData.EncodingData(sPassword))
        Writer.WriteAttributeString("RepositoryName", sRepositoryName)
        Writer.WriteAttributeString("FilePath", FilePath)
        Writer.WriteAttributeString("FileID", FileID)
        Writer.WriteAttributeString("Action", strAction)
        'закрываем элемент 
        Writer.WriteEndElement()
        'заносим данные в myMemoryStream 
        Writer.Flush()
        Writer.Close()
        ParamToXML = strXmlFile
    End Function

    Private Function RunDFSAdapterServer(ByVal strFileID As String, ByVal strFilePath As String, Optional ByVal strAction As String = "get") As String
        Try
            strPortForDFSAdapterServer = GetFreeNetPort(strPortForDFSAdapterServer)
            Dim tempFile As String = New FileInfo(Assembly.GetExecutingAssembly().Location).DirectoryName & "\DFSAdapterServer.exe"
            Dim DFSAdapterServer As Process = Process.Start(tempFile, strPortForDFSAdapterServer.ToString())
            Dim tcpClient As New System.Net.Sockets.TcpClient()
            tcpClient.Connect("127.0.0.1", strPortForDFSAdapterServer)
            Dim strPrmFilePath = ParamToXML(strAction, strFilePath, strFileID)
            Dim networkStream As NetworkStream = tcpClient.GetStream()
            If networkStream.CanWrite And networkStream.CanRead Then
                Dim sendBytes As [Byte]() = Encoding.ASCII.GetBytes(strPrmFilePath)
                networkStream.Write(sendBytes, 0, sendBytes.Length)
                Dim bytes(tcpClient.ReceiveBufferSize) As Byte
                networkStream.Read(bytes, 0, CInt(tcpClient.ReceiveBufferSize))
                ' Output the data received from the host to the console.
                Dim returndata As String = Encoding.ASCII.GetString(bytes)
                Return Replace(returndata, Chr(0), "")
            Else
                If Not networkStream.CanRead Then
                    Raise1CException("cannot not write data to this stream", "RunDFSAdapterServer()")
                Else
                    If Not networkStream.CanWrite Then
                        Raise1CException("cannot read data from this stream", "RunDFSAdapterServer()")
                    End If
                End If
                Return ""
            End If
            tcpClient.Close()
            'DFSAdapterServer.Kill() 
        Catch ex As Exception
            Raise1CException(ex.Message, "RunDFSAdapterServer()", ex.StackTrace)
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
        propPortForDFSAdapterServer = 6
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
            Case "PortForDFSAdapterServer", "ПортДляDFSAdapterServer"
                plPropNum = Props.propPortForDFSAdapterServer
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
                Case Props.propPortForDFSAdapterServer
                    pvarPropVal = strPortForDFSAdapterServer
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
            Case Props.propPortForDFSAdapterServer
                strPortForDFSAdapterServer = varPropVal
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
        pboolPropWrite = True ' Все свойства доступны для записи
    End Sub

#End Region

#Region "Методы"

    '/////////////////////////////////////////////////////////////////////////////////
    Enum Methods
        'Числовые идентификаторы методов (процедур или функций) внешней компоненты
        methDownloadFileFromRepository = 0 'Получить файл из репозитория
        methUploadFileToRepository = 1 'Отправить файл в репозиторий

        LastMethod = 2
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
                plParams = 2
            Case Methods.methUploadFileToRepository
                plParams = 2
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
                        If paParams.Length = 0 Then
                            pvarRetValue = False
                            Raise1CException("Не задан ID файла!", "DownloadFileFromRepository()")
                        Else
                            pvarRetValue = RunDFSAdapterServer(paParams.GetValue(0).ToString(), paParams.GetValue(1).ToString(), "get")
                        End If
                    Case Methods.methUploadFileToRepository
                        If paParams.Length = 0 Then
                            pvarRetValue = False
                            Raise1CException("Не указан файл!", "UploadFileToRepository()")
                        Else
                            pvarRetValue = RunDFSAdapterServer(paParams.GetValue(0).ToString(), paParams.GetValue(1).ToString(), "update")
                        End If
                End Select
            End If
        Catch ex As Exception 'Обрабатываем исключение (ошибку)
            Raise1CException(ex.Message, "CallAsFunc()", ex.ToString)
        End Try
    End Sub
#End Region
End Class
