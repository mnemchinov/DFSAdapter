Imports System.Runtime.InteropServices

<ComVisible(True), Guid("CF0C6DFE-F97C-4DB6-8298-0F86DF4507B6")> _
Public Class ClientPage
    Implements IPropertyPage

    Public Sub New()
        m_Page = New PropertyPageControl()
        m_Dirty = False
    End Sub

    Public Sub Activate(ByVal hWndParent As IntPtr, ByRef pRect As tagRECT, ByVal bModal As Boolean) Implements IPropertyPage.Activate
        If IntPtr.op_Equality(WinAPI.SetParent(m_Page.Handle, hWndParent), IntPtr.Zero) Then
            Throw New COMException("Cannot set parent frame", COMError.E_UNEXPECTED)
        Else
            Try
                m_Page.Left = pRect.left
                m_Page.Top = pRect.top
                m_Page.Width = pRect.right - pRect.left
                m_Page.Height = pRect.bottom - pRect.top

                m_Page.txtServer.Text = DFSAdapter.sServer
                m_Page.txtPort.Text = DFSAdapter.sPort
                m_Page.txtRepozitoryName.Text = DFSAdapter.sRepositoryName
                m_Page.txtUserName.Text = DFSAdapter.sUserName
                m_Page.txtPassword.Text = DFSAdapter.sPassword

            Catch
                Throw New COMException("Invalid page's size", COMError.E_POINTER)
            End Try
            m_Page.Show()
        End If
    End Sub

    Public Sub Apply() Implements IPropertyPage.Apply
        DFSAdapter.sServer = m_Page.txtServer.Text
        DFSAdapter.sPort = m_Page.txtPort.Text
        DFSAdapter.sRepositoryName = m_Page.txtRepozitoryName.Text
        DFSAdapter.sUserName = m_Page.txtUserName.Text
        DFSAdapter.sPassword = m_Page.txtPassword.Text
        DFSAdapter.WriteConfigSettings()
        m_Dirty = False
    End Sub

    Public Sub Deactivate() Implements IPropertyPage.Deactivate
        m_Page.Visible = False
        m_Page.Dispose()
    End Sub

    Public Sub GetPageInfo(ByRef pPageInfo As propPageInfo) Implements IPropertyPage.GetPageInfo
        Try
            pPageInfo.cb = Marshal.SizeOf(pPageInfo)
            pPageInfo.pszTitle = "DFS Adaper"
            pPageInfo.cx = 300
            pPageInfo.cy = 250
        Catch
            Throw New COMException("Invalid PageInfo structure", COMError.E_POINTER)
        End Try
    End Sub

    Public Sub Help(ByVal pszHelpDir As String) Implements IPropertyPage.Help

    End Sub

    Public Sub IsPageDirty() Implements IPropertyPage.IsPageDirty
        If Not m_Dirty Then Throw New COMException("", COMError.S_FALSE)
    End Sub

    Public Sub Move(ByRef pRect As tagRECT) Implements IPropertyPage.Move
        Try
            m_Page.Left = pRect.left
            m_Page.Top = pRect.top
            m_Page.Width = pRect.right - pRect.left
            m_Page.Height = pRect.bottom - pRect.top
        Catch
            Throw New COMException("Invalid pointer", COMError.E_POINTER)
        End Try
    End Sub

    Public Sub SetObjects(ByVal cObjects As Integer, ByRef ppUnk As Object) Implements IPropertyPage.SetObjects

    End Sub

    Public Sub SetPageSite(ByVal pPageSite As IPropertyPageSite) Implements IPropertyPage.SetPageSite
        If pPageSite Is Nothing Then
            m_PageSite = Nothing
        Else
            If m_PageSite Is Nothing Then
                m_PageSite = pPageSite
            Else
                Throw New COMException("Two nonnullable pointers on IPageSite interface are not allowed", COMError.E_UNEXPECTED)
            End If
        End If
    End Sub

    Public Sub Show(ByVal nCmdShow As Integer) Implements IPropertyPage.Show
        Select Case nCmdShow
            Case WinAPI.SW_HIDE
                m_Page.Hide()
            Case Else
                m_Page.Show()
        End Select
    End Sub

    Public Sub TranslateAccelerator(ByRef pMsg As tagMSG) Implements IPropertyPage.TranslateAccelerator

    End Sub

    Private WithEvents m_Page As PropertyPageControl
    Private m_Dirty As Boolean
    Private m_PageSite As IPropertyPageSite

    Private Sub m_Page_DirtyChanged() Handles m_Page.DirtyChanged
        m_Dirty = True
        m_PageSite.OnStatusChange(1)
    End Sub
End Class
