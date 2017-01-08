Imports System.Runtime.InteropServices
Imports System.IO

Public Class Form1

    Private Property pageready As Boolean = False
    Public Property ScriptErrorsSuppressed As Boolean

    <Runtime.InteropServices.DllImport("wininet.dll", SetLastError:=True)> _
    Private Shared Function InternetSetOption(ByVal hInternet As IntPtr, ByVal dwOption As Integer, ByVal lpBuffer As IntPtr, ByVal lpdwBufferLength As Integer) As Boolean
    End Function

    Public Structure Struct_INTERNET_PROXY_INFO
        Public dwAccessType As Integer
        Public proxy As IntPtr
        Public proxyBypass As IntPtr
    End Structure

    Private Sub UseProxy(ByVal strProxy As String)
        Const INTERNET_OPTION_PROXY As Integer = 38
        Const INTERNET_OPEN_TYPE_PROXY As Integer = 3

        Dim struct_IPI As Struct_INTERNET_PROXY_INFO

        struct_IPI.dwAccessType = INTERNET_OPEN_TYPE_PROXY
        struct_IPI.proxy = Marshal.StringToHGlobalAnsi(strProxy)
        struct_IPI.proxyBypass = Marshal.StringToHGlobalAnsi("local")

        Dim intptrStruct As IntPtr = Marshal.AllocCoTaskMem(Marshal.SizeOf(struct_IPI))

        Marshal.StructureToPtr(struct_IPI, intptrStruct, True)

        Dim iReturn As Boolean = InternetSetOption(IntPtr.Zero, INTERNET_OPTION_PROXY, intptrStruct, System.Runtime.InteropServices.Marshal.SizeOf(struct_IPI))
    End Sub

#Region "Page Loading Functions"
    Private Sub WaitForPageLoad()
        AddHandler WebBrowser1.DocumentCompleted, New WebBrowserDocumentCompletedEventHandler(AddressOf PageWaiter)
        While Not pageready
            Application.DoEvents()
        End While
        pageready = False
    End Sub

    Private Sub PageWaiter(ByVal sender As Object, ByVal e As WebBrowserDocumentCompletedEventArgs)
        If WebBrowser1.ReadyState = WebBrowserReadyState.Complete Then
            pageready = True
            RemoveHandler WebBrowser1.DocumentCompleted, New WebBrowserDocumentCompletedEventHandler(AddressOf PageWaiter)
        End If
    End Sub

#End Region

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        WebBrowser1.AllowNavigation = True
        WebBrowser1.ScriptErrorsSuppressed = True

        Dim userList As List(Of String) = GetUsers()
        Dim proxyList As List(Of String) = GetProxies()
        Dim i As Integer = 0
        While (i < userList.Count And i < proxyList.Count)
            UseProxy(proxyList(i))
            'UseProxy("185.28.193.95:8080")
            'WebBrowser1.Navigate("https://whatismyipaddress.com/")
            'WaitForPageLoad()
            Dim words As String() = userList(i).Split(New Char() {" "c})
            Dim user As String = words(0)
            Dim pass As String = words(1)
            Login(user, pass)
            Upvote("https://www.reddit.com/r/Guildwars2/comments/5eo56s/we_asian_mmo_now/")
            Logout()
            i += 1
        End While

    End Sub

    Private Sub Login(user As String, password As String)
        Dim elements As HtmlElementCollection
        Dim el As HtmlElement
        WebBrowser1.Navigate("https://www.reddit.com/login")
        WaitForPageLoad()
        el = WebBrowser1.Document.GetElementById("user_login")
        el.SetAttribute("value", user)
        el = WebBrowser1.Document.GetElementById("passwd_login")
        el.SetAttribute("value", password)
        elements = WebBrowser1.Document.GetElementsByTagName("button")
        For Each elem As HtmlElement In elements
            If (elem.GetAttribute("className") = "c-btn c-btn-primary c-pull-right") Then
                elem.InvokeMember("click")
            End If
        Next
        WaitForPageLoad()
    End Sub

    Private Sub Logout()
        WebBrowser1.Navigate("https://www.reddit.com/")
        WaitForPageLoad()
        WebBrowser1.Document.InvokeScript("eval", {"$(""form[action='https://www.reddit.com/logout']"").submit()"})
        WaitForPageLoad()
    End Sub

    Private Sub Upvote(url As String)
        WebBrowser1.Navigate(url)
        WaitForPageLoad()
        Dim elements As HtmlElementCollection
        elements = WebBrowser1.Document.GetElementsByTagName("div")
        For Each elem As HtmlElement In elements
            If (elem.GetAttribute("className") = "arrow up login-required access-required") Then
                elem.InvokeMember("click")
                Return
            End If
        Next
    End Sub

    Private Function GetProxies() As List(Of String)
        Dim list As New List(Of String)
        Using r As StreamReader = New StreamReader("C:\Python27\proxy.txt")
            Dim line As String
            line = r.ReadLine
            Do While (Not line Is Nothing)
                list.Add(line)
                Console.WriteLine(line)
                line = r.ReadLine
            Loop
        End Using
        Return list
    End Function

    Private Function GetUsers() As List(Of String)
        Dim list As New List(Of String)
        Using r As StreamReader = New StreamReader("C:\Python27\pass.txt")
            Dim line As String
            line = r.ReadLine
            Do While (Not line Is Nothing)
                list.Add(line)
                Console.WriteLine(line)
                line = r.ReadLine
            Loop
        End Using
        Return list
    End Function
End Class
