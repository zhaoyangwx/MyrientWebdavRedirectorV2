Imports System.Net
Imports System.Text
Imports System.Xml
Imports System.Globalization
Imports System.Threading
Imports System.Threading.Tasks
Imports HtmlAgilityPack

Module Module1
    Public StartTime As Date = Now
    Dim baseUrl As String = "https://myrient.erista.me/files"

    Public LogLock As New Object
    Public Sub Log(s As String, Optional ByVal AppendCrLf As Boolean = False)
        SyncLock LogLock
            My.Computer.FileSystem.WriteAllText(IO.Path.Combine("\\?\" & AppDomain.CurrentDomain.BaseDirectory, $"log_{StartTime.ToString("yyyyMMdd_HHmmss_fffffff")}.txt"), $"{s}{If(AppendCrLf, vbCrLf, "")}", True)
        End SyncLock
    End Sub
    ' 👉 限制 HEAD 并发
    Dim headSemaphore As New SemaphoreSlim(50)

    Sub Main()
        Dim csvpath As String = IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "import.csv")
        If IO.File.Exists(csvpath) Then
            Dim f As New IO.StreamReader(csvpath)
            Dim sb As New StringBuilder
            For i As Integer = 0 To 9999
                Dim t As String = f.ReadLine()
                If Not t.EndsWith(",,,") Then sb.AppendLine(f.ReadLine())
            Next
            System.Windows.Forms.Clipboard.SetText(sb.ToString())
            While Not f.EndOfStream
                Dim t As String = f.ReadLine
                If t.EndsWith(",,,") Then Continue While
                Dim td() As String = t.Split({","}, StringSplitOptions.None)
                If td.Length <> 7 Then Continue While
                Dim link As String = td(4)
                Dim size As Long = -1
                If Not Long.TryParse(td(5), size) Then Continue While
                GetContentLengthLimited(link, size)
                Console.WriteLine($"csvimport {size} {link}")
            End While
        End If

        Dim listener As New HttpListener()
        Dim prefix As String = "http://+:18088/"
        If IO.File.Exists(IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "listener.txt")) Then
            prefix = IO.File.ReadAllText(IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "listener.txt"))
        End If
        listener.Prefixes.Add(prefix)
        listener.Start()

        Console.WriteLine($"WebDAV running at {prefix}")

        While True
            Dim ctx = listener.GetContext()
            ThreadPool.QueueUserWorkItem(AddressOf HandleRequest, ctx)
        End While
    End Sub

    Sub HandleRequest(state As Object)
        Dim ctx = CType(state, HttpListenerContext)
        Dim req = ctx.Request
        Dim res = ctx.Response

        Dim subPath As String = req.RawUrl

        Try
            Select Case req.HttpMethod.ToUpper()

                Case "PROPFIND"
                    HandlePropfind(subPath, res)

                Case "GET"
                    res.StatusCode = 302
                    res.AddHeader("Location", baseUrl & subPath)
                    res.Close()

                Case "HEAD"
                    res.StatusCode = 302
                    res.AddHeader("Location", baseUrl & subPath)
                    res.Close()

                Case "OPTIONS"
                    res.StatusCode = 200
                    res.AddHeader("DAV", "1")
                    res.AddHeader("Allow", "OPTIONS, PROPFIND, GET, HEAD")
                    res.Close()

                Case Else
                    res.StatusCode = 403
                    res.Close()

            End Select

        Catch ex As Exception
            Console.WriteLine(ex.ToString())
            Log($"  > {subPath}{vbCrLf}")
            Log(ex.ToString, True)
            res.StatusCode = 500
            res.Close()
        End Try
    End Sub

    Sub HandlePropfind(subPath As String, res As HttpListenerResponse)

        Dim url As String = baseUrl.TrimEnd("/"c) & subPath

        Dim html As String
        Dim cachepath As String = IO.Path.Combine("\\?\" & AppDomain.CurrentDomain.BaseDirectory, "myrient.erista.me\files", EncodeInvalidChars(Uri.UnescapeDataString(subPath).TrimStart("/").TrimEnd("/")).Replace("/", "\"))
        If Not IO.Directory.Exists(cachepath) Then
            IO.Directory.CreateDirectory(cachepath)
        End If
        Dim cachefile As String = IO.Path.Combine(cachepath, "MWR_Cache.html")
        If IO.File.Exists(cachefile) Then
            html = IO.File.ReadAllText(cachefile)
        Else
            Using client As New WebClient()
                client.Headers.Add("User-Agent", "Mozilla/5.0")
                client.Encoding = Encoding.UTF8
                html = client.DownloadString(url)
            End Using
            IO.File.WriteAllText(cachefile, html)
        End If

        Dim doc As New HtmlDocument()
        doc.LoadHtml(html)
        Dim nodes = doc.DocumentNode.SelectNodes("/html/body/div[3]/table/tbody/tr")

        res.StatusCode = 207
        res.ContentType = "application/xml; charset=utf-8"

        Using writer As XmlWriter = XmlWriter.Create(res.OutputStream, New XmlWriterSettings With {
            .Encoding = Encoding.UTF8,
            .Indent = True
        })

            writer.WriteStartDocument()
            writer.WriteStartElement("D", "multistatus", "DAV:")

            If nodes IsNot Nothing Then

                Dim tasks As New List(Of Task(Of Tuple(Of String, Long)))
                Dim fincount As Integer = 0, total As Integer = 0
                For i As Integer = 1 To nodes.Count - 1

                    Dim tr = nodes(i)
                    Dim tds = tr.SelectNodes("td")
                    If tds Is Nothing OrElse tds.Count < 3 Then Continue For

                    Dim linkNode = tds(0).SelectSingleNode(".//a")
                    If linkNode Is Nothing Then Continue For

                    Dim href As String = linkNode.GetAttributeValue("href", "")
                    If String.IsNullOrEmpty(href) OrElse href = "../" OrElse href = "./" Then Continue For

                    Dim sizeText As String = tds(1).InnerText.Trim()

                    Dim fullPath As String = (subPath.TrimEnd("/"c) & "/" & href).Replace("//", "/")

                    ' 👉 判断是否需要 HEAD
                    Dim approxBytes As Long = ParseSize(sizeText)
                    'If approxBytes > 100L * 1024 * 1024 AndAlso Not href.EndsWith("/") Then
                    If Not href.EndsWith("/") Then
                        Threading.Interlocked.Increment(total)
                        ' 异步获取大小
                        tasks.Add(Task.Run(Function()
                                               Dim size = GetContentLengthLimited(baseUrl & fullPath)
                                               Threading.Interlocked.Increment(fincount)
                                               Return Tuple.Create(fullPath, size)
                                           End Function))
                    End If

                Next
                Console.WriteLine($"**STATUS**>html len = {html.Length} validnode = {nodes.Count - 3} request = {tasks.Count}")
                ' 等待所有 HEAD 完成
                Dim fin As Boolean = False, thprogended As Boolean = False
                Dim th As New Threading.Thread(Sub()
                                                   While Not fin
                                                       Console.WriteLine($"{fincount}/{total}")
                                                       If fin Then thprogended = True
                                                       Threading.Thread.Sleep(1000)
                                                   End While
                                                   If fin Then thprogended = True
                                               End Sub)
                th.Start()
                Task.WaitAll(tasks.ToArray())
                fin = True
                Task.Run(Sub()
                             While Not thprogended
                                 Threading.Thread.Sleep(1)
                             End While
                             'Console.SetCursorPosition(posL, Console.CursorTop)
                             Console.WriteLine($"{fincount}/{total}")
                         End Sub)
                ' 转为字典
                Dim sizeDict As New Dictionary(Of String, Long)
                For Each t In tasks
                    If t.Result IsNot Nothing AndAlso t.Result.Item2 >= 0 Then
                        sizeDict(t.Result.Item1) = t.Result.Item2
                    End If
                Next


                ' ===== 再遍历一遍输出 =====
                For i As Integer = 1 To nodes.Count - 1

                    Dim tr = nodes(i)
                    Dim tds = tr.SelectNodes("td")
                    If tds Is Nothing OrElse tds.Count < 3 Then Continue For

                    Dim linkNode = tds(0).SelectSingleNode(".//a")
                    If linkNode Is Nothing Then Continue For

                    Dim href As String = linkNode.GetAttributeValue("href", "")
                    If String.IsNullOrEmpty(href) OrElse href = "../" OrElse href = "./" Then Continue For

                    Dim sizeText As String = tds(1).InnerText.Trim()
                    Dim timeText As String = tds(2).InnerText.Trim()

                    Dim fullPath As String = (subPath.TrimEnd("/"c) & "/" & href).Replace("//", "/")

                    writer.WriteStartElement("D", "response", Nothing)
                    writer.WriteElementString("D", "href", Nothing, fullPath)

                    writer.WriteStartElement("D", "propstat", Nothing)
                    writer.WriteStartElement("D", "prop", Nothing)

                    If href.EndsWith("/") Then
                        writer.WriteStartElement("D", "resourcetype", Nothing)
                        writer.WriteElementString("D", "collection", Nothing, "")
                        writer.WriteEndElement()
                    Else
                        writer.WriteElementString("D", "resourcetype", Nothing, "")

                        If sizeDict.ContainsKey(fullPath) Then
                            writer.WriteElementString("D", "getcontentlength", Nothing, sizeDict(fullPath).ToString())
                        End If
                    End If

                    ' 时间
                    Dim dt As DateTime
                    If DateTime.TryParseExact(timeText,
                        "dd-MMM-yyyy HH:mm",
                        CultureInfo.InvariantCulture,
                        DateTimeStyles.None,
                        dt) Then

                        writer.WriteElementString("D", "getlastmodified", Nothing,
                            dt.ToUniversalTime().ToString("R"))
                    End If

                    writer.WriteEndElement()
                    writer.WriteElementString("D", "status", Nothing, "HTTP/1.1 200 OK")
                    writer.WriteEndElement()
                    writer.WriteEndElement()
                    'Console.WriteLine($"{If(sizeDict.ContainsKey(fullPath), sizeDict(fullPath), sizeText).ToString().PadLeft(12)}{vbTab}{timeText}{vbTab}{fullPath}")
                    'Log($"{fullPath}{vbCrLf}")
                Next
            End If

            writer.WriteEndElement()
            writer.WriteEndDocument()
        End Using

        res.Close()
    End Sub

    ' ===== 限流 HEAD =====
    Function GetContentLengthLimited(url As String, Optional ByVal len As Long = -1) As Long
        Try
            If url.StartsWith(baseUrl) Then
                Dim cachefile As String = Uri.UnescapeDataString(url.Substring(baseUrl.Length)) & ".length"
                If cachefile.StartsWith("/") Then cachefile = cachefile.Substring(1)
                cachefile = EncodeInvalidChars(cachefile)
                Try
                    cachefile = IO.Path.Combine("\\?\" & AppDomain.CurrentDomain.BaseDirectory, "myrient.erista.me\files", cachefile).Replace("/", "\")
                Catch ex As Exception
                    Log(ex.ToString, True)
                    Log(cachefile, True)
                    Console.WriteLine(ex.ToString())
                    Console.WriteLine(cachefile)
                    cachefile = IO.Path.Combine("\\?\" & AppDomain.CurrentDomain.BaseDirectory, "myrient.erista.me\files", $"MWR_{cachefile}").Replace("/", "\")
                End Try
                Dim cachepath As String = New IO.FileInfo(cachefile).DirectoryName
                If Not IO.Directory.Exists(cachepath) Then IO.Directory.CreateDirectory(cachepath)
                If IO.File.Exists(cachefile) Then
                    Return Long.Parse(IO.File.ReadAllText(cachefile))
                Else
                    If len < 0 Then len = GetContentLength(url)
                    If len >= 0 Then
                        Try
                            IO.File.WriteAllText(cachefile, len.ToString())
                            Console.WriteLine($"+ Cached: {cachefile}")
                        Catch ex As Exception
                            cachefile = cachefile
                            Console.WriteLine(ex.ToString())
                            Log(ex.ToString, True)
                        End Try
                    End If
                    Return len
                End If
            Else
                len = GetContentLength(url)
                Console.WriteLine($"Query fin:{url}")
                Return len
            End If


        Catch ex As Exception
            Console.WriteLine(ex.ToString())
            Log(ex.ToString, True)
            Return -1
        End Try
    End Function
    Public Function EncodeInvalidChars(input As String) As String
        If String.IsNullOrEmpty(input) Then Return input

        Dim invalidChars As String = ":*?""<>|" & Chr(&H10)
        Dim sb As New System.Text.StringBuilder(input.Length)

        ' 如果以 \\?\ 开头，跳过前4个字符
        Dim startIndex As Integer = If(input.StartsWith("\\?\"), 7, 0)

        ' 先保留前缀
        sb.Append(input.Substring(0, startIndex))

        ' 处理剩余部分
        For i = startIndex To input.Length - 1
            Dim c = input(i)

            If invalidChars.Contains(c) Then
                sb.Append("%" & AscW(c).ToString("X2"))
            Else
                sb.Append(c)
            End If
        Next

        Return sb.ToString()
    End Function
    Function GetContentLength(url As String) As Long
        headSemaphore.Wait()
        Try
            Dim req = CType(WebRequest.Create(url), HttpWebRequest)
            req.Method = "HEAD"
            req.AllowAutoRedirect = True
            req.Timeout = 60000
            Using resp = CType(req.GetResponse(), HttpWebResponse)
                Return resp.ContentLength
            End Using
        Catch ex As Exception
            Console.WriteLine(ex.ToString())
            Log(ex.ToString, True)
            Return -1
        Finally
            headSemaphore.Release()
        End Try
    End Function

    Function ParseSize(sizeText As String) As Long
        If sizeText = "-" Then Return 0

        Dim parts = sizeText.Split(" "c)
        If parts.Length <> 2 Then Return 0

        Dim value As Double
        If Not Double.TryParse(parts(0), NumberStyles.Any, CultureInfo.InvariantCulture, value) Then Return 0

        Select Case parts(1).ToUpper()
            Case "B", "BYTE", "BYTES" : Return CLng(value)
            Case "KB", "KIB" : Return CLng(value * 1024)
            Case "MB", "MIB" : Return CLng(value * 1024 * 1024)
            Case "GB", "GIB" : Return CLng(value * 1024 ^ 3)
            Case "TB", "TIB" : Return CLng(value * 1024 ^ 4)
        End Select

        Return 0
    End Function

End Module