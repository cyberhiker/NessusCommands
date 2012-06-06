Imports System.Net
Imports System.Security.Cryptography.X509Certificates
Imports System.Net.Security
Imports System.Text
Imports System.Xml
Imports System.Xml.Xsl
Imports System.IO
Imports System.Security
Imports System.Runtime.InteropServices

Module NessusCommands
    Public DebugFlag As Boolean = False

    Sub Main()
        Dim arrArgs = My.Application.CommandLineArgs
        Dim CarryOn As Boolean = False
        Dim strOutputXML As String = String.Empty, strReportName As String = String.Empty, strServerAddress As String = String.Empty, _
            strUsername As String = String.Empty, strPassword As String = String.Empty

        'If arrArgs.Count = 4 Or arrArgs.Count = 5 Then
        For i = 0 To arrArgs.Count - 1
            Select Case True
                Case arrArgs(i).ToLower.StartsWith("/a")
                    Try
                        Dim colonPlace As Integer = arrArgs.Item(i).IndexOf(":") + 1
                        strServerAddress = arrArgs.Item(i).Substring(colonPlace, arrArgs.Item(i).Length - colonPlace)

                        DebugWrite(strServerAddress)
                    Catch e As Exception
                        Console.WriteLine("Server address not correctly specified.")
                        ShowHelp()
                        Environment.ExitCode = -1
                        Exit Sub
                    End Try

                Case arrArgs(i).ToLower.StartsWith("/r")
                    Try
                        Dim colonPlace As Integer = arrArgs.Item(i).IndexOf(":") + 1
                        strReportName = arrArgs.Item(i).Substring(colonPlace, arrArgs.Item(i).Length - colonPlace)

                    Catch e As Exception
                        Console.WriteLine("Report name not correctly identified.")
                        ShowHelp()
                        Environment.ExitCode = -2
                        Exit Sub
                    End Try

                Case arrArgs(i).ToLower.StartsWith("/o")
                    Try
                        Dim colonPlace As Integer = arrArgs.Item(i).IndexOf(":") + 1
                        strOutputXML = arrArgs.Item(i).Substring(colonPlace, arrArgs.Item(i).Length - colonPlace)

                    Catch e As Exception
                        Console.WriteLine("XML file path not correctly identified.")
                        ShowHelp()
                        Environment.ExitCode = -3
                        Exit Sub
                    End Try

                    If File.Exists(strOutputXML) Then
                        Console.WriteLine("File (" & strOutputXML & ") already exists.")
                        ShowHelp()
                        Environment.ExitCode = -4
                        Exit Sub
                    End If


                Case arrArgs(i).ToLower.StartsWith("/s")
                    Try
                        Dim colonPlace As Integer = arrArgs.Item(i).IndexOf(":") + 1
                        strOutputXML = arrArgs.Item(i).Substring(colonPlace, arrArgs.Item(i).Length - colonPlace)

                    Catch e As Exception
                        Console.WriteLine("XML file path not correctly identified.")
                        ShowHelp()
                        Environment.ExitCode = -5
                        Exit Sub
                    End Try

                    If File.Exists(strOutputXML) Then
                        Console.WriteLine("File (" & strOutputXML & ") already exists.")
                        ShowHelp()
                        Environment.ExitCode = -6
                        Exit Sub
                    End If

                Case arrArgs(i).ToLower.StartsWith("/u")    'Check for Username
                    Try
                        Dim colonPlace As Integer = arrArgs.Item(i).IndexOf(":") + 1
                        strUsername = arrArgs.Item(i).Substring(colonPlace, arrArgs.Item(i).Length - colonPlace)

                    Catch e As Exception
                        Console.WriteLine("Username not correctly specified.")
                        ShowHelp()
                        Environment.ExitCode = -7
                        Exit Sub
                    End Try

                Case arrArgs(i).ToLower.StartsWith("/p")    'Check for Password
                    Try
                        Dim colonPlace As Integer = arrArgs.Item(i).IndexOf(":") + 1
                        strPassword = arrArgs.Item(i).Substring(colonPlace, arrArgs.Item(i).Length - colonPlace)

                    Catch e As Exception
                        Console.WriteLine("Password not correctly specified.")
                        ShowHelp()
                        Environment.ExitCode = -8
                        Exit Sub
                    End Try

                Case arrArgs(i).ToLower.StartsWith("/y")
                    CarryOn = True

                Case arrArgs(i).ToLower.StartsWith("/d")
                    DebugFlag = True

                Case Else
                    ShowHelp()
                    Exit Sub

            End Select
        Next

        'If the /y is not selected. then do this.

        Dim line As String
        Console.WriteLine()

        If strServerAddress = String.Empty Then                                '/a
            Console.WriteLine("No server address supplied.")
            ShowHelp()
            Exit Sub
        Else
            Console.WriteLine("Using Address: " & strServerAddress)
        End If

        If strReportName = String.Empty Then                                '/r
            Console.WriteLine("Scan Name not specified, downloading all of them.")
            strReportName = "0"
        Else
            Console.WriteLine("Downloading Report: " & strReportName)
        End If

        If strOutputXML = "" Then                                          '/o
            Console.WriteLine("File Output Path not specified.  Using scan name and current directory.")
            strOutputXML = "0"
        Else
            Console.WriteLine("File Output Path: " & strOutputXML)
        End If

        'Check for Username, prompt if it isn't specified.
        If strUsername = "" Then                                          '/u
            Do
                Console.Write("Enter Username:" & vbTab)
                line = Console.ReadLine()
            Loop While line Is Nothing
        Else
            Console.WriteLine("Using Username: " & strUsername)
        End If

        'Check for Password, prompt if it isn't specified.
        If strPassword = "" Then                                          '/p
            strPassword = GetPassword()
        Else
            Console.WriteLine("Password is specified, but is not displayed.")
        End If

        If Not CarryOn Then
            Do
                Console.Write("Are these values correct? (y/n) ")
                line = Console.ReadLine()
                If Not line.ToLower.StartsWith("y") Then
                    Console.WriteLine("Answer is not yes... Exiting.")
                    Exit Sub
                End If
            Loop While line Is Nothing
        End If

        Dim strToken As String = PerformLogin(strServerAddress, strUsername, strPassword)

        'Check to see if the login was good, if not bomb out.
        If strToken = "0" Then
            Console.WriteLine("Server Unreachable.")
            Environment.ExitCode = -100
            Exit Sub

        ElseIf strToken = "1" Then
            Console.WriteLine("Login Unsuccessful.  Check Username and Password, then try again.")
            Environment.ExitCode = -101
            Exit Sub

        Else
            Console.WriteLine("Login Successful.")
        End If

        DownloadReport(strServerAddress, strToken, strReportName, strOutputXML)

        'Dim XslFile As String = "C:\Users\Chris\Desktop\Chris\html.xsl"

        ''PerformLogoff(strServerAddress)

        'ElseIf arrArgs.Count <= 2 Then
        'ShowHelp()
        'Exit Sub
        'End If
    End Sub

    Private Sub DownloadReport(ByVal ServerAddress As String, ByVal token As String, ByVal ReportName As String, Optional ByVal XmlSaveFile As String = "0",
                               Optional ByVal ExportFormat As String = "v2")
        Dim myDoc As New XmlDocument
        Dim XmlOutputFile As String
        Dim uuid As String = String.Empty
        Dim xmlReportNode As XmlNode
        Dim reportNodes As XmlNodeList

        myDoc.LoadXml(MakeRequest(ServerAddress & "/report/list", "token=" & token))

        DebugWrite(ReportName)

        If ReportName <> "0" Then
            Try
                reportNodes = myDoc.SelectNodes("//report[readableName='" & ReportName & "']")

            Catch ex As Exception
                DebugWrite(ex.Message)
                Console.WriteLine("Hmmm ... Looks like there is a problem with the name. Check spelling and give it another shot.")
                Exit Sub
            End Try

            Dim ReportTime As DateTime = Now
            Dim line As String = "0"

            If reportNodes.Count > 1 Then
                Console.WriteLine("There is more than one report with this name.  Selecting the newest.")
                Dim i As Integer = 0
                Dim timestamp As Double = 0

                For Each ThisNode As XmlNode In reportNodes
                    If timestamp < CDbl(ThisNode("timestamp").InnerText) Then
                        timestamp = CDbl(ThisNode("timestamp").InnerText)
                        line = i
                    End If

                    DebugWrite(i & "." & vbTab & ThisNode("readableName").InnerText & _
                                      vbTab & UNIXTimeToDateTime(ThisNode("timestamp").InnerText))
                    i = i + 1
                Next
                ReportTime = UNIXTimeToDateTime(timestamp)

                xmlReportNode = reportNodes.Item(line)
                DebugWrite(xmlReportNode.OuterXml)

                '
                'This is here if I wanted to add the date to the filename.
                '
                'Dim oFile As FileInfo = New FileInfo(XmlSaveFile)
                'XmlSaveFile = oFile.DirectoryName & "\" & FormattingFilename(ReportTime) & " " & oFile.Name

                DebugWrite("Attempting to write out filename: " & XmlSaveFile)

            ElseIf reportNodes.Count = 1 Then
                xmlReportNode = reportNodes.Item(0)
                DebugWrite(xmlReportNode.OuterXml)

            Else
                Console.WriteLine("There was no report by that name.")
                Exit Sub

            End If

            uuid = xmlReportNode("name").InnerText

            DebugWrite("Report Unique ID: " & uuid)

            myDoc.LoadXml(MakeRequest(ServerAddress & "/file/report/download", "token=" & token & "&report=" & uuid))

            If XmlSaveFile = "0" Then
                XmlOutputFile = ReportName.Replace("/", "_").Replace("\", "_").Replace(":", "_") & ".nessus"
            Else
                XmlOutputFile = XmlSaveFile
            End If

            DebugWrite("Output File: " & XmlOutputFile)

            File.WriteAllText(XmlOutputFile, myDoc.OuterXml)

        Else
            reportNodes = myDoc.SelectNodes("//report")

            Dim line As String = "0"

            Do
                Console.Write("You are about to download all of the reports, continue? (y/n) ")
                line = Console.ReadLine()
                If Not line.ToLower.StartsWith("y") Then
                    Console.WriteLine("Answer is not yes... Exiting.")
                    Exit Sub
                End If
            Loop While line Is Nothing

            For Each ThisNode As XmlNode In reportNodes
                uuid = ThisNode("name").InnerText
                Console.WriteLine(ThisNode("readableName").InnerText)

                XmlSaveFile = ThisNode("readableName").InnerText.Replace("/", "_").Replace("\", "_").Replace(":", "_").Replace(".", "_") & ".nessus"

                Dim ReportTime As DateTime = UNIXTimeToDateTime(ThisNode("timestamp").InnerText)

                DebugWrite(ThisNode.OuterXml)

                XmlSaveFile = FormattingFilename(ReportTime) & " " & XmlSaveFile

                Dim myXML As String = MakeRequest(ServerAddress & "/file/report/download", "token=" & token & "&report=" & uuid)

                DebugWrite("Output File: " & XmlSaveFile)
                File.WriteAllText(XmlSaveFile, myXML)
            Next
        End If

    End Sub

    Private Function ValidateRemoteCertificate(ByVal sender As Object, ByVal certificate As X509Certificate, ByVal chain As X509Chain, ByVal policyErrors As SslPolicyErrors) As Boolean
        ' allow any old dodgy certificate...
        Return True
    End Function

    Private Function MakeRequest(ByVal uri As String, ByVal postdata As String) As String
        Dim encoding As New ASCIIEncoding()

        ' allows for validation of SSL conversations
        ServicePointManager.ServerCertificateValidationCallback = New RemoteCertificateValidationCallback(AddressOf ValidateRemoteCertificate)

        Dim byte1 As Byte() = encoding.GetBytes(postdata)
        Dim myHttpWebRequest As HttpWebRequest = DirectCast(WebRequest.Create(uri), HttpWebRequest)
        myHttpWebRequest.AllowAutoRedirect = True
        myHttpWebRequest.Method = "POST"

        ' Set the content type of the data being posted.
        myHttpWebRequest.ContentType = "application/x-www-form-urlencoded"

        ' Set the content length of the string being posted.
        myHttpWebRequest.ContentLength = byte1.Length

        Dim newStream As IO.Stream = myHttpWebRequest.GetRequestStream()
        newStream.Write(byte1, 0, byte1.Length)

        Dim response As HttpWebResponse = Nothing
        Try
            response = DirectCast(myHttpWebRequest.GetResponse(), HttpWebResponse)

            Using s As IO.Stream = response.GetResponseStream()
                Using sr As New IO.StreamReader(s)
                    Return sr.ReadToEnd()
                End Using
            End Using

        Catch e As Exception
            Console.WriteLine(e.Message)
            Return Nothing
        Finally
            If response IsNot Nothing Then
                response.Close()
            End If
        End Try

    End Function

    Public Function PerformLogin(ByVal ServerAddress As String, ByVal Username As String, ByVal Password As String) As String
        Dim myDoc As New XmlDocument
        Dim myXML As String

        'Attempt to reach the server and login
        Try
            myXML = MakeRequest(ServerAddress & "/login", "login=" & Username & "&password=" & Password)

        Catch ex As Exception
            'This means that the server was not available for some reason, could be a bad address or a variety of issues.
            PerformLogin = "0"
            Exit Function
        End Try

        'Get the response if the server is working.
        myDoc.LoadXml(myXML)

        'Check to see if the login was good, if not bomb out.
        If myDoc.SelectSingleNode("//status").InnerText.ToLower <> "ok" Then
            PerformLogin = "1"
            Exit Function

        Else
            'Retrieve the authentication token
            Dim token As String = myDoc.SelectSingleNode("//token").InnerText

            DebugWrite("Session Token: " & token)

            Return token
        End If

    End Function

    Public Function PerformLogoff(ByVal ServerAddress As String) As String
        Dim myDoc As New XmlDocument
        Dim myXML As String

        'Attempt to reach the server and login
        Try
            myXML = MakeRequest(ServerAddress & "/logoff", "seq=2686")
            DebugWrite(myXML)

        Catch ex As Exception
            'This means that the server was not available for some reason, could be a bad address or a variety of issues.
            PerformLogoff = "0"
            Exit Function
        End Try

        'Get the response if the server is working.
        Try
            myDoc.LoadXml(myXML)
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            PerformLogoff = "1"
            Exit Function
        End Try

        'Check to see if the login was good, if not bomb out.
        If myDoc.SelectSingleNode("//status").InnerText.ToLower <> "ok" Then
            PerformLogoff = "2"
        Else
            Return "Logoff Successful"
        End If

    End Function

    Public Sub PerformTransform(ByVal outputFile As String, ByVal xmlFile As String, ByVal xslFile As String)

        If xslFile = Nothing Then
            Console.WriteLine("No XSL Specified, not running any additional conversions.")
        End If

        Dim stream As FileStream = File.Open(outputFile, FileMode.Create)

        ' Create XsltCommand and compile stylesheet.
        Dim processor As New XslCompiledTransform()
        processor.Load(xslFile)

        ' Transform the file.
        processor.Transform(xmlFile, Nothing, stream)

        stream.Close()
        stream.Dispose()

    End Sub

    Sub ShowHelp()
        Console.WriteLine("")
        Console.WriteLine("Nessus to Excel Conversion Utility with exclusions support.")
        Console.WriteLine("Required Switches (must include colon, case insensitive): ")
        Console.WriteLine("  /a:<address>" & vbTab & "URL to Nessus Server. (https://localhost:8834), be sure that you include the https.")
        Console.WriteLine("  /u:<username>" & vbTab & "Nessus Username")
        Console.WriteLine("  /p:<password>" & vbTab & "Nessus Password")
        Console.WriteLine("  /s:<scan>" & vbTab & "Scan Name, name of scan to be started or interrogated.")
        Console.WriteLine()
        Console.WriteLine("Scanning Switches (must include colon): ")
        Console.WriteLine("  /p:<path>" & vbTab & "Nessus Policy File Path")
        Console.WriteLine("  /y" & vbTab & vbTab & "No Prompting, answer yes.")
        Console.WriteLine()
        Console.WriteLine("Statusing Switches (must include colon): ")
        Console.WriteLine("  /o:<path>" & vbTab & "Excel Output File Path")
        Console.WriteLine("  /y" & vbTab & vbTab & "No Prompting, answer yes.")
        Console.WriteLine()
        Console.WriteLine("Reporting Switches (must include colon): ")
        Console.WriteLine("  /o:<path>" & vbTab & "Output File Path, if not specified then scan name will be used.")
        Console.WriteLine("  /x:<path>" & vbTab & "export format, if not specified then nessu v2 format will be used.")
        Console.WriteLine("  /y" & vbTab & vbTab & "No Prompting, answer yes.")
        Console.WriteLine()
        Console.WriteLine("Optional Switches (must include colon): ")
        Console.WriteLine("  /d" & vbTab & "Display Debug Output")
        Console.WriteLine("  /y" & vbTab & vbTab & "No Prompting, answer yes.")
    End Sub

    Private Function GetPassword() As String
        Dim password As New SecureString
        Console.Write("Password: ")
        'get the first character of the password
        Dim nextKey As ConsoleKeyInfo = Console.ReadKey(True)

        Do While nextKey.Key <> ConsoleKey.Enter

            If nextKey.Key = ConsoleKey.Backspace Then

                If (password.Length > 0) Then

                    password.RemoveAt(password.Length - 1)

                    'erase the last * as well
                    Console.Write(nextKey.KeyChar)
                    Console.Write(" ")
                    Console.Write(nextKey.KeyChar)
                End If

            Else

                password.AppendChar(nextKey.KeyChar)
                Console.Write("*")
            End If

            nextKey = Console.ReadKey(True)
        Loop

        password.MakeReadOnly()

        Dim bstr As IntPtr = Marshal.SecureStringToBSTR(password)

        Try
        Finally
            Marshal.ZeroFreeBSTR(bstr)
        End Try
        Return bstr

    End Function

    Private Sub DebugWrite(ByVal inputString As String)
        If DebugFlag Then
            Console.WriteLine(inputString)
        End If
    End Sub

    Private Function UNIXTimeToDateTime(ByVal unixTime As Double) As DateTime
        Return New DateTime(1970, 1, 1, 0, 0, 0, 0, DateTimeKind.Utc).AddSeconds(unixTime).ToLocalTime()
    End Function

    Private Function FileNaming(ByVal Filename As String) As String
        Dim Folder As String
        Folder = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly.Location)

        Dim fileToCreate As System.IO.FileInfo = New System.IO.FileInfo(Filename)
        Dim i As Integer = 1

        Dim myName As String = fileToCreate.Name.Replace(fileToCreate.Extension, "")

        While (fileToCreate.Exists())
            fileToCreate = New System.IO.FileInfo(Folder & myName & " (" & i.ToString() & ")" & fileToCreate.Extension)
        End While

        Return fileToCreate.Name
    End Function

    Private Function FormattingFilename(ByVal DateToFormat As DateTime) As String
        Dim NewFormat As String

        NewFormat = Year(DateToFormat)

        If Month(DateToFormat) < 10 Then
            NewFormat = NewFormat & "-0" & Month(DateToFormat)
        Else
            NewFormat = NewFormat & "-" & Month(DateToFormat)
        End If

        If Day(DateToFormat) < 10 Then
            NewFormat = NewFormat & "-0" & Day(DateToFormat)
        Else
            NewFormat = NewFormat & "-" & Day(DateToFormat)
        End If

        Return NewFormat
    End Function

End Module
