Imports System.IO
Public Class Main_Screen

    Private lastinputline As String = ""
    Private inputlines As Long = 0
    Private highestPercentageReached As Integer = 0
    Private inputlinesprecount As Long = 0
    Private pretestdone As Boolean = False
    Private primary_PercentComplete As Integer = 0
    Private percentComplete As Integer

    Private Cancelled As Boolean

    Private Sub Error_Handler(ByVal ex As Exception, Optional ByVal identifier_msg As String = "")
        Try
            If My.Computer.FileSystem.FileExists((Application.StartupPath & "\Sounds\UHOH.WAV").Replace("\\", "\")) = True Then
                My.Computer.Audio.Play((Application.StartupPath & "\Sounds\UHOH.WAV").Replace("\\", "\"), AudioPlayMode.Background)
            End If
            If ex.Message.IndexOf("Thread was being aborted") < 0 Then
                Dim Display_Message1 As New Display_Message()
                Display_Message1.Message_Textbox.Text = "The Application encountered the following problem: " & vbCrLf & identifier_msg & ":" & ex.Message.ToString
                Display_Message1.Timer1.Interval = 1000
                Display_Message1.ShowDialog()
                Dim dir As System.IO.DirectoryInfo = New System.IO.DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs")
                If dir.Exists = False Then
                    dir.Create()
                End If
                dir = Nothing
                Dim filewriter As System.IO.StreamWriter = New System.IO.StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
                filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy hh:mm:ss tt") & " - " & identifier_msg & ":" & ex.ToString)
                filewriter.Flush()
                filewriter.Close()
                filewriter = Nothing
            End If
        Catch exc As Exception
            MsgBox("An error occurred in the application's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            Label5.Text = "Initializing Operation Variables"
            If FolderBrowserDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Textbox1.Text = FolderBrowserDialog1.SelectedPath
                If FolderBrowserDialog2.ShowDialog() = Windows.Forms.DialogResult.OK Then
                    Textbox2.Text = FolderBrowserDialog2.SelectedPath
                    
                    Button1.Enabled = False
                    Button2.Visible = True
                    Button3.Visible = True
                    Button2.Enabled = True
                    Button3.Enabled = True
                    Label6.Text = ""
                    Me.ControlBox = False
                    Cancelled = False

                    lastinputline = ""
                    inputlines = 0
                    highestPercentageReached = 0
                    inputlinesprecount = 0
                    pretestdone = False
                    primary_PercentComplete = 0
                    percentComplete = 0


                    BackgroundWorker1.RunWorkerAsync()
                Else
                    Label5.Text = "Operation Request Cancelled"
                End If
            Else
                Label5.Text = "Operation Request Cancelled"
            End If

        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Dim resultfoldername As String = (Textbox2.Text & "\CodeUnit Projects").Replace("\\", "\")
            ProgressBar1.Value = 0
            If Cancelled = False Then
                Try
                    If My.Computer.FileSystem.DirectoryExists(resultfoldername) Then
                        Me.WindowState = FormWindowState.Minimized
                        My.Computer.FileSystem.DeleteDirectory(resultfoldername, Microsoft.VisualBasic.FileIO.UIOption.AllDialogs, FileIO.RecycleOption.SendToRecycleBin)
                        Me.WindowState = FormWindowState.Normal
                    End If
                Catch ex As Exception
                    Button2.Enabled = False
                    Cancelled = True
                    BackgroundWorker1.CancelAsync()
                    Me.WindowState = FormWindowState.Normal
                    Exit Sub
                End Try
            Else
                Exit Sub
            End If

            If Cancelled = True Then
                Exit Sub
            End If

            My.Computer.FileSystem.CreateDirectory(resultfoldername)

            Dim dinfo As DirectoryInfo = New DirectoryInfo(Textbox1.Text)
            inputlinesprecount = dinfo.GetDirectories.Length
            inputlines = 0

            ' Report progress as a percentage of the total task.
            percentComplete = 0
            If inputlinesprecount > 0 Then
                percentComplete = CSng(inputlines) / CSng(inputlinesprecount) * 100
            Else
                percentComplete = 100
            End If
            primary_PercentComplete = percentComplete
            If percentComplete > highestPercentageReached Then
                highestPercentageReached = percentComplete
                BackgroundWorker1.ReportProgress(percentComplete)
            End If

            If My.Computer.FileSystem.FileExists((Application.StartupPath & "\CodeUnit_Banner_200x49.jpg").Replace("\\", "\")) = True Then
                My.Computer.FileSystem.CopyFile((Application.StartupPath & "\CodeUnit_Banner_200x49.jpg").Replace("\\", "\"), (resultfoldername & "\CodeUnit_Banner_200x49.jpg").Replace("\\", "\"), True)
            End If

            Dim mainwriter As StreamWriter = New StreamWriter((resultfoldername & "\default.htm").Replace("\\", "\"), False)
            mainwriter.WriteLine("<html>")
            mainwriter.WriteLine("<head>")
            mainwriter.WriteLine("<title>CodeUnit Projects</title>")
            mainwriter.WriteLine("<link rel=""stylesheet"" type=""text/css"" href=""CodeUnit_Website.css"">")
            mainwriter.WriteLine("</head>")
            mainwriter.WriteLine("<frameset rows=""64,*"" framespacing=""0"" border=""0"" frameborder=""0"">")
            mainwriter.WriteLine("<frame name=""banner"" scrolling=""no"" noresize target=""contents"" src=""header.htm"" marginwidth=""1"">")
            mainwriter.WriteLine("<frameset cols=""392,*"">")
            mainwriter.WriteLine("<frame name=""contents"" target=""main"" src=""default_menu.htm"">")
            mainwriter.WriteLine("<frame name=""main"" src=""blank.htm"" target=""main"">")
            mainwriter.WriteLine("</frameset>")
            mainwriter.WriteLine("<noframes>")
            mainwriter.WriteLine("<body>")
            mainwriter.WriteLine("<p>This page uses frames, but your browser doesn't support them.</p>")
            mainwriter.WriteLine("</body>")
            mainwriter.WriteLine("</noframes>")
            mainwriter.WriteLine("</frameset>")
            mainwriter.WriteLine("</html>")

            mainwriter.Flush()
            mainwriter.Close()

            Dim blankwriter As StreamWriter = New StreamWriter((resultfoldername & "\blank.htm").Replace("\\", "\"), False)
            blankwriter.WriteLine("<html>")
            blankwriter.WriteLine("<head>")
            blankwriter.WriteLine("<title>CodeUnit Projects</title>")
            blankwriter.WriteLine("<link rel=""stylesheet"" type=""text/css"" href=""CodeUnit_Website.css"">")
            blankwriter.WriteLine("</head>")
            blankwriter.WriteLine("<body>")
            blankwriter.WriteLine("</body>")
            blankwriter.WriteLine("</html>")
            blankwriter.Flush()
            blankwriter.Close()

            Dim csswriter As StreamWriter = New StreamWriter((resultfoldername & "\CodeUnit_Website.css").Replace("\\", "\"), False)

            csswriter.WriteLine("a:link {color: #003366;")
            csswriter.WriteLine("font-family: verdana;")
            csswriter.WriteLine("font-size: 8pt;")
            csswriter.WriteLine("FONT-WEIGHT: bold;")
            csswriter.WriteLine("text-decoration: none}")

            csswriter.WriteLine("a:visited {color: #003366;")
            csswriter.WriteLine("font-family: verdana;")
            csswriter.WriteLine("font-size: 8pt;")
            csswriter.WriteLine("FONT-WEIGHT: bold;")
            csswriter.WriteLine("text-decoration: none}")

            csswriter.WriteLine("a:hover{color: ff6600;")
            csswriter.WriteLine("font-family: verdana;")
            csswriter.WriteLine("font-size: 8pt;")
            csswriter.WriteLine("text-decoration: none;")
            csswriter.WriteLine("}")



            csswriter.WriteLine("H1.error {	FONT-WEIGHT: bold; FONT-SIZE: 16px; COLOR: #ff6600; FONT-FAMILY: Verdana, Arial, Helvetica, sans-serif}")
            csswriter.WriteLine("H1 {	FONT-WEIGHT: bold; FONT-SIZE: 16px; COLOR: #003366; FONT-FAMILY: Verdana, Arial, Helvetica, sans-serif}")
            csswriter.WriteLine("H2 {	FONT-WEIGHT: bold; FONT-SIZE: 14px; COLOR: #006699; FONT-FAMILY: Verdana, Arial, Helvetica, sans-serif}")
            csswriter.WriteLine("H3 {	FONT-WEIGHT: bold; FONT-SIZE: 12px; COLOR: #000000; FONT-FAMILY: Verdana, Arial, Helvetica, sans-serif}")
            csswriter.WriteLine("H3.error {	FONT-WEIGHT: bold; FONT-SIZE: 12px; COLOR: #F8AB06; FONT-FAMILY: Verdana, Arial, Helvetica, sans-serif}")
            csswriter.WriteLine("H4 {	FONT-WEIGHT: bold; FONT-SIZE: 12px; COLOR: #808080; FONT-FAMILY: Verdana, Arial, Helvetica, sans-serif}")

            csswriter.WriteLine("body {")
            csswriter.WriteLine("margin: 10px")
            csswriter.WriteLine("}")

            csswriter.WriteLine("p { font-family: Verdana, Tahoma, Arial; color: #000000; font-weight: normal; font-size: 8pt }")
            csswriter.Flush()
            csswriter.Close()


            Dim headerwriter As StreamWriter = New StreamWriter((resultfoldername & "\header.htm").Replace("\\", "\"), False)
            headerwriter.WriteLine("<html>")
            headerwriter.WriteLine("<head>")
            headerwriter.WriteLine("<title>CodeUnit Projects</title>")
            headerwriter.WriteLine("<link rel=""stylesheet"" type=""text/css"" href=""CodeUnit_Website.css"">")
            headerwriter.WriteLine("<base target=""contents"">")
            headerwriter.WriteLine("</head>")
            headerwriter.WriteLine("<body>")
            headerwriter.WriteLine("<h1><img border=""0"" src=""CodeUnit_Banner_200x49.jpg"" width=""200"" height=""49""></h1>")
            headerwriter.WriteLine("</body>")
            headerwriter.WriteLine("</html>")

            headerwriter.Flush()
            headerwriter.Close()

            Dim filewriter As StreamWriter = New StreamWriter((resultfoldername & "\default_menu.htm").Replace("\\", "\"), False)

            filewriter.WriteLine("<html>")
            filewriter.WriteLine("<head>")
            filewriter.WriteLine("<title>CodeUnit Projects</title>")
            filewriter.WriteLine("<link rel=""stylesheet"" type=""text/css"" href=""CodeUnit_Website.css"">")
            filewriter.WriteLine("<base target=""main"">")
            filewriter.WriteLine("</head>")
            filewriter.WriteLine("<body>")
            filewriter.WriteLine("<h1>CodeUnit Projects</h1>")
            filewriter.WriteLine("<p>There are currently " & inputlinesprecount & " CodeUnit projects listed (as of " & Format(Now(), "dd/MM/yyyy") & "). Click on a project below to pull up its details and downloads.</p>")
            filewriter.WriteLine("<p>Sort list by: <a target=""_self"" href=""default_menu.htm"">[Name]</a>")
            filewriter.WriteLine("<a target=""_self"" href=""default_menu_create.htm"">[Creation Date]</a>")
            filewriter.WriteLine("<a target=""_self"" href=""default_menu_build.htm"">[Build Number]</a></p>")
            filewriter.WriteLine("<h2><b>Project Listing <font color=""#C0C0C0"">(by Name)</font></b></h2><ul>")

            Dim shortproject, shortbuild, shortcreation As String

            Dim shortcreationList, shortbuildList As ArrayList
            shortcreationList = New ArrayList
            shortbuildList = New ArrayList

            For Each projectinfo As DirectoryInfo In dinfo.GetDirectories()
                Try
                    If projectinfo.Name.IndexOf(" - ") <> -1 Then


                        shortproject = ""
                        shortbuild = ""
                        shortcreation = ""

                        lastinputline = projectinfo.Name
                        shortproject = projectinfo.Name.Remove(projectinfo.Name.LastIndexOf("-") - 1, projectinfo.Name.Length - projectinfo.Name.LastIndexOf("-") + 1)

                        Dim dte As Date
                        dte = dte.Parse(projectinfo.Name.Substring(projectinfo.Name.LastIndexOf("-") - 1, projectinfo.Name.Length - projectinfo.Name.LastIndexOf("-") + 1).Remove(0, 3))
                        shortcreation = (Format(dte, "yyyy/MM/dd"))



                        If My.Computer.FileSystem.DirectoryExists((projectinfo.FullName & "\Release").Replace("\\", "\")) = True Then
                            My.Computer.FileSystem.CopyDirectory((projectinfo.FullName & "\Release").Replace("\\", "\"), (resultfoldername & "\" & projectinfo.Name & "\Release").Replace("\\", "\"))

                            filewriter.WriteLine("<li><a target=""main"" href=""" & projectinfo.Name & ".htm"">" & shortproject & "</a>")

                            Dim projectwriter As StreamWriter = New StreamWriter((resultfoldername & "\" & projectinfo.Name & ".htm").Replace("\\", "\"), False)
                            projectwriter.WriteLine("<html><head><link rel=""stylesheet"" type=""text/css"" href=""CodeUnit_Website.css""></head><body><h2>" & projectinfo.Name & "</h2>")
                            If My.Computer.FileSystem.FileExists((projectinfo.FullName & "\Build.txt").Replace("\\", "\")) Then
                                My.Computer.FileSystem.CopyFile((projectinfo.FullName & "\Build.txt").Replace("\\", "\"), (resultfoldername & "\" & projectinfo.Name & "\Build.txt").Replace("\\", "\"), True)
                                projectwriter.WriteLine("<h3>Build: </h3><p>" & My.Computer.FileSystem.ReadAllText((resultfoldername & "\" & projectinfo.Name & "\Build.txt").Replace("\\", "\")) & "</p>")
                                shortbuild = My.Computer.FileSystem.ReadAllText((resultfoldername & "\" & projectinfo.Name & "\Build.txt").Replace("\\", "\"))
                                filewriter.WriteLine("<br/><font size=""1"" face=""Verdana"" color=""#99CCFF"">BUILD " & My.Computer.FileSystem.ReadAllText((resultfoldername & "\" & projectinfo.Name & "\Build.txt").Replace("\\", "\")) & "</font><font size=""1"" face=""Verdana"" color=""#CCCCCC""> (" & shortcreation & ")</font></li>")
                            Else
                                filewriter.WriteLine("<br/><font size=""1"" face=""Verdana"" color=""#99CCFF"">BUILD UNKNOWN</font><font size=""1"" face=""Verdana"" color=""#CCCCCC""> (" & shortcreation & ")</font></li>")
                                shortbuild = "UNKNOWN"
                            End If
                            If My.Computer.FileSystem.FileExists((projectinfo.FullName & "\Description.txt").Replace("\\", "\")) Then
                                My.Computer.FileSystem.CopyFile((projectinfo.FullName & "\Description.txt").Replace("\\", "\"), (resultfoldername & "\" & projectinfo.Name & "\Description.txt").Replace("\\", "\"), True)
                                projectwriter.WriteLine("<h3>Description: </h3><p>" & My.Computer.FileSystem.ReadAllText((resultfoldername & "\" & projectinfo.Name & "\Description.txt").Replace("\\", "\")).Replace(vbCrLf, "<br/>") & "</p>")
                            End If
                            If My.Computer.FileSystem.FileExists((projectinfo.FullName & "\Preview_Image.jpg").Replace("\\", "\")) Then
                                My.Computer.FileSystem.CopyFile((projectinfo.FullName & "\Preview_Image.jpg").Replace("\\", "\"), (resultfoldername & "\" & projectinfo.Name & "\Preview_Image.jpg").Replace("\\", "\"), True)
                                projectwriter.WriteLine("<h3>Screen Shot:</h3>")
                                projectwriter.WriteLine("<p><img border=""1"" src=""" & projectinfo.Name & "\Preview_Image.jpg" & """></p>")
                            End If
                            projectwriter.WriteLine("</body></html>")
                            projectwriter.WriteLine("<h3>Installer Downloads:</h3>")
                            recursivewrite((resultfoldername & "\" & projectinfo.Name & "\Release").Replace("\\", "\"), projectinfo.Name & "/", projectwriter)
                            projectwriter.Flush()
                            projectwriter.Close()
                            projectwriter = Nothing


                        End If
                        shortcreationList.Add(shortcreation & "||" & shortbuild & "||" & projectinfo.FullName)
                        shortbuildList.Add(shortbuild & "||" & shortcreation & "||" & projectinfo.FullName)



                        projectinfo = Nothing

                       
                    End If
Catch ex As Exception
                    Error_Handler(ex, "Generating Project Details")
                End Try
                Try
                    inputlines = inputlines + 1
                    ' Report progress as a percentage of the total task.
                    percentComplete = 0
                    If inputlinesprecount > 0 Then
                        percentComplete = CSng(inputlines) / CSng(inputlinesprecount) * 100
                    Else
                        percentComplete = 100
                    End If
                    primary_PercentComplete = percentComplete
                    If percentComplete > highestPercentageReached Then
                        highestPercentageReached = percentComplete
                        BackgroundWorker1.ReportProgress(percentComplete)
                    End If

                    If Cancelled = True Then
                        Exit For
                    End If
                Catch ex As Exception
                    Error_Handler(ex, "Generating Project Details")
                End Try

            Next
            shortbuildList.Sort()
            shortbuildList.Reverse()

            Dim filewriter2 As StreamWriter = New StreamWriter((resultfoldername & "\default_menu_build.htm").Replace("\\", "\"), False)
            filewriter2.WriteLine("<html>")
            filewriter2.WriteLine("<head>")
            filewriter2.WriteLine("<title>CodeUnit Projects</title>")
            filewriter2.WriteLine("<link rel=""stylesheet"" type=""text/css"" href=""CodeUnit_Website.css"">")
            filewriter2.WriteLine("<base target=""main"">")
            filewriter2.WriteLine("</head>")
            filewriter2.WriteLine("<body>")
            filewriter2.WriteLine("<h1>CodeUnit Projects</h1>")
            filewriter2.WriteLine("<p>There are currently " & inputlinesprecount & " CodeUnit projects listed (as of " & Format(Now(), "dd/MM/yyyy") & "). Click on a project below to pull up its details and downloads.</p>")
            filewriter2.WriteLine("<p>Sort list by: <a target=""_self"" href=""default_menu.htm"">[Name]</a>")
            filewriter2.WriteLine("<a target=""_self"" href=""default_menu_create.htm"">[Creation Date]</a>")
            filewriter2.WriteLine("<a target=""_self"" href=""default_menu_build.htm"">[Build Number]</a></p>")
            filewriter2.WriteLine("<h2><b>Project Listing <font color=""#C0C0C0"">(by Build)</font></b></h2><ul>")
            For Each str As String In shortbuildList
                Dim strarr As String() = str.Split("||")
                shortbuild = strarr(0)
                shortcreation = strarr(2)
                shortproject = strarr(4)

                Dim proj As DirectoryInfo = New DirectoryInfo(shortproject)

                shortproject = proj.Name.Remove(proj.Name.LastIndexOf("-") - 1, proj.Name.Length - proj.Name.LastIndexOf("-") + 1)
                filewriter2.WriteLine("<li><a target=""main"" href=""" & proj.Name & ".htm"">" & shortproject & "</a>")
                filewriter2.WriteLine("<br/><font size=""1"" face=""Verdana"" color=""#99CCFF"">BUILD " & shortbuild & "</font><font size=""1"" face=""Verdana"" color=""#CCCCCC""> (" & shortcreation & ")</font></li>")
            Next
            filewriter2.WriteLine("</ul>")
            filewriter2.WriteLine("</body></html>")
            filewriter2.Flush()
            filewriter2.Close()
            filewriter2 = Nothing


            shortcreationList.Sort()
            shortcreationList.Reverse()

            Dim filewriter3 As StreamWriter = New StreamWriter((resultfoldername & "\default_menu_create.htm").Replace("\\", "\"), False)
            filewriter3.WriteLine("<html>")
            filewriter3.WriteLine("<head>")
            filewriter3.WriteLine("<title>CodeUnit Projects</title>")
            filewriter3.WriteLine("<link rel=""stylesheet"" type=""text/css"" href=""CodeUnit_Website.css"">")
            filewriter3.WriteLine("<base target=""main"">")
            filewriter3.WriteLine("</head>")
            filewriter3.WriteLine("<body>")
            filewriter3.WriteLine("<h1>CodeUnit Projects</h1>")
            filewriter3.WriteLine("<p>There are currently " & inputlinesprecount & " CodeUnit projects listed (as of " & Format(Now(), "dd/MM/yyyy") & "). Click on a project below to pull up its details and downloads.</p>")
            filewriter3.WriteLine("<p>Sort list by: <a target=""_self"" href=""default_menu.htm"">[Name]</a>")
            filewriter3.WriteLine("<a target=""_self"" href=""default_menu_create.htm"">[Creation Date]</a>")
            filewriter3.WriteLine("<a target=""_self"" href=""default_menu_build.htm"">[Build Number]</a></p>")
            filewriter3.WriteLine("<h2><b>Project Listing <font color=""#C0C0C0"">(by Creation Date)</font></b></h2><ul>")
            For Each str As String In shortcreationList
                Dim strarr As String() = str.Split("||")
                shortbuild = strarr(2)
                shortcreation = strarr(0)
                shortproject = strarr(4)

                Dim proj As DirectoryInfo = New DirectoryInfo(shortproject)

                shortproject = proj.Name.Remove(proj.Name.LastIndexOf("-") - 1, proj.Name.Length - proj.Name.LastIndexOf("-") + 1)
                filewriter3.WriteLine("<li><a target=""main"" href=""" & proj.Name & ".htm"">" & shortproject & "</a>")
                filewriter3.WriteLine("<br/><font size=""1"" face=""Verdana"" color=""#99CCFF"">BUILD " & shortbuild & "</font><font size=""1"" face=""Verdana"" color=""#CCCCCC""> (" & shortcreation & ")</font></li>")
            Next
            filewriter3.WriteLine("</ul>")
            filewriter3.WriteLine("</body></html>")
            filewriter3.Flush()
            filewriter3.Close()
            filewriter3 = Nothing

            shortbuildList.Clear()
            shortbuildList = Nothing
            shortcreationList.Clear()
            shortcreationList = Nothing
            filewriter.WriteLine("</ul>")
            filewriter.WriteLine("</body></html>")
            filewriter.Flush()
            filewriter.Close()
            filewriter = Nothing
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub



    Private Sub ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        Try
            ProgressBar1.Value = e.ProgressPercentage
            Label5.Text = "Added " & lastinputline
            Label6.Text = inputlines & "/" & inputlinesprecount
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            If e.Cancelled = True Or Cancelled = True Then

                If My.Computer.FileSystem.FileExists((Application.StartupPath & "\Sounds\HEEY.WAV").Replace("\\", "\")) = True Then
                    My.Computer.Audio.Play((Application.StartupPath & "\Sounds\HEEY.WAV").Replace("\\", "\"), AudioPlayMode.Background)
                End If
                MsgBox("Operation Cancelled. You will now be prompted to remove the half-generated download folder")
                Button1.Enabled = True
                Button2.Visible = False
                Button3.Visible = False
                Button2.Enabled = False
                Button3.Enabled = False
                Me.ControlBox = True
                Dim resultfoldername As String = (Textbox2.Text & "\CodeUnit Projects").Replace("\\", "\")
                Try
                    If My.Computer.FileSystem.DirectoryExists(resultfoldername) Then
                        Me.WindowState = FormWindowState.Minimized
                        My.Computer.FileSystem.DeleteDirectory(resultfoldername, Microsoft.VisualBasic.FileIO.UIOption.AllDialogs, FileIO.RecycleOption.SendToRecycleBin)
                        Me.WindowState = FormWindowState.Normal
                        Label5.Text = "Operation Cancelled. Incomplete Generated Folder Removed"
                    End If
                Catch ex As Exception
                    Label5.Text = "Operation Cancelled. Incomplete Generated Folder Retained"
                    Me.WindowState = FormWindowState.Normal
                End Try
                ProgressBar1.Value = 100
            Else
                If My.Computer.FileSystem.FileExists((Application.StartupPath & "\Sounds\VICTORY.WAV").Replace("\\", "\")) = True Then
                    My.Computer.Audio.Play((Application.StartupPath & "\Sounds\VICTORY.WAV").Replace("\\", "\"), AudioPlayMode.Background)
                End If
                Label5.Text = "Operation Completed"
                MsgBox("Operation Completed", MsgBoxStyle.Information, "Operation Completed")
                Button1.Enabled = True
                Button2.Visible = False
                Button3.Visible = False
                Button2.Enabled = False
                Button3.Enabled = False
                Me.ControlBox = True
                Shell("explorer " & (Textbox2.Text & "\CodeUnit Projects\default.htm").Replace("\\", "\"), AppWinStyle.NormalFocus, False, 10000)
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub Form1_Closing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        Try
            Label5.Text = "Application Shutdown Initiated"
            My.Settings("SourceDirectory") = Textbox1.Text
            My.Settings("TargetDirectory") = Textbox2.Text
            My.Settings.Save()
        Catch ex As Exception
            Error_Handler(ex, "Application Closed")
        End Try
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.Text = My.Application.Info.ProductName & " " & My.Application.Info.Version.Major & Format(My.Application.Info.Version.Minor, "00") & Format(My.Application.Info.Version.Build, "00") & "." & My.Application.Info.Version.Revision
            Control.CheckForIllegalCrossThreadCalls = False
            If Not My.Settings Is Nothing Then
                If My.Settings("SourceDirectory").ToString.Length > 0 Then
                    If My.Computer.FileSystem.DirectoryExists(My.Settings("SourceDirectory")) = True Then
                        Textbox1.Text = My.Settings("SourceDirectory")
                        FolderBrowserDialog1.SelectedPath = My.Settings("SourceDirectory")
                    End If
                End If
                If My.Settings("TargetDirectory").ToString.Length > 0 Then
                    If My.Computer.FileSystem.DirectoryExists(My.Settings("TargetDirectory")) = True Then
                        Textbox2.Text = My.Settings("TargetDirectory")
                        FolderBrowserDialog2.SelectedPath = My.Settings("TargetDirectory")
                    End If
                End If
            End If
            Label5.Text = "Application Successfully Loaded"
        Catch ex As Exception
            Error_Handler(ex, "Application Load")
        End Try

    End Sub

    Private Sub recursivewrite(ByVal SourceFolder As String, ByVal downloadPath As String, ByVal fileWriter As StreamWriter)
        Try
            If Cancelled = True Then
                Exit Sub
            End If

            Dim dinfo As DirectoryInfo = New DirectoryInfo(SourceFolder)
            If dinfo.Exists = True Then
                FolderWalker(downloadPath, dinfo.FullName, fileWriter)
            End If
            dinfo = Nothing

            If Cancelled = True Then
                Exit Sub
            End If

        Catch ex As Exception
            Error_Handler(ex, "Recursive Write")
        End Try
    End Sub

    Private Function filesize(ByVal size As Long) As String
        Try
            Dim result As String
            Dim switch As String = "bytes"
            If size >= 1024 Then switch = "KB"
            If size >= 1048576 Then switch = "MB"
            If size >= 1073741824 Then switch = "GB"

            Select Case switch
                Case "KB"
                    result = Math.Round((size / 1024), 1).ToString & " KB"
                Case "MB"
                    result = Math.Round((size / 1048576), 1).ToString & " MB"
                Case "GB"
                    result = Math.Round((size / 1073741824), 1).ToString & " GB"
                Case Else
                    result = size.ToString & " Bytes"
            End Select



            result = "<font size=""1"" color=""#5378C3""> (" & result & ")</font>"
            Return result
        Catch ex As Exception
            Error_Handler(ex, "Displaying File Size")
            Return ""
        End Try
    End Function

    Private Sub FolderWalker(ByVal rootdir As String, ByVal targetdir As String, ByVal filewriter As StreamWriter)
        Try
            Dim dinfo As DirectoryInfo = New DirectoryInfo(targetdir)
            Dim finfo As FileInfo
            rootdir = (rootdir & dinfo.Name & "/").Replace("//", "/")
            filewriter.WriteLine("</ul>" & vbCrLf & "<p><b><font color=""#CCCCCC"">" & (rootdir).Remove(rootdir.Length - 1, 1) & "</font></b></p>" & vbCrLf & "<ul>" & vbCrLf)
            For Each finfo In dinfo.GetFiles()
                If Cancelled = True Then
                    Exit For
                End If
                filewriter.WriteLine("<li>" & ((finfo.Name.Remove(finfo.Name.Length - finfo.Extension.Length, finfo.Extension.Length))) & " - Click <a target=""_blank"" href=""" & (rootdir & "/" & finfo.Name).Replace("//", "/") & """>here</a> <font size=""1"" color=""#008080"">(" & finfo.Extension.ToUpper & ")</font> " & filesize(finfo.Length) & "<font size=""1"" color=""#808080""> (" & Format(Now(), "dd/MM/yyyy") & ") </font></li>")
            Next
            finfo = Nothing
            Dim d2info As DirectoryInfo
            For Each d2info In dinfo.GetDirectories()
                If Cancelled = True Then
                    Exit For
                End If
                FolderWalker(rootdir, d2info.FullName, filewriter)
            Next
            d2info = Nothing
            If Cancelled = True Then
                Exit Sub
            End If
        Catch ex As Exception
            Error_Handler(ex, "Folder Walker")
        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Try
            Me.WindowState = FormWindowState.Minimized
        Catch ex As Exception
            Error_Handler(ex, "Minimize Window")
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            Button2.Enabled = False
            BackgroundWorker1.CancelAsync()
            Cancelled = True
            Label5.Text = "Operation Cancelled"
        Catch ex As Exception
            Error_Handler(ex, "Operation Cancel")
        End Try
    End Sub
End Class
