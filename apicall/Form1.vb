Imports System.IO
Imports System.Net
Imports ChoETL
Imports Newtonsoft.Json

Public Class Form1
    Dim sKey As String
    Dim sSite As String
    Dim sFunctions As String
    Dim aFunctions() As String
    Dim sReports As String
    Dim aReports() As String
    Dim sCarsFile As String
    Dim sFrom As String
    Dim sTo As String
    Dim sCar As String
    Dim sReport As String
    Dim aCars() As String
    Dim sInterval As String
    Dim sFtp As String
    Dim sUsername As String
    Dim sPassword As String
    Dim sFolder As String = "DATA"
    Dim sDelete As String
    Dim sSep As String = ";"

    Dim sDeleteFilesAfterUpload As Boolean = False
    Dim sApiCall As String
    Dim jsonRead As String
    Dim countSeconds As Integer = 0
    Dim carIndex As Integer = 0
    Dim reportIndex As Integer = 0
    Dim functionIndex As Integer = 0
    Dim flagProcessingData As Boolean = False
    Dim flagSendingData As Boolean = False
    Dim flagReportsCalls As Boolean = False


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Not My.Computer.FileSystem.FileExists("settings.ini") Then
            If MsgBox("Settings file not found!!! Application could not run!", vbOKOnly, "4maini - api call") = vbOK Then
                End
            End If
        Else

            Dim My_Ini As New Ini("settings.ini")

            'example of calling functions for ini file
            'My_Ini.SetValue("Settings", "key", "230", True)
            'My_Ini.DeleteValue("Settings", "site", True)
            'My_Ini.DeleteSection("Settings", True)

            sKey = My_Ini.GetValue("Settings", "key")
            sSite = My_Ini.GetValue("Settings", "site")
            sFunctions = My_Ini.GetValue("Settings", "functions")
            aFunctions = Split(sFunctions, ",")
            sReports = My_Ini.GetValue("Settings", "reports")
            aReports = Split(sReports, ",")
            sCarsFile = My_Ini.GetValue("Settings", "cars")
            sFrom = My_Ini.GetValue("Settings", "from")
            sTo = My_Ini.GetValue("Settings", "to")
            sInterval = My_Ini.GetValue("Settings", "interval")
            sFtp = My_Ini.GetValue("Settings", "ftp")
            sUsername = My_Ini.GetValue("Settings", "username")
            sPassword = My_Ini.GetValue("Settings", "password")
            sDelete = My_Ini.GetValue("Settings", "delete")
            sSep = My_Ini.GetValue("Settings", "sep")

            If sFtp.Length > 0 And sUsername = "" Then
                If MsgBox("username not defined in settings.ini file!!! Application could not run!", vbOKOnly, "4maini - api call") = vbOK Then
                    End
                End If
            End If

            If sKey = "" Then
                sKey = "396b026cf8695cf23a63cad01ec712b7"
            End If
            If sSite = "" Then
                sSite = "https://www.gpstracking.ro/api"
            End If
            If sFunctions = "" Then
                sFunctions = "getVehicleJourneys,getDailyActivity,getFilledQuestionnaires,getVehicleSummary"
                aFunctions = Split(sFunctions, ",")
            End If
            If sReports = "" Then
                sReports = "40,91"
                aReports = Split(sReports, ",")
            End If

            If sCarsFile = "" Then
                sCarsFile = "cars.txt"
            End If
            If sFrom = "" Then
                sFrom = "2000-12-01"
            End If
            If sTo = "" Then
                sTo = Format(Now, "yyyy-MM-dd")
            End If
            If sInterval = "" Then
                sInterval = "30500"
            End If
            If sDelete = "true" Or sDelete = "1" Then
                sDeleteFilesAfterUpload = True
            Else
                sDeleteFilesAfterUpload = False
            End If
            If sSep = "" Then
                sSep = ","
            End If


            aCars = File.ReadAllLines(sCarsFile)
            carIndex = 0
            functionIndex = 0

            Debug.Print("key:" + sKey)
            Debug.Print("site:" + sSite)
            Debug.Print("functions:" + sFunctions)
            Debug.Print("cars:" + sCarsFile)
            Debug.Print("from:" + sFrom)
            Debug.Print("to:" + sTo)
            Debug.Print("aFunctions:" + aFunctions(0) + aFunctions(1) + aFunctions(2) + aFunctions(3))
            Timer1.Interval = "1"
            countSeconds = 0
            Debug.Print("Creating data directory")
            Try
                My.Computer.FileSystem.CreateDirectory(sFolder)
            Catch
                If MsgBox("Error creating data folder! Application could not run!" + " Data folder: " + sFolder, vbOKOnly, "4maini - api call") = vbOK Then
                    End
                End If
            End Try
            If Now.Month < 0 Then
                Me.Text = Me.Text & " - Beta version"
                If MsgBox("Versiune beta! Contactati furnizorul pentru versiunea cu licenta!", vbOKOnly, "4maini - api call") = vbOK Then
                    End
                End If
            End If
        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If flagProcessingData Or flagSendingData Then
            Exit Sub
        End If
        If carIndex < aCars.Length Then
            If Trim(aCars(carIndex)).Length = 0 Then 'rand gol in fisierul cars - treci la urmatoarea masina de pe randul urmator
                carIndex = carIndex + 1
                Exit Sub
            End If
            sCar = Trim(aCars(carIndex)).Replace(" ", "%20")
            If aFunctions(functionIndex) = "getFilledQuestionnaires" Then
                'pentru fiecare report fa un api call
                Dim report As String
                'de rezolvat iteratia - la fiecare tick se parcurge un report
                If reportIndex = aReports.Length Then
                    reportIndex = 0
                    carIndex = carIndex + 1
                    flagReportsCalls = False
                Else
                    flagReportsCalls = True
                    report = aReports(reportIndex)
                    sApiCall = sSite + "?" + aFunctions(functionIndex) + "=" + sKey + "&cars[]=" + sCar + "&quest[]=" + report + "&from=" + sFrom + "&to=" + sTo
                    Debug.Print("sApiCall:" + sApiCall)
                    'citeste json de la site pentru fiecare report
                    If ReadJson(sApiCall, report) Then
                        ListBox1.Items.Add(sApiCall)
                        ListBox1.SelectedIndex = ListBox1.Items.Count - 1
                        Label2.Text = aFunctions(functionIndex)
                        Label3.Text = aCars(carIndex) + " - " + report
                        reportIndex = reportIndex + 1
                    End If
                End If
            Else
                sApiCall = sSite + "?" + aFunctions(functionIndex) + "=" + sKey + "&cars[]=" + sCar + "&from=" + sFrom + "&to=" + sTo
                Debug.Print("sApiCall:" + sApiCall)
                'citeste json de la site
                If ReadJson(sApiCall) Then
                    ListBox1.Items.Add(sApiCall)
                    ListBox1.SelectedIndex = ListBox1.Items.Count - 1
                    Label2.Text = aFunctions(functionIndex)
                    Label3.Text = aCars(carIndex)
                    carIndex = carIndex + 1
                End If
            End If
        Else
            'reinitializeaza datele
            carIndex = 0
            functionIndex = functionIndex + 1
            sTo = Format(Now, "yyyy-MM-dd")
            'grupeaza datele intr-un fisier pentru fiecare functie
            ' trimite datele ------------------------------------pentru functia anterioara si toti soferii
            If SendData(aFunctions(functionIndex - 1), sFtp, sUsername, sPassword, sDeleteFilesAfterUpload) Then
                Debug.Print("Data sent!")
                ToolStripLabel1.ForeColor = Color.Black
                ToolStripLabel1.Text = "Data sent!"
            Else
                Debug.Print("Data not sent!!! Error in sending data to the ftp server!")
                ToolStripLabel1.ForeColor = Color.Red
                ToolStripLabel1.Text = "Data not sent! Ftp not configured or bad ftp settings."
            End If

            'tragem datele doar pentru 4 functii - de la capat
            If functionIndex > 3 Then
                functionIndex = 0
                ListBox1.Items.Clear()
            End If
        End If
        Timer1.Interval = CLng(sInterval)
    End Sub
    'afiseaza secundele pana la urmatoarea citire
    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        If flagProcessingData Or flagSendingData Then
            Label1.Text = "..."
            Exit Sub
        End If
        Dim value As Integer

        countSeconds = countSeconds + 1
        If countSeconds = CInt(Timer1.Interval / 1000) Then
            countSeconds = 0
        End If
        value = CInt(Timer1.Interval / 1000) - countSeconds
        If value < 0 Then
            value = 0
        End If
        Label1.Text = value.ToString
    End Sub


    'reload settings
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Form1_Load(sender, e)
    End Sub

    'functia apeleaza site-ul si salveaza datele primite
    Private Function ReadJson(ByVal sUrl As String, Optional report As String = "") As Boolean
        Try
            flagProcessingData = True
            Dim request As WebRequest = WebRequest.Create(sUrl)
            Dim response As WebResponse = request.GetResponse()
            ' Display the status.  
            Console.WriteLine(CType(response, HttpWebResponse).StatusDescription)
            'show status
            ToolStripLabel1.ForeColor = Color.Black
            ToolStripLabel1.Text = CType(response, HttpWebResponse).StatusDescription
            ' Get the stream containing content returned by the server.  
            Dim dataStream As Stream = response.GetResponseStream()
            ' Open the stream using a StreamReader for easy access.  
            Dim reader As New StreamReader(dataStream)
            ' Read the content.  
            Dim responseFromServer As String = reader.ReadToEnd()
            ' Display the content.  
            'Console.WriteLine(responseFromServer)
            ' Clean up the streams and the response.  
            reader.Close()
            response.Close()

            'wait the right time for the request - by default, definded by site, 30 seconds
            If responseFromServer.IndexOf("Too many requests") > 0 Then
                flagProcessingData = True
                ToolStripLabel1.Text = "Please wait 30 seconds! Too many api requests!"
                If MsgBox(responseFromServer, vbOKOnly, "4maini - api call") = vbOK Then
                    ToolStripLabel1.Text = "Please wait 30 seconds! Too many api requests!"
                    Me.Refresh()
                    Threading.Thread.Sleep(30500)
                    Me.Button1.PerformClick()
                    flagProcessingData = False
                End If
            End If
            jsonRead = JsonConvert.DeserializeObject(responseFromServer).ToString()

            If report.Length = 0 Then 'fara report
                File.WriteAllText(sFolder + "\" + aFunctions(functionIndex) + "_" + carIndex.ToString + "_data.json", jsonRead)

                'converteste datele pentru fiecare sofer
                Dim r = New ChoJSONReader(sFolder + "\" + aFunctions(functionIndex) + "_" + carIndex.ToString + "_data.json").WithFlatToNestedObjectSupport
                Dim w = New ChoCSVWriter(sFolder + "\" + aFunctions(functionIndex) + "_" + carIndex.ToString + ".csv").WithFirstLineHeader().WithDelimiter(sSep).QuoteAllFields(True)
                w.Write(r)

                w.Close()
                r.Close()

                'aplica transformare fisier csv numai pentru functiile getVehicleJourneys, getDailyActivity
                If aFunctions(functionIndex) = "getVehicleJourneys" Or aFunctions(functionIndex) = "getDailyActivity" Then
                    AdjustCSVFile(sFolder + "\" + aFunctions(functionIndex) + "_" + carIndex.ToString + ".csv", aFunctions(functionIndex))
                End If
            Else 'report is present
                File.WriteAllText(sFolder + "\" + aFunctions(functionIndex) + "_" + carIndex.ToString + "_" + report + "_data.json", jsonRead)

                'converteste datele pentru fiecare sofer
                Dim r = New ChoJSONReader(sFolder + "\" + aFunctions(functionIndex) + "_" + carIndex.ToString + "_" + report + "_data.json").WithFlatToNestedObjectSupport
                Dim w = New ChoCSVWriter(sFolder + "\" + aFunctions(functionIndex) + "_" + carIndex.ToString + "_" + report + ".csv").WithFirstLineHeader().WithDelimiter(sSep).QuoteAllFields(True)
                w.Write(r)

                w.Close()
                r.Close()

                'aplica transformare fisier csv pentru functia getFilledQuestionnaires - grupeaza valorile 'Cu ce putem ajuta?' intr-un singur camp - daca raportul este 91
                If aFunctions(functionIndex) = "getFilledQuestionnaires" And report = "91" Then
                    AdjustCSVFile(sFolder + "\" + aFunctions(functionIndex) + "_" + carIndex.ToString + "_" + report + ".csv", aFunctions(functionIndex))
                End If
            End If
                flagProcessingData = False
            Return True
        Catch
            Return False
        End Try
    End Function

    Private Function SendData(ByVal functionName As String, ByVal ftpSite As String, ByVal user As String, ByVal pass As String, ByVal deleteAfterUpload As Boolean) As Boolean

        Dim fileToSend As String = sFolder + "\" + functionName + ".csv"
        Dim car As Integer
        Dim report As Integer
        Dim readFirst As Boolean
        Dim carChanged As Boolean = False
        Dim reportChanged As Boolean = False
        Dim line As String

        flagSendingData = True
        'grupeaza datele mai intai si apoi trimite fisierul ce contine toti soferii
        'grupeaza toate fisierele functionName_[0..aCars.Length-1].csv in functionName.csv
        'mai intai sterge fisierele vechi daca exista
        If File.Exists(Application.StartupPath + "\" + sFolder + "\" + functionName + ".csv") Then
            File.Delete(Application.StartupPath + "\" + sFolder + "\" + functionName + ".csv")
        End If

        For report = 0 To aReports.Length - 1
            If File.Exists(Application.StartupPath + "\" + sFolder + "\" + functionName + "_" + aReports(report).ToString + ".csv") Then
                File.Delete(Application.StartupPath + "\" + sFolder + "\" + functionName + "_" + aReports(report).ToString + ".csv")
            End If
        Next


        'cumuleaza fisierele cu reports pentru functia cu rapoarte - prelucrare suplimentara a fisierelor cu rapoarte
        If functionName = "getFilledQuestionnaires" Then
            For report = 0 To aReports.Length - 1
                reportChanged = True
                For car = 0 To aCars.Length - 1
                    readFirst = True
                    carChanged = True
                    'read line by line
                    Dim info As New FileInfo(Application.StartupPath + "\" + sFolder + "\" + functionName + "_" + car.ToString + "_" + aReports(report) + ".csv")
                    Dim length As Long = info.Length
                    If length > 0 Then
                        For Each line In File.ReadAllLines(Application.StartupPath + "\" + sFolder + "\" + functionName + "_" + car.ToString + "_" + aReports(report) + ".csv")
                            If Not readFirst Then
                                File.AppendAllText(Application.StartupPath + "\" + sFolder + "\" + functionName + "_" + aReports(report).ToString + ".csv", line + vbCrLf)
                            Else
                                If Not carChanged Then
                                    File.AppendAllText(Application.StartupPath + "\" + sFolder + "\" + functionName + "_" + aReports(report).ToString + ".csv", line + vbCrLf)
                                End If
                            End If
                            If reportChanged Then
                                File.AppendAllText(Application.StartupPath + "\" + sFolder + "\" + functionName + "_" + aReports(report).ToString + ".csv", line + vbCrLf)
                                reportChanged = False
                            End If
                            readFirst = False
                            carChanged = False
                        Next
                    Else
                        File.AppendAllText(Application.StartupPath + "\" + sFolder + "\" + functionName + "_" + aReports(report).ToString + ".csv", "")
                    End If
                Next
            Next
        Else
            'cumuleaza fisierele cu soferi intr-unul singur pentru fiecare functie
            For car = 0 To aCars.Length - 1
                readFirst = True
                'read line by line
                For Each line In File.ReadAllLines(Application.StartupPath + "\" + sFolder + "\" + functionName + "_" + car.ToString + ".csv")
                    If car = 0 Then
                        File.AppendAllText(Application.StartupPath + "\" + sFolder + "\" + functionName + ".csv", line + vbCrLf)
                    Else
                        If Not readFirst Then
                            File.AppendAllText(Application.StartupPath + "\" + sFolder + "\" + functionName + ".csv", line + vbCrLf)
                        End If
                    End If
                    readFirst = False
                Next
            Next
        End If

        'trimite si datele daca avem ftpSite
        If ftpSite.Length > 0 Then
            If UploadFTPFiles(ftpSite, user, pass, fileToSend, deleteAfterUpload) Then
                flagSendingData = False
                Return True
            Else
                flagSendingData = False
                Return False
            End If
        Else
            flagSendingData = False
            Return False
        End If
    End Function

    Private Function UploadFTPFiles(ftpAddress As String, ftpUser As String, ftpPassword As String, fileToUpload As String, deleteAfterUpload As Boolean) As Boolean

        'daca fisierul nu exista nu am ce uploada, returneaza fals - de implementat trimiterea fisierelor de tip <nume functie>_<report>.csv - acestea nu se pun pe ftp momentan
        If Not File.Exists(fileToUpload) Then Return False

        Dim credential As NetworkCredential

        Try
            credential = New NetworkCredential(ftpUser, ftpPassword)

            Dim sFtpFile As String = fileToUpload

            Dim request As FtpWebRequest = DirectCast(WebRequest.Create("ftp://" & ftpAddress & "/" & Path.GetFileName(sFtpFile)), FtpWebRequest)

            request.KeepAlive = False
            request.Method = WebRequestMethods.Ftp.UploadFile
            request.Credentials = credential
            request.UsePassive = False
            request.Timeout = (60 * 1000) * 3 '3 mins

            Using reader As New FileStream(fileToUpload, FileMode.Open)

                Dim buffer(Convert.ToInt32(reader.Length - 1)) As Byte
                reader.Read(buffer, 0, buffer.Length)
                reader.Close()

                request.ContentLength = buffer.Length
                Dim stream As Stream = request.GetRequestStream
                stream.Write(buffer, 0, buffer.Length)
                stream.Close()

                Using response As FtpWebResponse = DirectCast(request.GetResponse, FtpWebResponse)

                    If deleteAfterUpload Then
                        My.Computer.FileSystem.DeleteFile(fileToUpload)
                    End If

                    response.Close()
                End Using

            End Using

            Return True

        Catch ex As Exception
            'ExceptionInfo = ex
            Return False
        Finally

        End Try

    End Function

    Private Sub AdjustCSVFile(file As String, functionName As String)
        If functionName = "getFilledQuestionnaires" Then
            Using MyReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(file)
                MyReader.TextFieldType = FileIO.FieldType.Delimited
                MyReader.SetDelimiters(sSep)
                Dim currentRow As String()
                Dim index As Integer = 0
                Dim curentElement As Integer = 0
                Dim firstRow As Boolean = True
                Dim lineRead As String = ""
                Dim startCuceputemajuta As Boolean = False
                Dim CuCePutemAjuta As String = ""


                Dim f As System.IO.StreamWriter
                f = My.Computer.FileSystem.OpenTextFileWriter(file + "_", True)

                While Not MyReader.EndOfData
                    Try
                        currentRow = MyReader.ReadFields()
                        Dim currentField As String
                        Dim i As Integer
                        For Each currentField In currentRow
                            If Not firstRow Then
                                If currentField = "Alte Observatii" Then
                                    startCuceputemajuta = False
                                    lineRead = lineRead + Chr(34) + CuCePutemAjuta + Chr(34) + sSep + "826" + sSep + currentField + sSep
                                    Continue For
                                End If
                                If startCuceputemajuta Then
                                    If Not Integer.TryParse(currentField, i) Then 'campul este valoare string
                                        If CuCePutemAjuta.Length > 0 Then
                                            CuCePutemAjuta = CuCePutemAjuta & "+" & currentField
                                        Else
                                            CuCePutemAjuta = currentField
                                        End If
                                    End If
                                    Continue For
                                End If
                                If currentField = "Il putem ajuta cu ceva?" Then
                                    startCuceputemajuta = True
                                    CuCePutemAjuta = ""
                                    lineRead = lineRead + Chr(34) + currentField + Chr(34) + sSep
                                    Continue For
                                End If

                                lineRead = lineRead + Chr(34) + currentField + Chr(34) + sSep
                            Else 'primul rand, ignora tot ce contine fields_7_val_
                                If InStr(currentField, "fields_7_val") > 0 Then
                                Else
                                    lineRead = lineRead + Chr(34) + currentField + Chr(34) + sSep
                                End If
                            End If
                        Next
                        firstRow = False
                        f.WriteLine(lineRead)
                        lineRead = ""
                    Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
                        MsgBox("Line " & ex.Message & "is not valid and will be skipped.")
                    End Try
                End While
                f.Close()
            End Using
            My.Computer.FileSystem.MoveFile(file + "_", file, True)
            Exit Sub
        End If

        If functionName = "getVehicleJourneys" Then
            Using MyReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(file)
                MyReader.TextFieldType = FileIO.FieldType.Delimited
                MyReader.SetDelimiters(sSep)
                Dim currentRow As String()
                Dim index As Integer = 0
                Dim secondRow As Integer = 0
                Dim curentElement As Integer = 0
                Dim firstRow As Boolean = True
                Dim lineRead As String = ""
                Dim primele3Elemente(3) As String
                Dim done As Boolean = False

                Dim f As System.IO.StreamWriter
                f = My.Computer.FileSystem.OpenTextFileWriter(file + "_", True)

                While Not MyReader.EndOfData
                    Try
                        currentRow = MyReader.ReadFields()
                        Dim currentField As String
                        For Each currentField In currentRow
                            If Not firstRow Then
                                'citeste randul 2 in calupuri de index elemente
                                'scrie in fisier elementele citite pe fiecare linie
                                lineRead = lineRead + Chr(34) + currentField + Chr(34) + sSep
                                If curentElement < 3 And Not done Then
                                    primele3Elemente(curentElement) = Chr(34) + currentField + Chr(34)
                                    If curentElement = 2 Then
                                        done = True
                                    Else
                                        done = False
                                    End If
                                End If
                                curentElement = curentElement + 1
                                If curentElement = index Then
                                    'Debug.Print(lineRead)
                                    f.WriteLine(lineRead)
                                    'lineRead = sSep & sSep & sSep
                                    lineRead = primele3Elemente(0) + sSep + primele3Elemente(1) + sSep + primele3Elemente(2) + sSep
                                    curentElement = 3
                                    If secondRow = 2 Then
                                        index = index - 3
                                    End If
                                End If
                                secondRow = secondRow + 1
                            Else
                                If currentField = "journeys_1_start_time" Then
                                    'read next line
                                    f.WriteLine(lineRead)
                                    lineRead = ""
                                    secondRow = secondRow + 1
                                    Exit For
                                Else
                                    index = index + 1
                                    lineRead = lineRead + Chr(34) + currentField + Chr(34) + sSep
                                    'scrie elementul citit in fisier pe prima linie
                                End If
                            End If
                            'Debug.Print(lineRead)
                        Next
                        firstRow = False
                    Catch ex As Microsoft.VisualBasic.
                                FileIO.MalformedLineException
                        MsgBox("Line " & ex.Message & "is not valid and will be skipped.")
                    End Try
                End While
                f.Close()
            End Using
            My.Computer.FileSystem.MoveFile(file + "_", file, True)
            Exit Sub
        End If

        If functionName = "getDailyActivity" Then
            Using MyReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(file)
                MyReader.TextFieldType = FileIO.FieldType.Delimited
                MyReader.SetDelimiters(sSep)
                Dim currentRow As String()
                Dim index As Integer = 0
                Dim secondRow As Integer = 0
                Dim curentElement As Integer = 0
                Dim firstRow As Boolean = True
                Dim lineRead As String = ""
                Dim arrayOfIndex(10000) As Integer
                Dim row As Integer = 0
                Dim primele3Elemente(3) As String
                Dim done As Boolean = False

                Dim f As System.IO.StreamWriter
                f = My.Computer.FileSystem.OpenTextFileWriter(file + "_", True)

                While Not MyReader.EndOfData
                    Try
                        currentRow = MyReader.ReadFields()
                        Dim currentField As String
                        For Each currentField In currentRow
                            If Not firstRow Then
                                'citeste randul 2 in calupuri de index elemente
                                'scrie in fisier elementele citite pe fiecare linie, dar ignora campul cu numele soferului din finalul randului
                                If InStr(currentField.ToString, " ") > 0 And curentElement > 2 And Not InStr(currentField.ToString, "-") > 0 Then
                                    'skip record if it is a name (of the driver), not on the second line
                                Else
                                    lineRead = lineRead + Chr(34) + currentField + Chr(34) + sSep
                                End If
                                If curentElement < 3 And Not done Then
                                    primele3Elemente(curentElement) = Chr(34) + currentField + Chr(34)
                                    If curentElement = 2 Then
                                        done = True
                                    Else
                                        done = False
                                    End If
                                End If
                                curentElement = curentElement + 1
                                If row < 1 Then
                                    index = arrayOfIndex(row)
                                Else
                                    index = arrayOfIndex(row) + 2
                                End If
                                If curentElement = index Then
                                    'Debug.Print(lineRead)
                                    f.WriteLine(lineRead)
                                    row = row + 1
                                    'lineRead = sSep & sSep & sSep
                                    lineRead = primele3Elemente(0) + sSep + primele3Elemente(1) + sSep + primele3Elemente(2) + sSep
                                    index = arrayOfIndex(row)
                                    curentElement = 2 'arrayOfIndex(row) - arrayOfIndex(row - 1)
                                End If
                            Else 'first row
                                If currentField = "days_" + (secondRow).ToString + "_mark" Then
                                    'read next line
                                    If secondRow = 0 Then
                                        f.WriteLine(lineRead + Chr(34) + currentField + Chr(34))
                                    End If
                                    secondRow = secondRow + 1
                                    arrayOfIndex(secondRow - 1) = index + 1
                                    index = 0
                                    lineRead = ""
                                Else
                                    index = index + 1
                                    If secondRow = 0 Then
                                        lineRead = lineRead + Chr(34) + currentField + Chr(34) + sSep
                                    Else
                                        lineRead = ""
                                    End If
                                    arrayOfIndex(secondRow) = index
                                End If

                            End If
                        Next
                        firstRow = False
                    Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
                        MsgBox("Line " & ex.Message & "is not valid and will be skipped.")
                    End Try
                End While
                f.Close()
                arrayOfIndex = Nothing
            End Using
            My.Computer.FileSystem.MoveFile(file + "_", file, True)
            Exit Sub
        End If
    End Sub

End Class
