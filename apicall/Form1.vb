Imports System.IO
Imports System.Net
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports ChoETL
Imports Newtonsoft.Json
Imports CsvHelper
Imports System.Xml

Public Class Form1
    Dim sKey As String
    Dim sSite As String
    Dim sFunctions As String
    Dim aFunctions() As String
    Dim sCarsFile As String
    Dim sFrom As String
    Dim sTo As String
    Dim sCar As String
    Dim aCars() As String
    Dim sInterval As String
    Dim sFtp As String
    Dim sUsername As String
    Dim sPassword As String
    Dim sFolder As String = "DATA"
    Dim sDelete As String

    Dim sDeleteFilesAfterUpload As Boolean = False
    Dim sApiCall As String
    Dim jsonRead As String
    Dim countSeconds As Integer = 0
    Dim carIndex As Integer = 0
    Dim functionIndex As Integer = 0
    Dim flagProcessingData As Boolean = False
    Dim flagSendingData As Boolean = False


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

            sCarsFile = My_Ini.GetValue("Settings", "cars")
            sFrom = My_Ini.GetValue("Settings", "from")
            sTo = My_Ini.GetValue("Settings", "to")
            sInterval = My_Ini.GetValue("Settings", "interval")
            sFtp = My_Ini.GetValue("Settings", "ftp")
            sUsername = My_Ini.GetValue("Settings", "username")
            sPassword = My_Ini.GetValue("Settings", "password")
            sDelete = My_Ini.GetValue("Settings", "delete")

            If sFtp = "" Or sUsername = "" Then
                If MsgBox("ftp or/and username not defined in settings.ini file!!! Application could not run!", vbOKOnly, "4maini - api call") = vbOK Then
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
                sFunctions = "getQuestionnaires,getFilledQuestionnaires,getVehicleJourneys,getDailyActivity"
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
                sInterval = 30500
            End If
            If sDelete = "true" Or sDelete = "1" Then
                sDeleteFilesAfterUpload = True
            Else
                sDeleteFilesAfterUpload = False
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
            Me.Text = Me.Text & " - Beta version"
            If Now.Month > 9 Then
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
            sCar = aCars(carIndex).Replace(" ", "%20")
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
        Else
            'reinitializeaza datele
            carIndex = 0
            functionIndex = functionIndex + 1
            'tragem datele doar pentru 4 functii - de la capat
            If functionIndex > 3 Then
                functionIndex = 0
                ListBox1.Items.Clear()
                ' trimite datele - in caz ca se trimit la final de functie
                ' ---------AICI----------
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
    Private Function ReadJson(ByVal sUrl As String) As Boolean
        Try
            flagProcessingData = True
            Dim request As WebRequest = WebRequest.Create(sUrl)
            Dim response As WebResponse = request.GetResponse()
            ' Display the status.  
            Console.WriteLine(CType(response, HttpWebResponse).StatusDescription)
            'show status
            Label4.Text = CType(response, HttpWebResponse).StatusDescription
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

            'Console.WriteLine(jsonRead)
            File.WriteAllText(sFolder + "/" + aFunctions(functionIndex) + "_" + carIndex.ToString + "_data.json", jsonRead)

            'converteste datele pentru fiecare sofer
            Dim r = New ChoJSONReader(sFolder + "/" + aFunctions(functionIndex) + "_" + carIndex.ToString + "_data.json")
            Dim w = New ChoCSVWriter(sFolder + "/" + aFunctions(functionIndex) + "_" + carIndex.ToString + "_export.csv")
            w.Write(r)

            w.Close()
            r.Close()

            ' trimite datele ------------------------------------
            If SendData(sFolder + "/" + aFunctions(functionIndex) + "_" + carIndex.ToString + "_export.csv", sFtp, sUsername, sPassword, sDeleteFilesAfterUpload) Then
                Debug.Print("Data sent!")
                ToolStripLabel1.ForeColor = Color.Black
                ToolStripLabel1.Text = "Data sent!"
            Else
                Debug.Print("Data not sent!!! Error in sending data to the ftp server!")
                ToolStripLabel1.ForeColor = Color.Red
                ToolStripLabel1.Text = "Data not sent!!! Check your ftp server settings in the settings file! "
            End If

            flagProcessingData = False
            Return True
        Catch
            Return False
        End Try
    End Function

    Private Function SendData(ByVal fileToSend As String, ByVal ftpSite As String, ByVal user As String, ByVal pass As String, ByVal deleteAfterUpload As Boolean) As Boolean
        flagSendingData = True


        If UploadFTPFiles(ftpSite, user, pass, fileToSend, deleteAfterUpload) Then
            flagSendingData = False
            Return True
        Else
            flagSendingData = False
            Return False
        End If
    End Function

    Private Function UploadFTPFiles(ftpAddress As String, ftpUser As String, ftpPassword As String,
                                   fileToUpload As String,
                                   deleteAfterUpload As Boolean) As Boolean

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
End Class
