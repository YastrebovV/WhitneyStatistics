Imports System.IO
Imports System.Text
Imports System.Runtime.InteropServices
Imports System.Net
Imports System.Net.Sockets
Imports System.Net.NetworkInformation
Imports System.Data.SQLite
Imports System.Threading
Imports Microsoft.VisualBasic.FileIO

Public Class Main
    <DllImport("user32.dll")>
    Private Shared Function SetWindowLong(ByVal window As IntPtr, ByVal index As Integer, ByVal value As Integer) As Integer
    End Function
    <DllImport("user32.dll")>
    Private Shared Function GetWindowLong(ByVal window As IntPtr, ByVal index As Integer) As Integer
    End Function
    Private Const GWL_EXSTYLE As Integer = -20
    Private Const WS_EX_TOOLWINDOW As Integer = &H80
    Sub HideFromAltTab(ByVal Handle As IntPtr)
        SetWindowLong(Handle, GWL_EXSTYLE, GetWindowLong(Handle, GWL_EXSTYLE) Or WS_EX_TOOLWINDOW)
    End Sub 'убирает приложение из меню TabAlt
    Public FlibHndl As Short, FlibHndl2 As Short, start As Boolean,
    _stop As Boolean, hold As Boolean,
    count As Int32, errinfo As Short,
    erralm As Short, odbah As Focas1.ODBAHIS,
    odbst1 As Focas1.ODBST, almstat As Boolean,
    caseSwitch As Short, alm As String,
    Err As Short, closeStat As Boolean,
    timecnc As Focas1.IODBTIMER, datecnc As Focas1.IODBTIMER, almcomparison As String,
    emgstat As Boolean, LoadIni As String(),
    numprog As Focas1.ODBPRO, seqnum As Focas1.ODBSEQ,
    numvar As String, macronum As Focas1.ODBM, station_Id As String, host As String, port As Int32, portStatus As Int32,
    port_local_to_server As Int32,
    client As TcpClient, ServerConnect As Boolean,
    t_status As String, num_cmd As Short, odbcommand As Focas1.ODBCMD, plasma_on As Boolean,
    odbsp As Focas1.ODBSPEED, time_speed_null As Int32, speed_null_bool As Boolean, stop_speed As Boolean,
    debugOnOff As Boolean, errorToServer As Boolean

    Function new_coding(ByVal str As String) As String
        Dim charw As Char,
            newstr As String = ""

        For i = 1 To Len(str)
            charw = Mid(str, i, 1)

            Select Case charw
                Case "Ж"
                    newstr += "Ё"
                Case "З"
                    newstr += "Ж"
                Case "И"
                    newstr += "З"
                Case "Й"
                    newstr += "И"
                Case "К"
                    newstr += "Й"
                Case "Л"
                    newstr += "К"
                Case "М"
                    newstr += "Л"
                Case "Н"
                    newstr += "М"
                Case "О"
                    newstr += "Н"
                Case "П"
                    newstr += "О"
                Case "Р"
                    newstr += "П"
                Case "С"
                    newstr += "Р"
                Case "Т"
                    newstr += "С"
                Case "У"
                    newstr += "Т"
                Case "Ф"
                    newstr += "У"
                Case "Х"
                    newstr += "Ф"
                Case "Ц"
                    newstr += "Х"
                Case "Ч"
                    newstr += "Ц"
                Case "Ш"
                    newstr += "Ч"
                Case "Щ"
                    newstr += "Ш"
                Case "Ъ"
                    newstr += "Щ"
                Case "Ы"
                    newstr += "Ъ"
                Case "Ь"
                    newstr += "Ы"
                Case "Э"
                    newstr += "Ь"
                Case "Ю"
                    newstr += "Э"
                Case "Я"
                    newstr += "Ю"
                Case "а"
                    newstr += "Я"
                Case Else
                    newstr += charw
            End Select
        Next

        Return newstr
    End Function 'перекодирование сообщений запрашиваемых с чпу(т.к. отсутствует буква Ё)
    Private Function ping_f() As Boolean
        Dim _Ping As Ping, reply As PingReply
        Try
            _Ping = New Ping()
            reply = _Ping.Send(host, 500)

            If (reply.Status = IPStatus.Success) Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function 'функция пинга сервера
    Sub ProgrammLog(ByVal str As String)
        Dim myFileInfo As New FileInfo("LOG\LogProg.txt"), dateNow As Date

        If myFileInfo.Length < 15000000 Then
            Using file As New StreamWriter("LOG\LogProg.txt", True)
                file.WriteLine(str)
                file.Close()
                file.Dispose()
            End Using
        Else
            dateNow = Today
            My.Computer.FileSystem.RenameFile("LOG\LogProg.txt", "LogProg_" + dateNow.Day.ToString() + dateNow.Month.ToString() + dateNow.Year.ToString + ".txt")
            Using file As New StreamWriter("LOG\LogProg.txt", True)
                file.WriteLine(str)
                file.Close()
                file.Dispose()
            End Using
        End If
    End Sub
    Private Sub SpeedNull_Tick(sender As Object, e As EventArgs) Handles SpeedNull.Tick
        time_speed_null += time_speed_null
    End Sub
    Private Sub Main_T_Tick(sender As Object, e As EventArgs) Handles Main_T.Tick
        Dim e_no, lenght As Integer, time As Int32, speed_0 As Boolean
        speed_0 = False
        Main_T.Enabled = False

        Try
            If debugOnOff <> True Then
                Err = Focas1.cnc_allclibhndl(FlibHndl) ' получение хэндла

                If (Err <> 0) Then
                    ProgrammLog("Error: " + Convert.ToString(Err))
                End If
                errinfo = Focas1.cnc_statinfo(FlibHndl, odbst1) ' получение информации о состоянии станка

                If (errinfo <> 0) Then
                    ProgrammLog(Convert.ToString(errinfo) + " " + Convert.ToString(FlibHndl))
                End If

                caseSwitch = odbst1.run
                ToolStripStatusLabel2.Text = "Case: " + Convert.ToString(caseSwitch) + " ErrorInfo: " + Convert.ToString(errinfo)

                timecnc.type = 1
                Focas1.cnc_gettimer(FlibHndl, timecnc) 'время с контроллера ЧПУ
                datecnc.type = 0
                Focas1.cnc_gettimer(FlibHndl, datecnc) 'дата с контроллера ЧПУ

                Focas1.cnc_rdprgnum(FlibHndl, numprog) ' номер программы  

                '********************************************************
                Focas1.cnc_rdspeed(FlibHndl, 0, odbsp) 'Запрос скорости подачи
                '********************************************************
            Else
                caseSwitch = CInt(Int((4 * Rnd()) + 1))
                numprog.data = 1
                numprog.mdata = 1
            End If
            '********************************************************
            'num_cmd = 1
            'Focas1.cnc_rdcommand(FlibHndl, 112, 1, num_cmd, odbcommand) ' запрос М команд

            Select Case caseSwitch
                Case 0
                    ToolStripStatusLabel1.Text = "STOP"
                    If (_stop = False) Then
                        time = Convert.ToInt32((DateTime.UtcNow - New DateTime(1970, 1, 1)).TotalSeconds)
                        ServerOrLocal(station_Id, time, "Stop cnc program", numprog, "*")
                        t_status = "2"
                        ServerOrLocal(station_Id, time, " ", numprog, t_status)
                        _stop = True
                        hold = False
                        start = False
                        speed_null_bool = False
                        stop_speed = False
                        SpeedNull.Enabled = False
                        time_speed_null = 0
                    End If
                Case 1
                    ToolStripStatusLabel1.Text = "STOP"
                    If (_stop = False) Then
                        time = Convert.ToInt32((DateTime.UtcNow - New DateTime(1970, 1, 1)).TotalSeconds)
                        ServerOrLocal(station_Id, time, "Stop cnc program", numprog, "*")
                        t_status = "2"
                        ServerOrLocal(station_Id, time, " ", numprog, t_status)
                        _stop = True
                        hold = False
                        start = False
                        speed_null_bool = False
                        stop_speed = False
                        SpeedNull.Enabled = False
                        time_speed_null = 0
                    End If
                Case 2
                    ToolStripStatusLabel1.Text = "HOLD"
                    If (hold = False And _stop = False) Then
                        time = Convert.ToInt32((DateTime.UtcNow - New DateTime(1970, 1, 1)).TotalSeconds)
                        ServerOrLocal(station_Id, time, "Hold cnc program", numprog, "*")
                        t_status = "2"
                        ServerOrLocal(station_Id, time, " ", numprog, t_status)
                        hold = True
                        _stop = False
                        start = False
                        speed_null_bool = False
                        stop_speed = False
                        SpeedNull.Enabled = False
                        time_speed_null = 0
                    End If
                Case 3
                    If debugOnOff <> True Then
                        If odbsp.actf.data = 0 And Not speed_null_bool Then ' если скорость перемещения равна 0, счетать это простоем
                            If SpeedNull.Enabled = False Then
                                SpeedNull.Enabled = True
                            End If
                        Else
                            If SpeedNull.Enabled = True Then
                                SpeedNull.Enabled = False
                                time_speed_null = 0
                            End If
                        End If

                        If odbsp.actf.data <> 0 And speed_null_bool Then 'если скорость не равна нулю сбросить флаг нулевой скорости
                            speed_null_bool = False
                        End If

                        If time_speed_null > 300 Then  ' после выставления скорости в ноль прошло пять минут 
                            speed_null_bool = True
                        End If

                        If speed_null_bool Then
                            ToolStripStatusLabel1.Text = "STOP"
                            If stop_speed = False Then
                                time = Convert.ToInt32((DateTime.UtcNow - New DateTime(1970, 1, 1)).TotalSeconds)
                                ServerOrLocal(station_Id, time, "Stop cnc program", numprog, "*")
                                t_status = "2"
                                ServerOrLocal(station_Id, time, " ", numprog, t_status)
                                stop_speed = True
                            End If
                        Else
                            If stop_speed Then
                                stop_speed = False
                                ToolStripStatusLabel1.Text = "START"
                                start = True
                            End If
                        End If
                    Else
                        ToolStripStatusLabel1.Text = "START"
                    End If
                    If (start = False) Then
                        time = Convert.ToInt32((DateTime.UtcNow - New DateTime(1970, 1, 1)).TotalSeconds)
                        ServerOrLocal(station_Id, time, "Start cnc program", numprog, "*")
                        t_status = "1"
                        ServerOrLocal(station_Id, time, " ", numprog, t_status)
                        start = True
                        hold = False
                        _stop = False
                    End If
            End Select
            If debugOnOff <> True Then
                If (odbst1.alarm = 1) Then
                    If (almstat = False) Then

                        Focas1.cnc_stopophis(FlibHndl) 'остановить сбор данных по ошибкам
                        Focas1.cnc_rdalmhisno(FlibHndl, e_no)
                        lenght = 54 * e_no

                        Focas1.cnc_rdalmhistry(FlibHndl, 2, e_no, lenght, odbah) 'получение данных по ошибкам
                        Focas1.cnc_startophis(FlibHndl) 'запустить сбор данных

                        If ((odbah.alm_his.data50.hour + odbah.alm_his.data50.minute) >= (odbah.alm_his.data49.hour + odbah.alm_his.data49.minute)) Then
                            alm = new_coding(Convert.ToString(odbah.alm_his.data50.day) + "." + Convert.ToString(odbah.alm_his.data50.month) +
                                "." + Convert.ToString(odbah.alm_his.data50.year) + " " + Convert.ToString(odbah.alm_his.data50.hour) + ":" + Convert.ToString(odbah.alm_his.data50.minute) + ":" +
                               Convert.ToString(odbah.alm_his.data50.second) + " :[" + Convert.ToString(odbah.alm_his.data50.alm_no) + "]  " + odbah.alm_his.data50.alm_msg)

                            If almcomparison <> alm Then
                                almcomparison = alm
                                time = Convert.ToInt32((DateTime.UtcNow - New DateTime(1970, 1, 1)).TotalSeconds)
                                ServerOrLocal(station_Id, time, alm, numprog, "*")
                                t_status = "10"
                                ServerOrLocal(station_Id, time, " ", numprog, t_status)
                            Else
                                time = Convert.ToInt32((DateTime.UtcNow - New DateTime(1970, 1, 1)).TotalSeconds)
                                ServerOrLocal(station_Id, time, "No message", numprog, "*")
                                t_status = "10"
                                ServerOrLocal(station_Id, time, " ", numprog, t_status)
                            End If

                        ElseIf ((odbah.alm_his.data49.hour + odbah.alm_his.data49.minute) >= (odbah.alm_his.data48.hour + odbah.alm_his.data48.minute)) Then
                            alm = new_coding(Convert.ToString(odbah.alm_his.data49.day) + "." + Convert.ToString(odbah.alm_his.data49.month) +
                           "." + Convert.ToString(odbah.alm_his.data49.year) + " " + Convert.ToString(odbah.alm_his.data49.hour) + ":" + Convert.ToString(odbah.alm_his.data49.minute) + ":" +
                          Convert.ToString(odbah.alm_his.data49.second) + " :[" + Convert.ToString(odbah.alm_his.data49.alm_no) + "]  " + odbah.alm_his.data49.alm_msg)

                            If almcomparison <> alm Then
                                almcomparison = alm
                                time = Convert.ToInt32((DateTime.UtcNow - New DateTime(1970, 1, 1)).TotalSeconds)
                                ServerOrLocal(station_Id, time, alm, numprog, "*")
                                t_status = "10"
                                ServerOrLocal(station_Id, time, " ", numprog, t_status)
                            Else
                                time = Convert.ToInt32((DateTime.UtcNow - New DateTime(1970, 1, 1)).TotalSeconds)
                                ServerOrLocal(station_Id, time, "No message", numprog, "*")
                                t_status = "10"
                                ServerOrLocal(station_Id, time, " ", numprog, t_status)
                            End If

                        Else
                            alm = new_coding(Convert.ToString(odbah.alm_his.data48.day) + "." + Convert.ToString(odbah.alm_his.data48.month) +
                          "." + Convert.ToString(odbah.alm_his.data48.year) + " " + Convert.ToString(odbah.alm_his.data48.hour) + ":" + Convert.ToString(odbah.alm_his.data48.minute) + ":" +
                         Convert.ToString(odbah.alm_his.data48.second) + " :[" + Convert.ToString(odbah.alm_his.data48.alm_no) + "]  " + odbah.alm_his.data48.alm_msg)

                            If almcomparison <> alm Then
                                almcomparison = alm
                                time = Convert.ToInt32((DateTime.UtcNow - New DateTime(1970, 1, 1)).TotalSeconds)
                                ServerOrLocal(station_Id, time, alm, numprog, "*")
                                t_status = "10"
                                ServerOrLocal(station_Id, time, " ", numprog, t_status)
                            Else
                                time = Convert.ToInt32((DateTime.UtcNow - New DateTime(1970, 1, 1)).TotalSeconds)
                                ServerOrLocal(station_Id, time, "No message", numprog, "*")
                                t_status = "10"
                                ServerOrLocal(station_Id, time, " ", numprog, t_status)
                            End If

                        End If
                        almstat = True
                    End If
                Else
                    almstat = False
                End If

                If odbst1.emergency = 1 Then
                    If emgstat = False Then
                        emgstat = True
                        time = Convert.ToInt32((DateTime.UtcNow - New DateTime(1970, 1, 1)).TotalSeconds)
                        ServerOrLocal(station_Id, time, "Аварийный останов", numprog, "*")
                        t_status = "10"
                        ServerOrLocal(station_Id, time, " ", numprog, t_status)
                    End If
                Else
                    emgstat = False
                End If
            End If
        Catch ex As Exception
            ProgrammLog(ex.Message + " - " + Convert.ToString(DateTime.Now) + " - Main_T_Tick" + " " + t_status)
        Finally
            If debugOnOff <> True Then
                Focas1.cnc_freelibhndl(FlibHndl)
            End If
            Main_T.Dispose()
            Main_T.Enabled = True
        End Try
    End Sub 'основной таймер в нем мы получаем состояние станка 
    Private Sub State_Plasma_T_Tick(sender As Object, e As EventArgs) Handles State_Plasma_T.Tick
        Dim Err As Short, time As Int32
        State_Plasma_T.Enabled = False
        If debugOnOff <> True Then
            Try
                Err = Focas1.cnc_allclibhndl(FlibHndl2) ' получение хэндла
                If (Err <> 0) Then
                    ProgrammLog("Error: " + Convert.ToString(Err))
                End If

                '********************************************************
                num_cmd = 1
                Focas1.cnc_rdcommand(FlibHndl2, 112, 1, num_cmd, odbcommand) ' запрос М команд
                '********************************************************

                If (odbcommand.cmd0.cmd_val = 17) Then
                    t_status = "3"
                    time = Convert.ToInt32((DateTime.UtcNow - New DateTime(1970, 1, 1)).TotalSeconds)
                    ServerOrLocal(station_Id, time, " ", numprog, t_status)
                End If

                If (odbcommand.cmd0.cmd_val = 18) Then
                    t_status = "1"
                    time = Convert.ToInt32((DateTime.UtcNow - New DateTime(1970, 1, 1)).TotalSeconds)
                    ServerOrLocal(station_Id, time, " ", numprog, t_status)
                End If

            Catch ex As Exception
                ProgrammLog(ex.Message + " - " + Convert.ToString(DateTime.Now) + " - State_Plasma_T" + " " + t_status)
            Finally
                Focas1.cnc_freelibhndl(FlibHndl2)
                State_Plasma_T.Dispose()
                State_Plasma_T.Enabled = True
            End Try
        End If
    End Sub
    Private Sub Status_Tick(sender As Object, e As EventArgs) Handles Status.Tick
        Dim time As Int32
        Status.Enabled = False
        Try
            time = Convert.ToInt32((DateTime.UtcNow - New DateTime(1970, 1, 1)).TotalSeconds)
            If t_status <> "" Then
                ServerOrLocal(station_Id, time, " ", numprog, t_status)
            End If
        Catch ex As Exception
            ProgrammLog(ex.Message + " - " + Convert.ToString(DateTime.Now) + " - Status_Tick")
        Finally
            Status.Enabled = True
        End Try
    End Sub 'отправляет на сервер своё состояние
    Private Function Send_To_Server(data As String, host As String, port As Int32) As Boolean
        Dim sw As StreamWriter, client As TcpClient, stateConn As IAsyncResult, timeOut As Boolean
        Try
            client = New TcpClient()
            'client.Connect(host, port)
            stateConn = client.BeginConnect(host, port, Nothing, Nothing)
            timeOut = stateConn.AsyncWaitHandle.WaitOne(TimeSpan.FromMilliseconds(500))

            If timeOut Then
                errorToServer = False
                sw = New StreamWriter(client.GetStream())
                sw.AutoFlush = True
                sw.WriteLine(data) 'отсылаем данные серверу
            Else
                client.Close()
                Return False
            End If
            client.Close()
            Return True
        Catch ex As Exception
            If errorToServer <> True Then
                ProgrammLog(ex.Message + " - " + Convert.ToString(DateTime.Now) + " - Send_To_Server")
                errorToServer = True
            End If
            Return False
        End Try
    End Function 'отправка данных серверу
    Private Sub InsertMainTable(station_id As String, time As Int32, text As String, NumProg As String)
        Dim dbConnection As SQLiteConnection, sqlCommand As SQLiteCommand
        Try
            dbConnection = New SQLiteConnection("Data Source=" + Application.StartupPath + "\Statistics.sqlite ;Version=3;")
            dbConnection.Open()
            sqlCommand = New SQLiteCommand()
            sqlCommand.Connection = dbConnection

            sqlCommand.CommandText = "INSERT INTO line_data([station_id], [time], [text], [ProgNum]) VALUES('" + station_id + "'," + time.ToString() + ", '" + text + "', '" + NumProg + "')"
            sqlCommand.ExecuteNonQuery()

        Catch ee As Exception
            ProgrammLog(ee.Message + " - " + Convert.ToString(DateTime.Now) + " - InsertMainTable")
        Finally
            If dbConnection IsNot Nothing Then
                dbConnection.Close()
            End If
        End Try
    End Sub 'добавление в локальную базу данных основных данных
    Private Sub InsertStatusTable(station_id As String, time As Int32, status As String)
        Dim dbConnection As SQLiteConnection, sqlCommand As SQLiteCommand
        Try
            dbConnection = New SQLiteConnection("Data Source=" + Application.StartupPath + "\Statistics.sqlite ;Version=3;")
            dbConnection.Open()
            sqlCommand = New SQLiteCommand()
            sqlCommand.Connection = dbConnection

            sqlCommand.CommandText = "INSERT INTO status([station_id], [time], [status]) VALUES('" + station_id + "'," + time.ToString() + ", '" + status + "')"
            sqlCommand.ExecuteNonQuery()

        Catch ee As Exception
            ProgrammLog(ee.Message + " - " + Convert.ToString(DateTime.Now) + " - InsertStatusTable")
        Finally
            If dbConnection IsNot Nothing Then
                dbConnection.Close()
            End If
        End Try
    End Sub 'добавление в локальную базу статуса
    Private Sub LocalDBToServerDB()
        Dim dbConnection As SQLiteConnection, sqlCommand As SQLiteCommand, countTable As Int32
        Try
            dbConnection = New SQLiteConnection("Data Source=" + Application.StartupPath + "\Statistics.sqlite ;Version=3;")
            dbConnection.Open()
            sqlCommand = New SQLiteCommand()
            sqlCommand.Connection = dbConnection

            sqlCommand.CommandText = "SELECT count(id) FROM  line_data"
            Using dr As IDataReader = sqlCommand.ExecuteReader()
                If dr.Read() Then
                    countTable = dr.GetInt32(0)
                End If
            End Using

            If countTable > 0 Then
                sqlCommand.CommandText = "SELECT [id], [station_id], [time], [text], [ProgNum] FROM  line_data"
                Using dr As SQLiteDataReader = sqlCommand.ExecuteReader()
                    While dr.Read()
                        Send_To_Server(dr.GetValue(1).ToString() + "," + dr.GetValue(2).ToString() + "," + dr.GetValue(3).ToString() + "," + dr.GetValue(4).ToString(), host, port_local_to_server)
                    End While
                    CliarLocalDB("line_data")
                End Using

                sqlCommand.CommandText = "SELECT [station_id], [time], [status] FROM  status"
                Using dr As SQLiteDataReader = sqlCommand.ExecuteReader()
                    If dr.StepCount > 0 Then
                        While dr.Read()
                            Send_To_Server(dr.GetValue(0).ToString() + "," + dr.GetValue(1).ToString() + "," + dr.GetValue(2).ToString(), host, port_local_to_server)
                        End While
                    End If
                    CliarLocalDB("status")
                End Using
            End If
        Catch ee As Exception
            ProgrammLog(ee.Message + " - " + Convert.ToString(DateTime.Now) + " - LocalDBToServerDB")
        Finally
            If dbConnection IsNot Nothing Then
                dbConnection.Close()
            End If
        End Try

    End Sub ' после того как появляется связь с сервером происходит передача данных с локальной базы в серверную базу
    Private Sub CliarLocalDB(tableName As String)
        Dim dbConnection As SQLiteConnection, sqlCommand As SQLiteCommand
        Try
            dbConnection = New SQLiteConnection("Data Source=" + Application.StartupPath + "\Statistics.sqlite ;Version=3;")
            dbConnection.Open()
            sqlCommand = New SQLiteCommand()
            sqlCommand.Connection = dbConnection

            sqlCommand.CommandText = "DELETE FROM " + tableName
            sqlCommand.ExecuteNonQuery()
            sqlCommand.CommandText = "UPDATE SQLITE_SEQUENCE SET SEQ=0"
            sqlCommand.ExecuteNonQuery()

        Catch ee As Exception
            ProgrammLog(ee.Message + " - " + Convert.ToString(DateTime.Now) + " - CliarLocalDB")
        Finally
            If dbConnection IsNot Nothing Then
                dbConnection.Close()
            End If
        End Try
    End Sub ' удаляет данные с локальной базы данных после передачи их на сервер
    Private Sub ServerOrLocal(station_Id As String, time As Int32, text As String, numprog As Focas1.ODBPRO, status As String)
        Dim str1 As String, stateSend As Boolean ', tempStr(2) As String

        Try
            If (ping_f()) Then
                If status <> "*" Then
                    stateSend = Send_To_Server(station_Id + "," + Convert.ToString(time) + "," + status, host, portStatus)
                    If stateSend = False Then
                        InsertStatusTable(station_Id, time, status)
                        ToolStripStatusLabel3.Text = "Server disconnected"
                    Else
                        ToolStripStatusLabel3.Text = "Server connected"
                    End If
                    ToolStripStatusLabel4.Text = status
                Else
                    str1 = station_Id + "," + Convert.ToString(time) + "," + text + "," + Convert.ToString(numprog.mdata) + "*" + Convert.ToString(numprog.data)
                    stateSend = Send_To_Server(str1, host, port)
                    If stateSend = False Then
                        InsertMainTable(station_Id, time, text, Convert.ToString(numprog.mdata) + "*" +
                         Convert.ToString(numprog.data))
                        ToolStripStatusLabel3.Text = "Server disconnected"
                    Else
                        LocalDBToServerDB()
                        ToolStripStatusLabel3.Text = "Server connected"
                    End If
                End If
            Else
                ToolStripStatusLabel3.Text = "Server disconnected"
                If status <> "*" Then
                    InsertStatusTable(station_Id, time, status)
                Else
                    InsertMainTable(station_Id, time, text, Convert.ToString(numprog.mdata) + "*" +
                    Convert.ToString(numprog.data))
                End If
            End If
        Catch ex As Exception
            ProgrammLog(ex.Message + " - " + Convert.ToString(DateTime.Now) + " - Ping_Tick")
        End Try

    End Sub ' отправка данных на сервер или в локальную базу данных

    Private Sub Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim time As Int32
        start = False
        _stop = False
        stop_speed = False
        hold = False
        almstat = False
        emgstat = False
        count = 0
        errinfo = 0
        erralm = 0
        closeStat = False
        num_cmd = 1
        plasma_on = False
        errorToServer = False

        debugOnOff = False   'Для тестов. В состоянии true, игнорирует библиотеку FOCAS

        Try
            LoadIni = File.ReadAllLines("IniFile.ini", Encoding.UTF8)
            station_Id = LoadIni(0)
            host = LoadIni(1)
            port = Convert.ToInt32(LoadIni(2))
            portStatus = Convert.ToInt32(LoadIni(3))
            port_local_to_server = Convert.ToInt32(LoadIni(4))
            Status.Interval = Convert.ToInt32(LoadIni(5))

            Me.WindowState = FormWindowState.Minimized
            If (Me.WindowState = FormWindowState.Minimized) Then
                Me.ShowInTaskbar = False
                NotifyIcon1.Visible = True
            End If
            ServerConnect = ping_f()
            If debugOnOff <> True Then
                Err = Focas1.cnc_allclibhndl(FlibHndl) ' получение хэндла
                If (Err = 0) Then
                    Focas1.cnc_rdprgnum(FlibHndl, numprog) ' номер программы   
                    time = Convert.ToInt32((DateTime.UtcNow - New DateTime(1970, 1, 1)).TotalSeconds)
                    ServerOrLocal(station_Id, time, "StartWhineyProgram", numprog, "*")
                End If
                Focas1.cnc_freelibhndl(FlibHndl)
            End If
        Catch ex As Exception
            ProgrammLog(ex.Message + " - " + Convert.ToString(DateTime.Now) + " - Start")
            ToolStripStatusLabel3.Text = "Server disconnected"
            ServerConnect = False
        Finally
            HideFromAltTab(Me.Handle)
            Main_T.Enabled = True
            Status.Enabled = True
            State_Plasma_T.Enabled = True
            SpeedNull.Enabled = True
        End Try
    End Sub 'ЗАГРУЗКА ФОРМЫ
    Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Dim tempStr(1) As String
        Dim style = MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton2 Or
            MsgBoxStyle.Critical
        Dim result As MsgBoxResult = MsgBox("Вы уверены, что хотите закрыть приложение???", style, "Warning")

        If (result = MsgBoxResult.Yes) And My.Computer.Keyboard.CtrlKeyDown Then
            ProgrammLog("Program closing")
            'tempStr(0) = Convert.ToString(countNoConn(0))
            ' tempStr(1) = Convert.ToString(countNoConn(1))
            ' File.WriteAllLines("saveLocalCount.txt", tempStr)
            Dispose()
        Else
            e.Cancel = True
        End If
    End Sub 'обработчик закрытия формы
    Private Sub StripMenu_Close_Form(sender As Object, e As EventArgs) Handles ContMenuCloseForm.Click, StripMenuCloseForm.Click
        Close()
    End Sub
    Private Sub ContMenuExpandForm_Click(sender As Object, e As EventArgs) Handles ContMenuExpandForm.Click
        If (Me.WindowState = FormWindowState.Minimized) Then
            Me.WindowState = FormWindowState.Normal
            NotifyIcon1.Visible = False
            Me.ShowInTaskbar = True
        End If
    End Sub 'развернуть из трэя
    Sub InTrey()
        Me.WindowState = FormWindowState.Minimized
        If (Me.WindowState = FormWindowState.Minimized) Then
            Me.ShowInTaskbar = False
            NotifyIcon1.Visible = True
        End If
        HideFromAltTab(Me.Handle)
    End Sub
    Private Sub StripMenuTurnForm_Click(sender As Object, e As EventArgs) Handles StripMenuTurnForm.Click
        InTrey()
    End Sub 'свернуть в трэй
End Class