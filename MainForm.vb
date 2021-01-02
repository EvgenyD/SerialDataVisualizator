Imports System

Public Class MainForm
    Dim sys_version As String = "Serial data visualization v0.7"
    Dim max_number_of_parameters As Integer = 16
    Dim min_size As New Point(480, 240)
    Dim list_of_serial_rates() As String = {"1200", "2400", "4800", "9600", "14400", "19200", "28800", "38400", "57600", "115200", "230400"}
    Dim sys_del As String = ","
    Dim sys_delims() As Char = {" ", ":", ";", ",", "/", vbTab, "_", "=", ">"}
    Dim sys_plotsize As Integer = 1000

    Dim min_update_time As Integer = 2000
    Dim ready_to_rescale As Boolean = True

    Dim tttext As ToolTip

    'add logging functionality
    Dim log_data_cb As CheckBox
    Dim ListOfFiles As ListBox
    Dim file_name As String = ""
    Dim file_path As String = Application.StartupPath
    Dim No_logging As Boolean
    Dim Log_Counter As Integer = 1
    Dim LastEntered_Name As String = "None"
    Dim New_file_bt As Button

    'interface objects
    ' containers
    Dim left_panel As Panel
    Dim top_panel As Panel

    'com ports
    Dim com_name_tb As ListBox
    Dim com_rfrsh As Button
    Dim com_rate_list As ComboBox
    Dim active_serial As New System.IO.Ports.SerialPort

    Dim output_buffer_tb As TextBox
    Dim input_buffer_tb As TextBox

    Dim output_msg_send_bt As Button
    Dim output_msg_rpt_cb As CheckBox
    Dim output_msg_period_nb As NumericUpDown
    Dim ClearPlot_bt As Button

    Dim param_start_tb As TextBox


    Dim ChannelShow_CB() As CheckBox
    ReadOnly ChannelColors() As Color = {Color.Black, Color.Red, Color.Green, Color.Blue, Color.DarkCyan, Color.Magenta, Color.Brown, Color.Gray, Color.Orange, Color.BlueViolet}
    ReadOnly ChannelColorsCompl() As Color = {Color.Gray, Color.Black, Color.Black, Color.Black, Color.Black, Color.Black, Color.Black, Color.Black, Color.Black, Color.Black}
    ReadOnly ChannelNames() As String = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "1", "2", "3", "4"}

    Dim ActiveChannels As Integer = 0
    Dim Current_Point As Integer = 0

    Dim MainPlot As DataVisualization.Charting.Chart
    Dim Annotes(30) As DataVisualization.Charting.ArrowAnnotation
    Dim NChannels As Integer = 30
    Dim RBFR() As Byte
    Dim BufferString As String
    Dim ReceivedMsg As String
    Dim IsDataReady As Boolean
    Dim LostPackets As Integer

    Dim receivedData() As String

    Dim output_msg_struct As String
    Dim input_msg_struct As String

    Dim RescaleTimer As Timer
    Dim UpdatePackage As Timer

    Dim SendAgainMessage_Timer As Timer

    Dim Can_try_again As Boolean

    Dim set_last_baudrate As String = "115200"
    Dim set_last_delimeter As String = ":"
    Dim set_last_startingchar As String = ">>"
    Dim set_last_visible As String = "ynnnnnnnnnnnnnnnnnnnnnnnnnnnnnn"



    Dim sys_step As Integer = 30
    Dim sys_offset As Integer = (MyBase.Height - MyBase.ClientSize.Height)
    Private Sub Scan_files_in_folder()
        Dim _all_files As New List(Of String)
        Dim ii As Integer

        ' Make sure LOG file exists
        file_path = Application.StartupPath + IO.Path.DirectorySeparatorChar + "Logs" + IO.Path.DirectorySeparatorChar
        If Not (IO.Directory.Exists(file_path)) Then
            Try
                Dim di As IO.DirectoryInfo = IO.Directory.CreateDirectory(file_path)
                MsgBox("Logs folder was created: \n" + file_path)
                No_logging = False
            Catch ex As Exception
                MsgBox("Could not create Logs folder: \n" + file_path)
                No_logging = True
            End Try

        End If


        If IO.Directory.Exists(file_path) Then
            Try
                For Each _file As String In IO.Directory.EnumerateFiles(file_path, "exp*.txt")
                    _all_files.Add(IO.Path.GetFileName(_file))
                Next
                _all_files.Sort()
                For ii = _all_files.Count - 1 To 0 Step -1
                    ListOfFiles.Items.Add(_all_files(ii).ToString)
                Next

            Catch ex As Exception

            End Try

        End If

    End Sub

    Private Sub Log_Changed(sender As Object, e As EventArgs)
        If log_data_cb.Checked Then
            Generate_new_file(New_file_bt, EventArgs.Empty)

        End If
    End Sub

    Private Sub CopyMsg(sender As Object, e As EventArgs)
        Dim tp As String = sender.GetType().Name
        If (tp = "TextBox") Then
            If (sender.text > "") Then
                System.Windows.Forms.Clipboard.SetText(sender.text)
                sender.text = "Copied"
            End If
        End If
    End Sub
    Private Sub Generate_new_file(sender As Object, e As EventArgs)
        If No_logging Then Exit Sub

        Dim localDate = DateTime.Now

        Dim board As String = ""

        file_name = "exp" + localDate.Year.ToString("00") + "" + localDate.Month.ToString("00") + "" + localDate.Day.ToString("00") + " " + localDate.Hour.ToString("00") + "-" + localDate.Minute.ToString("00") + "-" + localDate.Second.ToString("00") + ".txt"

        '' creates new file if needed, and starts writing to it
        Using outputFile As New IO.StreamWriter(file_path + IO.Path.DirectorySeparatorChar + file_name, True)
            outputFile.WriteLine("starting time:" + localDate)
        End Using

        Dim file_listed As Boolean = False

        If (ListOfFiles.Items.Count > 0) Then
            If StrComp(ListOfFiles.Items(0).ToString, file_name, CompareMethod.Text) <> 0 Then
                ListOfFiles.Items.Insert(0, file_name)
                ListOfFiles.SelectedIndex = 0

            End If
        Else
            ListOfFiles.Items.Insert(0, file_name)
            ListOfFiles.SelectedIndex = 0
        End If



        Clear_plots(MainPlot, EventArgs.Empty)

    End Sub
    Private Sub ListOfFiles_DoubleClick(sender As Object, e As EventArgs)
        Dim _fname As String
        _fname = ListOfFiles.SelectedItem.ToString
        Dim full_name As String
        full_name = file_path + IO.Path.DirectorySeparatorChar + _fname

        If (IO.File.Exists(full_name)) Then

            Process.Start(full_name)
        End If


    End Sub
    Private Sub ListOfFiles_MouseEnter(sender As Object, e As EventArgs)
        ListOfFiles.Height = (ClientSize.Height \ 2)
    End Sub
    Private Sub ListOfFiles_MouseLeave(sender As Object, e As EventArgs)
        ListOfFiles.Height = 43
    End Sub

    Private Sub ListOfFiles_MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs)
        Dim full_name As String
        full_name = file_path + IO.Path.DirectorySeparatorChar

        If (e.Button = MouseButtons.Right) Then
            If (IO.Directory.Exists(full_name)) Then

                Process.Start(full_name)
            End If
        End If
    End Sub


    Private Sub Build_Gui()
        Dim RData(sys_plotsize) As Single
        Dim Tdata(sys_plotsize) As Single

        tttext = New ToolTip

        left_panel = New Panel With {
             .Location = New Point(2, 2),
            .Size = New Point(150, MyBase.ClientSize.Height - 4),
            .Anchor = (AnchorStyles.Top + AnchorStyles.Left + AnchorStyles.Bottom)
        }

        top_panel = New Panel With {
            .Location = New Point(155, 2),
            .Size = New Point(MyBase.ClientSize.Width - 150 - 5, 44),
            .Anchor = (AnchorStyles.Top + AnchorStyles.Left + AnchorStyles.Right)
        }

        com_rate_list = New ComboBox With {
            .Size = New Point(left_panel.Width / 2, 25),
            .Location = New Point(1, 1),
            .Font = New Font("Consolas", 10),
            .Padding = New Padding(2)
        }

        com_rate_list.Items.AddRange(list_of_serial_rates)
        If com_rate_list.Items.Contains(set_last_baudrate) Then
            com_rate_list.SelectedItem = set_last_baudrate
        Else
            com_rate_list.Text = set_last_baudrate
        End If
        com_rate_list.SelectedItem = set_last_baudrate

        com_rfrsh = New Button With {
            .Text = "Refresh",
            .Size = New Point(left_panel.Width / 2 - 2, com_rate_list.Height),
            .Location = New Point(com_rate_list.Location.X + com_rate_list.Width, com_rate_list.Location.Y)
        }
        AddHandler com_rfrsh.Click, AddressOf COM_refreshB_Click

        com_name_tb = New ListBox With {
            .Anchor = AnchorStyles.Top + AnchorStyles.Left,
            .Size = New Point(left_panel.Width - 2, 60),
            .Location = New Point(1, com_rate_list.Height + com_rate_list.Location.Y),
            .IntegralHeight = False
        }
        AddHandler com_name_tb.DoubleClick, AddressOf com_name_tb_DoubleClick

        left_panel.Controls.Add(com_name_tb)
        left_panel.Controls.Add(com_rfrsh)
        left_panel.Controls.Add(com_rate_list)


        param_start_tb = New TextBox With {
            .Font = New Font("Consolas", 9),
            .Text = set_last_startingchar,
            .Size = New Point(49, 22),
            .Location = New Point(0, 0),
            .Anchor = AnchorStyles.Top + AnchorStyles.Left,
            .AutoSize = False
        }
        top_panel.Controls.Add(param_start_tb)
        AddHandler param_start_tb.TextChanged, AddressOf updateMsgStart

        input_buffer_tb = New TextBox With {
            .Font = New Font("Consolas", 9),
            .Size = New Point(top_panel.Width - param_start_tb.Location.X - param_start_tb.Width - 1, param_start_tb.Height), 'output_buffer_tb.Width + 50, 25),
            .Location = New Point(param_start_tb.Location.X + param_start_tb.Width, 0),
            .Anchor = AnchorStyles.Top + AnchorStyles.Right + AnchorStyles.Left,
            .ReadOnly = True,
            .AutoSize = False
        }
        top_panel.Controls.Add(input_buffer_tb)
        AddHandler input_buffer_tb.Click, AddressOf CopyMsg

        output_buffer_tb = New TextBox With {
            .Font = New Font("Consolas", 10),
            .Size = New Point(top_panel.Width - 200, 23),
            .Location = New Point(param_start_tb.Location.X, param_start_tb.Location.Y + param_start_tb.Height + 1),
            .Anchor = AnchorStyles.Top + AnchorStyles.Right + AnchorStyles.Left,
            .AutoSize = False
        }
        top_panel.Controls.Add(output_buffer_tb)

        output_msg_send_bt = New Button With {
            .Font = New Font("Consolas", 9),
            .Text = "SEND",
            .Size = New Point(50, 22),
            .Location = New Point(output_buffer_tb.Location.X + output_buffer_tb.Width, output_buffer_tb.Location.Y - 1),
            .Anchor = AnchorStyles.Top + AnchorStyles.Right,
            .AutoSize = False
        }
        AddHandler output_msg_send_bt.Click, AddressOf send_new_msg

        output_msg_rpt_cb = New CheckBox With {
            .Text = "Every",
            .Checked = False,
            .Size = New Point(54, output_msg_send_bt.Height),
            .Padding = New Padding(0),
            .Location = New Point(output_msg_send_bt.Location.X + output_msg_send_bt.Width + 2, output_msg_send_bt.Location.Y),
            .Anchor = AnchorStyles.Top + AnchorStyles.Right
            }
        AddHandler output_msg_rpt_cb.CheckedChanged, AddressOf output_msg_rpt_cb_changed

        output_msg_period_nb = New NumericUpDown With {
             .Font = New Font("Consolas", 9),
            .Minimum = 1,
            .Maximum = 600,
            .Increment = 1,
            .Size = New Point(44, 22),
            .Location = New Point(output_msg_rpt_cb.Location.X + output_msg_rpt_cb.Width, output_msg_rpt_cb.Location.Y),
            .Anchor = AnchorStyles.Top + AnchorStyles.Right,
            .AutoSize = False
        }
        AddHandler output_msg_period_nb.ValueChanged, AddressOf output_msg_period_nb_changed

        ClearPlot_bt = New Button With {
            .Font = New Font("Consolas", 9),
            .Text = "Clear",
            .Size = New Point(50, 23),
            .Location = New Point(output_msg_period_nb.Location.X + output_msg_period_nb.Width, output_buffer_tb.Location.Y - 1),
            .Anchor = AnchorStyles.Top + AnchorStyles.Right,
            .AutoSize = False
        }
        AddHandler ClearPlot_bt.Click, AddressOf Clear_plots

        top_panel.Controls.Add(ClearPlot_bt)


        log_data_cb = New CheckBox With {
            .Text = "LOG",
            .Font = New Font("Consolas", 10),
            .Appearance = Appearance.Button,
            .Size = New Point(left_panel.Width / 3, 25),
            .Location = New Point(1, com_name_tb.Location.Y + com_name_tb.Height),
            .Anchor = AnchorStyles.Top + AnchorStyles.Right
        }

        New_file_bt = New Button With {
            .Text = "New file",
            .Font = New Font("Consolas", 10),
            .Size = New Point(left_panel.Width / 3 * 2, log_data_cb.Height),
            .Location = New Point(left_panel.Width / 3, log_data_cb.Location.Y),
            .Anchor = AnchorStyles.Top + AnchorStyles.Right
        }

        AddHandler log_data_cb.CheckedChanged, AddressOf Log_Changed
        AddHandler New_file_bt.Click, AddressOf Generate_new_file
        left_panel.Controls.Add(log_data_cb)
        left_panel.Controls.Add(New_file_bt)



        ListOfFiles = New ListBox With {
            .Anchor = AnchorStyles.Top + AnchorStyles.Left + AnchorStyles.Bottom,
            .Size = New Point(left_panel.Width - 2, left_panel.Height - log_data_cb.Location.Y - log_data_cb.Height - 1),
            .Location = New Point(com_name_tb.Location.X, log_data_cb.Location.Y + log_data_cb.Height),
            .IntegralHeight = False,
            .SelectionMode = SelectionMode.One,
            .ScrollAlwaysVisible = True
        }

        AddHandler ListOfFiles.DoubleClick, AddressOf ListOfFiles_DoubleClick
        AddHandler ListOfFiles.MouseDown, AddressOf ListOfFiles_MouseDown
        left_panel.Controls.Add(ListOfFiles)


        ReDim ChannelShow_CB(30 - 1)
        For i = 0 To ChannelShow_CB.Length - 1
            ChannelShow_CB(i) = New CheckBox With {
            .Appearance = Appearance.Normal,
            .TextAlign = ContentAlignment.MiddleLeft,
            .Text = ChannelNames(i Mod 30),
            .ForeColor = ChannelColors(i Mod 10),
            .Checked = False,
            .FlatStyle = FlatStyle.Standard,
            .Padding = New Padding(2, 0, 0, 0),
            .Size = New Point(60, 25),
            .Location = New Point(top_panel.Location.X, top_panel.Location.Y + top_panel.Height + 25 * i), '.Location = New Point(input_buffer_tb.Location.X + 25 * i, 55),
            .Anchor = AnchorStyles.Top + AnchorStyles.Left,
            .Visible = False
            }
            If (set_last_visible(i) = "y") Then ChannelShow_CB(i).Checked = True

            AddHandler ChannelShow_CB(i).CheckedChanged, AddressOf update_visibility
            tttext.SetToolTip(ChannelShow_CB(i), "Show/Hide plot for the parameter")
            MyBase.Controls.Add(ChannelShow_CB(i))
        Next

        top_panel.Controls.Add(output_msg_send_bt)
        top_panel.Controls.Add(output_msg_rpt_cb)
        top_panel.Controls.Add(output_msg_period_nb)

        MyBase.Controls.Add(left_panel)
        MyBase.Controls.Add(top_panel)
        Dim marg As Integer = 1

        MainPlot = New DataVisualization.Charting.Chart With {
            .Size = New Size(ClientSize.Width - left_panel.Width - left_panel.Location.X - ChannelShow_CB(0).Width - 2 * marg, ClientSize.Height - top_panel.Height - top_panel.Location.Y - 2 * marg),
            .Location = New Point(left_panel.Location.X + left_panel.Width + ChannelShow_CB(0).Width + marg, top_panel.Location.Y + top_panel.Height + marg),
            .Anchor = AnchorStyles.Left + AnchorStyles.Right + AnchorStyles.Bottom + AnchorStyles.Top,
            .Enabled = False,
            .BackColor = MyBase.BackColor
         }
        MyBase.Controls.Add(MainPlot)

        MainPlot.ChartAreas.Add("PAYLOAD")
        With MainPlot.ChartAreas(0)
            .Position.Auto = False
            .Position.X = 0
            .Position.Y = 0
            .Position.Height = 100
            .Position.Width = 100
            .Axes(0).Minimum = 0
            .Axes(0).Maximum = sys_plotsize
            .Axes(1).IsLogarithmic = False
            .Axes(1).MinorTickMark.Enabled = True
            .Axes(1).MinorTickMark.Size = 0.2
            .BackColor = MyBase.BackColor
        End With


        For ii = 0 To NChannels - 1
            MainPlot.Annotations.Add(New DataVisualization.Charting.ArrowAnnotation)
            MainPlot.Annotations(ii).ForeColor = ChannelColors(ii Mod 10)
            MainPlot.Annotations(ii).BackColor = ChannelColors(ii Mod 10)

            MainPlot.Annotations(ii).AxisX = MainPlot.ChartAreas(0).Axes(0)
            MainPlot.Annotations(ii).AxisY = MainPlot.ChartAreas(0).Axes(1)
            MainPlot.Annotations(ii).Height = 0
            MainPlot.Annotations(ii).Width = 1
            MainPlot.Annotations(ii).X = 0
            MainPlot.Annotations(ii).Y = 0

            MainPlot.Annotations(ii).Visible = False

        Next

        For jj = 0 To sys_plotsize - 1
            Tdata(jj) = jj
            RData(jj) = 0.0
        Next
        For ii = 0 To NChannels - 1
            MainPlot.Series.Add("Channel #" + (ii + 1).ToString)
            MainPlot.Series(ii).ChartType = DataVisualization.Charting.SeriesChartType.FastLine
            MainPlot.Series(ii).Color = ChannelColors(ii Mod 10)
            MainPlot.Series(ii).Points.DataBindXY(Tdata, RData)
            MainPlot.Series(ii).Enabled = False
        Next

        tttext.SetToolTip(com_name_tb, "Use doubleclick to open or close ports")
        tttext.SetToolTip(com_rfrsh, "Search for COM ports")
        tttext.SetToolTip(com_rate_list, "Select com port baudrate or enter custom value")
        tttext.SetToolTip(param_start_tb, "Output/Input message starting phrase" + vbNewLine + "in incolming lines, everything before will be ignored")

        tttext.SetToolTip(output_buffer_tb, "message to be send to the active com port")
        tttext.SetToolTip(input_buffer_tb, "last received message")

        tttext.SetToolTip(output_msg_send_bt, "Send current string")
        tttext.SetToolTip(output_msg_rpt_cb, "Send message periodically")
        tttext.SetToolTip(output_msg_period_nb, "Period of the message in seconds")

        tttext.SetToolTip(log_data_cb, "Save incomming messages into the log file")
        tttext.SetToolTip(New_file_bt, "Create a new log file")

        tttext.SetToolTip(ListOfFiles, "List of log files" + vbNewLine + "DoubleClick to open the file" + vbNewLine + "RightClick to open the log folder")

        MyBase.MinimumSize = min_size
    End Sub


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MyBase.Text = sys_version
        MyBase.AutoScaleMode = AutoScaleMode.Dpi

        If (Not String.IsNullOrEmpty(My.Settings.BaudRate)) Then set_last_baudrate = My.Settings.BaudRate
        If (Not String.IsNullOrEmpty(My.Settings.Delimeter)) Then set_last_delimeter = My.Settings.Delimeter
        'If (Not String.IsNullOrEmpty(My.Settings.ParamConfig)) Then set_last_parameters = My.Settings.ParamConfig
        If (Not String.IsNullOrEmpty(My.Settings.StartingChar)) Then set_last_startingchar = My.Settings.StartingChar
        If (Not String.IsNullOrEmpty(My.Settings.SelectedPlots)) Then set_last_visible = My.Settings.SelectedPlots

        Build_Gui()
        Update_Serial_Ports()
        Scan_files_in_folder()

        AddHandler active_serial.DataReceived, AddressOf Serial_DataReceived

        UpdatePackage = New Timer With {
            .Interval = 5,
            .Enabled = True
         }
        AddHandler UpdatePackage.Tick, AddressOf UpdatePackage_Tick

        RescaleTimer = New Timer With {
        .Enabled = True,
        .Interval = min_update_time
        }

        AddHandler RescaleTimer.Tick, AddressOf RescaleTimer_Tick

        SendAgainMessage_Timer = New Timer With {
        .Enabled = False,
        .Interval = 1000
        }
        AddHandler SendAgainMessage_Timer.Tick, AddressOf send_new_msg 'SendAgainMessage_Timer_Tick
    End Sub

    Private Sub COM_refreshB_Click(sender As Object, e As EventArgs)
        Update_Serial_Ports()
    End Sub
    Private Function Update_Serial_Ports(Optional ByVal check_serial As Boolean = False) As String()
        Dim list_of_ports(0) As String
        Dim ii As Integer

        Dim myPorts() As String = IO.Ports.SerialPort.GetPortNames()

        If myPorts.Length > 0 Then
            For Each port As String In myPorts
                ReDim Preserve list_of_ports(list_of_ports.Length)
                list_of_ports(list_of_ports.Length - 2) = port
            Next
            ReDim Preserve list_of_ports(list_of_ports.Length - 2)
        Else
            ReDim Preserve list_of_ports(0)
            list_of_ports(0) = "No serial devices"
        End If

        com_name_tb.Items.Clear()
        com_name_tb.Items.AddRange(list_of_ports)

        ' select back active port if any
        If active_serial.IsOpen Then
            For ii = 0 To com_name_tb.Items.Count - 1
                If StrComp(com_name_tb.Items(ii).ToString, active_serial.PortName.ToString, CompareMethod.Text) = 0 Then
                    com_name_tb.SelectedIndex = ii
                End If
            Next
        End If


        Return list_of_ports
    End Function
    Private Sub com_name_tb_DoubleClick(sender As Object, e As EventArgs)
        If IsNothing(com_name_tb.SelectedItem) Then Exit Sub
        If CChar(com_name_tb.SelectedItem.ToString()) <> "C"c Then Exit Sub
        Dim newport As String = com_name_tb.SelectedItem.ToString()
        Dim currentport As String = ""

        If active_serial.IsOpen Then
            currentport = active_serial.PortName
        End If

        If StrComp(newport, currentport, CompareMethod.Text) <> 0 Then 'if strings are different
            active_serial.Close()
            active_serial.PortName = newport
            active_serial.BaudRate = com_rate_list.Text
            'active_serial.RtsEnable = True
            'active_serial.DtrEnable = True
            active_serial.Parity = IO.Ports.Parity.None
            active_serial.DataBits = 8
            active_serial.StopBits = 1
            'active_serial.Handshake = IO.Ports.Handshake.None

            Try
                active_serial.Open()
                com_name_tb.BackColor = Color.LightGreen
                'ControlPanel.Enabled = True
                file_name = ""
            Catch ex As Exception
                com_name_tb.BackColor = SystemColors.Window
            End Try

        Else
            active_serial.Close()
            com_name_tb.BackColor = SystemColors.Window
        End If
    End Sub


    Private Sub send_new_msg(sender As Object, e As EventArgs)
        Dim msg As String = output_buffer_tb.Text

        If (active_serial.IsOpen) Then
            Try
                active_serial.WriteLine(msg)

            Catch ex As Exception
            End Try
        End If
    End Sub

    Private Sub Serial_DataReceived(sender As System.Object, e As System.IO.Ports.SerialDataReceivedEventArgs)
        Dim alen As Integer

        ' read all the data into array and then process it
        alen = active_serial.BytesToRead
        ReDim RBFR(alen - 1)
        active_serial.Read(RBFR, 0, alen)
        ' put the received data into string
        BufferString += System.Text.Encoding.Default.GetString(RBFR)
        ' cut begginning untill the message starts
        If Not (String.IsNullOrEmpty(set_last_startingchar)) Then

            Dim mstart As Integer = InStr(BufferString, set_last_startingchar, CompareMethod.Text)
            If (mstart < 1) Then
                Exit Sub
            End If
            BufferString = Mid(BufferString, mstart)

        End If

        ' Process all packets in the buffer
        Dim mstop As Integer = InStr(BufferString, vbLf, CompareMethod.Text)
        Do While mstop > 0
            If IsDataReady Then
                LostPackets += 1
            Else
                IsDataReady = 1
            End If
            ReceivedMsg = Mid(BufferString, 1, mstop + 1)
            If BufferString.Length > mstop + 1 Then
                BufferString = Mid(BufferString, mstop + 2)
            Else
                BufferString = ""
            End If
            mstop = InStr(BufferString, vbLf, CompareMethod.Text)
        Loop

    End Sub
    Private Sub RescaleTimer_Tick(sender As Object, e As EventArgs)
        ready_to_rescale = True
    End Sub
    Private Sub updateMsgStart(sender As Object, e As EventArgs)
        set_last_startingchar = param_start_tb.Text
    End Sub

    Private Sub UpdatePackage_Tick(sender As Object, e As EventArgs)

        If IsDataReady Then
            Recived_Report(ReceivedMsg)
            IsDataReady = False
        End If

    End Sub

    Private Sub Recived_Report(_msg As String)
        Dim num As Integer = 0

        input_buffer_tb.Text = _msg
        receivedData = _msg.Split(sys_delims, StringSplitOptions.RemoveEmptyEntries)

        For i = 0 To receivedData.Length - 1
            If IsNumeric(receivedData(i)) Then
                MainPlot.Series(num).Points.Item(Current_Point).YValues(0) = CDbl(receivedData(i))
                MainPlot.Series(num).Points.Item(Current_Point).IsEmpty = False
                If Current_Point < sys_plotsize Then
                    MainPlot.Series(num).Points.Item(Current_Point + 1).IsEmpty = True
                End If
                MainPlot.Annotations(num).X = Current_Point
                MainPlot.Annotations(num).Y = MainPlot.Series(num).Points.Item(Current_Point).YValues(0)

                num += 1

            End If
        Next

        Current_Point = Current_Point + 1
        If Current_Point >= sys_plotsize Then Current_Point = 0

        If ready_to_rescale Then
            ready_to_rescale = False
            MainPlot.ChartAreas(0).RecalculateAxesScale()
        End If

        set_Visible_channels(num, receivedData)
        If log_data_cb.Checked Then
            Try
                IO.File.AppendAllText(file_path + IO.Path.DirectorySeparatorChar + file_name, _msg)
            Catch ex As Exception

            End Try
        End If

    End Sub
    Private Sub set_Visible_channels(_num As Integer, ByRef _income() As String)
        If _num = ActiveChannels Then Exit Sub
        If _num < 1 Then Exit Sub
        If _num > ChannelShow_CB.Length - 1 Then _num = ChannelShow_CB.Length - 1

        For i = 0 To _num
            ChannelShow_CB(i).Visible = True
            MainPlot.Series(i).Enabled = ChannelShow_CB(i).Checked
            MainPlot.Annotations(i).Visible = ChannelShow_CB(i).Checked
        Next
        For i = _num To ChannelShow_CB.Length - 1
            ChannelShow_CB(i).Visible = False
            MainPlot.Series(i).Enabled = False
            MainPlot.Annotations(i).Visible = False
        Next

        ' Update variable names
        Dim num As Integer = 0
        Dim nam As String = ""
        Dim cnt As Integer = 1

        For i = 0 To receivedData.Length - 1
            If IsNumeric(receivedData(i)) Then
                If nam <> "" Then
                    ChannelShow_CB(num).Text = nam + cnt.ToString
                    cnt += 1
                    'nam = ""
                End If
                num += 1
            Else
                nam = receivedData(i)
                cnt = 1
            End If
        Next

        MainPlot.Refresh()
        ActiveChannels = _num
    End Sub
    Private Sub update_visibility(sender As Object, e As EventArgs)
        Dim _change_all As Integer = 1

        For i = 0 To ActiveChannels - 1
            MainPlot.Series(i).Enabled = ChannelShow_CB(i).Checked
            MainPlot.Annotations(i).Visible = ChannelShow_CB(i).Checked
        Next

        MainPlot.Update()
        MainPlot.Refresh()
    End Sub

    Private Sub output_msg_rpt_cb_changed(sender As Object, e As EventArgs)

        SendAgainMessage_Timer.Enabled = output_msg_rpt_cb.Checked

    End Sub

    Private Sub output_msg_period_nb_changed(sender As Object, e As EventArgs)

        SendAgainMessage_Timer.Interval = output_msg_period_nb.Value * 1000

    End Sub

    Private Sub Clear_plots(sender As Object, e As EventArgs)
        Current_Point = 0
        For ii = 0 To NChannels - 1
            For jj = 0 To sys_plotsize - 1
                MainPlot.Series(ii).Points.Item(jj).IsEmpty = True
            Next
        Next
    End Sub

    Private Sub MainForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        My.Settings.BaudRate = com_rate_list.Text
        My.Settings.ParamConfig = ""
        My.Settings.StartingChar = param_start_tb.Text
        My.Settings.SelectedPlots = ""

        For i = 0 To ChannelShow_CB.Length - 1

            If ChannelShow_CB(i).Checked Then
                My.Settings.SelectedPlots += "y"
            Else
                My.Settings.SelectedPlots += "n"
            End If
        Next

    End Sub
End Class
