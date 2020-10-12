Imports System

Public Class MainForm
    Dim sys_version As String = "Serial data visualization v0.5"
    Dim max_number_of_parameters As Integer = 16
    Dim min_size As New Point(480, 240)
    Dim list_of_serial_rates() As String = {"1200", "2400", "4800", "9600", "14400", "19200", "28800", "38400", "57600", "115200", "230400"}
    Dim sys_del As String = ","
    Dim sys_delims() As Char = {" ", ":", ";", ",", "/", vbTab, "_", "="}
    Dim sys_plotsize As Integer = 1000

    Dim min_update_time As Integer = 1000
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
    Dim output_panel As Panel
    Dim input_panel As Panel
    Dim setup_panel As Panel

    'com ports
    Dim com_name_tb As ListBox
    'Dim com_connect_bt As Button
    Dim com_rfrsh As Button
    Dim com_rate_list As ComboBox
    Dim active_serial As New System.IO.Ports.SerialPort
    '   Dim spstream As System.IO.Ports.serialportstream

    Dim output_buffer_tb As TextBox
    Dim input_buffer_tb As TextBox

    Dim output_msg_send_bt As Button
    Dim output_msg_rpt_cb As CheckBox
    Dim output_msg_period_nb As NumericUpDown

    Dim param_delim_cb As ComboBox
    Dim param_start_tb As TextBox

    Dim param_num() As NumericUpDown
    Dim param_txt() As TextBox
    Dim send_param() As CheckBox

    Dim ChannelShow_CB() As CheckBox
    Dim ChannelColors() As Color = {Color.Black, Color.Red, Color.Green, Color.Blue, Color.DarkCyan, Color.Magenta, Color.Brown, Color.Gray, Color.Orange, Color.BlueViolet}
    Dim ChannelColorsCompl() As Color = {Color.Gray, Color.Black, Color.Black, Color.Black, Color.Black, Color.Black, Color.Black, Color.Black, Color.Black, Color.Black}
    Dim ChannelNames() As String = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "1", "2", "3", "4"}
    Dim Min_Value As Single = 0.0
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

    Dim add_num_bt As Button
    Dim add_txt_bt As Button
    Dim rem_par_bt As Button

    Dim output_msg_struct As String
    Dim input_msg_struct As String

    Dim RescaleTimer As Timer
    Dim UpdatePackage As Timer

    Dim SendAgainMessage_Timer As Timer

    Dim Can_try_again As Boolean
    'Dim IsSendingNewSettings As Boolean

    Dim set_last_baudrate As String = "115200"
    Dim set_last_delimeter As String = ":"
    Dim set_last_startingchar As String = ">>"
    Dim set_last_parameters As String = "#100" + vbNewLine + "@default" + vbNewLine
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
    Private Sub Generate_new_file(sender As Object, e As EventArgs)
        If No_logging Then Exit Sub

        Dim localDate = DateTime.Now

        Dim board As String = ""
        'If CurrentData.Board_ID >= 0 Then board = " Board" + CurrentData.Board_ID.ToString("00")

        file_name = "exp" + localDate.Year.ToString("0000") + "" + localDate.Month.ToString("00") + "" + localDate.Day.ToString("00") + " " + localDate.Hour.ToString("00") + "-" + localDate.Minute.ToString("00") + "-" + localDate.Second.ToString("00") + ".txt"
        ''ToolStripStatusLabel1.Text = file_name
        'Status_filename.Text = file_name
        ''status_logging.Text = "new file generated"

        '' creates new file if needed, and starts writing to it
        Using outputFile As New IO.StreamWriter(file_path + IO.Path.DirectorySeparatorChar + file_name, True)
            outputFile.WriteLine("starting time:" + localDate)
        End Using

        'SplitFileTimer.Enabled = False
        'SplitFileTimer.Enabled = True
        ''        LogTimer.Enabled = False
        ''       LogTimer.Enabled = True


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


        Current_Point = 0
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
        'Dim _fname As String
        '_fname = ListOfFiles.SelectedItem.ToString
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

        '.BackColor = Color.Salmon,
        '.ForeColor = Color.Green,
        output_panel = New Panel With {
             .Location = New Point(2, 2),
            .Size = New Point(103, MyBase.ClientSize.Height - 4),
            .Anchor = (AnchorStyles.Top + AnchorStyles.Left + AnchorStyles.Bottom)
        }
        setup_panel = New Panel With {
            .Location = New Point(105, 2),
            .Size = New Point(MyBase.ClientSize.Width - 107 - 150, 43),
            .Anchor = (AnchorStyles.Top + AnchorStyles.Left + AnchorStyles.Right)
        }

        com_rate_list = New ComboBox With {
            .Size = New Point(output_panel.Width - 30, 25),
            .Location = New Point(1, 1)
        }

        com_rate_list.Items.AddRange(list_of_serial_rates)
        If com_rate_list.Items.Contains(set_last_baudrate) Then
            com_rate_list.SelectedItem = set_last_baudrate
        Else
            com_rate_list.Text = set_last_baudrate
        End If
        com_rate_list.SelectedItem = set_last_baudrate

        com_rfrsh = New Button With {
            .Text = "R",
            .Size = New Point(25, com_rate_list.Height),
            .Location = New Point(com_rate_list.Location.X + com_rate_list.Width, com_rate_list.Location.Y)
        }
        AddHandler com_rfrsh.Click, AddressOf COM_refreshB_Click

        com_name_tb = New ListBox With {
            .Anchor = AnchorStyles.Top + AnchorStyles.Left,
            .Size = New Point(output_panel.Width - 5, 60),
            .Location = New Point(1, com_rate_list.Height + com_rate_list.Location.Y),
            .IntegralHeight = False
        }
        AddHandler com_name_tb.DoubleClick, AddressOf com_name_tb_DoubleClick


        param_start_tb = New TextBox With {
            .Font = New Font("Consolas", 10),
            .Text = set_last_startingchar,
            .Size = New Point(49, 25),
            .Location = New Point(com_name_tb.Location.X, com_name_tb.Location.Y + com_name_tb.Height)
        }
        AddHandler param_start_tb.TextChanged, AddressOf updateMsgStart
        param_delim_cb = New ComboBox With {
            .Font = New Font("Consolas", 10),
              .Size = New Point(49, 25),
              .Location = New Point(param_start_tb.Location.X + param_start_tb.Width, param_start_tb.Location.Y)
        }
        For Each d In sys_delims
            param_delim_cb.Items.Add(d.ToString)
        Next

        If param_delim_cb.Items.Contains(set_last_delimeter) Then
            param_delim_cb.SelectedItem = set_last_delimeter
        Else
            param_delim_cb.Text = set_last_delimeter
        End If
        param_delim_cb.SelectedItem = set_last_delimeter


        AddHandler param_delim_cb.TextChanged, AddressOf UpdateMessage
        AddHandler param_start_tb.TextChanged, AddressOf UpdateMessage

        add_num_bt = New Button With {
            .AutoSize = False,
            .Size = New Point(50, 25),
            .Location = New Point(com_name_tb.Location.X, param_delim_cb.Location.Y + param_delim_cb.Height),
            .Anchor = AnchorStyles.Top + AnchorStyles.Left,
            .Text = "+123",
            .Font = New Font("Consolas", 10),
            .TextAlign = ContentAlignment.MiddleCenter,
            .Margin = New Padding(0, 0, 0, 0),
            .Padding = New Padding(0, 0, 0, 0)
        }
        add_txt_bt = New Button With {
            .AutoSize = False,
            .Size = New Point(50, 25),
            .Location = New Point(add_num_bt.Width, add_num_bt.Location.Y),
            .Anchor = AnchorStyles.Top + AnchorStyles.Left,
            .Text = "+ABC",
            .Font = New Font("Consolas", 10),
            .TextAlign = ContentAlignment.MiddleCenter,
            .Margin = New Padding(0, 0, 0, 0),
            .Padding = New Padding(0, 0, 0, 0)
        }

        rem_par_bt = New Button With {
            .AutoSize = False,
            .Size = New Point(20, 22),
            .Location = New Point(-50, -50),
            .Anchor = AnchorStyles.Top + AnchorStyles.Left,
            .Text = "X",
            .Font = New Font("Consolas", 10),
            .TextAlign = ContentAlignment.MiddleCenter,
            .Margin = New Padding(2, 2, 2, 2),
            .Padding = New Padding(2, 0, 0, 2)
        }




        AddHandler add_num_bt.Click, AddressOf Add_new_parameter
        AddHandler add_txt_bt.Click, AddressOf Add_new_parameter
        AddHandler rem_par_bt.Click, AddressOf Remove_last_parameter
        output_panel.Controls.Add(add_num_bt)
        output_panel.Controls.Add(add_txt_bt)
        output_panel.Controls.Add(rem_par_bt)
        output_panel.Controls.Add(com_name_tb)
        output_panel.Controls.Add(param_delim_cb)
        output_panel.Controls.Add(param_start_tb)

        output_panel.Controls.Add(com_rfrsh)
        output_panel.Controls.Add(com_rate_list)





        output_buffer_tb = New TextBox With {
            .Size = New Point(setup_panel.Width - 150, 25),
            .Location = New Point(1, com_name_tb.Location.Y),
            .Anchor = AnchorStyles.Top + AnchorStyles.Right + AnchorStyles.Left
        }

        output_msg_send_bt = New Button With {
            .Text = "SEND",
            .Size = New Point(50, com_rfrsh.Height),
            .Location = New Point(output_buffer_tb.Location.X + output_buffer_tb.Width, output_buffer_tb.Location.Y),
            .Anchor = AnchorStyles.Top + AnchorStyles.Right
        }
        AddHandler output_msg_send_bt.Click, AddressOf send_new_msg

        output_msg_rpt_cb = New CheckBox With {
            .Text = "Every",
            .Checked = False,
            .Size = New Point(55, output_msg_send_bt.Height),
            .Location = New Point(output_msg_send_bt.Location.X + output_msg_send_bt.Width + 2, output_msg_send_bt.Location.Y),
            .Anchor = AnchorStyles.Top + AnchorStyles.Right
            }
        AddHandler output_msg_rpt_cb.CheckedChanged, AddressOf output_msg_rpt_cb_changed

        output_msg_period_nb = New NumericUpDown With {
            .Minimum = 1,
            .Maximum = 600,
            .Increment = 1,
            .Size = New Point(42, output_msg_send_bt.Height),
            .Location = New Point(output_msg_rpt_cb.Location.X + output_msg_rpt_cb.Width, output_msg_rpt_cb.Location.Y),
            .Anchor = AnchorStyles.Top + AnchorStyles.Right
        }
        AddHandler output_msg_period_nb.ValueChanged, AddressOf output_msg_period_nb_changed

        input_buffer_tb = New TextBox With {
            .Size = New Point(output_buffer_tb.Width + 50, 25),
            .Location = New Point(output_buffer_tb.Location.X, 1),
            .Anchor = AnchorStyles.Top + AnchorStyles.Right + AnchorStyles.Left,
            .ReadOnly = True
        }

        log_data_cb = New CheckBox With {
            .Text = "LOG",
            .Size = New Point(output_msg_rpt_cb.Width, output_msg_rpt_cb.Height),
            .Location = New Point(output_msg_rpt_cb.Location.X, 1),
            .Anchor = AnchorStyles.Top + AnchorStyles.Right
        }

        New_file_bt = New Button With {
            .Text = "NEW",
            .Size = New Point(output_msg_period_nb.Width, output_msg_period_nb.Height),
            .Location = New Point(output_msg_period_nb.Location.X, log_data_cb.Location.Y),
            .Anchor = AnchorStyles.Top + AnchorStyles.Right
        }

        AddHandler New_file_bt.Click, AddressOf Generate_new_file


        ListOfFiles = New ListBox With {
            .Anchor = AnchorStyles.Top + AnchorStyles.Right,
            .Size = New Point(150, 43),
            .Location = New Point(setup_panel.Location.X + setup_panel.Width, setup_panel.Location.Y),
            .IntegralHeight = False,
            .SelectionMode = SelectionMode.One,
            .ScrollAlwaysVisible = True
        }
        AddHandler ListOfFiles.DoubleClick, AddressOf ListOfFiles_DoubleClick
        AddHandler ListOfFiles.MouseDown, AddressOf ListOfFiles_MouseDown
        AddHandler ListOfFiles.MouseEnter, AddressOf ListOfFiles_MouseEnter
        AddHandler ListOfFiles.MouseLeave, AddressOf ListOfFiles_MouseLeave


        '.BackColor = ChannelColors(i Mod 10),
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
            .Location = New Point(setup_panel.Location.X, setup_panel.Location.Y + setup_panel.Height + 25 * i), '.Location = New Point(input_buffer_tb.Location.X + 25 * i, 55),
            .Anchor = AnchorStyles.Top + AnchorStyles.Left,
            .Visible = False
            }
            If (set_last_visible(i) = "y") Then ChannelShow_CB(i).Checked = True

            AddHandler ChannelShow_CB(i).CheckedChanged, AddressOf update_visibility
            'AddHandler ChannelShow_CB(i).D, AddressOf hide_all_plots
            tttext.SetToolTip(ChannelShow_CB(i), "Show/Hide plot for the parameter")
            MyBase.Controls.Add(ChannelShow_CB(i))
        Next

        MyBase.Controls.Add(ListOfFiles)
        setup_panel.Controls.Add(output_buffer_tb)
        setup_panel.Controls.Add(output_msg_send_bt)
        setup_panel.Controls.Add(output_msg_rpt_cb)
        setup_panel.Controls.Add(output_msg_period_nb)
        setup_panel.Controls.Add(input_buffer_tb)
        setup_panel.Controls.Add(log_data_cb)
        setup_panel.Controls.Add(New_file_bt)

        MyBase.Controls.Add(output_panel)
        MyBase.Controls.Add(setup_panel)
        Dim marg As Integer = 1

        MainPlot = New DataVisualization.Charting.Chart With {
            .Size = New Size(ClientSize.Width - output_panel.Width - output_panel.Location.X - ChannelShow_CB(0).Width - 2 * marg, ClientSize.Height - setup_panel.Height - setup_panel.Location.Y - 2 * marg),
            .Location = New Point(output_panel.Location.X + output_panel.Width + ChannelShow_CB(0).Width + marg, setup_panel.Location.Y + setup_panel.Height + marg),
            .Anchor = AnchorStyles.Left + AnchorStyles.Right + AnchorStyles.Bottom + AnchorStyles.Top,
            .Enabled = False,
            .BackColor = Color.White 'MyBase.BackColor
         }
        MyBase.Controls.Add(MainPlot)
        'AddHandler MainPlot.MouseUp, AddressOf ShowCoordinates

        MainPlot.ChartAreas.Add("PAYLOAD")
        With MainPlot.ChartAreas(0)
            .Position.Auto = False
            .Position.X = 0
            .Position.Y = 0
            .Position.Height = 100
            .Position.Width = 100
            .Axes(0).Minimum = 0
            .Axes(0).Maximum = sys_plotsize
            '.Axes(1).Minimum = -32768
            '.Axes(1).Maximum = 32767
            .Axes(1).IsLogarithmic = False
            '.Axes(1).LabelStyle.Format = "0.E+0"
            .Axes(1).MinorTickMark.Enabled = True
            '.Axes(1).MinorTickMark.Interval = 1
            .Axes(1).MinorTickMark.Size = 0.2
            .BackColor = MyBase.BackColor
        End With
        'MainPlot.ChartAreas(0).CursorX.Position = 0
        'MainPlot.ChartAreas(0).CursorY.Position = 0

        MainPlot.ChartAreas(0).RecalculateAxesScale()

        For ii = 0 To NChannels - 1
            MainPlot.Annotations.Add(New DataVisualization.Charting.ArrowAnnotation)
            MainPlot.Annotations(ii).ForeColor = ChannelColors(ii Mod 10)
            MainPlot.Annotations(ii).BackColor = ChannelColors(ii Mod 10)

            MainPlot.Annotations(ii).AxisX = MainPlot.ChartAreas(0).Axes(0)
            MainPlot.Annotations(ii).AxisY = MainPlot.ChartAreas(0).Axes(1)
            MainPlot.Annotations(ii).Height = 0
            MainPlot.Annotations(ii).Width = 1
            MainPlot.Annotations(ii).X = 0
            MainPlot.Annotations(ii).Y = Min_Value

            MainPlot.Annotations(ii).Visible = False

        Next

        For jj = 0 To sys_plotsize - 1
            Tdata(jj) = jj
            RData(jj) = Min_Value
        Next
        For ii = 0 To NChannels - 1
            MainPlot.Series.Add("Channel #" + (ii + 1).ToString)
            MainPlot.Series(ii).ChartType = DataVisualization.Charting.SeriesChartType.FastLine
            MainPlot.Series(ii).Color = ChannelColors(ii Mod 10)
            MainPlot.Series(ii).Points.DataBindXY(Tdata, RData)
            MainPlot.Series(ii).Enabled = False
        Next
        AddHandler MainPlot.DoubleClick, AddressOf Clear_plots
        'Me.Controls.Add(MainPlot)

        tttext.SetToolTip(com_name_tb, "Use doubleclick to open or close ports")
        tttext.SetToolTip(com_rfrsh, "Search for COM ports")
        tttext.SetToolTip(com_rate_list, "Select com port baudrate or enter custom value")
        tttext.SetToolTip(param_start_tb, "Output/Input message starting phrase" + vbNewLine + "in incolming lines, everything before will be ignored")
        tttext.SetToolTip(param_delim_cb, "Delimeter between parameters of outgoing message ")

        tttext.SetToolTip(add_num_bt, "Add new numerical parameter (0:255)")
        tttext.SetToolTip(add_txt_bt, "Add new string parameter")

        tttext.SetToolTip(rem_par_bt, "Remove last parameter")
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
        If (Not String.IsNullOrEmpty(My.Settings.ParamConfig)) Then set_last_parameters = My.Settings.ParamConfig
        If (Not String.IsNullOrEmpty(My.Settings.StartingChar)) Then set_last_startingchar = My.Settings.StartingChar
        If (Not String.IsNullOrEmpty(My.Settings.SelectedPlots)) Then set_last_visible = My.Settings.SelectedPlots

        Build_Gui()
        Update_Serial_Ports()
        Scan_files_in_folder()

        'load set of parameters from last time
        Dim listOFParams() As String
        listOFParams = set_last_parameters.Split(vbNewLine.ToCharArray, StringSplitOptions.RemoveEmptyEntries)
        For i = 0 To listOFParams.Length - 1
            listOFParams(i).Replace(vbLf, "")

            If listOFParams(i)(0) = "#" Then
                listOFParams(i) = listOFParams(i).Remove(0, 1)
                Add_new_parameter(add_num_bt, EventArgs.Empty)
                param_num(i).Value = Convert.ToInt32(listOFParams(i))
            ElseIf listOFParams(i)(0) = "@" Then
                listOFParams(i) = listOFParams(i).Remove(0, 1)
                Add_new_parameter(add_txt_bt, EventArgs.Empty)
                param_txt(i).Text = listOFParams(i)
            End If

        Next

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


    Private Sub Add_new_parameter(sender As Object, e As EventArgs)
        Dim num_of_param As Integer
        If IsNothing(send_param) Then
            num_of_param = 0
        Else
            num_of_param = send_param.Length
        End If

        If num_of_param > max_number_of_parameters Then Exit Sub

        ReDim Preserve send_param(num_of_param)
        ReDim Preserve param_num(num_of_param)
        ReDim Preserve param_txt(num_of_param)

        send_param(send_param.Length - 1) = New CheckBox With {
            .Text = "",
            .AutoSize = False,
            .Location = New Point(5, add_num_bt.Location.Y + send_param.Length * sys_step),
            .Size = New Point(20, 20),
            .Checked = True
        }
        param_num(send_param.Length - 1) = New NumericUpDown With {
            .Visible = False,
            .Increment = 1,
            .Minimum = 0,
            .Maximum = 256,
            .Value = 128,
            .AutoSize = False,
            .Size = New Point(50, 20),
            .Location = New Point(30, send_param(send_param.Length - 1).Location.Y)
        }
        param_txt(send_param.Length - 1) = New TextBox With {
            .Visible = False,
            .Text = "a",
            .Multiline = False,
            .MaxLength = 16,
            .AutoSize = False,
            .Size = New Point(50, 20),
            .Location = New Point(30, send_param(send_param.Length - 1).Location.Y)
        }
        AddHandler send_param(send_param.Length - 1).CheckStateChanged, AddressOf UpdateMessage
        AddHandler param_num(send_param.Length - 1).ValueChanged, AddressOf UpdateMessage
        AddHandler param_txt(send_param.Length - 1).TextChanged, AddressOf UpdateMessage


        output_panel.Controls.Add(send_param(send_param.Length - 1))
        output_panel.Controls.Add(param_num(send_param.Length - 1))
        output_panel.Controls.Add(param_txt(send_param.Length - 1))

        tttext.SetToolTip(send_param(send_param.Length - 1), "include the value into the message")

        If (sender Is add_num_bt) Then
            'add_num_bt.ForeColor = Color.Red
            param_num(send_param.Length - 1).Visible = True
        Else
            'add_txt_bt.ForeColor = Color.Red
            param_txt(send_param.Length - 1).Visible = True
        End If
        ' move remove button

        rem_par_bt.Location = New Point(30 + 50, send_param(send_param.Length - 1).Location.Y - 1)

        If (send_param.Length > 1) Then
            param_num(send_param.Length - 2).Size = New Point(70, 20)
            param_txt(send_param.Length - 2).Size = New Point(70, 20)
        End If

        UpdateMessage(send_param(send_param.Length - 1), EventArgs.Empty)

        If (add_num_bt.Location.Y + (send_param.Length + 1) * sys_step + sys_offset > min_size.Y) Then
            MyBase.MinimumSize = New Point(MyBase.MinimumSize.Width, add_num_bt.Location.Y + (send_param.Length + 1) * sys_step + sys_offset)
        Else
            MyBase.MinimumSize = New Point(MyBase.MinimumSize.Width, min_size.Y)
        End If
        'output_panel.MinimumSize = New Point(0, 0)
        'MyBase.Text = MyBase.MinimumSize.Height.ToString
        'If MyBase.
    End Sub

    Private Sub Remove_last_parameter(sender As Object, e As EventArgs)
        Dim num_of_param As Integer
        If IsNothing(send_param) Then
            num_of_param = 0
        Else
            num_of_param = send_param.Length
        End If

        If num_of_param > 0 Then
            send_param(num_of_param - 1).Visible = False
            param_num(num_of_param - 1).Visible = False
            param_txt(num_of_param - 1).Visible = False

            ReDim Preserve send_param(num_of_param - 2)
            ReDim Preserve param_num(num_of_param - 2)
            ReDim Preserve param_txt(num_of_param - 2)

        End If




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
            active_serial.RtsEnable = True
            'active_serial.DtrEnable = True
            'active_serial.Parity = IO.Ports.Parity.Odd
            'active_serial.DataBits = 8
            'active_serial.StopBits = 1
            'active_serial.Handshake = IO.Ports.Handshake.None

            Try
                active_serial.Open()
                com_name_tb.BackColor = Color.LightGreen
                'ControlPanel.Enabled = True
                file_name = ""
                Generate_new_file(New_file_bt, EventArgs.Empty)
            Catch ex As Exception
                com_name_tb.BackColor = SystemColors.Window
                'ControlPanel.Enabled = False
            End Try

        Else
            active_serial.Close()
            com_name_tb.BackColor = SystemColors.Window
            'ControlPanel.Enabled = False
        End If
    End Sub


    Private Sub send_new_msg(sender As Object, e As EventArgs)
        'BuiltMsg()
        Dim msg As String = output_buffer_tb.Text

        If (active_serial.IsOpen) Then
            Try
                active_serial.WriteLine(msg)


            Catch ex As Exception
                'status_info.Text = "Failed To send data"
            End Try
        End If
    End Sub
    Private Sub UpdateMessage(sender As Object, e As EventArgs)
        BuiltMsg()
    End Sub
    Private Sub BuiltMsg()
        Dim _dlm As String = param_delim_cb.Text
        Dim _msg As String = param_start_tb.Text
        Dim _par As String = ""

        If IsNothing(send_param) Then Exit Sub
        If (send_param.Length < 1) Then Exit Sub


        For i = 0 To send_param.Length - 1
            If (param_num(i).Visible) Then
                _par = param_num(i).Value.ToString
            Else
                _par = param_txt(i).Text
            End If
            If send_param(i).Checked Then
                _msg += _par
            End If
            _msg += _dlm
        Next
        '_msg += vbLf
        output_buffer_tb.Text = _msg
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

        ' make sure the beggining is correct
        '        If StrComp(">>", Mid(BufferString, 1, 2), CompareMethod.Text) <> 0 Then
        '       Exit Sub
        '      End If


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
        Dim _i As Integer
        'If Status_lost.Text <> LostPackets.ToString("##000") Then
        '    Status_lost.Text = LostPackets.ToString("##000")
        'End If

        If IsDataReady Then
            Recived_Report(ReceivedMsg)
            IsDataReady = False
            'Can_try_again = True
        End If



        'If Can_try_again Then
        '    Can_try_again = False
        '    'For _i = 0 To 7
        '    '    'IsSendingNewSettings = True
        '    '    If (DP_Hr_needUpdating(_i) Or Timer_needUpdate) Then 'Or On_needUpdate(_i)
        '    '        Send_New_Settings()
        '    '        Exit For
        '    '    End If
        '    '    '  Can_try_again = True
        '    '    'IsSendingNewSettings = False
        '    'Next

        'End If

    End Sub

    Private Sub Recived_Report(_msg As String)
        Dim num As Integer = 0
        'Dim a As DateTime = Now
        '
        'Dim nam As String = ""
        input_buffer_tb.Text = _msg
        receivedData = _msg.Split(sys_delims, StringSplitOptions.RemoveEmptyEntries)
        'Dim b As DateTime = Now

        'output_buffer_tb.Text = ""
        For i = 0 To receivedData.Length - 1
            If IsNumeric(receivedData(i)) Then
                'output_buffer_tb.Text += receivedData(i) + "^"
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

        'Dim c As DateTime = Now


        Current_Point = Current_Point + 1
        If Current_Point >= sys_plotsize Then Current_Point = 0

        If ready_to_rescale Then
            ready_to_rescale = False
            MainPlot.ChartAreas(0).RecalculateAxesScale()
        End If

        set_Visible_channels(num, receivedData)
        If log_data_cb.Checked Then
            Try
                'Using outputFile As New StreamWriter(file_path + "\" + file_name, True)
                ' outputFile.WriteLine(twf)
                ' End Using
                IO.File.AppendAllText(file_path + IO.Path.DirectorySeparatorChar + file_name, _msg)

                'status_info.Text = (start_time - stop_time).TotalMilliseconds.ToString("###0") + "ms"
                'lastSaveTime = Now
                'status_logging.Text = "last log: " + lastSaveTime.ToString("HH:mm:ss")
            Catch ex As Exception
                'status_logging.Text = "Log File In use!"
            End Try
        End If
        'Dim d As DateTime = Now

        'Dim msg As String = (b - a).TotalMilliseconds.ToString + " / " + (c - b).TotalMilliseconds.ToString + " / " + (d - c).TotalMilliseconds.ToString
        'output_buffer_tb.Text = msg
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
        'If My.Computer.Keyboard.ShiftKeyDown Then
        '    For i = 0 To ActiveChannels - 1
        '        ChannelShow_CB(i).Checked = ChannelShow_CB(0).Checked
        '    Next
        'End If
        For i = 0 To ActiveChannels - 1
            MainPlot.Series(i).Enabled = ChannelShow_CB(i).Checked
            MainPlot.Annotations(i).Visible = ChannelShow_CB(i).Checked
        Next
        'For i = ActiveChannels To ChannelShow_CB.Length - 1
        '    ChannelShow_CB(i).Visible = False
        '    MainPlot.Series(i).Enabled = False
        '    MainPlot.Annotations(i).Visible = False
        'Next
        MainPlot.Update()
        MainPlot.Refresh()
    End Sub

    Private Sub output_msg_rpt_cb_changed(sender As Object, e As EventArgs)

        SendAgainMessage_Timer.Enabled = output_msg_rpt_cb.Checked

    End Sub

    Private Sub output_msg_period_nb_changed(sender As Object, e As EventArgs)

        SendAgainMessage_Timer.Interval = output_msg_period_nb.Value * 1000

    End Sub

    'Private Sub ShowCoordinates(sender As Object, e As MouseEventArgs)
    '    Dim i As Integer

    '    MainPlot.ChartAreas(0).CursorX.SetCursorPixelPosition(New Point(e.X, e.Y), True)
    '    MainPlot.ChartAreas(0).CursorY.SetCursorPixelPosition(New Point(e.X, e.Y), True)

    '    Dim pX As Double = MainPlot.ChartAreas(0).CursorX.Position
    '    Dim pY As Double = MainPlot.ChartAreas(0).CursorY.Position

    'End Sub
    Private Sub Clear_plots(sender As Object, e As EventArgs)
        'Dim RData(sys_plotsize) As Single
        'Dim Tdata(sys_plotsize) As Single

        'For jj = 0 To sys_plotsize - 1
        '    Tdata(jj) = jj
        '    RData(jj) = Min_Value
        'Next

        Current_Point = 0
        For ii = 0 To NChannels - 1
            For jj = 0 To sys_plotsize - 1

                '                MainPlot.Series(ii).Points.Item(jj).YValues(0) = Min_Value
                MainPlot.Series(ii).Points.Item(jj).IsEmpty = True
            Next
        Next


    End Sub



    Private Sub Update_Plots()
        Dim _i As Integer

        'For _i = 0 To 7
        '    MainPlot.Series(_i).Points.Item(Current_point).YValues(0) = CurrentData.R_x(_i)
        '    MainPlot.Series(_i).Points.Item(Current_point).IsEmpty = False
        '    If Current_point < 1000 Then
        '        MainPlot.Series(_i).Points.Item(Current_point + 1).IsEmpty = True
        '    End If
        '    MainPlot.Annotations(_i).X = Current_point
        '    MainPlot.Annotations(_i).Y = CurrentData.R_x(_i)
        'Next

        'MainPlot.Series(8).Points.Item(Current_point).YValues(0) = CurrentData.Rgas
        'MainPlot.Series(8).Points.Item(Current_point).IsEmpty = False
        'If Current_point < 1000 Then
        '    MainPlot.Series(8).Points.Item(Current_point + 1).IsEmpty = True
        'End If
        'MainPlot.Annotations(8).X = Current_point
        'MainPlot.Annotations(8).Y = CurrentData.Rgas


        'Current_point = Current_point + 1
        'If Current_point > 1000 Then Current_point = 0

        MainPlot.Refresh()
    End Sub

    Private Sub MainForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        My.Settings.BaudRate = com_rate_list.Text
        My.Settings.Delimeter = param_delim_cb.Text
        My.Settings.ParamConfig = ""
        My.Settings.StartingChar = param_start_tb.Text
        My.Settings.SelectedPlots = ""

        If Not (IsNothing(send_param)) Then
            For i = 0 To send_param.Length - 1
                If param_num(i).Visible Then
                    My.Settings.ParamConfig += "#" + param_num(i).Value.ToString + vbNewLine
                Else
                    My.Settings.ParamConfig += "@" + param_txt(i).Text + vbNewLine
                End If
            Next
        End If


        For i = 0 To ChannelShow_CB.Length - 1

            If ChannelShow_CB(i).Checked Then
                My.Settings.SelectedPlots += "y"
            Else
                My.Settings.SelectedPlots += "n"
            End If
        Next

    End Sub
End Class
