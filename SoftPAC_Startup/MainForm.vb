Imports System.Net.Sockets
Public Class MainForm


    Private Function RunningTasks() As Integer
        Dim Rx(127) As Byte
        _NetworkStream.Write(ANYTasks_B, 0, ANYTasks_B.Length)
        Dim dStart As Date = DateAdd(DateInterval.Second, 2, Now)
        Do Until (_NetworkStream.DataAvailable Or (Now > dStart))
            System.Threading.Thread.Sleep(10)
        Loop
        _NetworkStream.Read(Rx, 0, Rx.Length)
        If Rx.Length > 0 Then
            Dim sX As String = BytesToString(Rx).Trim
            Return CInt(Val(sX))
        End If
        Return -1
    End Function

    Private Comm As TcpClient = Nothing
    Private LastStart As Date = Now
    Private _NetworkStream As NetworkStream = Nothing
    Private CommState As String = String.Empty

    Private Function OpenComm() As Boolean
        Try
            Comm = New TcpClient("127.0.0.1", 22001)
            CommState = "Connected"
        Catch ex As Exception
            CommState = ex.ToString
            Return False
        End Try

        If Not Comm.Connected Then Return False
        _NetworkStream = Comm.GetStream()
        _NetworkStream.WriteTimeout = 1000
        _NetworkStream.ReadTimeout = 1000
        Return True
    End Function

    Private Sub StartSP()
        _NetworkStream.Write(Run_B, 0, Run_B.Length)

    End Sub

    Private Shared Function BytesToString(B() As Byte) As String
        Dim c(-1) As Char

        For iX As Integer = 0 To B.GetUpperBound(0)
            If B(iX) <> 0 Then
                ReDim Preserve c(c.Count)
                c(c.GetUpperBound(0)) = Chr(B(iX))
            End If
        Next
        Return New String(c)
    End Function

    Private Shared Function StringToBytes(S As String) As Byte()
        Dim b(-1) As Byte
        ReDim b(S.Length)
        System.Buffer.BlockCopy(S.ToCharArray, 0, b, 0, S.Length - 1)
        b(S.Length) = 13
        Return b

    End Function

    Private _RunB(-1) As Byte
    Private ReadOnly Property Run_B As Byte()
        Get
            If _RunB.GetUpperBound(0) < 0 Then
                Dim sRUN As String = "_RUN"
                Dim B(4) As Byte
                For iX As Integer = 0 To sRUN.Length - 1
                    B(iX) = CByte(Asc(sRUN(iX)))
                Next
                B(4) = CByte(13)
                _RunB = B

            End If
            Return _RunB
        End Get
    End Property


    Private _ANYTasksB(-1) As Byte
    Private ReadOnly Property ANYTasks_B As Byte()
        Get
            If _ANYTasksB.GetUpperBound(0) < 0 Then
                Dim AnyTasks As String = "ANY.TASKS?"
                Dim B(10) As Byte
                For iX As Integer = 0 To AnyTasks.Length - 1
                    B(iX) = CByte(Asc(AnyTasks(iX)))
                Next
                B(10) = CByte(13)
                _ANYTasksB = B

            End If
            Return _ANYTasksB
        End Get
    End Property
    Private _FileNameB(-1) As Byte
    Private ReadOnly Property FileName_B As Byte()
        Get
            If _ANYTasksB.GetUpperBound(0) < 0 Then
                Dim sX As String = "FILENAME"
                Dim B(8) As Byte
                For iX As Integer = 0 To sX.Length - 1
                    B(iX) = CByte(Asc(sX(iX)))
                Next
                B(8) = CByte(13)
                _FileNameB = B

            End If
            Return _FileNameB
        End Get
    End Property

    Private ReadOnly Property InitialCountdown As Integer
        Get
            Dim iX As Integer = CInt(Val(GetSetting("Supervisor Startup", "Initialization", "Time", "0")))
            If iX = 0 Then
                iX = 40
                SaveSetting("Supervisor Startup", "Initialization", "Time", iX.ToString)
            End If
            Return iX
        End Get
    End Property

    Private _Countdown As Integer = -32768
    Public Property Countdown As Integer
        Get
            If _Countdown = -32768 Then
                _Countdown = InitialCountdown + New Random(CInt(My.Computer.FileSystem.GetDriveInfo("C:\").AvailableFreeSpace Mod Integer.MaxValue)).Next(0, 40)
            End If
            Return _Countdown
        End Get
        Set(value As Integer)
            _Countdown = value
        End Set
    End Property

    Private CDStarted As Integer = -1
    Private Taskcount As Integer = -1
    Private StartupState As String = String.Empty
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If CDStarted = -1 Then CDStarted = Countdown
        Countdown -= 1
        Try
            If (Countdown <= CDStarted - 10) And (LastStart < Now) And (_NetworkStream Is Nothing) Then
                OpenComm()
                LastStart = DateAdd(DateInterval.Second, 1, Now)
            End If
            If Countdown Mod 10 = 0 AndAlso (_NetworkStream IsNot Nothing) Then
                Taskcount = RunningTasks()
            End If
        Catch ex As Exception
            CommState = "Error: " & ex.ToString
        End Try

        If Countdown = 0 Then
            If Taskcount > 0 Then
                StartupState = "Already Running"
            ElseIf _NetworkStream Is Nothing Then
                StartupState = "Unable to Communicate"
            Else
                Try
                    StartSP()
                    StartupState = "Started"
                    Taskcount = RunningTasks()
                Catch ex As Exception
                    StartupState = "Start Failed"
                End Try
            End If
        End If

        If Countdown = -20 Then
            If Taskcount <= 0 Then
                Timer1.Enabled = False
                Label1.Height = 200
                Button1.Visible = True
                Me.Height = 311
                Label1.Text = CommState
            Else
                Me.Close()
            End If
        Else
            Dim sX As String = String.Empty
            If Taskcount > 0 Then
                If StartupState <> String.Empty Then sX = StartupState & Environment.NewLine
                sX &= "Connected, Running (" & Taskcount.ToString & " Tasks)" & Environment.NewLine
                sX &= "Closing in " & ((Countdown + 20) / 10).ToString("0.0") & " Seconds"
            ElseIf Countdown <= 0 And Countdown >= -10 Then
                sX = StartupState
            ElseIf Taskcount = 0 Then
                If StartupState IsNot String.Empty Then
                    sX = StartupState & Environment.NewLine
                Else
                    sX = "Connected, Not Running" & Environment.NewLine
                End If
                If Countdown >= 0 Then
                    sX &= "Attempting in " & (Countdown / 10).ToString("0.0") & " Seconds"

                End If
            ElseIf Countdown >= 0 Then
                sX = "Not Connected" & Environment.NewLine
                sX &= "Starting in " & (Countdown / 10).ToString("0.0") & " Seconds"
            ElseIf Taskcount < 0 Then
                sX = CommState
            End If

            Label1.Text = sX
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
End Class
