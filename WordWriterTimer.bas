Attribute VB_Name = "WordWriterTimer"

' Word 码字计时插件
' 功能：统计总字数、写入字数、写入时间、写入速度、总时间、总速度
' 说明：粘贴、删除、输入都会反映到写入字数中

Public isRunning As Boolean
Public isStarted As Boolean
Public startTime As Double
Public elapsedSeconds As Double
Public startTotalTime As Date
Public baselineWords As Long
Public currentWords As Long
Public nextUpdate As Date

' 开始/继续计时
Sub StartWriting()
    If isRunning Then Exit Sub
    If Not isStarted Then
        baselineWords = ActiveDocument.ComputeStatistics(Statistic:=wdStatisticWords)
        currentWords = baselineWords
        startTotalTime = Now
        elapsedSeconds = 0
        isStarted = True
    End If
    isRunning = True
    startTime = Timer
    UpdateStatus
    ScheduleUpdate
End Sub

' 暂停计时
Sub PauseWriting()
    If Not isRunning Then Exit Sub
    elapsedSeconds = elapsedSeconds + (Timer - startTime)
    isRunning = False
    CancelUpdate
    UpdateStatus
End Sub

' 重置计时
Sub ResetWriting()
    isRunning = False
    isStarted = False
    elapsedSeconds = 0
    baselineWords = ActiveDocument.ComputeStatistics(Statistic:=wdStatisticWords)
    currentWords = baselineWords
    startTotalTime = Now
    CancelUpdate
    UpdateStatus
End Sub

' 显示浮动统计窗口
Sub ShowStats()
    On Error Resume Next
    frmStatsSimple.Show vbModeless
    RefreshStatsForm
End Sub

' 刷新浮动窗口
Sub RefreshStatsForm()
    On Error Resume Next
    Dim netWords As Long
    netWords = currentWords - baselineWords

    Dim writeTimeSeconds As Double
    If isRunning Then
        writeTimeSeconds = elapsedSeconds + (Timer - startTime)
    Else
        writeTimeSeconds = elapsedSeconds
    End If
    If writeTimeSeconds < 0 Then writeTimeSeconds = writeTimeSeconds + 86400

    Dim totalTimeSeconds As Double
    If isStarted Then
        totalTimeSeconds = DateDiff("s", startTotalTime, Now)
    Else
        totalTimeSeconds = 0
    End If
    If totalTimeSeconds < 0 Then totalTimeSeconds = 0

    frmStatsSimple.lblTotal.Caption = "总字数: " & Format(currentWords, "#,##0")
    frmStatsSimple.lblNet.Caption = "写入字数: " & Format(netWords, "#,##0")
    frmStatsSimple.lblWriteTime.Caption = "写入时间: " & FormatSeconds(writeTimeSeconds)
    frmStatsSimple.lblWriteSpeed.Caption = "写入速度: " & Format(GetWriteSpeed, "#,##0.0") & " 字/小时"
    frmStatsSimple.lblTime.Caption = "总时间: " & FormatSeconds(totalTimeSeconds)
    frmStatsSimple.lblSpeed.Caption = "总速度: " & Format(GetTotalSpeed, "#,##0.0") & " 字/小时"
End Sub

' 更新状态栏和浮动窗口
Sub UpdateStatus()
    On Error Resume Next
    currentWords = ActiveDocument.ComputeStatistics(Statistic:=wdStatisticWords)

    Dim writeTimeSeconds As Double
    If isRunning Then
        writeTimeSeconds = elapsedSeconds + (Timer - startTime)
    Else
        writeTimeSeconds = elapsedSeconds
    End If
    If writeTimeSeconds < 0 Then writeTimeSeconds = writeTimeSeconds + 86400

    Dim totalTimeSeconds As Double
    If isStarted Then
        totalTimeSeconds = DateDiff("s", startTotalTime, Now)
    Else
        totalTimeSeconds = 0
    End If
    If totalTimeSeconds < 0 Then totalTimeSeconds = 0

    Dim netWords As Long
    netWords = currentWords - baselineWords

    Application.StatusBar = "总字数:" & Format(currentWords, "#,##0") & _
        " | 写入字数:" & Format(netWords, "#,##0") & _
        " | 写入时间:" & FormatSeconds(writeTimeSeconds) & _
        " | 写入速度:" & Format(GetWriteSpeed, "#,##0.0") & _
        " | 总时间:" & FormatSeconds(totalTimeSeconds) & _
        " | 总速度:" & Format(GetTotalSpeed, "#,##0.0") & "字/小时"

    RefreshStatsForm
End Sub

' 写入速度（字/小时）
Function GetWriteSpeed() As Double
    Dim writeTimeSeconds As Double
    If isRunning Then
        writeTimeSeconds = elapsedSeconds + (Timer - startTime)
    Else
        writeTimeSeconds = elapsedSeconds
    End If
    If writeTimeSeconds <= 0 Then
        GetWriteSpeed = 0
    Else
        Dim netWords As Long
        netWords = currentWords - baselineWords
        GetWriteSpeed = netWords / (writeTimeSeconds / 3600)
    End If
End Function

' 总速度（字/小时）
Function GetTotalSpeed() As Double
    Dim totalTimeSeconds As Double
    If isStarted Then
        totalTimeSeconds = DateDiff("s", startTotalTime, Now)
    Else
        totalTimeSeconds = 0
    End If
    If totalTimeSeconds <= 0 Then
        GetTotalSpeed = 0
    Else
        Dim netWords As Long
        netWords = currentWords - baselineWords
        GetTotalSpeed = netWords / (totalTimeSeconds / 3600)
    End If
End Function

' 安排下一次刷新
Sub ScheduleUpdate()
    nextUpdate = Now + TimeValue("00:00:01")
    Application.OnTime When:=nextUpdate, Name:="WordWriterTimer.UpdateStatusLoop"
End Sub

' 取消刷新（Word 不支持真正取消，靠 isRunning 控制）
Sub CancelUpdate()
    ' Word OnTime cannot be cancelled; UpdateStatusLoop checks isRunning
End Sub

' 定时循环
Sub UpdateStatusLoop()
    If Not isRunning Then Exit Sub
    UpdateStatus
    ScheduleUpdate
End Sub

' 清除状态栏
Sub ClearStatus()
    Application.StatusBar = False
End Sub

' 秒数格式化为 HH:MM:SS
Function FormatSeconds(seconds As Double) As String
    Dim total As Long
    total = Int(seconds)
    Dim h As Long, m As Long, s As Long
    h = total \ 3600
    m = (total Mod 3600) \ 60
    s = total Mod 60
    FormatSeconds = Format(h, "00") & ":" & Format(m, "00") & ":" & Format(s, "00")
End Function
