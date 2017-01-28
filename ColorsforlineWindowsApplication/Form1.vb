Imports System.Threading

Public Class Form1
    Dim CflL As New ColorsforlineLibrary.Main
    'Dim CflL As New Main
    Dim theSizeOfArray As Integer = 7
    Dim SelectBtn(1) As Integer

    Dim apppath As String = Application.StartupPath + "\config.ini"
    Dim highscore As Integer

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        savegame()
        Me.Dispose(True)
    End Sub

    '如果为空
    '    选中状态不为空
    '        取出之前选中的并判断是否可以移动
    '            可以移动
    '                刷新并清空选中状态
    '            不可以移动
    '                结束
    '    选中状态为空
    '        结束
    '如果不为空
    '    将目前的位置放入选中状态
    '    结束
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        libverToolStripMenuItem.Text = "lib v" & CflL.libver
        highscore = GetINI("board", "highscore", "0", apppath)

        If System.IO.File.Exists(apppath) = True Then
            loadgame()
        Else
            CflL.initialize()
        End If

        For Each con As Control In Me.Controls
            If TypeOf con Is Button Then
                con.Enabled = False
                CType(con, Button).Text = ""
            End If
        Next

        'CheckForIllegalCrossThreadCalls = False
        Refresh()
        SelectBtn(0) = -1
        SelectBtn(1) = -1

    End Sub
    Sub savegame()
        WriteINI("board", "score", CflL.score, apppath)
        If CflL.score >= highscore Then
            WriteINI("board", "highscore", CflL.score, apppath)
        End If
        For i = 0 To CflL.theSizeOfArray
            WriteINI("save", "line" & i, CflL.save(i), apppath)
        Next
        Dim color As String = ""
        For j = 0 To 2
            color &= CflL.color(j)
        Next
        WriteINI("save", "color", color, apppath)
    End Sub
    Sub loadgame()
        CflL.score = GetINI("board", "score", "", apppath)
        ToolStripStatusLabel3.Text = "最高分数:" & highscore
        For i = 0 To CflL.theSizeOfArray
            CflL.load(GetINI("save", "line" & i, "", apppath), i)
        Next
        For j = 0 To 2
            CflL.color(j) = Mid(GetINI("save", "color", "0", apppath), j + 1, 1)
        Next
    End Sub

    Private Sub 重新开始ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 重新开始ToolStripMenuItem.Click
        CflL.initialize()
        Refresh()
    End Sub
    Private Sub 关于ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 关于ToolStripMenuItem.Click
        MsgBox("15/5/31 v1.0 基本功能可用" & vbCrLf & "15/5/31晚 v1.1 添加预览下一步颜色功能" & vbCrLf &
               "15/6/7 v1.2 完成下一步生成三个随机时不会检查是否已超过或等于五个 以及 分数放入主类中" & vbCrLf &
               "15/6/20 v1.3 添加存档功能" & vbCrLf &
               "15/7/6 v1.4 棋子图片美化" & vbCrLf & "17/1/28 v1.5 因GDI绘图效率低下全面转向网页"_
               , MsgBoxStyle.OkOnly, "关于五彩连珠")
    End Sub

    Sub Button(ByVal name As String) ', ByVal color As System.Drawing.Color)
        Dim x As Integer = Mid(name, 7, 1)
        Dim y As Integer = Mid(name, 8, 1)
        If CflL.array(x, y) = 0 Then
            If SelectBtn(0) = -1 And SelectBtn(1) = -1 Then
            Else
                '先列再行
                Dim canmove As Boolean, eliminateChecklineScore As Integer
                canmove = CflL.canMove(SelectBtn(0), SelectBtn(1), x, y)
                If canmove = True Then
                    CflL.Move(SelectBtn(0), SelectBtn(1), x, y)
                    If CflL.canNextStep() = False Then
                        MsgBox("游戏结束", MsgBoxStyle.OkOnly)
                        CflL.initialize()
                    Else
                        eliminateChecklineScore = CflL.checkline(x, y)
                        If eliminateChecklineScore = 0 Then CflL.nextStep()
                        'Else
                        '    score += eliminateChecklineScore * 2
                        'End If

                    End If
                    SelectBtn(0) = -1
                    SelectBtn(1) = -1
                Else
                    If CflL.canNextStep() = False Then
                        MsgBox("游戏结束", MsgBoxStyle.OkOnly)
                        CflL.initialize()
                    End If
                End If
            End If
        Else
            SelectBtn(0) = Microsoft.VisualBasic.Mid(name, 7, 1)
            SelectBtn(1) = Microsoft.VisualBasic.Mid(name, 8, 1)
        End If
        Refresh()
    End Sub
    Overrides Sub Refresh()
        'If BackgroundWorker1.IsBusy Then BackgroundWorker1.CancelAsync()
        'BackgroundWorker1.RunWorkerAsync()

        Dim Color As String
        For i = 1 To (CflL.theSizeOfArray + 1) ^ 2
            For Each con As Control In Me.Controls
                If TypeOf con Is Button Then
                    Color = CflL.array(Mid(con.Name, 7, 1), Mid(con.Name, 8, 1))
                    If Color <> 0 Then
                        CType(con, Button).BackColor = InttoColor(Color)
                        'CType(con, Button).Image = InttoColor(Color)
                    Else
                        CType(con, Button).BackColor = Nothing
                        'CType(con, Button).Image = Nothing
                    End If
                    'CType(con, Button).BackColor = System.Drawing.Color.Transparent
                    'CType(con, Button).BackColor = InttoColor(Color)
                End If
            Next
        Next

        'ToolStripStatusLabel1.Text = "已选中" & SelectBtn(0) & SelectBtn(1)
        ToolStripStatusLabel2.Text = "分数:" & CflL.score
        'PictureBox1.Image = InttoColor(CflL.color(0))
        'PictureBox2.Image = InttoColor(CflL.color(1))
        'PictureBox3.Image = InttoColor(CflL.color(2))
        PictureBox1.BackColor = InttoColor(CflL.color(0))
        PictureBox2.BackColor = InttoColor(CflL.color(1))
        PictureBox3.BackColor = InttoColor(CflL.color(2))

        If highscore < CflL.score Then
            ToolStripStatusLabel3.Text = "最高分数:" & CflL.score
        End If
    End Sub

    'Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
    '    Dim Color As String
    '    For i = 1 To (CflL.theSizeOfArray + 1) ^ 2
    '        For Each con As Control In Me.Controls
    '            If TypeOf con Is Button Then
    '                '如果请求取消 就结束线程
    '                If BackgroundWorker1.CancellationPending Then
    '                    e.Cancel = True
    '                End If

    '                Color = CflL.array(Mid(con.Name, 7, 1), Mid(con.Name, 8, 1))
    '                CType(con, Button).Text = ""
    '                If Color <> 0 Then
    '                    CType(con, Button).Image = InttoColor(Color)
    '                Else
    '                    CType(con, Button).Image = Nothing
    '                End If
    '                'CType(con, Button).BackColor = System.Drawing.Color.Transparent
    '                'CType(con, Button).BackColor = InttoColor(Color)
    '            End If
    '        Next
    '    Next
    'End Sub

    Private Sub Form1_Click(sender As Object, e As EventArgs) Handles Me.Click
        'ToolStripStatusLabel1.Text = CursorToBtn()
        Dim temp As String = CursorToBtn()
        If Mid(temp, 1, 6) = "button" Then
            Button(temp)
        End If
    End Sub
    Function CursorToBtn()
        Dim CurX As Integer = System.Windows.Forms.Cursor.Position.X
        Dim CurY As Integer = System.Windows.Forms.Cursor.Position.Y
        Dim FormX As Integer = Me.Left + Button00.Location.X + 9
        Dim FormY As Integer = Me.Top + Button00.Location.Y + Me.Height - Me.ClientSize.Height - 5
        If CurX > FormX And CurX < Me.Left + Button70.Location.X + 50 And
             CurY > FormY And CurY < Me.Top + Button07.Location.Y + 50 + Me.Height - Me.ClientSize.Height - 5 Then
            Dim x As Integer = System.Math.Truncate((CurX - FormX) / (50 + 6))
            Dim y As Integer = System.Math.Truncate((CurY - FormY) / (50 + 6))
            If CurX - FormX - x * 56 > 47 Or CurY - FormY - y * 56 > 47 Then
                Return ("未点到")
            Else
                Return ("button" & x & y)
            End If
        End If
        Return ("界外")
    End Function

    Function InttoColor(coordinate)
        Dim color As System.Drawing.Color
        Select Case coordinate
            Case 0
                color = Drawing.Color.White
            Case 1
                color = Drawing.Color.Green
            Case 2
                color = Drawing.Color.Blue
            Case 3
                color = Drawing.Color.Red
            Case 4
                'color = Drawing.Color.Yellow
                color = Drawing.Color.Aqua
            Case 5
                color = Drawing.Color.Purple
            Case Else
                color = Drawing.Color.Black
        End Select
        Return (color)
        'Dim color As System.Drawing.Bitmap
        'Select Case coordinate
        '    Case 0
        '        color = My.Resources.Background
        '    Case 1
        '        color = My.Resources.Green
        '    Case 2
        '        color = My.Resources.Purple
        '    Case 3
        '        color = My.Resources.Red
        '    Case 4
        '        color = My.Resources.Yellow
        '    Case 5
        '        color = My.Resources.Blue
        '    Case Else
        '        color = My.Resources.Grey
        'End Select
        'Return (color)

    End Function

    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Int32, ByVal lpFileName As String) As Int32
    Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Int32
    '定义读取配置文件函数
    Public Function GetINI(ByVal Section As String, ByVal AppName As String, ByVal lpDefault As String, ByVal FileName As String) As String
        Dim Str As String = LSet(Str, 256)
        GetPrivateProfileString(Section, AppName, lpDefault, Str, Len(Str), FileName)
        Return Microsoft.VisualBasic.Left(Str, InStr(Str, Chr(0)) - 1)
    End Function
    '定义写入配置文件函数
    Public Function WriteINI(ByVal Section As String, ByVal AppName As String, ByVal lpDefault As String, ByVal FileName As String) As Long
        WriteINI = WritePrivateProfileString(Section, AppName, lpDefault, FileName)
    End Function

End Class
