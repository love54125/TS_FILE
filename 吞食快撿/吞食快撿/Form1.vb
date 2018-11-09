Public Class Form1
    Dim sock1 As New Net.Sockets.TcpClient
    Dim opassword As String
    Dim chklog As Boolean = False
    Dim logout As Boolean = True
    Dim pick(), times As Byte, count As Long

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If My.Computer.FileSystem.FileExists("SERVER.INI") Then
            Dim server() As String, i As Short
            ComboBox1.Items.Clear()
            server = Split(My.Computer.FileSystem.ReadAllText("SERVER.INI", System.Text.Encoding.Default), vbCrLf)
            For i = 0 To UBound(server)
                ComboBox1.Items.Add(server(i))
            Next
        End If
        ComboBox1.SelectedIndex = 11
        WebBrowser1.DocumentText = "<html><head><style>body{margin:0px,0px,0px,0px;}</style></head><body><a href=" & Chr(34) & "ymsgr:sendIM?jses88001" & Chr(34) & "><img border=0 src=" & Chr(34) & "http://opi.yahoo.com/online?u=jses88001&m=g&t=1" & Chr(34) & " alt=" & Chr(34) & "jses88001" & Chr(34) & "></a></body></html>"
        Timer3.Stop()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim ip() As String
        ip = Split(ComboBox1.SelectedItem, "*")
        If TextBox2.Text = "" Then
            MsgBox("帳號不可為空！", 48)
            TextBox2.Focus()
            Exit Sub
        ElseIf opassword = "" Then
            MsgBox("密碼不可為空！", 48)
            TextBox3.Focus()
            Exit Sub
        End If
        If sock1.Connected = True Then
            sock1.Client.Disconnect(False)
            sock1.Close()
            sock1 = New Net.Sockets.TcpClient
            Label2.Text = "登入狀態：斷開連接"
            chklog = False
        Else
            sock1.NoDelay = True
            sock1.Connect(ip(1), 6414)
            Label2.Text = "登入狀態：正在連接"
            Login()
            chklog = True
        End If
        Timer1.Start()
    End Sub

    Sub Login()
        If sock1.Connected = False Then
            Dim ip() As String
            ip = Split(ComboBox1.SelectedItem, "*")
            sock1.Connect(ip(1), 6414)
            sock1.NoDelay = True
        End If

        Dim pack() As Byte, tmp, length, country, version, id, password As String, password_tmp() As Char
        Dim i As Integer
        country = Buwei(Hex(Asc(TextBox1.Text.Substring(0, 1))), 2) & Buwei(Hex(Asc(TextBox1.Text.Substring(1, 1))), 2)
        version = Buwei(Hex(NumericUpDown1.Value), 4)
        id = Buwei(Hex(Val(TextBox2.Text)), 8)
        password_tmp = opassword.ToCharArray
        password = ""
        For i = 0 To password_tmp.GetUpperBound(0)
            password &= Buwei(Hex(Asc(password_tmp(i))), 2)
        Next
        length = Buwei(Hex(Len(password) \ 2 + 10), 4)
        tmp = "59E9ACADAD59E9" & length & "AC" & Buwei(Hex(Len(password) \ 2), 2) & id & country & version & password
        ReDim pack(Len(tmp) \ 2 - 1)
        For i = 0 To Len(tmp) \ 2 - 1
            pack(i) = Val("&h" & Mid(tmp, 1 + i * 2, 2))
        Next
        If sock1.Connected = True Then
            sock1.Client.Send(pack)
        End If
        Timer3.Start()
    End Sub

    Function Buwei(ByVal data As String, ByVal length As Integer) As String
        Dim i As Integer
        Dim tmp As String = ""
        Do Until Len(data) = length
            data = "0" & data
        Loop
        For i = 1 To length \ 2
            If Len(Hex(Val("&h" & Mid(data, 1 + (i - 1) * 2, 2)) Xor 173)) = 1 Then
                tmp = "0" & Hex(Val("&h" & Mid(data, 1 + (i - 1) * 2, 2)) Xor 173) & tmp
            Else
                tmp = Hex(Val("&h" & Mid(data, 1 + (i - 1) * 2, 2)) Xor 173) & tmp
            End If
        Next
        Buwei = tmp
    End Function

    Private Sub TextBox2_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox2.LostFocus
        TextBox2.Text = Val(TextBox2.Text)
    End Sub

    Private Sub TextBox3_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox3.GotFocus
        TextBox3.Text = ""
    End Sub

    Private Sub TextBox3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = 13 Then Button1.PerformClick() : Button1.Focus()
    End Sub

    Private Sub TextBox3_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox3.LostFocus
        TextBox3.Text = "**********"
    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged
        If TextBox3.Text <> "" AndAlso TextBox3.Text <> "**********" Then opassword = TextBox3.Text
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Timer1.Stop()
        chklogin()
    End Sub

    Sub chklogin()
        If chklog = True AndAlso logout = False Then
            Button1.Text = "登出"
            ComboBox1.Enabled = False
            TextBox1.Enabled = False
            NumericUpDown1.Enabled = False
            TextBox2.Enabled = False
            TextBox3.Enabled = False
        Else
            Button1.Text = "登入"
            ComboBox1.Enabled = True
            TextBox1.Enabled = True
            NumericUpDown1.Enabled = True
            TextBox2.Enabled = True
            TextBox3.Enabled = True
        End If
    End Sub

    Private Sub Timer3_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer3.Tick
        If sock1.Connected = True AndAlso chklog = True Then
            If sock1.Client.Available = 0 Then Exit Sub
            Dim tmp() As Byte, data As String, i As Integer
            ReDim tmp(sock1.Client.Available - 1)
            sock1.Client.Receive(tmp)
            data = ""
            For i = 0 To tmp.GetUpperBound(0)
                data &= Hex(tmp(i))
            Next
            If InStr(data, "59E9AFADB9A5") Then
                chklog = True
                logout = False
                Label2.Text = "登入狀態：成功登入"
            Else
                If InStr(data, "59E9AFADACAB") Then
                    chklog = False
                    logout = True
                    Label2.Text = "登入狀態：帳密錯誤"
                    sock1.Client.Disconnect(False)
                    sock1.Close()
                    sock1 = New Net.Sockets.TcpClient
                Else
                    If InStr(data, "59E9AFADAD") Then
                        chklog = False
                        logout = True
                        Dim error_id As Integer = (Val("&h" & data.Substring(InStr(data, "59E9AFADAD") + 9, 2)) Xor 173)
                        Label2.Text = "登入狀態：錯誤" & error_id
                        sock1.Client.Disconnect(False)
                        sock1.Close()
                        sock1 = New Net.Sockets.TcpClient
                    End If
                End If
            End If
            chklogin()
        End If
    End Sub

    Private Sub NumericUpDown3_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown3.ValueChanged
        If NumericUpDown3.Value > NumericUpDown4.Value Then NumericUpDown4.Value = NumericUpDown3.Value
    End Sub

    Private Sub NumericUpDown4_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown4.ValueChanged
        If NumericUpDown4.Value < NumericUpDown3.Value Then NumericUpDown3.Value = NumericUpDown4.Value
    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        NumericUpDown5.Enabled = Not RadioButton1.Checked
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If sock1.Connected = False Then Exit Sub
        Timer2.Interval = NumericUpDown2.Value
        times = NumericUpDown5.Value
        count = 0
        If Timer2.Enabled = True Then
            Timer2.Enabled = False
            NumericUpDown2.Enabled = True
            NumericUpDown3.Enabled = True
            NumericUpDown4.Enabled = True
            NumericUpDown5.Enabled = Not RadioButton1.Checked
            RadioButton1.Enabled = True
            RadioButton2.Enabled = True
            CheckBox1.Enabled = True
            Button2.Text = "發送"
            Label13.ForeColor = Color.Red
        Else
            Timer2.Enabled = True
            NumericUpDown2.Enabled = False
            NumericUpDown3.Enabled = False
            NumericUpDown4.Enabled = False
            NumericUpDown5.Enabled = False
            RadioButton1.Enabled = False
            RadioButton2.Enabled = False
            CheckBox1.Enabled = False
            Button2.Text = "停止"
        End If
        Dim tmp As String
        Dim i As Integer
        tmp = ""
        For i = NumericUpDown3.Value To NumericUpDown4.Value
            tmp &= "59E9A9ADBAAF" & Buwei(Hex(i), 4)
        Next
        ReDim pick(Len(tmp) \ 2 - 1)
        For i = 0 To Len(tmp) \ 2 - 1
            pick(i) = Val("&h" & Mid(tmp, 1 + i * 2, 2))
        Next
    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        If (count = times AndAlso RadioButton2.Checked = True) Or sock1.Connected = False Then
            Timer2.Enabled = False
            NumericUpDown2.Enabled = True
            NumericUpDown3.Enabled = True
            NumericUpDown4.Enabled = True
            NumericUpDown5.Enabled = Not RadioButton1.Checked
            RadioButton1.Enabled = True
            RadioButton2.Enabled = True
            CheckBox1.Enabled = True
            Button2.Text = "發送"
            Label13.ForeColor = Color.Red
            Exit Sub
        End If
        If sock1.Connected = True Then
            sock1.Client.Send(pick)
            Label13.ForeColor = Color.Green
        End If
        If RadioButton2.Checked = True Then count += 1
    End Sub

    Private Sub Timer4_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer4.Tick
        WebBrowser1.DocumentText = "<html><head><style>body{margin:0px,0px,0px,0px;}</style></head><body><a href=" & Chr(34) & "ymsgr:sendIM?jses88001" & Chr(34) & "><img border=0 src=" & Chr(34) & "http://opi.yahoo.com/online?u=jses88001&m=g&t=1" & Chr(34) & " alt=" & Chr(34) & "jses88001" & Chr(34) & "></a></body></html>"
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            NumericUpDown3.Value = 1
            NumericUpDown4.Value = 255
            NumericUpDown3.Enabled = False
            NumericUpDown4.Enabled = False
        Else
            NumericUpDown3.Enabled = True
            NumericUpDown4.Enabled = True
        End If

    End Sub

    Private Sub Label15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label15.Click
        Process.Start("http://kkjs880603.myweb.hinet.net/tf/")
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim pack() As Byte, tmp As String
        Dim who As String, id As String
        Dim i As Integer
        who = InputBox("請輸入對方帳號的數字部分")
        id = Buwei(Hex(Val(who)), 8)
        tmp = "59E9ABADA0AC" & id
        ReDim pack(Len(tmp) \ 2 - 1)
        For i = 0 To Len(tmp) \ 2 - 1
            pack(i) = Val("&h" & Mid(tmp, 1 + i * 2, 2))
        Next
        If sock1.Connected = True Then sock1.Client.Send(pack)
        Timer3.Start()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim pack() As Byte, tmp As String
        Dim i As Integer
        tmp = "59E9A9ADB9A5ADAD"
        ReDim pack(Len(tmp) \ 2 - 1)
        For i = 0 To Len(tmp) \ 2 - 1
            pack(i) = Val("&h" & Mid(tmp, 1 + i * 2, 2))
        Next
        If chklog = True AndAlso logout = False Then
            sock1.Client.Send(pack)
            sock1.Client.Disconnect(False)
        End If
        MsgBox(sock1.Connected)
        chklogin()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim pack() As Byte, tmp As String
        Dim i As Integer
        tmp = "59E9AEAD8DACD0"
        ReDim pack(Len(tmp) \ 2 - 1)
        For i = 0 To Len(tmp) \ 2 - 1
            pack(i) = Val("&h" & Mid(tmp, 1 + i * 2, 2))
        Next
        If sock1.Connected = True Then sock1.Client.Send(pack)
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim pack() As Byte, tmp As String
        Dim i As Integer
        tmp = "59E9ABADBABDACADADAD"
        ReDim pack(Len(tmp) \ 2 - 1)
        For i = 0 To Len(tmp) \ 2 - 1
            pack(i) = Val("&h" & Mid(tmp, 1 + i * 2, 2))
        Next
        If sock1.Connected = True Then sock1.Client.Send(pack)
    End Sub

    Private Sub CheckBox1_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox1.Disposed
        sock1.Close()
    End Sub
End Class
