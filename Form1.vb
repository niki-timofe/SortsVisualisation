Public Class Form1

    Dim Array(49) As Integer
    Dim Bars(49) As Object

    Sub ReMass(ByRef Mass() As Integer)
        Mass(0) = 47
        Mass(1) = 79
        Mass(2) = 127
        Mass(3) = 40
        Mass(4) = 66
        Mass(5) = 77
        Mass(6) = 14
        Mass(7) = 184
        Mass(8) = 73
        Mass(9) = 2
        Mass(10) = 56
        Mass(11) = 57
        Mass(12) = 103
        Mass(13) = 7
        Mass(14) = 5
        Mass(15) = 167
        Mass(16) = 154
        Mass(17) = 163
        Mass(18) = 64
        Mass(19) = 67
        Mass(20) = 75
        Mass(21) = 183
        Mass(22) = 101
        Mass(23) = 115
        Mass(24) = 109
        Mass(25) = 5
        Mass(26) = 116
        Mass(27) = 72
        Mass(28) = 79
        Mass(29) = 18
        Mass(30) = 156
        Mass(31) = 161
        Mass(32) = 33
        Mass(33) = 140
        Mass(34) = 81
        Mass(35) = 164
        Mass(36) = 136
        Mass(37) = 145
        Mass(38) = 180
        Mass(39) = 181
        Mass(40) = 3
        Mass(41) = 43
        Mass(42) = 32
        Mass(43) = 26
        Mass(44) = 99
        Mass(45) = 109
        Mass(46) = 32
        Mass(47) = 89
        Mass(48) = 172
        Mass(49) = 157
    End Sub

    Sub BarsToBBars()
        Bars(0) = PictureBox1
        Bars(1) = PictureBox2
        Bars(2) = PictureBox3
        Bars(3) = PictureBox4
        Bars(4) = PictureBox5
        Bars(5) = PictureBox6
        Bars(6) = PictureBox7
        Bars(7) = PictureBox8
        Bars(8) = PictureBox9
        Bars(9) = PictureBox10
        Bars(10) = PictureBox11
        Bars(11) = PictureBox12
        Bars(12) = PictureBox13
        Bars(13) = PictureBox14
        Bars(14) = PictureBox15
        Bars(15) = PictureBox16
        Bars(16) = PictureBox17
        Bars(17) = PictureBox18
        Bars(18) = PictureBox19
        Bars(19) = PictureBox20
        Bars(20) = PictureBox21
        Bars(21) = PictureBox22
        Bars(22) = PictureBox23
        Bars(23) = PictureBox24
        Bars(24) = PictureBox25
        Bars(25) = PictureBox26
        Bars(26) = PictureBox27
        Bars(27) = PictureBox28
        Bars(28) = PictureBox29
        Bars(29) = PictureBox30
        Bars(30) = PictureBox31
        Bars(31) = PictureBox32
        Bars(32) = PictureBox33
        Bars(33) = PictureBox34
        Bars(34) = PictureBox35
        Bars(35) = PictureBox36
        Bars(36) = PictureBox37
        Bars(37) = PictureBox38
        Bars(38) = PictureBox39
        Bars(39) = PictureBox40
        Bars(40) = PictureBox41
        Bars(41) = PictureBox42
        Bars(42) = PictureBox43
        Bars(43) = PictureBox44
        Bars(44) = PictureBox45
        Bars(45) = PictureBox46
        Bars(46) = PictureBox47
        Bars(47) = PictureBox48
        Bars(48) = PictureBox49
        Bars(49) = PictureBox50

    End Sub

    Sub ArrayToBars()
        For i = 0 To UBound(Bars)
            Bars(i).Width = Array(i)
        Next
    End Sub

    Sub Swap(ByRef Val1, ByRef Val2, ByVal Num, ByVal Num2)
        If Num <> Num2 Then
            Bars(Num).BackColor = Color.GreenYellow
            Bars(Num).BackgroundImage = PictureBox52.BackgroundImage
            Bars(Num2).BackColor = Color.GreenYellow
            Bars(Num2).BackgroundImage = PictureBox52.BackgroundImage
        End If

        Dim Proc
        Proc = Val1
        Val1 = Val2
        Val2 = Proc
        ArrayToBars()
        System.Threading.Thread.Sleep(TrackBar1.Value)
        If Num <> Num2 Then
            Bars(Num).BackColor = Color.Red
            Bars(Num).BackgroundImage = PictureBox51.BackgroundImage
            Bars(Num2).BackColor = Color.Red
            Bars(Num2).BackgroundImage = PictureBox51.BackgroundImage
        End If
    End Sub

    Sub Sort(ByRef a() As Integer)
        Dim S As Integer
        For k = 0 To UBound(a) - 1
            S = 0
            For i = 0 To UBound(a) - k - 1
                If a(i) > a(i + 1) Then
                    S += 1
                    Swap(a(i), a(i + 1), i, i + 1)
                End If
            Next
            If S = 0 Then Exit For
        Next
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ReMass(Array)
        ArrayToBars()
        Button1.Enabled = False
        Button2.Enabled = False
        Select Case True
            Case RadioButton1.Checked
                BackgroundWorker1.RunWorkerAsync()
            Case RadioButton2.Checked
                BackgroundWorker2.RunWorkerAsync()
        End Select
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Sort(Array)
        Button1.Enabled = True
        Button2.Enabled = True
        ArrayToBars()
        ArrayToBars()
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ReMass(Array)
        BarsToBBars()
        ArrayToBars()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        ReMass(Array)
        ArrayToBars()
        TrackBar1.Value = 20
    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        qSort(Array)
        Button1.Enabled = True
        Button2.Enabled = True
        ArrayToBars()
        ArrayToBars()
    End Sub

    'сортировка qSort>>>

    Function partition(ByRef a() As Integer, ByVal left As Integer, ByVal right As Integer, ByRef pivot As Integer)
        Dim i
        Dim piv
        Dim store
        piv = a(pivot)
        Swap(a(right - 1), a(pivot), right - 1, pivot)
        store = left
        For i = left To right - 2
            If a(i) <= piv Then
                Swap(a(store), a(i), store, i)
                store = store + 1
            End If
        Next
        Swap(a(right - 1), a(store), right - 1, store)

        Return store
    End Function

    Function getpivot(ByRef a() As Integer, ByVal left As Integer, ByVal right As Integer)
        Return New System.Random().Next(left, right - 1)
    End Function

    Sub quicksort(ByRef a() As Integer, ByVal left As Integer, ByVal right As Integer)
        Dim pivot As Integer
        If right - left > 1 Then
            pivot = getpivot(a, left, right)
            pivot = partition(a, left, right, pivot)
            quicksort(a, left, pivot)
            quicksort(a, pivot + 1, right)
        End If
    End Sub

    Sub qSort(ByVal a() As Integer)
        Dim i
        Dim ii
        For i = 0 To a.Length() - 1
            ii = New System.Random().Next(0, a.Length() - 1)
            If i <> ii Then
                Swap(a(i), a(ii), i, ii)
            End If
        Next

        quicksort(a, 0, a.Length())
    End Sub
    '<<<сортировка qSort


    Private Sub RadioButton1_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged

        Select Case True
            Case RadioButton1.Checked
                LinkLabel1.Text = "Пузырьком"
                Label3.Text = "Это простой алгоритм сортировки."
            Case RadioButton2.Checked
                LinkLabel1.Text = "qSort"
                Label3.Text = "Это широко известный алгоритм" & vbCrLf & "сортировки."
        End Select

    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Select Case True
            Case RadioButton1.Checked
                About_Sorts.TabControl1.SelectedTab = About_Sorts.TabPage1
            Case RadioButton2.Checked
                About_Sorts.TabControl1.SelectedTab = About_Sorts.TabPage2
        End Select
        About_Sorts.Show()
    End Sub
End Class
