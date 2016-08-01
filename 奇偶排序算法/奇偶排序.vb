Public Class 奇偶排序
    Private Const NumberCount As Byte = 19
    Private Const NumberMax As Byte = 200
    Dim NumberList(0 To NumberCount - 1) As Byte

    Private Sub 按钮_生成随机数(sender As Object, e As EventArgs) Handles Button1.Click
        生成新的随机数组()
        刷新显示列表()
    End Sub

    Private Sub 按钮_奇偶排序(sender As Object, e As EventArgs) Handles Button2.Click
        奇偶排序()
        刷新显示列表()
    End Sub

    Private Sub 奇偶排序()
        Dim Sorted As Boolean = True

        Do While Sorted
            Sorted = False
            '遍历奇数
            For Index = 1 To NumberCount - 1 Step 2
                If NumberList(Index) > NumberList(Index + 1) Then
                    异或算法交换值(NumberList(Index), NumberList(Index + 1))
                    Sorted = True
                End If
            Next
            '遍历偶数
            For Index = 0 To NumberCount - 2 Step 2
                If NumberList(Index) > NumberList(Index + 1) Then
                    异或算法交换值(NumberList(Index), NumberList(Index + 1))
                    Sorted = True
                End If
            Next
        Loop
    End Sub

    Private Sub 生成新的随机数组()
        For Index As Byte = 0 To NumberCount - 1
            NumberList(Index) = CByte(VBMath.Rnd() * NumberMax)
        Next
    End Sub

    Private Sub 刷新显示列表()
        ListBox1.Items.Clear()
        For Index As Byte = 0 To NumberCount - 1
            ListBox1.Items.Add(NumberList(Index))
        Next
    End Sub

    Private Sub 异或算法交换值(ByRef Number1, ByRef Number2)
        Number1 = Number1 Xor Number2
        Number2 = Number2 Xor Number1
        Number1 = Number1 Xor Number2
    End Sub

End Class
