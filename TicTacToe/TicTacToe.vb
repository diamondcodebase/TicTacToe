Public Class TicTacToe
    Dim intzet As Integer = 0
    Dim round As Integer = 1
    Dim Result As String = ""
    Dim Level As String = ""
    Dim probra(23) As Integer

    Private Sub TicTacToe_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ResetB(Button1)
        ResetB(Button2)
        ResetB(Button3)
        ResetB(Button4)
        ResetB(Button5)
        ResetB(Button6)
        ResetB(Button7)
        ResetB(Button8)
        ResetB(Button9)
        ComboBoxLv.Text = "Idle"
    End Sub

    Private Sub ButtonStart_Click(sender As Object, e As EventArgs) Handles ButtonStart.Click
        LabelTurns.Text = "Turn : " & round
        LabelTurns.Visible = True
        Button1.Enabled = True
        Button2.Enabled = True
        Button3.Enabled = True
        Button4.Enabled = True
        Button5.Enabled = True
        Button6.Enabled = True
        Button7.Enabled = True
        Button8.Enabled = True
        Button9.Enabled = True
        Level = ComboBoxLv.Text
        ComboBoxLv.Enabled = False

    End Sub

    Private Sub ButtonReset_Click(sender As Object, e As EventArgs) Handles ButtonReset.Click
        ResetB(Button1)
        ResetB(Button2)
        ResetB(Button3)
        ResetB(Button4)
        ResetB(Button5)
        ResetB(Button6)
        ResetB(Button7)
        ResetB(Button8)
        ResetB(Button9)
        intzet = 0
        round = 1
        Result = ""
        LabelTurns.Visible = False
        ComboBoxLv.text = "Beginner"
        ComboBoxLv.Enabled = True

    End Sub

    Private Sub YouPlays(sender As Object, s As System.EventArgs) Handles Button8.Click, Button7.Click, Button9.Click, Button5.Click, Button4.Click, Button6.Click, Button2.Click, Button1.Click, Button3.Click
        LabelTurns.Text = "Turn : " & round

        If Level = "Expert" Then
            'Comp player first
            Computerplays_Expert()

            'You play
            sender.text = "O"
            sender.enabled = False
            sender.BackColor = Color.Blue
            intzet += 1
            CheckRound()
            If Result = "Y" Or Result = "C" Or Result = "D" Then
                Exit Sub
            End If
            round += 1

        Else
            'You play
            sender.text = "O"
            sender.enabled = False
            sender.BackColor = Color.Blue
            intzet += 1
            CheckRound()
            If Result = "Y" Or Result = "C" Or Result = "D" Then
                Exit Sub
            End If

            'Computer plays
            If Level = "Idle" Then
                Computerplays_Idle()
            ElseIf Level = "Intermediate" Then
                Computerplays_Intermediate()
            End If

            intzet += 1
            CheckRound()
            If Result = "Y" Or Result = "C" Or Result = "D" Then
                Exit Sub
            End If
            round += 1

        End If

    End Sub


    Sub Computerplays_Idle()
        Dim objRandom As New Random
        Dim intRandom As Integer

        Do While True
            intRandom = objRandom.Next(1, 10)

            If intRandom = 1 And Button1.Enabled = True Then
                Comp_Take(Button1)
                Exit Sub
            End If
            If intRandom = 2 And Button2.Enabled = True Then
                Comp_Take(Button2)
                Exit Sub
            End If
            If intRandom = 3 And Button3.Enabled = True Then
                Comp_Take(Button3)
                Exit Sub
            End If
            If intRandom = 4 And Button4.Enabled = True Then
                Comp_Take(Button4)
                Exit Sub
            End If
            If intRandom = 5 And Button5.Enabled = True Then
                Comp_Take(Button5)
                Exit Sub
            End If
            If intRandom = 6 And Button6.Enabled = True Then
                Comp_Take(Button6)
                Exit Sub
            End If
            If intRandom = 7 And Button7.Enabled = True Then
                Comp_Take(Button7)
                Exit Sub
            End If
            If intRandom = 8 And Button8.Enabled = True Then
                Comp_Take(Button8)
                Exit Sub
            End If
            If intRandom = 9 And Button9.Enabled = True Then
                Comp_Take(Button9)
                Exit Sub
            End If
        Loop

    End Sub

    Sub Computerplays_Intermediate()
        Dim objRandom As New Random
        Dim intRandom As Integer

        Do While True
            If round = 1 Then
                intRandom = objRandom.Next(1, 10)

                If intRandom = 1 And Button1.Enabled = True Then
                    Comp_Take(Button1)
                    Exit Sub
                End If
                If intRandom = 2 And Button2.Enabled = True Then
                    Comp_Take(Button2)
                    Exit Sub
                End If
                If intRandom = 3 And Button3.Enabled = True Then
                    Comp_Take(Button3)
                    Exit Sub
                End If
                If intRandom = 4 And Button4.Enabled = True Then
                    Comp_Take(Button4)
                    Exit Sub
                End If
                If intRandom = 5 And Button5.Enabled = True Then
                    Comp_Take(Button5)
                    Exit Sub
                End If
                If intRandom = 6 And Button6.Enabled = True Then
                    Comp_Take(Button6)
                    Exit Sub
                End If
                If intRandom = 7 And Button7.Enabled = True Then
                    Comp_Take(Button7)
                    Exit Sub
                End If
                If intRandom = 8 And Button8.Enabled = True Then
                    Comp_Take(Button8)
                    Exit Sub
                End If
                If intRandom = 9 And Button9.Enabled = True Then
                    Comp_Take(Button9)
                    Exit Sub
                End If

            ElseIf round = 2 Then
                ' 2nd round Defence approach
                If Comp_StrategyD2(Button1, Button2, Button3) Then
                    Exit Sub
                ElseIf Comp_StrategyD2(Button1, Button4, Button7) Then
                    Exit Sub
                ElseIf Comp_StrategyD2(Button1, Button5, Button9) Then
                    Exit Sub
                ElseIf Comp_StrategyD2(Button2, Button5, Button8) Then
                    Exit Sub
                ElseIf Comp_StrategyD2(Button3, Button6, Button9) Then
                    Exit Sub
                ElseIf Comp_StrategyD2(Button3, Button5, Button7) Then
                    Exit Sub
                ElseIf Comp_StrategyD2(Button4, Button5, Button6) Then
                    Exit Sub
                ElseIf Comp_StrategyD2(Button7, Button8, Button9) Then
                    Exit Sub
                End If

                '2nd round Aggressive approach
                If Comp_StrategyA2(Button1, Button2, Button3) Then
                    Exit Sub
                ElseIf Comp_StrategyA2(Button1, Button4, Button7) Then
                    Exit Sub
                ElseIf Comp_StrategyA2(Button1, Button5, Button9) Then
                    Exit Sub
                ElseIf Comp_StrategyA2(Button2, Button5, Button8) Then
                    Exit Sub
                ElseIf Comp_StrategyA2(Button3, Button6, Button9) Then
                    Exit Sub
                ElseIf Comp_StrategyA2(Button3, Button5, Button7) Then
                    Exit Sub
                ElseIf Comp_StrategyA2(Button4, Button5, Button6) Then
                    Exit Sub
                ElseIf Comp_StrategyA2(Button7, Button8, Button9) Then
                    Exit Sub
                End If

            ElseIf round = 3 Then
                '3rd round Defence
                If Comp_StrategyD3(Button1, Button2, Button3) Then
                    Exit Sub
                ElseIf Comp_StrategyD3(Button1, Button4, Button7) Then
                    Exit Sub
                ElseIf Comp_StrategyD3(Button1, Button5, Button9) Then
                    Exit Sub
                ElseIf Comp_StrategyD3(Button2, Button5, Button8) Then
                    Exit Sub
                ElseIf Comp_StrategyD3(Button3, Button6, Button9) Then
                    Exit Sub
                ElseIf Comp_StrategyD3(Button3, Button5, Button7) Then
                    Exit Sub
                ElseIf Comp_StrategyD3(Button4, Button5, Button6) Then
                    Exit Sub
                ElseIf Comp_StrategyD3(Button7, Button8, Button9) Then
                    Exit Sub
                Else
                    Do While True
                        intRandom = objRandom.Next(1, 10)

                        If intRandom = 1 And Button1.Enabled = True Then
                            Comp_Take(Button1)
                            Exit Sub
                        End If
                        If intRandom = 2 And Button2.Enabled = True Then
                            Comp_Take(Button2)
                            Exit Sub
                        End If
                        If intRandom = 3 And Button3.Enabled = True Then
                            Comp_Take(Button3)
                            Exit Sub
                        End If
                        If intRandom = 4 And Button4.Enabled = True Then
                            Comp_Take(Button4)
                            Exit Sub
                        End If
                        If intRandom = 5 And Button5.Enabled = True Then
                            Comp_Take(Button5)
                            Exit Sub
                        End If
                        If intRandom = 6 And Button6.Enabled = True Then
                            Comp_Take(Button6)
                            Exit Sub
                        End If
                        If intRandom = 7 And Button7.Enabled = True Then
                            Comp_Take(Button7)
                            Exit Sub
                        End If
                        If intRandom = 8 And Button8.Enabled = True Then
                            Comp_Take(Button8)
                            Exit Sub
                        End If
                        If intRandom = 9 And Button9.Enabled = True Then
                            Comp_Take(Button9)
                            Exit Sub
                        End If
                    Loop
                End If

                '3rd round Aggressive
                If Comp_StrategyA3(Button1, Button2, Button3) Then
                    Exit Sub
                ElseIf Comp_StrategyA3(Button1, Button4, Button7) Then
                    Exit Sub
                ElseIf Comp_StrategyA3(Button1, Button5, Button9) Then
                    Exit Sub
                ElseIf Comp_StrategyA3(Button2, Button5, Button8) Then
                    Exit Sub
                ElseIf Comp_StrategyA3(Button3, Button6, Button9) Then
                    Exit Sub
                ElseIf Comp_StrategyA3(Button3, Button5, Button7) Then
                    Exit Sub
                ElseIf Comp_StrategyA3(Button4, Button5, Button6) Then
                    Exit Sub
                ElseIf Comp_StrategyA3(Button7, Button8, Button9) Then
                    Exit Sub
                End If

            ElseIf round = 4 Then
                If Comp_StrategyD3(Button1, Button2, Button3) Then
                    Exit Sub
                ElseIf Comp_StrategyD3(Button1, Button4, Button7) Then
                    Exit Sub
                ElseIf Comp_StrategyD3(Button1, Button5, Button9) Then
                    Exit Sub
                ElseIf Comp_StrategyD3(Button2, Button5, Button8) Then
                    Exit Sub
                ElseIf Comp_StrategyD3(Button3, Button6, Button9) Then
                    Exit Sub
                ElseIf Comp_StrategyD3(Button3, Button5, Button7) Then
                    Exit Sub
                ElseIf Comp_StrategyD3(Button4, Button5, Button6) Then
                    Exit Sub
                ElseIf Comp_StrategyD3(Button7, Button8, Button9) Then
                    Exit Sub
                End If
            Else
                intRandom = objRandom.Next(1, 10)

                If intRandom = 1 And Button1.Enabled = True Then
                    Comp_Take(Button1)
                    Exit Sub
                End If
                If intRandom = 2 And Button2.Enabled = True Then
                    Comp_Take(Button2)
                    Exit Sub
                End If
                If intRandom = 3 And Button3.Enabled = True Then
                    Comp_Take(Button3)
                    Exit Sub
                End If
                If intRandom = 4 And Button4.Enabled = True Then
                    Comp_Take(Button4)
                    Exit Sub
                End If
                If intRandom = 5 And Button5.Enabled = True Then
                    Comp_Take(Button5)
                    Exit Sub
                End If
                If intRandom = 6 And Button6.Enabled = True Then
                    Comp_Take(Button6)
                    Exit Sub
                End If
                If intRandom = 7 And Button7.Enabled = True Then
                    Comp_Take(Button7)
                    Exit Sub
                End If
                If intRandom = 8 And Button8.Enabled = True Then
                    Comp_Take(Button8)
                    Exit Sub
                End If
                If intRandom = 9 And Button9.Enabled = True Then
                    Comp_Take(Button9)
                    Exit Sub
                End If
            End If
        Loop
    End Sub

    Sub Computerplays_Expert()
        Dim objRandom As New Random
        Dim intRandom As Integer

        Do While True
            If round = 1 Then
                intRandom = objRandom.Next(1, 10)

                If intRandom = 1 And Button1.Enabled = True Then
                    Comp_Take(Button1)
                    Exit Sub
                End If
                If intRandom = 2 And Button2.Enabled = True Then
                    Comp_Take(Button2)
                    Exit Sub
                End If
                If intRandom = 3 And Button3.Enabled = True Then
                    Comp_Take(Button3)
                    Exit Sub
                End If
                If intRandom = 4 And Button4.Enabled = True Then
                    Comp_Take(Button4)
                    Exit Sub
                End If
                If intRandom = 5 And Button5.Enabled = True Then
                    Comp_Take(Button5)
                    Exit Sub
                End If
                If intRandom = 6 And Button6.Enabled = True Then
                    Comp_Take(Button6)
                    Exit Sub
                End If
                If intRandom = 7 And Button7.Enabled = True Then
                    Comp_Take(Button7)
                    Exit Sub
                End If
                If intRandom = 8 And Button8.Enabled = True Then
                    Comp_Take(Button8)
                    Exit Sub
                End If
                If intRandom = 9 And Button9.Enabled = True Then
                    Comp_Take(Button9)
                    Exit Sub
                End If

            ElseIf round = 2 Then
                ' 2nd round Defence approach
                If Comp_StrategyD2(Button1, Button2, Button3) Then
                    Exit Sub
                ElseIf Comp_StrategyD2(Button1, Button4, Button7) Then
                    Exit Sub
                ElseIf Comp_StrategyD2(Button1, Button5, Button9) Then
                    Exit Sub
                ElseIf Comp_StrategyD2(Button2, Button5, Button8) Then
                    Exit Sub
                ElseIf Comp_StrategyD2(Button3, Button6, Button9) Then
                    Exit Sub
                ElseIf Comp_StrategyD2(Button3, Button5, Button7) Then
                    Exit Sub
                ElseIf Comp_StrategyD2(Button4, Button5, Button6) Then
                    Exit Sub
                ElseIf Comp_StrategyD2(Button7, Button8, Button9) Then
                    Exit Sub
                End If

                '2nd round Aggressive approach
                If Comp_StrategyA2(Button1, Button2, Button3) Then
                    Exit Sub
                ElseIf Comp_StrategyA2(Button1, Button4, Button7) Then
                    Exit Sub
                ElseIf Comp_StrategyA2(Button1, Button5, Button9) Then
                    Exit Sub
                ElseIf Comp_StrategyA2(Button2, Button5, Button8) Then
                    Exit Sub
                ElseIf Comp_StrategyA2(Button3, Button6, Button9) Then
                    Exit Sub
                ElseIf Comp_StrategyA2(Button3, Button5, Button7) Then
                    Exit Sub
                ElseIf Comp_StrategyA2(Button4, Button5, Button6) Then
                    Exit Sub
                ElseIf Comp_StrategyA2(Button7, Button8, Button9) Then
                    Exit Sub
                End If

            ElseIf round = 3 Then
                '3rd round Defence
                If Comp_StrategyD3(Button1, Button2, Button3) Then
                    Exit Sub
                ElseIf Comp_StrategyD3(Button1, Button4, Button7) Then
                    Exit Sub
                ElseIf Comp_StrategyD3(Button1, Button5, Button9) Then
                    Exit Sub
                ElseIf Comp_StrategyD3(Button2, Button5, Button8) Then
                    Exit Sub
                ElseIf Comp_StrategyD3(Button3, Button6, Button9) Then
                    Exit Sub
                ElseIf Comp_StrategyD3(Button3, Button5, Button7) Then
                    Exit Sub
                ElseIf Comp_StrategyD3(Button4, Button5, Button6) Then
                    Exit Sub
                ElseIf Comp_StrategyD3(Button7, Button8, Button9) Then
                    Exit Sub
                End If

                '3rd round Aggressive
                If Comp_StrategyA3(Button1, Button2, Button3) Then
                    Exit Sub
                ElseIf Comp_StrategyA3(Button1, Button4, Button7) Then
                    Exit Sub
                ElseIf Comp_StrategyA3(Button1, Button5, Button9) Then
                    Exit Sub
                ElseIf Comp_StrategyA3(Button2, Button5, Button8) Then
                    Exit Sub
                ElseIf Comp_StrategyA3(Button3, Button6, Button9) Then
                    Exit Sub
                ElseIf Comp_StrategyA3(Button3, Button5, Button7) Then
                    Exit Sub
                ElseIf Comp_StrategyA3(Button4, Button5, Button6) Then
                    Exit Sub
                ElseIf Comp_StrategyA3(Button7, Button8, Button9) Then
                    Exit Sub
                Else
                    Do While True
                        intRandom = objRandom.Next(1, 10)

                        If intRandom = 1 And Button1.Enabled = True Then
                            Comp_Take(Button1)
                            Exit Sub
                        End If
                        If intRandom = 2 And Button2.Enabled = True Then
                            Comp_Take(Button2)
                            Exit Sub
                        End If
                        If intRandom = 3 And Button3.Enabled = True Then
                            Comp_Take(Button3)
                            Exit Sub
                        End If
                        If intRandom = 4 And Button4.Enabled = True Then
                            Comp_Take(Button4)
                            Exit Sub
                        End If
                        If intRandom = 5 And Button5.Enabled = True Then
                            Comp_Take(Button5)
                            Exit Sub
                        End If
                        If intRandom = 6 And Button6.Enabled = True Then
                            Comp_Take(Button6)
                            Exit Sub
                        End If
                        If intRandom = 7 And Button7.Enabled = True Then
                            Comp_Take(Button7)
                            Exit Sub
                        End If
                        If intRandom = 8 And Button8.Enabled = True Then
                            Comp_Take(Button8)
                            Exit Sub
                        End If
                        If intRandom = 9 And Button9.Enabled = True Then
                            Comp_Take(Button9)
                            Exit Sub
                        End If
                    Loop
                End If

            ElseIf round = 4 Then
                If Comp_StrategyD3(Button1, Button2, Button3) Then
                    Exit Sub
                ElseIf Comp_StrategyD3(Button1, Button4, Button7) Then
                    Exit Sub
                ElseIf Comp_StrategyD3(Button1, Button5, Button9) Then
                    Exit Sub
                ElseIf Comp_StrategyD3(Button2, Button5, Button8) Then
                    Exit Sub
                ElseIf Comp_StrategyD3(Button3, Button6, Button9) Then
                    Exit Sub
                ElseIf Comp_StrategyD3(Button3, Button5, Button7) Then
                    Exit Sub
                ElseIf Comp_StrategyD3(Button4, Button5, Button6) Then
                    Exit Sub
                ElseIf Comp_StrategyD3(Button7, Button8, Button9) Then
                    Exit Sub
                End If
            Else
                intRandom = objRandom.Next(1, 10)

                If intRandom = 1 And Button1.Enabled = True Then
                    Comp_Take(Button1)
                    Exit Sub
                End If
                If intRandom = 2 And Button2.Enabled = True Then
                    Comp_Take(Button2)
                    Exit Sub
                End If
                If intRandom = 3 And Button3.Enabled = True Then
                    Comp_Take(Button3)
                    Exit Sub
                End If
                If intRandom = 4 And Button4.Enabled = True Then
                    Comp_Take(Button4)
                    Exit Sub
                End If
                If intRandom = 5 And Button5.Enabled = True Then
                    Comp_Take(Button5)
                    Exit Sub
                End If
                If intRandom = 6 And Button6.Enabled = True Then
                    Comp_Take(Button6)
                    Exit Sub
                End If
                If intRandom = 7 And Button7.Enabled = True Then
                    Comp_Take(Button7)
                    Exit Sub
                End If
                If intRandom = 8 And Button8.Enabled = True Then
                    Comp_Take(Button8)
                    Exit Sub
                End If
                If intRandom = 9 And Button9.Enabled = True Then
                    Comp_Take(Button9)
                    Exit Sub
                End If
            End If
        Loop

    End Sub



    Sub CheckRound()
        If intzet < 9 Then
            If (Button1.Text = "O" And Button2.Text = "O" And Button3.Text = "O") _
            Or (Button4.Text = "O" And Button5.Text = "O" And Button6.Text = "O") _
            Or (Button7.Text = "O" And Button8.Text = "O" And Button9.Text = "O") _
            Or (Button1.Text = "O" And Button4.Text = "O" And Button7.Text = "O") _
            Or (Button2.Text = "O" And Button5.Text = "O" And Button8.Text = "O") _
            Or (Button3.Text = "O" And Button6.Text = "O" And Button9.Text = "O") _
            Or (Button1.Text = "O" And Button5.Text = "O" And Button9.Text = "O") _
            Or (Button3.Text = "O" And Button5.Text = "O" And Button7.Text = "O") Then
                MessageBox.Show("You Win~")
                Result = "Y"
                Exit Sub

            ElseIf (Button1.Text = "X" And Button2.Text = "X" And Button3.Text = "X") _
            Or (Button4.Text = "X" And Button5.Text = "X" And Button6.Text = "X") _
            Or (Button7.Text = "X" And Button8.Text = "X" And Button9.Text = "X") _
            Or (Button1.Text = "X" And Button4.Text = "X" And Button7.Text = "X") _
            Or (Button2.Text = "X" And Button5.Text = "X" And Button8.Text = "X") _
            Or (Button3.Text = "X" And Button6.Text = "X" And Button9.Text = "X") _
            Or (Button1.Text = "X" And Button5.Text = "X" And Button9.Text = "X") _
            Or (Button3.Text = "X" And Button5.Text = "X" And Button7.Text = "X") Then
                MessageBox.Show("You Lose...")
                Result = "C"
                Exit Sub
            End If
        Else
            MessageBox.Show("Draw Game~")
            Result = "D"
            Exit Sub
        End If

    End Sub

    Function ResetB(ByVal button As Button) As Integer
        button.Text = ""
        button.Enabled = False
        button.BackColor = Color.White

        Return 1
    End Function

    Function Comp_Take(ByRef C As Button) As Boolean
        C.Text = "X"
        C.Enabled = False
        C.BackColor = Color.Red
        Return True
    End Function

    Function Comp_StrategyA2(ByRef A As Button, ByRef B As Button, ByRef C As Button) As Boolean
        If A.Text = "X" And B.Enabled = True And C.Enabled = True Then
            Comp_Take(B)
            Return True
        ElseIf B.Text = "X" And A.Enabled = True And C.Enabled = True Then
            Comp_Take(C)
            Return True
        ElseIf C.Text = "X" And A.Enabled = True And B.Enabled = True Then
            Comp_Take(A)
            Return True
        Else
            Return False
        End If
    End Function

    Function Comp_StrategyD2(ByRef A As Button, ByRef B As Button, ByRef C As Button) As Boolean
        If A.Text = "O" And B.Text = "O" And C.Enabled = True Then
            Comp_Take(C)
            Return True
        ElseIf B.Text = "O" And C.Text = "O" And A.Enabled = True Then
            Comp_Take(A)
            Return True
        ElseIf A.Text = "O" And C.Text = "O" And B.Enabled = True Then
            Comp_Take(B)
            Return True
        Else
            Return False
        End If
    End Function

    Function Comp_StrategyA3(ByRef A As Button, ByRef B As Button, ByRef C As Button) As Boolean
        If A.Text = "X" And B.Text = "X" And C.Enabled = True Then
            Comp_Take(C)
            Return True
        ElseIf B.Text = "X" And C.Text = "X" And A.Enabled = True Then
            Comp_Take(A)
            Return True
        ElseIf C.Text = "X" And A.Text = "X" And B.Enabled = True Then
            Comp_Take(B)
            Return True
        Else
            Return False
        End If
    End Function

    Function Comp_StrategyD3(ByRef A As Button, ByRef B As Button, ByRef C As Button) As Boolean
        If A.Text = "O" And B.Text = "O" And C.Enabled = True Then
            Comp_Take(C)
            Return True
        ElseIf B.Text = "O" And C.Text = "O" And A.Enabled = True Then
            Comp_Take(A)
            Return True
        ElseIf A.Text = "O" And C.Text = "O" And B.Enabled = True Then
            Comp_Take(B)
            Return True
        Else
            Return False
        End If
    End Function

End Class
