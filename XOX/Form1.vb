Public Class frmMain

    Dim toggle As Boolean
    Dim scoreA, scoreB As Integer
    Dim btnA, btnB, btnC As Button

    Dim counter As Integer = 0
    Dim colorToggle = False
    Dim prevBackColor As Color
    Dim winBackColor As New Color

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        ResetBoard()
        ResetScore()
        UpdateScore()
    End Sub

    Private Sub btnReset_Click1(sender As Object, e As EventArgs) Handles btnReset.MouseDown
        btnReset.ForeColor = Color.Green
    End Sub
    Private Sub btnReset_Click2(sender As Object, e As EventArgs) Handles btnReset.MouseUp
        btnReset.ForeColor = Color.White
    End Sub

    Private Sub btn1_Click(sender As Object, e As EventArgs) _
        Handles btn1.Click, btn2.Click, btn3.Click, _
                btn4.Click, btn5.Click, btn6.Click, _
                btn7.Click, btn8.Click, btn9.Click

        If sender.Text <> "" Then Return

        If toggle Then sender.Text = "X" Else sender.Text = "O"

        If HasHorizontalMatch() Then
        ElseIf HasVerticalMatch() Then
        ElseIf HasDiagonalMatch() Then
        Else
            If IsDraw() Then
                MsgBox("It's a draw!")
                ResetBoard()
            End If
        End If

        toggle = Not toggle
        NotifyWhosTurn()

    End Sub

    Function HasHorizontalMatch() As Boolean
        If IsSame(btn1, btn2, btn3) Then
            EndMatch(btn1, btn2, btn3)
            Return True
        End If
        If IsSame(btn4, btn5, btn6) Then
            EndMatch(btn4, btn5, btn6)
            Return True
        End If
        If IsSame(btn7, btn8, btn9) Then
            EndMatch(btn7, btn8, btn9)
            Return True
        End If
        Return False
    End Function

    Function HasVerticalMatch() As Boolean
        If IsSame(btn1, btn4, btn7) Then
            EndMatch(btn1, btn4, btn7)
            Return True
        End If
        If IsSame(btn2, btn5, btn8) Then
            EndMatch(btn2, btn5, btn8)
            Return True
        End If
        If IsSame(btn3, btn6, btn9) Then
            EndMatch(btn3, btn6, btn9)
            Return True
        End If
        Return False
    End Function

    Function HasDiagonalMatch()
        If IsSame(btn1, btn5, btn9) Then
            EndMatch(btn1, btn5, btn9)
            Return True
        End If
        Return False
    End Function

    Function IsDraw()
        Return Not (Empty(btn1) Or Empty(btn2) Or Empty(btn3) _
                 Or Empty(btn4) Or Empty(btn5) Or Empty(btn6) _
                 Or Empty(btn7) Or Empty(btn8) Or Empty(btn9))
    End Function

    Function Empty(ByVal button As Button) As Boolean
        Return button.Text = ""
    End Function

    Sub EndMatch(ByRef a As Button, ByRef b As Button, ByRef c As Button)

        btnA = a
        btnB = b
        btnC = c

        gbButtons.Enabled = False

        Timer1.Enabled = True
        Timer1.Start()
    End Sub

    Sub ResetScore()
        scoreA = 0
        scoreB = 0
    End Sub

    Sub WhoWin()
        If toggle Then scoreB += 1 Else scoreA += 1
    End Sub

    Sub UpdateScore()
        lblScore.Text = scoreA & " : " & scoreB
    End Sub
    Sub Notify(ByVal msg As String)
        lblNotify.Text = msg
    End Sub

    Sub NotifyWhosTurn()
        If toggle Then
            Notify("Player 2 turn")
        Else
            Notify("Player 1 turn")
        End If
    End Sub

    Function GetWinner() As String
        If toggle Then Return "Player 2" Else Return "Player 1"
    End Function

    Sub ResetBoard()
        Const DEF_VAL = ""
        btn1.Text = DEF_VAL
        btn2.Text = DEF_VAL
        btn3.Text = DEF_VAL
        btn4.Text = DEF_VAL
        btn5.Text = DEF_VAL
        btn6.Text = DEF_VAL
        btn7.Text = DEF_VAL
        btn8.Text = DEF_VAL
        btn9.Text = DEF_VAL
    End Sub

    Function IsSame(ByVal a As Button, ByVal b As Button, ByVal c As Button)
        If a.Text = "" Or b.Text = "" Or c.Text = "" Then Return False

        Return a.Text = b.Text And a.Text = c.Text
    End Function

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ResetBoard()
        prevBackColor = btn1.BackColor
        winBackColor = Color.DarkGreen
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        If colorToggle Then
            btnA.BackColor = prevBackColor
            btnB.BackColor = prevBackColor
            btnC.BackColor = prevBackColor
        Else
            btnA.BackColor = winBackColor
            btnB.BackColor = winBackColor
            btnC.BackColor = winBackColor
        End If


        colorToggle = Not colorToggle
        counter += 1

        If counter >= 4 Then
            counter = 0
            colorToggle = False

            btnA.BackColor = prevBackColor
            btnB.BackColor = prevBackColor
            btnC.BackColor = prevBackColor

            gbButtons.Enabled = True
            Timer1.Enabled = False
            Timer1.Stop()

            WhoWin()
            UpdateScore()
            ResetBoard()
        End If

    End Sub
End Class
