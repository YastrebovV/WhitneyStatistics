Public Class Pass
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "gpp" Then
            Main.CloseStat = True
            Close()
            Main.Close()
        Else
            MsgBox("Пароль введён не верно", 16, "Попробуй ещё")
            Main.InTrey()
            Close()
        End If
    End Sub
    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            If TextBox1.Text = "gpp" Then
                Main.CloseStat = True
                Close()
                Main.Close()
            Else
                MsgBox("Пароль введён не верно", 16, "Попробуй ещё")
                Main.InTrey()
                Close()
            End If
        End If
    End Sub
End Class