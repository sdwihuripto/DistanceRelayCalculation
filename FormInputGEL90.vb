Public Class FormInputGEL90

    Private Sub AddBtn1_Click(sender As Object, e As EventArgs) Handles AddBtn1.Click
        FormAddInfo.AddInfo1.Text = "Additional Info1 (L1)"
        FormAddInfo.ShowDialog()
    End Sub

    Private Sub AddBtn2_Click(sender As Object, e As EventArgs) Handles AddBtn2.Click
        FormAddInfo.txtCCC1.Visible = False
        FormAddInfo.AddInfo1.Text = "Additional Info1 (L2)"
        FormAddInfo.ShowDialog()
    End Sub

    Private Sub AddBtn3_Click(sender As Object, e As EventArgs) Handles AddBtn3.Click
        FormAddInfo.txtCCC1.Visible = False
        FormAddInfo.AddInfo1.Text = "Additional Info1 (L3)"
        FormAddInfo.ShowDialog()
    End Sub

    Private Sub AddBtn4_Click(sender As Object, e As EventArgs) Handles AddBtn4.Click
        FormAddInfo.txtCCC1.Visible = False
        FormAddInfo.AddInfo1.Text = "Additional Info1 (L4)"
        FormAddInfo.ShowDialog()
    End Sub
End Class