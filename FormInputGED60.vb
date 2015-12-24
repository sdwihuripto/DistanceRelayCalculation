Public Class FormInputGED60

    Private Sub FormInput_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub AddBtn1_Click(sender As Object, e As EventArgs) Handles AddBtn1.Click
        FormAddInfo.AddInfo1.Text = "Additional Info1 (L1)"
        FormAddInfo.AddInfo2.Text = "Additional Info2 (L1)"
        FormAddInfo.AddInfo3.Text = "Additional Info3 (L1)"
        FormAddInfo.ShowDialog()
    End Sub

    Private Sub AddBtn2_Click(sender As Object, e As EventArgs) Handles AddBtn2.Click
        FormAddInfo.txtCCC1.Visible = False
        FormAddInfo.TxtCCC2.Visible = False
        FormAddInfo.txtCCC3.Visible = False
        FormAddInfo.AddInfo1.Text = "Additional Info1 (L2)"
        FormAddInfo.AddInfo2.Text = "Additional Info2 (L2)"
        FormAddInfo.AddInfo3.Text = "Additional Info3 (L2)"
        FormAddInfo.ShowDialog()
    End Sub

    Private Sub AddBtn3_Click(sender As Object, e As EventArgs) Handles AddBtn3.Click
        FormAddInfo.txtCCC1.Visible = False
        FormAddInfo.TxtCCC2.Visible = False
        FormAddInfo.TxtCCC3.Visible = False
        FormAddInfo.AddInfo1.Text = "Additional Info1 (L3)"
        FormAddInfo.AddInfo2.Text = "Additional Info2 (L3)"
        FormAddInfo.AddInfo3.Text = "Additional Info3 (L3)"
        FormAddInfo.ShowDialog()
    End Sub

    Private Sub AddBtn4_Click(sender As Object, e As EventArgs) Handles AddBtn4.Click
        FormAddInfo.txtCCC1.Visible = False
        FormAddInfo.TxtCCC2.Visible = False
        FormAddInfo.TxtCCC3.Visible = False
        FormAddInfo.AddInfo1.Text = "Additional Info1 (L4)"
        FormAddInfo.AddInfo2.Text = "Additional Info2 (L4)"
        FormAddInfo.AddInfo3.Text = "Additional Info3 (L4)"
        FormAddInfo.ShowDialog()
    End Sub

    Private Sub cmdCalculate_Click(sender As Object, e As EventArgs) Handles cmdCalculate.Click

    End Sub
End Class