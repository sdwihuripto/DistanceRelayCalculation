Public Class HomePage

    Private Sub HomePage_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub PickDistanceButton_Click(sender As Object, e As EventArgs) Handles PickDistanceButton.Click
        If DistanceRelay.SelectedItem = Nothing Then
            MessageBox.Show("Please Pick Distance Relay Model", "No entry",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
            With Me.Focus()
            End With
        Else
            If DistanceRelay.SelectedItem = "GE D60" Then
                FormInputGED60.ShowDialog()
            End If
            If DistanceRelay.SelectedItem = "Siemens 7SA522" Then
                FormInputSiemens.ShowDialog()
            End If
            If DistanceRelay.SelectedItem = "ALSTOM MiCOM P442" Then
                FormInputAlstom.ShowDialog()
            End If
        End If

    End Sub
End Class
