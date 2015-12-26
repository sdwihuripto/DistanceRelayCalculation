Public Class ResultPage

    Public Sub New(ByVal a As Double, ByVal b As Double, ByVal c As Double, ByVal d As Double, ByVal e As Double, ByVal f As Double, ByVal g As Double, ByVal h As Double, ByVal i As Double, ByVal j As Double, ByVal k As Double, ByVal l As Double, ByVal m As Double, ByVal n As Double, ByVal o As Double)
        Me.InitializeComponent()
        Label16.Text = a
        Label17.Text = b
        Label18.Text = c
        Label19.Text = d
        Label20.Text = e
        Label21.Text = f
        Label22.Text = g
        Label23.Text = h
        Label24.Text = i
        Label25.Text = j
        Label26.Text = k
        Label27.Text = l
        Label28.Text = m
        Label29.Text = n
        Label30.Text = o
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
End Class