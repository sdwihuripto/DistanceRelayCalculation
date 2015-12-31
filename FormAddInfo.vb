Public Class FormAddInfo

    Public lineTemp As Double
    Public formTemp As String

    Public Sub New(ByVal form As String, ByVal line As Double)
        Me.InitializeComponent()

        If line = 1 Then
            AddInfo1.Text = "Additional Info (L1)"
        ElseIf line = 2 Then
            Label1.Visible = False
            txtCCC1.Visible = False
            AddInfo1.Text = "Additional Info (L2)"
        ElseIf line = 3 Then
            Label1.Visible = False
            txtCCC1.Visible = False
            AddInfo1.Text = "Additional Info (L3)"
        ElseIf line = 4 Then
            Label1.Visible = False
            txtCCC1.Visible = False
            AddInfo1.Text = "Additional Info (L4)"
        End If

        lineTemp = line
        formTemp = form

        If formTemp.Equals("GED60") Then
            If lineTemp = 1 Then
                txtCCC1.Text = GlobalVariables.GED60L1CCC
                txtResistansi1.Text = GlobalVariables.GED60L1R
                txtReaktansi1.Text = GlobalVariables.GED60L1X
                txtPanjang1.Text = GlobalVariables.GED60L1L
            ElseIf lineTemp = 2 Then
                txtResistansi1.Text = GlobalVariables.GED60L2R
                txtReaktansi1.Text = GlobalVariables.GED60L2X
                txtPanjang1.Text = GlobalVariables.GED60L2L
            ElseIf lineTemp = 3 Then
                txtResistansi1.Text = GlobalVariables.GED60L3R
                txtReaktansi1.Text = GlobalVariables.GED60L3X
                txtPanjang1.Text = GlobalVariables.GED60L3L
            ElseIf lineTemp = 4 Then
                txtResistansi1.Text = GlobalVariables.GED60L4R
                txtReaktansi1.Text = GlobalVariables.GED60L4X
                txtPanjang1.Text = GlobalVariables.GED60L4L
            End If
        ElseIf formTemp.Equals("Siemens") Then
            If lineTemp = 1 Then
                txtCCC1.Text = GlobalVariables.SiemensL1CCC
                txtResistansi1.Text = GlobalVariables.SiemensL1R
                txtReaktansi1.Text = GlobalVariables.SiemensL1X
                txtPanjang1.Text = GlobalVariables.SiemensL1L
            ElseIf lineTemp = 2 Then
                txtResistansi1.Text = GlobalVariables.SiemensL2R
                txtReaktansi1.Text = GlobalVariables.SiemensL2X
                txtPanjang1.Text = GlobalVariables.SiemensL2L
            ElseIf lineTemp = 3 Then
                txtResistansi1.Text = GlobalVariables.SiemensL3R
                txtReaktansi1.Text = GlobalVariables.SiemensL3X
                txtPanjang1.Text = GlobalVariables.SiemensL3L
            ElseIf lineTemp = 4 Then
                txtResistansi1.Text = GlobalVariables.SiemensL4R
                txtReaktansi1.Text = GlobalVariables.SiemensL4X
                txtPanjang1.Text = GlobalVariables.SiemensL4L
            End If
        ElseIf formTemp.Equals("Alstom") Then
            If lineTemp = 1 Then
                txtCCC1.Text = GlobalVariables.AlstomL1CCC
                txtResistansi1.Text = GlobalVariables.AlstomL1R
                txtReaktansi1.Text = GlobalVariables.AlstomL1X
                txtPanjang1.Text = GlobalVariables.AlstomL1L
            ElseIf lineTemp = 2 Then
                txtResistansi1.Text = GlobalVariables.AlstomL2R
                txtReaktansi1.Text = GlobalVariables.AlstomL2X
                txtPanjang1.Text = GlobalVariables.AlstomL2L
            ElseIf lineTemp = 3 Then
                txtResistansi1.Text = GlobalVariables.AlstomL3R
                txtReaktansi1.Text = GlobalVariables.AlstomL3X
                txtPanjang1.Text = GlobalVariables.AlstomL3L
            ElseIf lineTemp = 4 Then
                txtResistansi1.Text = GlobalVariables.AlstomL4R
                txtReaktansi1.Text = GlobalVariables.AlstomL4X
                txtPanjang1.Text = GlobalVariables.AlstomL4L
            End If
        End If
    End Sub

    Private Sub Back_Click(sender As Object, e As EventArgs) Handles Back.Click
        Me.Close()
    End Sub

    Private Sub Save_Click(sender As Object, e As EventArgs) Handles Save.Click
        If formTemp.Equals("GED60") Then
            If lineTemp = 1 Then
                GlobalVariables.GED60L1CCC = Double.Parse(txtCCC1.Text)
                GlobalVariables.GED60L1R = Double.Parse(txtResistansi1.Text)
                GlobalVariables.GED60L1X = Double.Parse(txtReaktansi1.Text)
                GlobalVariables.GED60L1L = Double.Parse(txtPanjang1.Text)
            ElseIf lineTemp = 2 Then
                GlobalVariables.GED60L2R = Double.Parse(txtResistansi1.Text)
                GlobalVariables.GED60L2X = Double.Parse(txtReaktansi1.Text)
                GlobalVariables.GED60L2L = Double.Parse(txtPanjang1.Text)
            ElseIf lineTemp = 3 Then
                GlobalVariables.GED60L3R = Double.Parse(txtResistansi1.Text)
                GlobalVariables.GED60L3X = Double.Parse(txtReaktansi1.Text)
                GlobalVariables.GED60L3L = Double.Parse(txtPanjang1.Text)
            ElseIf lineTemp = 4 Then
                GlobalVariables.GED60L4R = Double.Parse(txtResistansi1.Text)
                GlobalVariables.GED60L4X = Double.Parse(txtReaktansi1.Text)
                GlobalVariables.GED60L4L = Double.Parse(txtPanjang1.Text)
            End If
        ElseIf formTemp.Equals("Siemens") Then
            If lineTemp = 1 Then
                GlobalVariables.SiemensL1CCC = Double.Parse(txtCCC1.Text)
                GlobalVariables.SiemensL1R = Double.Parse(txtResistansi1.Text)
                GlobalVariables.SiemensL1X = Double.Parse(txtReaktansi1.Text)
                GlobalVariables.SiemensL1L = Double.Parse(txtPanjang1.Text)
            ElseIf lineTemp = 2 Then
                GlobalVariables.SiemensL2R = Double.Parse(txtResistansi1.Text)
                GlobalVariables.SiemensL2X = Double.Parse(txtReaktansi1.Text)
                GlobalVariables.SiemensL2L = Double.Parse(txtPanjang1.Text)
            ElseIf lineTemp = 3 Then
                GlobalVariables.SiemensL3R = Double.Parse(txtResistansi1.Text)
                GlobalVariables.SiemensL3X = Double.Parse(txtReaktansi1.Text)
                GlobalVariables.SiemensL3L = Double.Parse(txtPanjang1.Text)
            ElseIf lineTemp = 4 Then
                GlobalVariables.SiemensL4R = Double.Parse(txtResistansi1.Text)
                GlobalVariables.SiemensL4X = Double.Parse(txtReaktansi1.Text)
                GlobalVariables.SiemensL4L = Double.Parse(txtPanjang1.Text)
            End If
        ElseIf formTemp.Equals("Alstom") Then
            If lineTemp = 1 Then
                GlobalVariables.AlstomL1CCC = Double.Parse(txtCCC1.Text)
                GlobalVariables.AlstomL1R = Double.Parse(txtResistansi1.Text)
                GlobalVariables.AlstomL1X = Double.Parse(txtReaktansi1.Text)
                GlobalVariables.AlstomL1L = Double.Parse(txtPanjang1.Text)
            ElseIf lineTemp = 2 Then
                GlobalVariables.AlstomL2R = Double.Parse(txtResistansi1.Text)
                GlobalVariables.AlstomL2X = Double.Parse(txtReaktansi1.Text)
                GlobalVariables.AlstomL2L = Double.Parse(txtPanjang1.Text)
            ElseIf lineTemp = 3 Then
                GlobalVariables.AlstomL3R = Double.Parse(txtResistansi1.Text)
                GlobalVariables.AlstomL3X = Double.Parse(txtReaktansi1.Text)
                GlobalVariables.AlstomL3L = Double.Parse(txtPanjang1.Text)
            ElseIf lineTemp = 4 Then
                GlobalVariables.AlstomL4R = Double.Parse(txtResistansi1.Text)
                GlobalVariables.AlstomL4X = Double.Parse(txtReaktansi1.Text)
                GlobalVariables.AlstomL4L = Double.Parse(txtPanjang1.Text)
            End If
        End If

        MessageBox.Show("Additional for Line " & lineTemp & " Saved")
        Me.Close()
    End Sub
End Class