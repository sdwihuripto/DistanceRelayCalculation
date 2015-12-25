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

        Dim RL1 As Double = Double.Parse(txtResistansiL1.Text) * Double.Parse(txtPanjangL1.Text)
        Dim RL2 As Double = Double.Parse(txtResistansiL2.Text) * Double.Parse(txtPanjangL2.Text)
        Dim RL3 As Double = Double.Parse(txtResistansiL3.Text) * Double.Parse(txtPanjangL3.Text)
        Dim RL4 As Double = Double.Parse(txtResistansiL4.Text) * Double.Parse(txtPanjangL4.Text)

        Dim XL1 As Double = Double.Parse(txtReaktansiL1.Text) * Double.Parse(txtPanjangL1.Text)
        Dim XL2 As Double = Double.Parse(txtReaktansiL2.Text) * Double.Parse(txtPanjangL2.Text)
        Dim XL3 As Double = Double.Parse(txtReaktansiL3.Text) * Double.Parse(txtPanjangL3.Text)
        Dim XL4 As Double = Double.Parse(txtReaktansiL4.Text) * Double.Parse(txtPanjangL4.Text)

        Dim XTrf As Double = (Double.Parse(impedance.Text) * Double.Parse(voltageLevel.Text)) / (Double.Parse(mvaRating.Text) * 100)
        Dim CT As Double = Double.Parse(CTp.Text) / Double.Parse(CTs.Text)
        Dim PT As Double = Double.Parse(PTp.Text) / Double.Parse(PTs.Text)
        Dim n As Double = CT / PT

        Dim K As Double = Double.Parse(infeed.Text)

        Dim ZL1Abs As Double = Math.Sqrt(Math.Pow(RL1, 2) + Math.Pow(XL1, 2))
        Dim ZL2Abs As Double = Math.Sqrt(Math.Pow(RL2, 2) + Math.Pow(XL2, 2))
        Dim ZL3Abs As Double = Math.Sqrt(Math.Pow(RL3, 2) + Math.Pow(XL3, 2))
        Dim ZL4Abs As Double = Math.Sqrt(Math.Pow(RL4, 2) + Math.Pow(XL4, 2))

        Dim Z1PAbs As Double = 0.8 * ZL1Abs
        Dim Z1SAbs As Double = n * Z1PAbs
        ' THIS IS THE OUTPUT, TRUE

        Dim ThetaPH1 As Double = Math.Atan(XL1 / RL1) * (180 / Math.PI)
        ' THIS IS THE OUTPUT, TRUE

        Dim Z2minAbs As Double = 1.2 * ZL1Abs
        Dim Z2mak1Abs As Double = 0.8 * (ZL1Abs + (K * 0.8 * ZL2Abs))
        Dim ZTrfAbs As Double = 0.8 * Math.Sqrt(Math.Pow(RL1, 2) + Math.Pow((XL1 + (0.5 * XTrf)), 2))

        Dim Z21makAbs As Double
        If Z2mak1Abs > Z2minAbs Then
            Z21makAbs = Z2mak1Abs
        Else
            Z21makAbs = Z2minAbs
        End If

        Dim Z2PAbs As Double
        If Z21makAbs < ZTrfAbs Then
            Z2PAbs = Z21makAbs
        Else
            Z2PAbs = ZTrfAbs
        End If

        Dim Z2SAbs As Double = n * Z2PAbs
        ' THIS IS THE OUTPUT, FALSE

        Dim Z3minAbs As Double = 1.2 * (ZL1Abs + K * ZL3Abs)
        Dim Z3mak1Abs As Double = 0.8 * (ZL1Abs + (K * 1.2 * ZL3Abs))
        Dim Z3mak2Abs As Double = 0.8 * (ZL1Abs + (0.8 * (ZL3Abs + 0.8 * ZL4Abs) * K))
        Dim Z3TrfAbs As Double = 0.8 * Math.Sqrt(Math.Pow(RL1, 2) + Math.Pow(XL1 + (0.8 * XTrf), 2))

        Dim Z31Abs As Double
        If Z3mak1Abs > Z3mak2Abs Then
            Z31Abs = Z3mak1Abs
        Else
            Z31Abs = Z3mak2Abs
        End If

        Dim Z32Abs As Double
        If Z31Abs > Z3minAbs Then
            Z32Abs = Z31Abs
        Else
            Z32Abs = Z3minAbs
        End If

        Dim Z3PAbs As Double
        If Z32Abs > Z3TrfAbs Then
            Z3PAbs = Z3TrfAbs
        Else
            Z3PAbs = Z32Abs
        End If

        Dim Z3SAbs As Double = n * Z3PAbs
        ' THIS IS THE OUTPUT, FALSE

        If ComboBox4.SelectedIndex = 0 Then
            My.Application.Log.WriteEntry(ThetaPH1)
            My.Application.Log.WriteEntry("Z1SAbs: " & Z1SAbs & " Z2SAbs: " & Z2SAbs & " Z3SAbs: " & Z3SAbs)
        ElseIf ComboBox4.SelectedIndex = 1 Then
        End If
    End Sub
End Class