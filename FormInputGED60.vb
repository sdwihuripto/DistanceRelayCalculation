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

        Dim ZL1Abs As Double
        Dim ZL2Abs As Double
        Dim ZL3Abs As Double
        Dim ZL4Abs As Double

        Dim RL10 As Double
        Dim RL20 As Double
        Dim RL30 As Double
        Dim RL40 As Double

        Dim XL10 As Double
        Dim XL20 As Double
        Dim XL30 As Double
        Dim XL40 As Double

        Dim ZL10 As Double
        Dim ZL20 As Double
        Dim ZL30 As Double
        Dim ZL40 As Double

        Dim ThetaPH1 As Double

        If ComboBox4.SelectedIndex = 0 Then
            ZL1Abs = Math.Sqrt(Math.Pow(RL1, 2) + Math.Pow(XL1, 2))
            ZL2Abs = Math.Sqrt(Math.Pow(RL2, 2) + Math.Pow(XL2, 2))
            ZL3Abs = Math.Sqrt(Math.Pow(RL3, 2) + Math.Pow(XL3, 2))
            ZL4Abs = Math.Sqrt(Math.Pow(RL4, 2) + Math.Pow(XL4, 2))

            ThetaPH1 = Math.Atan(XL1 / RL1) * (180 / Math.PI)
            ' THIS IS THE OUTPUT

            RL10 = (Double.Parse(txtResistansiL1.Text) + 0.15) * Double.Parse(txtPanjangL1.Text)
            RL20 = (Double.Parse(txtResistansiL2.Text) + 0.15) * Double.Parse(txtPanjangL2.Text)
            RL30 = (Double.Parse(txtResistansiL3.Text) + 0.15) * Double.Parse(txtPanjangL3.Text)
            RL40 = (Double.Parse(txtResistansiL4.Text) + 0.15) * Double.Parse(txtPanjangL4.Text)

            XL10 = 3 * Double.Parse(txtReaktansiL1.Text) * Double.Parse(txtPanjangL1.Text)
            XL20 = 3 * Double.Parse(txtReaktansiL2.Text) * Double.Parse(txtPanjangL2.Text)
            XL30 = 3 * Double.Parse(txtReaktansiL3.Text) * Double.Parse(txtPanjangL3.Text)
            XL40 = 3 * Double.Parse(txtReaktansiL4.Text) * Double.Parse(txtPanjangL4.Text)

            ZL10 = Math.Sqrt(Math.Pow(RL10, 2) + Math.Pow(XL10, 2))
            ZL20 = Math.Sqrt(Math.Pow(RL20, 2) + Math.Pow(XL20, 2))
            ZL30 = Math.Sqrt(Math.Pow(RL30, 2) + Math.Pow(XL30, 2))
            ZL40 = Math.Sqrt(Math.Pow(RL40, 2) + Math.Pow(XL40, 2))
        ElseIf ComboBox4.SelectedIndex = 1 Then


        End If

        Dim MVA As Double = Double.Parse(mvaRating.Text)
        Dim kV As Double = Double.Parse(voltageLevel.Text)
        Dim impedance As Double = Double.Parse(impedanceBox.Text)

        Dim XTrf As Double = (impedance * Math.Pow(kV, 2)) / (MVA * 100)
        Dim CT As Double = Double.Parse(CTp.Text) / Double.Parse(CTs.Text)
        Dim PT As Double = Double.Parse(PTp.Text) / Double.Parse(PTs.Text)
        Dim n As Double = CT / PT

        Dim K As Double = Double.Parse(infeed.Text)
        Dim Rf As Double = Double.Parse(ResistTower.Text)
        Dim iLoad As Double = Double.Parse(iLoadBox.Text)

        Dim Z1PAbs As Double = 0.8 * ZL1Abs
        Dim Z1SAbs As Double = n * Z1PAbs
        ' THIS IS THE OUTPUT

        Dim ThetaPH10 As Double = Math.Atan(XL10 / RL10) * (180 / Math.PI)

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
        ' THIS IS THE OUTPUT

        Dim Z3minAbs As Double = 1.2 * (ZL1Abs + ZL3Abs)
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
        ' THIS IS THE OUTPUT

        Dim Z1PgAbs As Double = 0.8 * ZL1Abs + Rf
        Dim Z1SgAbs As Double = n * Z1PAbs
        ' THIS IS THE OUTPUT

        Dim Z2mingAbs As Double = 1.2 * (ZL1Abs + Rf)
        Dim Z2mak1gAbs As Double = 0.8 * ((ZL1Abs + Rf) + (K * 0.8 * ZL2Abs))
        Dim ZTrfgAbs As Double = 0.8 * (Math.Sqrt(Math.Pow(RL1 + Rf, 2) + Math.Pow(XL1 + (0.5 * XTrf), 2)))

        Dim Z21makgAbs As Double
        If Z2mak1gAbs > Z2mingAbs Then
            Z21makgAbs = Z2mak1gAbs
        Else
            Z21makgAbs = Z2mingAbs
        End If

        Dim Z2PgAbs As Double
        If Z21makgAbs < ZTrfgAbs Then
            Z2PgAbs = Z21makgAbs
        Else
            Z2PgAbs = ZTrfgAbs
        End If

        Dim Z2SgAbs As Double = n * Z2PgAbs
        ' THIS IS THE OUTPUT

        Dim Z3mingAbs As Double = 1.2 * (ZL1Abs + Rf + ZL3Abs)
        Dim Z3mak1gAbs As Double = 0.8 * (ZL1Abs + Rf + (K * 1.2 * ZL3Abs))
        Dim Z3mak2gAbs As Double = 0.8 * (ZL1Abs + Rf + (0.8 * (ZL3Abs + 0.8 * ZL4Abs) * K))
        Dim Z3TrfgAbs As Double = 0.8 * (Math.Sqrt(Math.Pow(RL1, 2) + Math.Pow(XL1 + (0.8 * XTrf), 2)))

        Dim Z31gAbs As Double
        If Z3mak1gAbs > Z3mak2gAbs Then
            Z31gAbs = Z3mak1gAbs
        Else
            Z31gAbs = Z3mak2gAbs
        End If

        Dim Z32gAbs As Double
        If Z31gAbs > Z3mingAbs Then
            Z32Abs = Z31gAbs
        Else
            Z32gAbs = Z3mingAbs
        End If

        Dim Z3PgAbs As Double
        If Z32gAbs > Z3TrfgAbs Then
            Z3PgAbs = Z3TrfgAbs
        Else
            Z3PgAbs = Z32gAbs
        End If

        Dim Z3SgAbs = n * Z3PgAbs
        ' THIS IS THE OUTPUT

        ' BELOW THIS LINE IS THE OUTPUT
        Dim Tk1ph As Double = 0.1
        Dim Tk2ph As Double = 0.4
        Dim Tk3ph As Double = 1.6
        Dim Tk1g As Double = 0.1
        Dim Tk2g As Double = 0.4
        Dim Tk3g As Double = 1.6
        ' ABOVE THIS LINE IS THE OUTPUT

        Dim KoAbs As Double = Math.Sqrt(Math.Pow(RL10, 2) + Math.Pow(XL10, 2)) / Math.Sqrt(Math.Pow(RL1, 2) + Math.Pow(XL1, 2))
        Dim ThetaphKo As Double = (Math.Atan(XL10 / RL10) - Math.Atan(XL1 / RL1)) * (180 / Math.PI)

        Dim R2 As Double = (kV * n * 0.5 * 1000) / (iLoad * Math.Sqrt(3))
        Dim ZL As Double = (kV * n * 0.5 * 1000) / (iLoad * Math.Sqrt(3))

        ' BELOW THIS LINE IS THE OUTPUT
        Dim BLD As Double = 0.9 * R2
        Dim ThetaBLD As Double = ThetaPH1
        Dim FORBL As Double = Z3SAbs * 1.5
        Dim INN As Double = ZL
        Dim OUT As Double = ZL
        Dim Td As Double = 50
        ' ABOVE THIS LINE IS THE OUTPUT

        ' WRITE RESULT TO CONSOLE
        My.Application.Log.WriteEntry("Z1S: " & Z1SAbs)
        My.Application.Log.WriteEntry("Theta ph 1: " & ThetaPH1)
        My.Application.Log.WriteEntry("Z1Sg: " & Z1SgAbs)
        My.Application.Log.WriteEntry("Theta ph 10: " & ThetaPH10)
        My.Application.Log.WriteEntry("Ko: " & KoAbs)
        My.Application.Log.WriteEntry("Theta ph Ko: " & ThetaphKo)
        My.Application.Log.WriteEntry("BLD: " & BLD)
        My.Application.Log.WriteEntry("Theta BLD: " & ThetaBLD)
        My.Application.Log.WriteEntry("Z2S: " & Z2SAbs)
        My.Application.Log.WriteEntry("Z2Sg: " & Z2SgAbs)
        My.Application.Log.WriteEntry("Z3S: " & Z3SAbs)
        My.Application.Log.WriteEntry("Z3Sg: " & Z3SgAbs)
        My.Application.Log.WriteEntry("FORBL: " & FORBL)
        My.Application.Log.WriteEntry("INN: " & INN)
        My.Application.Log.WriteEntry("OUT: " & OUT)

        ' OPEN RESULT PAGE
        Dim resultPage As New ResultPage(Z1SAbs, ThetaPH1, Z1SgAbs, ThetaPH10, KoAbs, ThetaphKo, BLD, ThetaBLD, Z2SAbs, Z2SgAbs, Z2SAbs, Z3SgAbs, FORBL, INN, OUT)
        resultPage.ShowDialog()

    End Sub
End Class