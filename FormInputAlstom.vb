Public Class FormInputAlstom

    Private Sub cmdCalculate_Click(sender As Object, e As EventArgs) Handles cmdCalculate.Click
        Dim GIA As String = FromL1.Text
        Dim GIB As String = ToL1.Text

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

        Dim ZL10Abs As Double
        Dim ZL20Abs As Double
        Dim ZL30Abs As Double
        Dim ZL40Abs As Double

        Dim ThetaPH1 As Double
        Dim ThetaPH2 As Double
        Dim ThetaPH3 As Double
        Dim ThetaPH4 As Double

        'If ComboBox4.SelectedIndex = 0 Then
        ZL1Abs = Math.Sqrt(Math.Pow(RL1, 2) + Math.Pow(XL1, 2))
        ZL2Abs = Math.Sqrt(Math.Pow(RL2, 2) + Math.Pow(XL2, 2))
        ZL3Abs = Math.Sqrt(Math.Pow(RL3, 2) + Math.Pow(XL3, 2))
        ZL4Abs = Math.Sqrt(Math.Pow(RL4, 2) + Math.Pow(XL4, 2))

        ThetaPH1 = Math.Atan(XL1 / RL1) * (180 / Math.PI)
        ThetaPH2 = Math.Atan(XL2 / RL2) * (180 / Math.PI)
        ThetaPH3 = Math.Atan(XL3 / RL3) * (180 / Math.PI)
        ThetaPH4 = Math.Atan(XL4 / RL4) * (180 / Math.PI)

        RL10 = (Double.Parse(txtResistansiL1.Text) + 0.15) * Double.Parse(txtPanjangL1.Text)
        RL20 = (Double.Parse(txtResistansiL2.Text) + 0.15) * Double.Parse(txtPanjangL2.Text)
        RL30 = (Double.Parse(txtResistansiL3.Text) + 0.15) * Double.Parse(txtPanjangL3.Text)
        RL40 = (Double.Parse(txtResistansiL4.Text) + 0.15) * Double.Parse(txtPanjangL4.Text)

        XL10 = 3 * Double.Parse(txtReaktansiL1.Text) * Double.Parse(txtPanjangL1.Text)
        XL20 = 3 * Double.Parse(txtReaktansiL2.Text) * Double.Parse(txtPanjangL2.Text)
        XL30 = 3 * Double.Parse(txtReaktansiL3.Text) * Double.Parse(txtPanjangL3.Text)
        XL40 = 3 * Double.Parse(txtReaktansiL4.Text) * Double.Parse(txtPanjangL4.Text)

        ZL10Abs = Math.Sqrt(Math.Pow(RL10, 2) + Math.Pow(XL10, 2))
        ZL20Abs = Math.Sqrt(Math.Pow(RL20, 2) + Math.Pow(XL20, 2))
        ZL30Abs = Math.Sqrt(Math.Pow(RL30, 2) + Math.Pow(XL30, 2))
        ZL40Abs = Math.Sqrt(Math.Pow(RL40, 2) + Math.Pow(XL40, 2))
        'ElseIf ComboBox4.SelectedIndex = 1 Then


        ' End If

        Dim ThetaPH10 As Double = Math.Atan(XL10 / RL10) * (180 / Math.PI)
        Dim ThetaPH20 As Double = Math.Atan(XL20 / RL20) * (180 / Math.PI)
        Dim ThetaPH30 As Double = Math.Atan(XL30 / RL30) * (180 / Math.PI)
        Dim ThetaPH40 As Double = Math.Atan(XL40 / RL40) * (180 / Math.PI)

        Dim MVA As Double = Double.Parse(mvaRating.Text)
        Dim kV As Double = Double.Parse(voltageLevel.Text)
        Dim impedance As Double = Double.Parse(impedanceBox.Text)

        Dim K As Double = Double.Parse(infeed.Text)
        Dim Lc As Double = Double.Parse(PhaseBox.Text)
        Dim lhs3f As Double = Double.Parse(shortCircuitBox.Text)

        Dim XTrf As Double = (impedance * Math.Pow(kV, 2)) / (MVA * 100)
        Dim CT As Double = Double.Parse(CTp.Text) / Double.Parse(CTs.Text)
        Dim PT As Double = Double.Parse(PTp.Text) / Double.Parse(PTs.Text)
        Dim n As Double = CT / PT

        Dim ZL As Double = ZL1Abs * n
        Dim Z1PAbs As Double = 0.8 * ZL1Abs
        Dim Z1SAbs As Double = n * Z1PAbs

        Dim Z2minAbs As Double = 1.2 * ZL1Abs
        Dim Z2mak1Abs As Double = 0.8 * (ZL1Abs + (K * 0.8 * ZL2Abs))
        Dim ZTrfAbs As Double = 0.8 * (Math.Sqrt(Math.Pow(RL1, 2) + Math.Pow(XL1 + (0.5 * XTrf), 2)))
        Dim Z2bAbs As Double = ZL1Abs + 0.8 * ZL2Abs

        Dim Z21makAbs As Double
        If Z2mak1Abs > Z2minAbs Then
            Z21makAbs = Z2mak1Abs
        Else
            Z21makAbs = Z2minAbs
        End If
        Dim Z22makAbs As Double
        If Z21makAbs < ZTrfAbs Then
            Z22makAbs = Z21makAbs
        Else
            Z22makAbs = ZTrfAbs
        End If
        Dim Z2SAbs As Double = n * Z22makAbs

        Dim Z3minAbs As Double = 1.2 * (ZL1Abs + K * ZL3Abs)
        Dim Z3mak1Abs As Double = 0.8 * (ZL1Abs + (K * 1.2 * ZL3Abs))
        Dim Z3mak2Abs As Double = 0.8 * (ZL1Abs + (0.8 * (ZL3Abs + 0.8 * ZL4Abs * K)))
        Dim Z3TrfAbs As Double = 0.8 * (Math.Sqrt(Math.Pow(RL1, 2) + Math.Pow(XL1 + (0.8 * XTrf), 2)))

        Dim Z3bAbs As Double = (ZL1Abs + (K * (0.8 * (ZL3Abs + 0.8 * ZL4Abs))))
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
        Dim Z3SAbs = n * Z3PAbs

        Dim T1 As Double = 0
        Dim T2 As Double
        If Z2bAbs > Z22makAbs Then
            T2 = 0.4
        Else
            T2 = 0.8
        End If
        Dim T3 As Double
        If Z3bAbs > Z32Abs Then
            T3 = 1.2
        Else
            T3 = 1.6
        End If

        Dim kZ0 As Double = (ZL10Abs - ZL1Abs) / (3 * ZL1Abs)
        Dim ThetakZ0 As Double = Math.Atan((XL10 - XL1) / (RL10 - RL1)) - Math.Atan((3 * XL1) / (3 * RL1))
        ThetakZ0 = ThetakZ0 * (180 / Math.PI)

        Dim Vn As Double
        If Double.Parse(PTs.Text) > 99 Then
            Vn = Double.Parse(PTs.Text) / Math.Sqrt(3)
        Else
            Vn = Double.Parse(PTs.Text) * 100 / Math.Sqrt(3)
        End If

        Dim InVar As Double = Double.Parse(CTs.Text)
        Dim Zloadmin As Double = Vn / InVar
        Dim MRphmax As Double = 0.4 * Zloadmin
        Dim MRgmax As Double = 0.2 * Zloadmin

        Dim Rphmax As Double = Zloadmin - MRphmax
        Dim Rgmax As Double = Zloadmin - MRgmax
        Dim lhs2f As Double = (Math.Sqrt(3) / 2) * lhs3f
        Dim Ra As Double = (28710 * Lc) / Math.Pow(lhs2f, 1.4)
        Dim Rphmin As Double = Ra * n
        Dim Rgmin As Double = 20 * n

        Dim R3ph As Double = 0.8 * Rphmax
        Dim R3g As Double = 0.96 * Rgmax
        Dim R2ph As Double = 0.8 * R3ph
        Dim R2g As Double = 0.96 * R3g
        Dim R1ph As Double = 0.8 * R2ph
        Dim R1g As Double = 0.96 * R2g

        Dim Zld As Double = (kV * 1000 * n) / (Double.Parse(CCCL1.Text) * Math.Sqrt(3))
        Dim Thetald As Double = 30
        Dim ZB As Double = Zld * Math.Cos(30) * 0.51

        My.Application.Log.WriteEntry("YES BERHASIL")
    End Sub
End Class