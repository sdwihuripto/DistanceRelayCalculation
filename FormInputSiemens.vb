Public Class FormInputSiemens

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

        Dim Larc As Double = Double.Parse(RodBox.Text)
        Dim larcKecil As Double = Double.Parse(ArcBox.Text)
        Dim Rfood As Double = Double.Parse(ResistanceBox.Text)
        Dim linfeed As Double = Double.Parse(InfeedBox.Text)

        Dim XTrf As Double = (impedance * Math.Pow(kV, 2)) / (MVA * 100)
        Dim CT As Double = Double.Parse(CTp.Text) / Double.Parse(CTs.Text)
        Dim PT As Double = Double.Parse(PTp.Text) / Double.Parse(PTs.Text)
        Dim n As Double = CT / PT

        Dim ln As Double = 1
        Dim Vn As Double = 100

        Dim XLine As Double = (XL1 / Double.Parse(txtPanjangL1.Text)) * n
        Dim R01 As Double = (1 / 3) * ((RL10 / RL1) - 1)
        Dim X01 As Double = (1 / 3) * ((XL10 / XL1) - 1)
        Dim R0B5 As Double = (1 / 3) * ((RL20 / RL2) - 1)
        Dim X0B5 As Double = (1 / 3) * ((XL20 / XL2) - 1)

        Dim ZldAbs As Double = ((kV * 1000) / (Math.Sqrt(3) * Double.Parse(CCCL1.Text))) * n
        Dim Thetald As Double = ThetaPH1 + 5
        Dim Rld As Double = ZldAbs * Math.Cos(Thetald) * 0.5
        Dim Xld As Double = ZldAbs * Math.Sin(Thetald) * 0.5
        Dim RfS As Double = Rfood * n
        Dim ZloadminAbs As Double = ZldAbs * 0.8
        Dim Z1safetyAbs As Double = ZldAbs * 0.5
        Dim Rgmax As Double = Z1safetyAbs
        Dim Rgmin As Double = RfS
        Dim RLphE As Double = Rgmax
        Dim ThetaLphE As Double = Thetald
        Dim RLphph As Double = Rgmax
        Dim ThetaLphph As Double = Thetald

        Dim Rarc1 As Double = (28700 * Larc) / (Math.Pow(larcKecil, 1.4))
        Dim Rarc As Double
        If Math.Abs(Rarc1) > Math.Abs(RL1) Then
            Rarc = 0
        Else
            Rarc = Rarc1
        End If

        Dim RZ1P As Double = (0.8 * RL1) + (0.5 * Rarc)
        Dim RZ1Abs As Double = RZ1P * n
        Dim XZ1P As Double = 0.8 * XL1
        Dim XZ1sAbs As Double = XZ1P * n
        Dim RGZ1P As Double = ((0.8 * RL1) + Rfood + Rarc)
        Dim RGZ1Abs As Double = RGZ1P * n
        Dim RZ1BP As Double = (0.8 * (RL1 + 0.8 * RL2 * linfeed) + (0.5 * Rarc))
        Dim RZ1BAbs As Double = RZ1BP * n
        Dim RZ2Abs As Double = RZ1BAbs
        Dim XZ1BPmin As Double = 1.2 * XL1
        Dim XZ1BPmax1 As Double = 0.8 * (XL1 + 0.8 * linfeed * XL2)
        Dim XZ1BPmax2 As Double = 0.8 * (XL1 + 0.5 * XTrf * linfeed)
        Dim X21Bmak As Double
        If XZ1BPmin > XZ1BPmax1 Then
            X21Bmak = XZ1BPmin
        Else
            X21Bmak = XZ1BPmax1
        End If
        Dim X1BP As Double
        If X21Bmak < XZ1BPmax2 Then
            X1BP = X21Bmak
        Else
            X1BP = XZ1BPmax2
        End If
        Dim XZ1BAbs As Double = X1BP * n
        Dim XZ2Abs As Double = XZ1BAbs
        Dim RGZ1BP As Double = (0.8 * (RL1 + 0.8 * RL2 * linfeed) + Rarc + 2 * Rfood)
        Dim RGZ1BAbs As Double = RGZ1BP * n
        Dim RGZ2Abs As Double = RGZ1BAbs

        Dim RZ3P As Double = (1.2 * (RL1 + RL3)) + (0.5 * Rarc)
        Dim RZ3Abs As Double = RZ3P * n
        Dim XZ3Pmin As Double = 1.2 * (XL1 + XL3)
        Dim XZ3Pmax1 As Double = 0.8 * (XL1 + 1.2 * linfeed * XL3)
        Dim XZ3Pmax2 As Double = 0.8 * (XL1 + (0.8 * linfeed * (XL3 + 0.8 * XL4)))
        Dim XZ3Pmax3 As Double = 0.8 * (XL1 + 0.8 * linfeed * XTrf)
        Dim X31 As Double
        If Math.Abs(XZ3Pmin) > Math.Abs(XZ3Pmax1) Then
            X31 = XZ3Pmin
        Else
            X31 = XZ3Pmax1
        End If
        Dim X32 As Double
        If Math.Abs(XZ3Pmax2) > Math.Abs(X31) Then
            X32 = XZ3Pmax2
        Else
            X32 = X31
        End If
        Dim X3P As Double
        If Math.Abs(X32) > Math.Abs(XZ3Pmax3) Then
            X3P = XZ3Pmax3
        Else
            X3P = X32
        End If

        Dim XZ3Abs As Double = X3P * n
        Dim RGZ3P As Double = (1.2 * (RL1 + RL3 * linfeed) + Rarc + 2 * Rfood)
        Dim RGZ3Abs As Double = RGZ3P * n

        Dim ZRZ1P As Double = 0.8 * ZL1Abs
        Dim ZRZ1Abs As Double = ZRZ1P * n
        Dim Z2min As Double = 1.2 * ZL1Abs
        Dim Z2mak1 As Double = 0.8 * (ZL1Abs + 0.8 * linfeed * ZL2Abs)
        Dim ZTrfAbs As Double = 0.8 * (Math.Sqrt(Math.Pow(RL1, 2) + Math.Pow(XL1 + (0.5 * XTrf), 2)))

        Dim Z21mak As Double
        If Math.Abs(Z2mak1) > Math.Abs(Z2min) Then
            Z21mak = Z2mak1
        Else
            Z21mak = Z2min
        End If
        Dim ZRZ1BP As Double
        If Math.Abs(Z21mak) < Math.Abs(ZTrfAbs) Then
            ZRZ1BP = Z21mak
        Else
            ZRZ1BP = ZTrfAbs
        End If

        Dim ZRZ1BAbs As Double = ZRZ1BP * n
        Dim ZRZ2Abs As Double = ZRZ1BAbs
        Dim Z3min As Double = 1.2 * (ZL1Abs + ZL3Abs)
        Dim Z3mak1 As Double = 0.8 * (ZL1Abs + 1.2 * linfeed * ZL3Abs)
        Dim Z3mak2 As Double = 0.8 * (ZL1Abs + (0.8 * linfeed * (ZL3Abs + 0.8 * ZL4Abs)))
        Dim Z3Trf As Double = 0.8 * (Math.Sqrt(Math.Pow(RL1, 2) + Math.Pow(XL1 + 0.8 * XTrf, 2)))

        Dim Z31 As Double
        If Math.Abs(Z3mak1) > Math.Abs(Z3mak2) Then
            Z31 = Z3mak1
        Else
            Z31 = Z3mak2
        End If
        Dim Z32 As Double
        If Math.Abs(Z31) > Math.Abs(Z3min) Then
            Z32 = Z31
        Else
            Z32 = Z3min
        End If
        Dim ZRZ3P As Double
        If Math.Abs(Z32) > Math.Abs(Z3Trf) Then
            ZRZ3P = Z3Trf
        Else
            ZRZ3P = Z32
        End If

        Dim ZRZ3Abs As Double = ZRZ3P * n

        Dim T1 As Double = 0
        Dim Z2PAbs As Double = Math.Sqrt(Math.Pow(RZ1BP, 2) + Math.Pow(X1BP, 2))
        Dim X1sgAbs As Double = 0.8 * XL2
        Dim T2 As Double
        If (Z2PAbs - Math.Abs(XL1)) < X1sgAbs Then
            T2 = 0.4
        Else
            T2 = 0.8
        End If
        Dim T3 As Double = 1.6

        My.Application.Log.WriteEntry("YES BERHASIL")
    End Sub
End Class