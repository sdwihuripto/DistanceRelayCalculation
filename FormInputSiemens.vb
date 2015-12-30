Public Class FormInputSiemens
    Private Sub FormInput_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        FromL1.Focus()
    End Sub

    Private Sub AddBtn1_Click(sender As Object, e As EventArgs) Handles AddBtn1.Click
        FormAddInfo.AddInfo1.Text = "Additional Info (L1)"
        FormAddInfo.ShowDialog()
    End Sub

    Private Sub AddBtn2_Click(sender As Object, e As EventArgs) Handles AddBtn2.Click
        FormAddInfo.txtCCC1.Visible = False
        FormAddInfo.AddInfo1.Text = "Additional Info (L2)"
        FormAddInfo.ShowDialog()
    End Sub

    Private Sub AddBtn3_Click(sender As Object, e As EventArgs) Handles AddBtn3.Click
        FormAddInfo.txtCCC1.Visible = False
        FormAddInfo.AddInfo1.Text = "Additional Info (L3)"
        FormAddInfo.ShowDialog()
    End Sub

    Private Sub AddBtn4_Click(sender As Object, e As EventArgs) Handles AddBtn4.Click
        FormAddInfo.txtCCC1.Visible = False
        FormAddInfo.AddInfo1.Text = "Additional Info (L4)"
        FormAddInfo.ShowDialog()
    End Sub
    Private Sub cmdCalculate_Click(sender As Object, e As EventArgs) Handles cmdCalculate.Click
        If FromL1.Text = String.Empty Or FromL1.Text = "From Substantion" Then
            MessageBox.Show("Please insert line parameter", "Invalid From Substantion",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
            FromL1.BackColor = Color.Red
            With FromL1.Focus()
            End With
        Else
            FromL1.BackColor = Color.White
            If ToL1.Text = String.Empty Or ToL1.Text = "To Substantion" Then
                MessageBox.Show("Please insert line parameter", "Invalid To Substantion",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
                ToL1.BackColor = Color.Red
                With ToL1.Focus()
                End With
            Else
                ToL1.BackColor = Color.White
                If IsNumeric(txtResistansiL1.Text) = False Then
                    MessageBox.Show("Please insert line parameter", "Invalid Resistance",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
                    txtResistansiL1.BackColor = Color.Red
                    With txtResistansiL1.Focus()
                    End With
                Else
                    txtResistansiL1.BackColor = Color.White
                    If txtReaktansiL1.Text = String.Empty Or txtReaktansiL1.Text = "Reactance (Ohm/Km)" Or
                        IsNumeric(txtReaktansiL1.Text) = False Then
                        MessageBox.Show("Please insert line parameter", "Invalid Reactance",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error)
                        txtReaktansiL1.BackColor = Color.Red
                        With txtReaktansiL1.Focus()
                        End With
                    Else
                        txtReaktansiL1.BackColor = Color.White
                        If txtPanjangL1.Text = String.Empty Or txtPanjangL1.Text = "Length (Km)" Or
                        IsNumeric(txtPanjangL1.Text) = False Then
                            MessageBox.Show("Please insert line parameter", "Invalid Length",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error)
                            txtPanjangL1.BackColor = Color.Red
                            With txtPanjangL1.Focus()
                            End With
                        Else
                            txtPanjangL1.BackColor = Color.White
                            If CCCL1.Text = String.Empty Or CCCL1.Text = "CCC (A)" Or
                            IsNumeric(CCCL1.Text) = False Then
                                MessageBox.Show("Please insert line parameter", "Invalid Current Carrying Capacity ",
                                        MessageBoxButtons.OK, MessageBoxIcon.Error)
                                CCCL1.BackColor = Color.Red
                                With CCCL1.Focus()
                                End With
                            Else
                                CCCL1.BackColor = Color.White
                                If txtTransformer.Text = String.Empty Then
                                    MessageBox.Show("Please insert line parameter", "Invalid Transformer Name ",
                                            MessageBoxButtons.OK, MessageBoxIcon.Error)
                                    txtTransformer.BackColor = Color.Red
                                    With txtTransformer.Focus()
                                    End With
                                Else
                                    txtTransformer.BackColor = Color.White
                                    If mvaRating.Text = String.Empty Or IsNumeric(mvaRating.Text) = False Then
                                        MessageBox.Show("Please insert line parameter", "Invalid MVA Rating",
                                                MessageBoxButtons.OK, MessageBoxIcon.Error)
                                        mvaRating.BackColor = Color.Red
                                        With mvaRating.Focus()
                                        End With
                                    Else
                                        mvaRating.BackColor = Color.White
                                        If voltageLevel.Text = String.Empty Or IsNumeric(voltageLevel.Text) = False Then
                                            MessageBox.Show("Please insert line parameter", "Invalid Voltage Level",
                                                    MessageBoxButtons.OK, MessageBoxIcon.Error)
                                            voltageLevel.BackColor = Color.Red
                                            With voltageLevel.Focus()
                                            End With
                                        Else
                                            voltageLevel.BackColor = Color.White
                                            If impedanceBox.Text = String.Empty Or IsNumeric(impedanceBox.Text) = False Then
                                                MessageBox.Show("Please insert line parameter", "Invalid Impedance",
                                                        MessageBoxButtons.OK, MessageBoxIcon.Error)
                                                impedanceBox.BackColor = Color.Red
                                                With impedanceBox.Focus()
                                                End With
                                            Else
                                                impedanceBox.BackColor = Color.White
                                                If CTp.Text = String.Empty Or IsNumeric(CTp.Text) = False Or
                                                    CTs.Text = String.Empty Or IsNumeric(CTs.Text) = False Then
                                                    MessageBox.Show("Please insert line parameter", "Invalid CT Ratio",
                                                            MessageBoxButtons.OK, MessageBoxIcon.Error)
                                                    CTp.BackColor = Color.Red
                                                    CTs.BackColor = Color.Red
                                                    With CTp.Focus()
                                                    End With
                                                Else
                                                    CTp.BackColor = Color.White
                                                    CTs.BackColor = Color.White
                                                    If PTp.Text = String.Empty Or IsNumeric(PTp.Text) = False Or
                                                        PTs.Text = String.Empty Or IsNumeric(PTs.Text) = False Then
                                                        MessageBox.Show("Please insert line parameter", "Invalid PT Ratio",
                                                                MessageBoxButtons.OK, MessageBoxIcon.Error)
                                                        PTp.BackColor = Color.Red
                                                        PTs.BackColor = Color.Red
                                                        With PTp.Focus()
                                                        End With
                                                    Else
                                                        PTp.BackColor = Color.White
                                                        PTs.BackColor = Color.White
                                                        If RodBox.Text = String.Empty Or IsNumeric(RodBox.Text) = False Then
                                                            MessageBox.Show("Please insert line parameter", "Invalid Rod Insulator Length",
                                                                    MessageBoxButtons.OK, MessageBoxIcon.Error)
                                                            RodBox.BackColor = Color.Red
                                                            With RodBox.Focus()
                                                            End With
                                                        Else
                                                            RodBox.BackColor = Color.White
                                                            If InfeedBox.Text = String.Empty Or IsNumeric(InfeedBox.Text) = False Then
                                                                MessageBox.Show("Please insert line parameter", "Invalid Infeed Factor",
                                                                        MessageBoxButtons.OK, MessageBoxIcon.Error)
                                                                InfeedBox.BackColor = Color.Red
                                                                With InfeedBox.Focus()
                                                                End With
                                                            Else
                                                                InfeedBox.BackColor = Color.White
                                                                If ArcBox.Text = String.Empty Or IsNumeric(ArcBox.Text) = False Then
                                                                    MessageBox.Show("Please insert line parameter", "Invalid Arc Current",
                                                                            MessageBoxButtons.OK, MessageBoxIcon.Error)
                                                                    ArcBox.BackColor = Color.Red
                                                                    With ArcBox.Focus()
                                                                    End With
                                                                Else
                                                                    ArcBox.BackColor = Color.White
                                                                    If ResistanceBox.Text = String.Empty Or IsNumeric(ResistanceBox.Text) = False Then
                                                                        MessageBox.Show("Please insert line parameter", "Invalid Foot Resistance of Tower",
                                                                                MessageBoxButtons.OK, MessageBoxIcon.Error)
                                                                        ResistanceBox.BackColor = Color.Red
                                                                        With ResistanceBox.Focus()
                                                                        End With
                                                                    Else
                                                                        ResistanceBox.BackColor = Color.White
                                                                        Dim GIA As String = FromL1.Text
                                                                        Dim GIB As String = ToL1.Text

                                                                        Dim R1 As Double
                                                                        Dim R2 As Double
                                                                        Dim R3 As Double
                                                                        Dim R4 As Double

                                                                        Dim X1 As Double
                                                                        Dim X2 As Double
                                                                        Dim X3 As Double
                                                                        Dim X4 As Double

                                                                        Dim L1 As Double
                                                                        Dim L2 As Double
                                                                        Dim L3 As Double
                                                                        Dim L4 As Double

                                                                        R1 = Double.Parse(txtResistansiL1.Text)
                                                                        Try
                                                                            R2 = Double.Parse(txtResistansiL2.Text)
                                                                        Catch ex As Exception
                                                                            R2 = 0
                                                                        End Try
                                                                        Try
                                                                            R3 = Double.Parse(txtResistansiL3.Text)
                                                                        Catch ex As Exception
                                                                            R3 = 0
                                                                        End Try
                                                                        Try
                                                                            R4 = Double.Parse(txtResistansiL4.Text)
                                                                        Catch ex As Exception
                                                                            R4 = 0
                                                                        End Try

                                                                        Try
                                                                            X1 = Double.Parse(txtReaktansiL1.Text)
                                                                        Catch ex As Exception
                                                                            X1 = 0
                                                                        End Try
                                                                        Try
                                                                            X2 = Double.Parse(txtReaktansiL2.Text)
                                                                        Catch ex As Exception
                                                                            X2 = 0
                                                                        End Try
                                                                        Try
                                                                            X3 = Double.Parse(txtReaktansiL3.Text)
                                                                        Catch ex As Exception
                                                                            X3 = 0
                                                                        End Try
                                                                        Try
                                                                            X4 = Double.Parse(txtReaktansiL4.Text)
                                                                        Catch ex As Exception
                                                                            X4 = 0
                                                                        End Try

                                                                        Try
                                                                            L1 = Double.Parse(txtPanjangL1.Text)
                                                                        Catch ex As Exception
                                                                            L1 = 0
                                                                        End Try
                                                                        Try
                                                                            L2 = Double.Parse(txtPanjangL2.Text)
                                                                        Catch ex As Exception
                                                                            L2 = 0
                                                                        End Try
                                                                        Try
                                                                            L3 = Double.Parse(txtPanjangL3.Text)
                                                                        Catch ex As Exception
                                                                            L3 = 0
                                                                        End Try
                                                                        Try
                                                                            L4 = Double.Parse(txtPanjangL4.Text)
                                                                        Catch ex As Exception
                                                                            L4 = 0
                                                                        End Try

                                                                        Dim RL1 As Double = R1 * L1
                                                                        Dim RL2 As Double = R2 * L2
                                                                        Dim RL3 As Double = R3 * L3
                                                                        Dim RL4 As Double = R4 * L4

                                                                        Dim XL1 As Double = X1 * L1
                                                                        Dim XL2 As Double = X2 * L2
                                                                        Dim XL3 As Double = X3 * L3
                                                                        Dim XL4 As Double = X4 * L4

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

                                                                        RL10 = (R1 + 0.15) * L1
                                                                        RL20 = (R2 + 0.15) * L2
                                                                        RL30 = (R3 + 0.15) * L3
                                                                        RL40 = (R4 + 0.15) * L4

                                                                        XL10 = 3 * X1 * L1
                                                                        XL20 = 3 * X2 * L2
                                                                        XL30 = 3 * X3 * L3
                                                                        XL40 = 3 * X4 * L4

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

                                                                        Dim Larc As Double
                                                                        Dim larcKecil As Double
                                                                        Dim Rfood As Double
                                                                        Dim linfeed As Double

                                                                        Try
                                                                            Larc = Double.Parse(RodBox.Text)
                                                                        Catch ex As Exception
                                                                            Larc = 7.5
                                                                        End Try
                                                                        Try
                                                                            larcKecil = Double.Parse(ArcBox.Text)
                                                                        Catch ex As Exception
                                                                            larcKecil = 2500
                                                                        End Try
                                                                        Try
                                                                            Rfood = Double.Parse(ResistanceBox.Text)
                                                                        Catch ex As Exception
                                                                            Rfood = 10
                                                                        End Try

                                                                        linfeed = Double.Parse(InfeedBox.Text)

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

                                                                        ' OPEN RESULT PAGE
                                                                        Dim resultPage As New FormOutputSiemens(GIA, GIB, CTp.Text, CTs.Text, PTp.Text, PTs.Text, ZL1Abs, ZL10Abs, txtPanjangL1.Text,
                                                                                                                ThetaPH1, XLine, L1, R01, X01, R0B5, X0B5,
                                                                                                                RLphE, ThetaLphE, RLphph, ThetaLphph,
                                                                                                                RZ1Abs, XZ1sAbs, RGZ1Abs,
                                                                                                                RZ1BAbs, XZ1BAbs, RGZ1BAbs,
                                                                                                                RZ2Abs, XZ2Abs, RGZ2Abs,
                                                                                                                RZ3Abs, XZ3Abs, RGZ3Abs,
                                                                                                                ZRZ1Abs, ZRZ1BAbs, ZRZ2Abs, ZRZ3Abs,
                                                                                                                T1, T1, T2, T2, T3)

                                                                        resultPage.ShowDialog()
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End Sub
End Class