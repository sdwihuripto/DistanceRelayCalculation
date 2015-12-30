Public Class FormInputAlstom
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
                                                        If infeed.Text = String.Empty Or IsNumeric(infeed.Text) = False Then
                                                            MessageBox.Show("Please insert line parameter", "Invalid Infeed Factor",
                                                                    MessageBoxButtons.OK, MessageBoxIcon.Error)
                                                            infeed.BackColor = Color.Red
                                                            With infeed.Focus()
                                                            End With
                                                        Else
                                                            infeed.BackColor = Color.White
                                                            If PhaseBox.Text = String.Empty Or IsNumeric(PhaseBox.Text) = False Then
                                                                MessageBox.Show("Please insert line parameter", "Invalid Phase Conductor Distance",
                                                                        MessageBoxButtons.OK, MessageBoxIcon.Error)
                                                                PhaseBox.BackColor = Color.Red
                                                                With infeed.Focus()
                                                                End With
                                                            Else
                                                                PhaseBox.BackColor = Color.White
                                                                If shortCircuitBox.Text = String.Empty Or IsNumeric(shortCircuitBox.Text) = False Then
                                                                    MessageBox.Show("Please insert line parameter", "Invalid 3-Phase Short Circuit Current",
                                                                            MessageBoxButtons.OK, MessageBoxIcon.Error)
                                                                    shortCircuitBox.BackColor = Color.Red
                                                                    With shortCircuitBox.Focus()
                                                                    End With
                                                                Else
                                                                    shortCircuitBox.BackColor = Color.White
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

                                                                    ' OPEN RESULT PAGE
                                                                    Dim resultPage As New FormOutputAlstom(GIA, GIB, CTp.Text, CTs.Text, PTp.Text, PTs.Text,
                                                                                                            Double.Parse(txtPanjangL1.Text), ZL, ThetaPH1,
                                                                                                            kZ0, ThetakZ0,
                                                                                                            Z1SAbs, T1, Z2SAbs, T2, Z3SAbs, T3,
                                                                                                            Rphmin, Rgmin, R3ph, R3g, R2ph, R2g, R1ph, R1g,
                                                                                                            ZB)

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
    End Sub
End Class