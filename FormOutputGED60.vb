Imports Microsoft.Office.Interop

Public Class FormOutputGED60

    Dim tempA As String, tempB As String
    Dim tempC As Double, tempD As Double
    Dim tempE As Double, tempF As Double
    Dim tempG As Double, tempH As Double
    Dim tempI As Double, tempJ As Double
    Dim tempK As Double, tempL As Double
    Dim tempM As Double, tempN As Double
    Dim tempO As Double, tempP As Double
    Dim tempQ As Double, tempR As Double
    Dim tempS As Double, tempT As Double
    Dim tempU As Double, tempV As Double
    Dim tempW As Double, tempX As Double
    Dim tempY As Double, tempZ As Double
    Dim tempA1 As Double, tempB1 As Double
    Dim tempC1 As Double, tempD1 As Double
    Dim tempE1 As Double, tempF1 As Double
    Dim tempG1 As Double, tempH1 As Double
    Dim tempI1 As Double, tempJ1 As Double
    Dim tempK1 As Double, tempL1 As Double
    Dim tempM1 As Double, tempN1 As Double
    Dim tempO1 As Double, tempP1 As Double
    Dim tempQ1 As Double, tempR1 As Double
    Dim tempS1 As Double, tempT1 As Double
    Dim tempU1 As Double, tempV1 As Double
    Dim tempW1 As Double, tempX1 As Double
    Dim tempY1 As Double, tempZ1 As Double

    Public Sub New(ByVal a As String, ByVal b As String,
                     ByVal c As Double, ByVal d As Double,
                     ByVal e As Double, ByVal f As Double,
                     ByVal g As Double, ByVal h As Double,
                     ByVal i As Double, ByVal j As Double,
                     ByVal k As Double, ByVal l As Double,
                     ByVal m As Double, ByVal n As Double,
                     ByVal o As Double, ByVal p As Double,
                     ByVal q As Double, ByVal r As Double,
                     ByVal s As Double, ByVal t As Double,
                     ByVal u As Double, ByVal v As Double,
                     ByVal w As Double, ByVal x As Double,
                     ByVal y As Double, ByVal z As Double,
                     ByVal a1 As Double, ByVal b1 As Double,
                     ByVal c1 As Double, ByVal d1 As Double,
                     ByVal e1 As Double, ByVal f1 As Double,
                     ByVal g1 As Double, ByVal h1 As Double,
                     ByVal i1 As Double, ByVal j1 As Double,
                     ByVal k1 As Double, ByVal l1 As Double,
                     ByVal m1 As Double, ByVal n1 As Double,
                     ByVal o1 As Double, ByVal p1 As Double,
                     ByVal q1 As Double, ByVal r1 As Double,
                     ByVal s1 As Double, ByVal t1 As Double,
                     ByVal u1 As Double, ByVal v1 As Double,
                     ByVal w1 As Double, ByVal x1 As Double,
                     ByVal y1 As Double, ByVal z1 As Double)
        Me.InitializeComponent()

        tempA = a
        tempB = b
        tempC = Math.Round(c, 2)
        tempD = Math.Round(d, 2)
        tempE = Math.Round(e, 2)
        tempF = Math.Round(f, 2)
        tempG = Math.Round(g, 2)
        tempH = Math.Round(h, 2)
        tempI = Math.Round(i, 2)
        tempJ = Math.Round(j, 2)
        tempK = Math.Round(k, 2)
        tempL = Math.Round(l, 2)
        tempM = Math.Round(m, 2)
        tempN = Math.Round(n, 2)
        tempO = Math.Round(o, 2)
        tempP = Math.Round(p, 2)
        tempQ = Math.Round(q, 2)
        tempR = Math.Round(r, 2)
        tempS = Math.Round(s, 2)
        tempT = Math.Round(t, 2)
        tempU = Math.Round(u, 2)
        tempV = Math.Round(v, 2)
        tempW = Math.Round(w, 2)
        tempX = Math.Round(x, 2)
        tempY = Math.Round(y, 2)
        tempZ = Math.Round(z, 2)
        tempA1 = Math.Round(a1, 2)
        tempB1 = Math.Round(b1, 2)
        tempC1 = Math.Round(c1, 2)
        tempD1 = Math.Round(d1, 2)
        tempE1 = Math.Round(e1, 2)
        tempF1 = Math.Round(f1, 2)
        tempG1 = Math.Round(g1, 2)
        tempH1 = Math.Round(h1, 2)
        tempI1 = Math.Round(i1, 2)
        tempJ1 = Math.Round(j1, 2)
        tempK1 = Math.Round(k1, 2)
        tempL1 = Math.Round(l1, 2)
        tempM1 = Math.Round(m1, 2)
        tempN1 = Math.Round(n1, 2)
        tempO1 = Math.Round(o1, 2)
        tempP1 = Math.Round(p1, 2)
        tempQ1 = Math.Round(q1, 2)
        tempR1 = Math.Round(r1, 2)
        tempS1 = Math.Round(s1, 2)
        tempT1 = Math.Round(t1, 2)
        tempU1 = Math.Round(u1, 2)
        tempV1 = Math.Round(v1, 2)
        tempW1 = Math.Round(w1, 2)
        tempX1 = Math.Round(x1, 2)
        tempY1 = Math.Round(y1, 2)
        tempZ1 = Math.Round(z1, 2)

        OutSum1.Text = tempA
        OutSum2.Text = tempB
        OutSum3.Text = tempC
        OutSum4.Text = tempD
        OutSum5.Text = tempE
        OutSum6.Text = tempF
        OutSum7.Text = tempG
        OutSum8.Text = tempH
        OutSum9.Text = tempI
        Out1A.Text = tempJ
        Out1B.Text = tempK
        Out1C.Text = tempL
        Out1D.Text = tempM
        Out2A.Text = tempN
        Out2B.Text = tempO
        Out2C.Text = tempP
        Out2D.Text = tempQ
        Out2E.Text = tempR
        Out2F.Text = tempS
        Out2G.Text = tempT
        Out2H.Text = tempU
        Out2I.Text = tempV
        Out3A.Text = tempW
        Out3B.Text = tempX
        Out3C.Text = tempY
        Out3D.Text = tempZ
        Out4A.Text = tempA1
        Out4B.Text = tempB1
        Out4C.Text = tempC1
        Out4D.Text = tempD1
        Out4E.Text = tempE1
        Out4F.Text = tempF1
        Out4G.Text = tempG1
        Out4H.Text = tempH1
        Out4I.Text = tempI1
        Out5A.Text = tempJ1
        Out5B.Text = tempK1
        Out5C.Text = tempL1
        Out5D.Text = tempM1
        Out6A.Text = tempN1
        Out6B.Text = tempO1
        Out6C.Text = tempP1
        Out6D.Text = tempQ1
        Out6E.Text = tempR1
        Out6F.Text = tempS1
        Out6G.Text = tempT1
        Out6H.Text = tempU1
        Out6I.Text = tempV1
        Out7A.Text = tempW1
        Out7B.Text = tempX1
        Out7C.Text = tempY1
        Out7D.Text = tempZ1
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim oWord As Word.Application
        Dim oDoc As Word.Document
        Dim oPara1 As Word.Paragraph, oPara2 As Word.Paragraph
        Dim oPara3 As Word.Paragraph, oPara4 As Word.Paragraph
        Dim oPara5 As Word.Paragraph, oPara6 As Word.Paragraph
        Dim oPara7 As Word.Paragraph, oPara8 As Word.Paragraph


        'Start Word and open the document template.
        oWord = CreateObject("Word.Application")
        oWord.Visible = True
        oDoc = oWord.Documents.Add
        oDoc = oWord.ActiveDocument

        Dim Table1 As Word.Table
        Dim Table2 As Word.Table
        Dim Table3 As Word.Table
        Dim Table4 As Word.Table
        Dim Table5 As Word.Table
        Dim Table6 As Word.Table
        Dim Table7 As Word.Table
        Dim Table8 As Word.Table

        oPara1 = oDoc.Content.Paragraphs.Add
        oPara1.Range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
        Table1 = oDoc.Content.Tables.Add(oPara1.Range, 7, 3)
        Table1.Columns(1).Cells(1).Range.Text = "Location"
        Table1.Columns(2).Cells(1).Range.Text = tempA
        Table1.Columns(1).Cells(2).Range.Text = "Line Bay To"
        Table1.Columns(2).Cells(2).Range.Text = tempB
        Table1.Columns(1).Cells(3).Range.Text = "CT Ratio"
        Table1.Columns(2).Cells(3).Range.Text = tempC & " / " & tempD
        Table1.Columns(3).Cells(3).Range.Text = "Ampere"
        Table1.Columns(1).Cells(4).Range.Text = "PT Ratio"
        Table1.Columns(2).Cells(4).Range.Text = tempE & " / " & tempF
        Table1.Columns(3).Cells(4).Range.Text = "Voltage"
        Table1.Columns(1).Cells(5).Range.Text = "Positive Sequence"
        Table1.Columns(2).Cells(5).Range.Text = tempG
        Table1.Columns(3).Cells(5).Range.Text = "Ohm/Phase"
        Table1.Columns(1).Cells(6).Range.Text = "Negative Sequence"
        Table1.Columns(2).Cells(6).Range.Text = tempH
        Table1.Columns(3).Cells(6).Range.Text = "Ohm/Phase"
        Table1.Columns(1).Cells(7).Range.Text = "Line Length"
        Table1.Columns(2).Cells(7).Range.Text = tempI
        Table1.Columns(3).Cells(7).Range.Text = "Kilometer"
        Table1.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        Table1.Borders.OutsideColor = Word.WdColor.wdColorBlack
        Table1.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
        Table1.Borders.InsideColor = Word.WdColor.wdColorBlack
        Table1.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
        oPara1.Format.SpaceAfter = 21
        oPara1.Range.InsertParagraphAfter()

        oPara2 = oDoc.Content.Paragraphs.Add
        oPara2.Range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
        Table2 = oDoc.Content.Tables.Add(oPara2.Range, 10, 3)
        Table2.Rows(1).Cells.Merge()
        Table2.Rows(1).Range.Text = "Phase Distance Z1 MHO"
        Table2.Rows(1).Range.Font.Bold = True
        Table2.Rows(1).Range.Font.Size = 16
        Table2.Rows(2).Cells.Split()
        Table2.Cell(2, 1).Range.Text = "Ph Dis Z1 Reach"
        Table2.Cell(2, 2).Range.Text = tempJ
        Table2.Cell(2, 3).Range.Text = "Ohm"
        Table2.Cell(3, 1).Range.Text = "Ph Dis Z1 Direction"
        Table2.Cell(3, 2).Range.Text = "FORWARD"
        Table2.Cell(4, 1).Range.Text = "Ph Dis Z1 Comp Limit"
        Table2.Cell(4, 2).Range.Text = tempK
        Table2.Cell(4, 3).Range.Text = "Degree"
        Table2.Cell(5, 1).Range.Text = "Ph Dis Z1 Delay"
        Table2.Cell(5, 2).Range.Text = tempL
        Table2.Cell(5, 3).Range.Text = "Sec"
        Table2.Cell(6, 1).Range.Text = "Ph Dis Z1 Supv"
        Table2.Cell(6, 2).Range.Text = "1.2"
        Table2.Cell(6, 3).Range.Text = "pu"
        Table2.Cell(7, 1).Range.Text = "RCA"
        Table2.Cell(7, 2).Range.Text = tempM
        Table2.Cell(7, 3).Range.Text = "Degree"
        Table2.Cell(8, 1).Range.Text = "COMPLIMIT"
        Table2.Cell(8, 2).Range.Text = "90"
        Table2.Cell(8, 3).Range.Text = "Degree"
        Table2.Cell(9, 1).Range.Text = "DIR RCA"
        Table2.Cell(9, 2).Range.Text = "80"
        Table2.Cell(9, 3).Range.Text = "Degree"
        Table2.Cell(10, 1).Range.Text = "DIR COMPLIMIT"
        Table2.Cell(10, 2).Range.Text = "90"
        Table2.Cell(10, 3).Range.Text = "Degree"
        Table2.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        Table2.Borders.OutsideColor = Word.WdColor.wdColorBlack
        Table2.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
        Table2.Borders.InsideColor = Word.WdColor.wdColorBlack
        Table2.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
        oPara2.Format.SpaceAfter = 21
        oPara2.Range.InsertParagraphAfter()
        oPara2.Range.InsertBreak(Word.WdBreakType.wdPageBreak)

        oPara3 = oDoc.Content.Paragraphs.Add
        Table3 = oDoc.Content.Tables.Add(oPara3.Range, 20, 3)
        Table3.Rows(1).Cells.Merge()
        Table3.Rows(1).Range.Text = "Ground Distance Z1 QUAD"
        Table3.Rows(1).Range.Font.Bold = True
        Table3.Rows(1).Range.Font.Size = 16
        Table3.Rows(2).Cells.Split()
        Table3.Cell(2, 1).Range.Text = "Gnd Dis Z1 Reach"
        Table3.Cell(2, 2).Range.Text = tempN
        Table3.Cell(2, 3).Range.Text = "Ohm"
        Table3.Cell(3, 1).Range.Text = "Gnd Dis Z1 Direction"
        Table3.Cell(3, 2).Range.Text = "FORWARD"
        Table3.Cell(4, 1).Range.Text = "Gnd Dis Z1 Comp Limit"
        Table3.Cell(4, 2).Range.Text = tempO
        Table3.Cell(4, 3).Range.Text = "Degree"
        Table3.Cell(5, 1).Range.Text = "Gnd Dis Z1 Delay"
        Table3.Cell(5, 2).Range.Text = tempP
        Table3.Cell(5, 3).Range.Text = "Sec"
        Table3.Cell(6, 1).Range.Text = "Gnd Diz Z1 Supv"
        Table3.Cell(6, 2).Range.Text = "0.2"
        Table3.Cell(6, 3).Range.Text = "Pu"
        Table3.Cell(7, 1).Range.Text = "Z0/Z1 Mag"
        Table3.Cell(7, 2).Range.Text = tempQ
        Table3.Cell(8, 1).Range.Text = "Z0/Z1 Ang"
        Table3.Cell(8, 2).Range.Text = tempR
        Table3.Cell(8, 3).Range.Text = "Degree"
        Table3.Cell(9, 1).Range.Text = "Z0M/Z1 Mag"
        Table3.Cell(9, 2).Range.Text = "0"
        Table3.Cell(10, 1).Range.Text = "Z0M/Z1 Ang"
        Table3.Cell(10, 2).Range.Text = "0"
        Table3.Cell(10, 3).Range.Text = "Degree"
        Table3.Cell(11, 1).Range.Text = "Right Blinder Magnitude"
        Table3.Cell(11, 2).Range.Text = tempS
        Table3.Cell(11, 3).Range.Text = "Ohm"
        Table3.Cell(12, 1).Range.Text = "Right Blinder Angle"
        Table3.Cell(12, 2).Range.Text = tempT
        Table3.Cell(12, 3).Range.Text = "Degree"
        Table3.Cell(13, 1).Range.Text = "Left Blinder Magnitude"
        Table3.Cell(13, 2).Range.Text = tempU
        Table3.Cell(13, 3).Range.Text = "Ohm"
        Table3.Cell(14, 1).Range.Text = "Left Blinder Angle"
        Table3.Cell(14, 2).Range.Text = tempV
        Table3.Cell(14, 3).Range.Text = "Degree"
        Table3.Cell(15, 1).Range.Text = "RCA"
        Table3.Cell(15, 2).Range.Text = "75"
        Table3.Cell(15, 3).Range.Text = "Degree"
        Table3.Cell(16, 1).Range.Text = "COMPLIMIT"
        Table3.Cell(16, 2).Range.Text = "90"
        Table3.Cell(16, 3).Range.Text = "Degree"
        Table3.Cell(17, 1).Range.Text = "DIR RCA"
        Table3.Cell(17, 2).Range.Text = "45"
        Table3.Cell(17, 3).Range.Text = "Degree"
        Table3.Cell(18, 1).Range.Text = "DIR COMPLIMIT"
        Table3.Cell(18, 2).Range.Text = "60"
        Table3.Cell(18, 3).Range.Text = "Degree"
        Table3.Cell(19, 1).Range.Text = "RIGHT BLINDER RCA"
        Table3.Cell(19, 2).Range.Text = "75"
        Table3.Cell(19, 3).Range.Text = "Degree"
        Table3.Cell(20, 1).Range.Text = "LEFT BLINDER RCA"
        Table3.Cell(20, 2).Range.Text = "75"
        Table3.Cell(20, 3).Range.Text = "Degree"
        Table3.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        Table3.Borders.OutsideColor = Word.WdColor.wdColorBlack
        Table3.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
        Table3.Borders.InsideColor = Word.WdColor.wdColorBlack
        Table3.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
        oPara3.Format.SpaceAfter = 21
        oPara3.Range.InsertParagraphAfter()
        oPara3.Range.InsertBreak(Word.WdBreakType.wdPageBreak)

        oPara4 = oDoc.Content.Paragraphs.Add
        oPara4.Range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
        Table4 = oDoc.Content.Tables.Add(oPara4.Range, 10, 3)
        Table4.Rows(1).Cells.Merge()
        Table4.Rows(1).Range.Text = "Phase Distance Z2 MHO"
        Table4.Rows(1).Range.Font.Bold = True
        Table4.Rows(1).Range.Font.Size = 16
        Table4.Rows(2).Cells.Split()
        Table4.Cell(2, 1).Range.Text = "Ph Dis Z2 Reach"
        Table4.Cell(2, 2).Range.Text = tempW
        Table4.Cell(2, 3).Range.Text = "Ohm"
        Table4.Cell(3, 1).Range.Text = "Ph Dis Z2 Direction"
        Table4.Cell(3, 2).Range.Text = "FORWARD"
        Table4.Cell(4, 1).Range.Text = "Ph Dis Z2 Comp Limit"
        Table4.Cell(4, 2).Range.Text = tempX
        Table4.Cell(4, 3).Range.Text = "Degree"
        Table4.Cell(5, 1).Range.Text = "Ph Dis Z2 Delay"
        Table4.Cell(5, 2).Range.Text = tempY
        Table4.Cell(5, 3).Range.Text = "Sec"
        Table4.Cell(6, 1).Range.Text = "Ph Dis Z2 Supv"
        Table4.Cell(6, 2).Range.Text = "1.2"
        Table4.Cell(6, 3).Range.Text = "Pu"
        Table4.Cell(7, 1).Range.Text = "RCA"
        Table4.Cell(7, 2).Range.Text = tempZ
        Table4.Cell(7, 3).Range.Text = "Degree"
        Table4.Cell(8, 1).Range.Text = "COMPLIMIT"
        Table4.Cell(8, 2).Range.Text = "90"
        Table4.Cell(8, 3).Range.Text = "Degree"
        Table4.Cell(9, 1).Range.Text = "DIR RCA"
        Table4.Cell(9, 2).Range.Text = "80"
        Table4.Cell(9, 3).Range.Text = "Degree"
        Table4.Cell(10, 1).Range.Text = "DIR COMPLIMIT"
        Table4.Cell(10, 2).Range.Text = "90"
        Table4.Cell(10, 3).Range.Text = "Degree"
        Table4.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        Table4.Borders.OutsideColor = Word.WdColor.wdColorBlack
        Table4.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
        Table4.Borders.InsideColor = Word.WdColor.wdColorBlack
        Table4.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
        oPara4.Format.SpaceAfter = 21
        oPara4.Range.InsertParagraphAfter()
        oPara4.Range.InsertBreak(Word.WdBreakType.wdPageBreak)

        oPara5 = oDoc.Content.Paragraphs.Add
        Table5 = oDoc.Content.Tables.Add(oPara5.Range, 20, 3)
        Table5.Rows(1).Cells.Merge()
        Table5.Rows(1).Range.Text = "Ground Distance Z2 QUAD"
        Table5.Rows(1).Range.Font.Bold = True
        Table5.Rows(1).Range.Font.Size = 16
        Table5.Rows(2).Cells.Split()
        Table5.Cell(2, 1).Range.Text = "Gnd Dis Z2 Reach"
        Table5.Cell(2, 2).Range.Text = tempA1
        Table5.Cell(2, 3).Range.Text = "Ohm"
        Table5.Cell(3, 1).Range.Text = "Gnd Dis Z2 Direction"
        Table5.Cell(3, 2).Range.Text = "FORWARD"
        Table5.Cell(4, 1).Range.Text = "Gnd Dis Z2 Comp Limit"
        Table5.Cell(4, 2).Range.Text = tempB1
        Table5.Cell(4, 3).Range.Text = "Degree"
        Table5.Cell(5, 1).Range.Text = "Gnd Dis Z2 Delay"
        Table5.Cell(5, 2).Range.Text = tempC1
        Table5.Cell(5, 3).Range.Text = "Sec"
        Table5.Cell(6, 1).Range.Text = "Gnd Diz Z2 Supv"
        Table5.Cell(6, 2).Range.Text = "0.2"
        Table5.Cell(6, 3).Range.Text = "Pu"
        Table5.Cell(7, 1).Range.Text = "Z0/Z1 Mag"
        Table5.Cell(7, 2).Range.Text = tempD1
        Table5.Cell(8, 1).Range.Text = "Z0/Z1 Ang"
        Table5.Cell(8, 2).Range.Text = tempE1
        Table5.Cell(8, 3).Range.Text = "Degree"
        Table5.Cell(9, 1).Range.Text = "Z0M/Z1 Mag"
        Table5.Cell(9, 2).Range.Text = "0"
        Table5.Cell(10, 1).Range.Text = "Z0M/Z1 Ang"
        Table5.Cell(10, 2).Range.Text = "0"
        Table5.Cell(10, 3).Range.Text = "Degree"
        Table5.Cell(11, 1).Range.Text = "Right Blinder Magnitude"
        Table5.Cell(11, 2).Range.Text = tempF1
        Table5.Cell(11, 3).Range.Text = "Ohm"
        Table5.Cell(12, 1).Range.Text = "Right Blinder Angle"
        Table5.Cell(12, 2).Range.Text = tempG1
        Table5.Cell(12, 3).Range.Text = "Degree"
        Table5.Cell(13, 1).Range.Text = "Left Blinder Magnitude"
        Table5.Cell(13, 2).Range.Text = tempH1
        Table5.Cell(13, 3).Range.Text = "Ohm"
        Table5.Cell(14, 1).Range.Text = "Left Blinder Angle"
        Table5.Cell(14, 2).Range.Text = tempI1
        Table5.Cell(14, 3).Range.Text = "Degree"
        Table5.Cell(15, 1).Range.Text = "RCA"
        Table5.Cell(15, 2).Range.Text = "75"
        Table5.Cell(15, 3).Range.Text = "Degree"
        Table5.Cell(16, 1).Range.Text = "COMPLIMIT"
        Table5.Cell(16, 2).Range.Text = "90"
        Table5.Cell(16, 3).Range.Text = "Degree"
        Table5.Cell(17, 1).Range.Text = "DIR RCA"
        Table5.Cell(17, 2).Range.Text = "45"
        Table5.Cell(17, 3).Range.Text = "Degree"
        Table5.Cell(18, 1).Range.Text = "DIR COMPLIMIT"
        Table5.Cell(18, 2).Range.Text = "60"
        Table5.Cell(18, 3).Range.Text = "Degree"
        Table5.Cell(19, 1).Range.Text = "RIGHT BLINDER RCA"
        Table5.Cell(19, 2).Range.Text = "75"
        Table5.Cell(19, 3).Range.Text = "Degree"
        Table5.Cell(20, 1).Range.Text = "LEFT BLINDER RCA"
        Table5.Cell(20, 2).Range.Text = "75"
        Table5.Cell(20, 3).Range.Text = "Degree"
        Table5.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        Table5.Borders.OutsideColor = Word.WdColor.wdColorBlack
        Table5.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
        Table5.Borders.InsideColor = Word.WdColor.wdColorBlack
        Table5.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
        oPara5.Format.SpaceAfter = 21
        oPara5.Range.InsertParagraphAfter()
        oPara5.Range.InsertBreak(Word.WdBreakType.wdPageBreak)

        oPara6 = oDoc.Content.Paragraphs.Add
        oPara6.Range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
        Table6 = oDoc.Content.Tables.Add(oPara6.Range, 10, 3)
        Table6.Rows(1).Cells.Merge()
        Table6.Rows(1).Range.Text = "Phase Distance Z3 MHO"
        Table6.Rows(1).Range.Font.Bold = True
        Table6.Rows(1).Range.Font.Size = 16
        Table6.Rows(2).Cells.Split()
        Table6.Cell(2, 1).Range.Text = "Ph Dis Z3 Reach"
        Table6.Cell(2, 2).Range.Text = tempJ1
        Table6.Cell(2, 3).Range.Text = "Ohm"
        Table6.Cell(3, 1).Range.Text = "Ph Dis Z3 Direction"
        Table6.Cell(3, 2).Range.Text = "FORWARD"
        Table6.Cell(4, 1).Range.Text = "Ph Dis Z3 Comp Limit"
        Table6.Cell(4, 2).Range.Text = tempK1
        Table6.Cell(4, 3).Range.Text = "Degree"
        Table6.Cell(5, 1).Range.Text = "Ph Dis Z3 Delay"
        Table6.Cell(5, 2).Range.Text = tempL1
        Table6.Cell(5, 3).Range.Text = "Sec"
        Table6.Cell(6, 1).Range.Text = "Ph Dis Z3 Supv"
        Table6.Cell(6, 2).Range.Text = "1.2"
        Table6.Cell(6, 3).Range.Text = "Pu"
        Table6.Cell(7, 1).Range.Text = "RCA"
        Table6.Cell(7, 2).Range.Text = tempM1
        Table6.Cell(7, 3).Range.Text = "Degree"
        Table6.Cell(8, 1).Range.Text = "COMPLIMIT"
        Table6.Cell(8, 2).Range.Text = "90"
        Table6.Cell(8, 3).Range.Text = "Degree"
        Table6.Cell(9, 1).Range.Text = "DIR RCA"
        Table6.Cell(9, 2).Range.Text = "80"
        Table6.Cell(9, 3).Range.Text = "Degree"
        Table6.Cell(10, 1).Range.Text = "DIR COMPLIMIT"
        Table6.Cell(10, 2).Range.Text = "90"
        Table6.Cell(10, 3).Range.Text = "Degree"
        Table6.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        Table6.Borders.OutsideColor = Word.WdColor.wdColorBlack
        Table6.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
        Table6.Borders.InsideColor = Word.WdColor.wdColorBlack
        Table6.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
        oPara6.Format.SpaceAfter = 21
        oPara6.Range.InsertParagraphAfter()
        oPara6.Range.InsertBreak(Word.WdBreakType.wdPageBreak)

        oPara7 = oDoc.Content.Paragraphs.Add
        Table7 = oDoc.Content.Tables.Add(oPara7.Range, 20, 3)
        Table7.Rows(1).Cells.Merge()
        Table7.Rows(1).Range.Text = "Ground Distance Z3 QUAD"
        Table7.Rows(1).Range.Font.Bold = True
        Table7.Rows(1).Range.Font.Size = 16
        Table7.Rows(2).Cells.Split()
        Table7.Cell(2, 1).Range.Text = "Gnd Dis Z3 Reach"
        Table7.Cell(2, 2).Range.Text = tempN1
        Table7.Cell(2, 3).Range.Text = "Ohm"
        Table7.Cell(3, 1).Range.Text = "Gnd Dis Z3 Direction"
        Table7.Cell(3, 2).Range.Text = "FORWARD"
        Table7.Cell(4, 1).Range.Text = "Gnd Dis Z3 Comp Limit"
        Table7.Cell(4, 2).Range.Text = tempO1
        Table7.Cell(4, 3).Range.Text = "Degree"
        Table7.Cell(5, 1).Range.Text = "Gnd Dis Z3 Delay"
        Table7.Cell(5, 2).Range.Text = tempP1
        Table7.Cell(5, 3).Range.Text = "Sec"
        Table7.Cell(6, 1).Range.Text = "Gnd Diz Z3 Supv"
        Table7.Cell(6, 2).Range.Text = "0.2"
        Table7.Cell(6, 3).Range.Text = "Pu"
        Table7.Cell(7, 1).Range.Text = "Z0/Z1 Mag"
        Table7.Cell(7, 2).Range.Text = tempQ1
        Table7.Cell(8, 1).Range.Text = "Z0/Z1 Ang"
        Table7.Cell(8, 2).Range.Text = tempR1
        Table7.Cell(8, 3).Range.Text = "Degree"
        Table7.Cell(9, 1).Range.Text = "Z0M/Z1 Mag"
        Table7.Cell(9, 2).Range.Text = "0"
        Table7.Cell(10, 1).Range.Text = "Z0M/Z1 Ang"
        Table7.Cell(10, 2).Range.Text = "0"
        Table7.Cell(10, 3).Range.Text = "Degree"
        Table7.Cell(11, 1).Range.Text = "Right Blinder Magnitude"
        Table7.Cell(11, 2).Range.Text = tempS1
        Table7.Cell(11, 3).Range.Text = "Ohm"
        Table7.Cell(12, 1).Range.Text = "Right Blinder Angle"
        Table7.Cell(12, 2).Range.Text = tempT1
        Table7.Cell(12, 3).Range.Text = "Degree"
        Table7.Cell(13, 1).Range.Text = "Left Blinder Magnitude"
        Table7.Cell(13, 2).Range.Text = tempU1
        Table7.Cell(13, 3).Range.Text = "Ohm"
        Table7.Cell(14, 1).Range.Text = "Left Blinder Angle"
        Table7.Cell(14, 2).Range.Text = tempV1
        Table7.Cell(14, 3).Range.Text = "Degree"
        Table7.Cell(15, 1).Range.Text = "RCA"
        Table7.Cell(15, 2).Range.Text = "75"
        Table7.Cell(15, 3).Range.Text = "Degree"
        Table7.Cell(16, 1).Range.Text = "COMPLIMIT"
        Table7.Cell(16, 2).Range.Text = "90"
        Table7.Cell(16, 3).Range.Text = "Degree"
        Table7.Cell(17, 1).Range.Text = "DIR RCA"
        Table7.Cell(17, 2).Range.Text = "45"
        Table7.Cell(17, 3).Range.Text = "Degree"
        Table7.Cell(18, 1).Range.Text = "DIR COMPLIMIT"
        Table7.Cell(18, 2).Range.Text = "60"
        Table7.Cell(18, 3).Range.Text = "Degree"
        Table7.Cell(19, 1).Range.Text = "RIGHT BLINDER RCA"
        Table7.Cell(19, 2).Range.Text = "75"
        Table7.Cell(19, 3).Range.Text = "Degree"
        Table7.Cell(20, 1).Range.Text = "LEFT BLINDER RCA"
        Table7.Cell(20, 2).Range.Text = "75"
        Table7.Cell(20, 3).Range.Text = "Degree"
        Table7.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        Table7.Borders.OutsideColor = Word.WdColor.wdColorBlack
        Table7.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
        Table7.Borders.InsideColor = Word.WdColor.wdColorBlack
        Table7.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
        oPara7.Format.SpaceAfter = 21
        oPara7.Range.InsertParagraphAfter()
        oPara7.Range.InsertBreak(Word.WdBreakType.wdPageBreak)

        oPara8 = oDoc.Content.Paragraphs.Add
        oPara8.Range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
        Table8 = oDoc.Content.Tables.Add(oPara6.Range, 5, 3)
        Table8.Rows(1).Cells.Merge()
        Table8.Rows(1).Range.Text = "Power Swing Element:"
        Table8.Rows(1).Range.Font.Bold = True
        Table8.Rows(1).Range.Font.Size = 16
        Table8.Rows(2).Cells.Split()
        Table8.Cell(2, 1).Range.Text = "Power Swing Reach"
        Table8.Cell(2, 2).Range.Text = tempW1
        Table8.Cell(2, 3).Range.Text = "Ohm"
        Table8.Cell(3, 1).Range.Text = "Power Swing Inner"
        Table8.Cell(3, 2).Range.Text = tempX1
        Table8.Cell(3, 3).Range.Text = "Ohm"
        Table8.Cell(4, 1).Range.Text = "Power Swing Reach"
        Table8.Cell(4, 2).Range.Text = tempY1
        Table8.Cell(4, 3).Range.Text = "Ohm"
        Table8.Cell(5, 1).Range.Text = "Delay Pickup"
        Table8.Cell(5, 2).Range.Text = tempZ1
        Table8.Cell(5, 3).Range.Text = "ms"
        Table8.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        Table8.Borders.OutsideColor = Word.WdColor.wdColorBlack
        Table8.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
        Table8.Borders.InsideColor = Word.WdColor.wdColorBlack
        Table8.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
        oPara8.Format.SpaceAfter = 21

        Me.Close()
        MessageBox.Show("Export to Word Complete", "Distance Relay Calculation",
                            MessageBoxButtons.OK)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
End Class