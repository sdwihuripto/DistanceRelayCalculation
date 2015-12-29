Imports Microsoft.Office.Interop

Public Class FormOutputSiemens

    '41 OUTPUT
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
    Dim tempO1 As Double

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
                     ByVal o1 As Double)
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
        Out1E.Text = tempN
        Out1F.Text = tempO
        Out1G.Text = tempP
        Out2A.Text = tempQ
        Out2B.Text = tempR
        Out2C.Text = tempS
        Out2D.Text = tempT
        Out3A.Text = tempU
        Out3B.Text = tempV
        Out3C.Text = tempW
        Out3D.Text = tempX
        Out3E.Text = tempY
        Out3F.Text = tempZ
        Out3G.Text = tempA1
        Out3H.Text = tempB1
        Out3I.Text = tempC1
        Out3J.Text = tempD1
        Out3K.Text = tempE1
        Out3L.Text = tempF1
        Out4A.Text = tempG1
        Out4B.Text = tempH1
        Out4C.Text = tempI1
        Out4D.Text = tempJ1
        Out5A.Text = tempK1
        Out5B.Text = tempL1
        Out5C.Text = tempM1
        Out5D.Text = tempN1
        Out5E.Text = tempO1
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
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
        Table2 = oDoc.Content.Tables.Add(oPara2.Range, 8, 4)
        Table2.Rows(1).Cells.Merge()
        Table2.Rows(1).Range.Text = "Power System Data"
        Table2.Rows(1).Range.Font.Bold = True
        Table2.Rows(1).Range.Font.Size = 16
        Table2.Rows(2).Cells.Split()
        Table2.Cell(2, 1).Range.Text = "1105"
        Table2.Cell(2, 2).Range.Text = "Line Angle"
        Table2.Cell(2, 3).Range.Text = tempJ
        Table2.Cell(2, 4).Range.Text = "Degree"
        Table2.Cell(3, 1).Range.Text = "1110"
        Table2.Cell(3, 2).Range.Text = "X'"
        Table2.Cell(3, 3).Range.Text = tempK
        Table2.Cell(3, 4).Range.Text = "Ohm/km"
        Table2.Cell(4, 1).Range.Text = "1111"
        Table2.Cell(4, 2).Range.Text = "Line Length"
        Table2.Cell(4, 3).Range.Text = tempL
        Table2.Cell(4, 4).Range.Text = "Km"
        Table2.Cell(5, 1).Range.Text = "1116"
        Table2.Cell(5, 2).Range.Text = "RE/RL (Z1)"
        Table2.Cell(5, 3).Range.Text = tempM
        Table2.Cell(6, 1).Range.Text = "1117"
        Table2.Cell(6, 2).Range.Text = "XE/XL (Z1)"
        Table2.Cell(6, 3).Range.Text = tempN
        Table2.Cell(7, 1).Range.Text = "1118"
        Table2.Cell(7, 2).Range.Text = "RE/RL (ZB.Z5)"
        Table2.Cell(7, 3).Range.Text = tempO
        Table2.Cell(8, 1).Range.Text = "1119"
        Table2.Cell(8, 2).Range.Text = "XE/XL (ZB.Z5)"
        Table2.Cell(8, 3).Range.Text = tempP
        Table2.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        Table2.Borders.OutsideColor = Word.WdColor.wdColorBlack
        Table2.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
        Table2.Borders.InsideColor = Word.WdColor.wdColorBlack
        Table2.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
        oPara2.Format.SpaceAfter = 21
        oPara2.Range.InsertParagraphAfter()
        oPara2.Range.InsertBreak(Word.WdBreakType.wdPageBreak)

        oPara3 = oDoc.Content.Paragraphs.Add
        oPara3.Range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
        Table3 = oDoc.Content.Tables.Add(oPara3.Range, 5, 4)
        Table3.Rows(1).Cells.Merge()
        Table3.Rows(1).Range.Text = "21 Distance Protection, General Setting"
        Table3.Rows(1).Range.Font.Bold = True
        Table3.Rows(1).Range.Font.Size = 16
        Table3.Rows(2).Cells.Split()
        Table3.Cell(2, 1).Range.Text = "1241"
        Table3.Cell(2, 2).Range.Text = "R load (ph-E)"
        Table3.Cell(2, 3).Range.Text = tempQ
        Table3.Cell(2, 4).Range.Text = "Ohm"
        Table3.Cell(3, 1).Range.Text = "1242"
        Table3.Cell(3, 2).Range.Text = "Phi load (ph-E)"
        Table3.Cell(3, 3).Range.Text = tempR
        Table3.Cell(3, 4).Range.Text = "Degree"
        Table3.Cell(4, 1).Range.Text = "1243"
        Table3.Cell(4, 2).Range.Text = "R load (ph-ph)"
        Table3.Cell(4, 3).Range.Text = tempS
        Table3.Cell(4, 4).Range.Text = "Ohm"
        Table3.Cell(5, 1).Range.Text = "1244"
        Table3.Cell(5, 2).Range.Text = "Phi load (ph-ph)"
        Table3.Cell(5, 3).Range.Text = tempT
        Table3.Cell(5, 4).Range.Text = "Degree"
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
        Table4 = oDoc.Content.Tables.Add(oPara4.Range, 21, 4)
        Table4.Rows(1).Cells.Merge()
        Table4.Rows(1).Range.Text = "21 Distance Zone Quadrilateral"
        Table4.Rows(1).Range.Font.Bold = True
        Table4.Rows(1).Range.Font.Size = 16
        Table4.Rows(2).Cells.Merge()
        Table4.Rows(2).Range.Text = "Group Zone1 Setting"
        Table4.Rows(3).Cells.Split()
        Table4.Cell(3, 1).Range.Text = "1301"
        Table4.Cell(3, 2).Range.Text = "Op. mode Z1"
        Table4.Cell(3, 3).Range.Text = "Forward"
        Table4.Cell(4, 1).Range.Text = "1302"
        Table4.Cell(4, 2).Range.Text = "R(Z1)"
        Table4.Cell(4, 3).Range.Text = tempU
        Table4.Cell(4, 4).Range.Text = "Ohm"
        Table4.Cell(5, 1).Range.Text = "1303"
        Table4.Cell(5, 2).Range.Text = "X(Z1)"
        Table4.Cell(5, 3).Range.Text = tempV
        Table4.Cell(5, 4).Range.Text = "Ohm"
        Table4.Cell(6, 1).Range.Text = "1302"
        Table4.Cell(6, 2).Range.Text = "RG(Z1)"
        Table4.Cell(6, 3).Range.Text = tempW
        Table4.Rows(7).Cells.Merge()
        Table4.Rows(7).Range.Text = "Group Zone1B.Setting"
        Table4.Rows(8).Cells.Split()
        Table4.Cell(8, 1).Range.Text = "1351"
        Table4.Cell(8, 2).Range.Text = "Op. mode Z1B"
        Table4.Cell(8, 3).Range.Text = "Forward"
        Table4.Cell(9, 1).Range.Text = "1352"
        Table4.Cell(9, 2).Range.Text = "R(Z1B)"
        Table4.Cell(9, 3).Range.Text = tempX
        Table4.Cell(9, 4).Range.Text = "Ohm"
        Table4.Cell(10, 1).Range.Text = "1353"
        Table4.Cell(10, 2).Range.Text = "X(Z1B)"
        Table4.Cell(10, 3).Range.Text = tempY
        Table4.Cell(10, 4).Range.Text = "Ohm"
        Table4.Cell(11, 1).Range.Text = "1352"
        Table4.Cell(11, 2).Range.Text = "RG(Z1B)"
        Table4.Cell(11, 3).Range.Text = tempZ
        Table4.Cell(11, 4).Range.Text = "Ohm"
        Table4.Rows(12).Cells.Merge()
        Table4.Rows(12).Range.Text = "Group Zone2 Setting"
        Table4.Rows(13).Cells.Split()
        Table4.Cell(13, 1).Range.Text = "1311"
        Table4.Cell(13, 2).Range.Text = "Op. mode Z2"
        Table4.Cell(13, 3).Range.Text = "Forward"
        Table4.Cell(14, 1).Range.Text = "1312"
        Table4.Cell(14, 2).Range.Text = "R(Z2)"
        Table4.Cell(14, 3).Range.Text = tempA1
        Table4.Cell(14, 4).Range.Text = "Ohm"
        Table4.Cell(15, 1).Range.Text = "1313"
        Table4.Cell(15, 2).Range.Text = "X(Z2)"
        Table4.Cell(15, 3).Range.Text = tempB1
        Table4.Cell(15, 4).Range.Text = "Ohm"
        Table4.Cell(16, 1).Range.Text = "1314"
        Table4.Cell(16, 2).Range.Text = "RG(Z2)"
        Table4.Cell(16, 3).Range.Text = tempC1
        Table4.Cell(16, 4).Range.Text = "Ohm"
        Table4.Rows(17).Cells.Merge()
        Table4.Rows(17).Range.Text = "Group Zone3 Setting"
        Table4.Rows(18).Cells.Split()
        Table4.Cell(18, 1).Range.Text = "1321"
        Table4.Cell(18, 2).Range.Text = "Op. mode Z3"
        Table4.Cell(18, 3).Range.Text = "Forward"
        Table4.Cell(19, 1).Range.Text = "1322"
        Table4.Cell(19, 2).Range.Text = "R(Z3)"
        Table4.Cell(19, 3).Range.Text = tempD1
        Table4.Cell(19, 4).Range.Text = "Ohm"
        Table4.Cell(20, 1).Range.Text = "1323"
        Table4.Cell(20, 2).Range.Text = "X(Z3)"
        Table4.Cell(20, 3).Range.Text = tempE1
        Table4.Cell(20, 4).Range.Text = "Ohm"
        Table4.Cell(21, 1).Range.Text = "1324"
        Table4.Cell(21, 2).Range.Text = "RG(Z3)"
        Table4.Cell(21, 3).Range.Text = tempF1
        Table4.Cell(21, 4).Range.Text = "Ohm"
        Table4.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        Table4.Borders.OutsideColor = Word.WdColor.wdColorBlack
        Table4.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
        Table4.Borders.InsideColor = Word.WdColor.wdColorBlack
        Table4.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
        oPara4.Format.SpaceAfter = 21
        oPara4.Range.InsertParagraphAfter()
        oPara4.Range.InsertBreak(Word.WdBreakType.wdPageBreak)

        oPara5 = oDoc.Content.Paragraphs.Add
        oPara5.Range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
        Table5 = oDoc.Content.Tables.Add(oPara5.Range, 13, 4)
        Table5.Rows(1).Cells.Merge()
        Table5.Rows(1).Range.Text = "21 Distance Zone MHO"
        Table5.Rows(1).Range.Font.Bold = True
        Table5.Rows(1).Range.Font.Size = 16
        Table5.Rows(2).Cells.Merge()
        Table5.Rows(2).Range.Text = "Group Zone1 (MHO) Setting"
        Table5.Rows(3).Cells.Split()
        Table5.Cell(3, 1).Range.Text = "1401"
        Table5.Cell(3, 2).Range.Text = "Op. mode Z1"
        Table5.Cell(3, 3).Range.Text = "Forward"
        Table5.Cell(4, 1).Range.Text = "1402"
        Table5.Cell(4, 2).Range.Text = "ZR(Z1)"
        Table5.Cell(4, 3).Range.Text = tempG1
        Table5.Cell(4, 4).Range.Text = "Ohm"
        Table5.Rows(5).Cells.Merge()
        Table5.Rows(5).Range.Text = "Group Zone1B-Extends (MHO) Setting"
        Table5.Rows(6).Cells.Split()
        Table5.Cell(6, 1).Range.Text = "1451"
        Table5.Cell(6, 2).Range.Text = "Op. mode Z1B"
        Table5.Cell(6, 3).Range.Text = "Forward"
        Table5.Cell(7, 1).Range.Text = "1452"
        Table5.Cell(7, 2).Range.Text = "ZR(Z1B)"
        Table5.Cell(7, 3).Range.Text = tempH1
        Table5.Cell(7, 4).Range.Text = "Ohm"
        Table5.Rows(8).Cells.Merge()
        Table5.Rows(8).Range.Text = "Group Zone2(MHO) Setting"
        Table5.Rows(9).Cells.Split()
        Table5.Cell(9, 1).Range.Text = "1411"
        Table5.Cell(9, 2).Range.Text = "Op. mode Z2"
        Table5.Cell(9, 3).Range.Text = "Forward"
        Table5.Cell(10, 1).Range.Text = "1412"
        Table5.Cell(10, 2).Range.Text = "ZR(Z2)"
        Table5.Cell(10, 3).Range.Text = tempI1
        Table5.Cell(10, 4).Range.Text = "Ohm"
        Table5.Rows(11).Cells.Merge()
        Table5.Rows(11).Range.Text = "Group Zone3(MHO) Setting"
        Table5.Rows(12).Cells.Split()
        Table5.Cell(12, 1).Range.Text = "1421"
        Table5.Cell(12, 2).Range.Text = "Op. mode Z3"
        Table5.Cell(12, 3).Range.Text = "Forward"
        Table5.Cell(13, 1).Range.Text = "1422"
        Table5.Cell(13, 2).Range.Text = "ZR(Z3)"
        Table5.Cell(13, 3).Range.Text = tempJ1
        Table5.Cell(13, 4).Range.Text = "Ohm"
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
        Table6 = oDoc.Content.Tables.Add(oPara6.Range, 6, 4)
        Table6.Rows(1).Cells.Merge()
        Table6.Rows(1).Range.Text = "21 Distance Protection, time delays"
        Table6.Rows(1).Range.Font.Bold = True
        Table6.Rows(1).Range.Font.Size = 16
        Table6.Rows(2).Cells.Split()
        Table6.Cell(2, 1).Range.Text = "1305"
        Table6.Cell(2, 2).Range.Text = "T1-phase"
        Table6.Cell(2, 3).Range.Text = tempK
        Table6.Cell(2, 4).Range.Text = "Sec"
        Table6.Cell(3, 1).Range.Text = "1306"
        Table6.Cell(3, 2).Range.Text = "T1-multiphase"
        Table6.Cell(3, 3).Range.Text = tempL
        Table6.Cell(3, 4).Range.Text = "Sec"
        Table6.Cell(4, 1).Range.Text = "1315"
        Table6.Cell(4, 2).Range.Text = "T2-1phase"
        Table6.Cell(4, 3).Range.Text = tempM
        Table6.Cell(4, 4).Range.Text = "Sec"
        Table6.Cell(5, 1).Range.Text = "1316"
        Table6.Cell(5, 2).Range.Text = "T2-multiphase"
        Table6.Cell(5, 3).Range.Text = tempN
        Table6.Cell(5, 4).Range.Text = "Sec"
        Table6.Cell(6, 1).Range.Text = "1325"
        Table6.Cell(6, 2).Range.Text = "T3 Delay"
        Table6.Cell(6, 3).Range.Text = tempO
        Table6.Cell(6, 4).Range.Text = "Sec"
        Table6.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        Table6.Borders.OutsideColor = Word.WdColor.wdColorBlack
        Table6.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
        Table6.Borders.InsideColor = Word.WdColor.wdColorBlack
        Table6.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
        oPara6.Format.SpaceAfter = 21
        oPara6.Range.InsertParagraphAfter()
        oPara6.Range.InsertBreak(Word.WdBreakType.wdPageBreak)

        Me.Close()
        MessageBox.Show("Export to Word Complete", "Distance Relay Calculation",
                            MessageBoxButtons.OK)
    End Sub
End Class