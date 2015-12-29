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
        'oPara2.Range.InsertBreak(Word.WdBreakType.wdPageBreak)

        oPara3 = oDoc.Content.Paragraphs.Add
        Table3 = oDoc.Content.Tables.Add(oPara3.Range, 10, 3)
        Table3.Rows(1).Cells.Merge()
        Table3.Rows(1).Range.Text = "Phase Distance Z1 MHO"
        Table3.Rows(1).Range.Font.Bold = True
        Table3.Rows(1).Range.Font.Size = 16
        Table3.Rows(2).Cells.Split()
        Table3.Cell(2, 1).Range.Text = "Ph Dis Z1 Reach"
        Table3.Cell(2, 2).Range.Text = tempJ
        Table3.Cell(2, 3).Range.Text = "Ohm"
        Table3.Cell(3, 1).Range.Text = "Ph Dis Z1 Direction"
        Table3.Cell(3, 2).Range.Text = "FORWARD"
        Table3.Cell(4, 1).Range.Text = "Ph Dis Z1 Comp Limit"
        Table3.Cell(4, 2).Range.Text = tempK
        Table3.Cell(4, 3).Range.Text = "Degree"
        Table3.Cell(5, 1).Range.Text = "Ph Dis Z1 Delay"
        Table3.Cell(5, 2).Range.Text = tempL
        Table3.Cell(5, 3).Range.Text = "Sec"
        Table3.Cell(6, 1).Range.Text = "Ph Dis Z1 Supv"
        Table3.Cell(6, 2).Range.Text = "1.2"
        Table3.Cell(6, 3).Range.Text = "pu"
        Table3.Cell(7, 1).Range.Text = "RCA"
        Table3.Cell(7, 2).Range.Text = tempM
        Table3.Cell(7, 3).Range.Text = "Degree"
        Table3.Cell(8, 1).Range.Text = "COMPLIMIT"
        Table3.Cell(8, 2).Range.Text = "90"
        Table3.Cell(8, 3).Range.Text = "Degree"
        Table3.Cell(9, 1).Range.Text = "DIR RCA"
        Table3.Cell(9, 2).Range.Text = "80"
        Table3.Cell(9, 3).Range.Text = "Degree"
        Table3.Cell(10, 1).Range.Text = "DIR COMPLIMIT"
        Table3.Cell(10, 2).Range.Text = "90"
        Table3.Cell(10, 3).Range.Text = "Degree"
        Table3.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        Table3.Borders.OutsideColor = Word.WdColor.wdColorBlack
        Table3.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
        Table3.Borders.InsideColor = Word.WdColor.wdColorBlack
        Table3.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
        oPara3.Format.SpaceAfter = 21
        oPara3.Range.InsertParagraphAfter()

        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
End Class