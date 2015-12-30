Imports Microsoft.Office.Interop

Public Class FormOutputAlstom

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
                     ByVal y As Double, ByVal z As Double)
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

        OutSum1.Text = tempA
        OutSum2.Text = tempB
        OutSum3.Text = tempC
        OutSum4.Text = tempD
        OutSum5.Text = tempE
        OutSum6.Text = tempF
        Out1.Text = tempG
        Out2.Text = tempH
        Out3.Text = tempI
        Out4.Text = tempJ
        Out5.Text = tempK
        Out6.Text = tempL
        Out7.Text = tempM
        Out8.Text = tempN
        Out9.Text = tempO
        Out10.Text = tempP
        Out11.Text = tempQ
        Out12.Text = tempR
        Out13.Text = tempS
        Out14.Text = tempT
        Out15.Text = tempU
        Out16.Text = tempV
        Out17.Text = tempW
        Out18.Text = tempX
        Out19.Text = tempY
        Out10.Text = tempZ
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
        Table1 = oDoc.Content.Tables.Add(oPara1.Range, 4, 2)
        Table1.Columns(1).Cells(1).Range.Text = "Location"
        Table1.Columns(2).Cells(1).Range.Text = tempA
        Table1.Columns(1).Cells(2).Range.Text = "Line Bay To"
        Table1.Columns(2).Cells(2).Range.Text = tempB
        Table1.Columns(1).Cells(3).Range.Text = "CT Ratio"
        Table1.Columns(2).Cells(3).Range.Text = tempC & " / " & tempD
        Table1.Columns(1).Cells(4).Range.Text = "PT Ratio"
        Table1.Columns(2).Cells(4).Range.Text = tempE & " / " & tempF
        Table1.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        Table1.Borders.OutsideColor = Word.WdColor.wdColorBlack
        Table1.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
        Table1.Borders.InsideColor = Word.WdColor.wdColorBlack
        Table1.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
        oPara1.Format.SpaceAfter = 21
        oPara1.Range.InsertParagraphAfter()

        oPara2 = oDoc.Content.Paragraphs.Add
        oPara2.Range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
        Table2 = oDoc.Content.Tables.Add(oPara2.Range, 4, 3)
        Table2.Rows(1).Cells.Merge()
        Table2.Rows(1).Range.Text = "Line Setting"
        Table2.Rows(1).Range.Font.Bold = True
        Table2.Rows(1).Range.Font.Size = 16
        Table2.Rows(2).Cells.Split()
        Table2.Cell(2, 1).Range.Text = "L"
        Table2.Cell(2, 2).Range.Text = tempG
        Table2.Cell(2, 3).Range.Text = "Km"
        Table2.Cell(3, 1).Range.Text = "ZL"
        Table2.Cell(3, 2).Range.Text = tempH
        Table2.Cell(3, 3).Range.Text = "Ohm"
        Table2.Cell(4, 1).Range.Text = "Theta L"
        Table2.Cell(4, 2).Range.Text = tempI
        Table2.Cell(4, 3).Range.Text = "Degree"
        Table2.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        Table2.Borders.OutsideColor = Word.WdColor.wdColorBlack
        Table2.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
        Table2.Borders.InsideColor = Word.WdColor.wdColorBlack
        Table2.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
        oPara2.Format.SpaceAfter = 21
        oPara2.Range.InsertParagraphAfter()

        oPara3 = oDoc.Content.Paragraphs.Add
        oPara3.Range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
        Table3 = oDoc.Content.Tables.Add(oPara3.Range, 3, 3)
        Table3.Rows(1).Cells.Merge()
        Table3.Rows(1).Range.Text = "Ground Fault Compensation Setting"
        Table3.Rows(1).Range.Font.Bold = True
        Table3.Rows(1).Range.Font.Size = 16
        Table3.Rows(2).Cells.Split()
        Table3.Cell(2, 1).Range.Text = "kZ0"
        Table3.Cell(2, 2).Range.Text = tempJ
        Table3.Cell(3, 1).Range.Text = "Theta kZ0"
        Table3.Cell(3, 2).Range.Text = tempK
        Table3.Cell(3, 3).Range.Text = "Degree"
        Table3.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        Table3.Borders.OutsideColor = Word.WdColor.wdColorBlack
        Table3.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
        Table3.Borders.InsideColor = Word.WdColor.wdColorBlack
        Table3.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
        oPara3.Format.SpaceAfter = 21
        oPara3.Range.InsertParagraphAfter()
        oPara3.Range.InsertBreak(Word.WdBreakType.wdPageBreak)

        oPara3 = oDoc.Content.Paragraphs.Add
        oPara3.Range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
        Table3 = oDoc.Content.Tables.Add(oPara3.Range, 9, 3)
        Table3.Rows(1).Cells.Merge()
        Table3.Rows(1).Range.Text = "Zone1"
        Table3.Rows(1).Range.Font.Bold = True
        Table3.Rows(1).Range.Font.Size = 16
        Table3.Rows(2).Cells.Split()
        Table3.Cell(2, 1).Range.Text = "Z1"
        Table3.Cell(2, 2).Range.Text = tempL
        Table3.Cell(2, 3).Range.Text = "Ohm"
        Table3.Cell(3, 1).Range.Text = "tZ1"
        Table3.Cell(3, 2).Range.Text = tempM
        Table3.Cell(3, 3).Range.Text = "sec"
        Table3.Rows(4).Cells.Merge()
        Table3.Rows(4).Range.Text = "Zone2"
        Table3.Rows(5).Cells.Split()
        Table3.Cell(5, 1).Range.Text = "Z2"
        Table3.Cell(5, 2).Range.Text = tempN
        Table3.Cell(5, 3).Range.Text = "Ohm"
        Table3.Cell(6, 1).Range.Text = "tZ2"
        Table3.Cell(6, 2).Range.Text = tempO
        Table3.Cell(6, 3).Range.Text = "sec"
        Table3.Rows(7).Cells.Merge()
        Table3.Rows(7).Range.Text = "Zone3"
        Table3.Rows(8).Cells.Split()
        Table3.Cell(8, 1).Range.Text = "Z3"
        Table3.Cell(8, 2).Range.Text = tempP
        Table3.Cell(8, 3).Range.Text = "Ohm"
        Table3.Cell(9, 1).Range.Text = "tZ3"
        Table3.Cell(9, 2).Range.Text = tempQ
        Table3.Cell(9, 3).Range.Text = "sec"
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
        Table4 = oDoc.Content.Tables.Add(oPara4.Range, 9, 3)
        Table4.Rows(1).Cells.Merge()
        Table4.Rows(1).Range.Text = "Setting Resistive Reach"
        Table4.Rows(1).Range.Font.Bold = True
        Table4.Rows(1).Range.Font.Size = 16
        Table4.Rows(2).Cells.Split()
        Table4.Cell(2, 1).Range.Text = "Rphmin"
        Table4.Cell(2, 2).Range.Text = tempR
        Table4.Cell(2, 3).Range.Text = "Ohm"
        Table4.Cell(3, 1).Range.Text = "Rgmin"
        Table4.Cell(3, 2).Range.Text = tempS
        Table4.Cell(3, 3).Range.Text = "Ohm"
        Table4.Cell(4, 1).Range.Text = "R3ph"
        Table4.Cell(4, 2).Range.Text = tempT
        Table4.Cell(4, 3).Range.Text = "Ohm"
        Table4.Cell(5, 1).Range.Text = "R3g"
        Table4.Cell(5, 2).Range.Text = tempU
        Table4.Cell(5, 3).Range.Text = "Ohm"
        Table4.Cell(6, 1).Range.Text = "R2ph"
        Table4.Cell(6, 2).Range.Text = tempV
        Table4.Cell(6, 3).Range.Text = "Ohm"
        Table4.Cell(7, 1).Range.Text = "R2g"
        Table4.Cell(7, 2).Range.Text = tempW
        Table4.Cell(7, 3).Range.Text = "Ohm"
        Table4.Cell(8, 1).Range.Text = "R1ph"
        Table4.Cell(8, 2).Range.Text = tempX
        Table4.Cell(8, 3).Range.Text = "Ohm"
        Table4.Cell(9, 1).Range.Text = "R1g"
        Table4.Cell(9, 2).Range.Text = tempY
        Table4.Cell(9, 3).Range.Text = "Ohm"
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
        Table5 = oDoc.Content.Tables.Add(oPara5.Range, 2, 3)
        Table5.Rows(1).Cells.Merge()
        Table5.Rows(1).Range.Text = "BLINDER"
        Table5.Rows(1).Range.Font.Bold = True
        Table5.Rows(1).Range.Font.Size = 16
        Table5.Rows(2).Cells.Split()
        Table5.Cell(2, 1).Range.Text = "ZB"
        Table5.Cell(2, 2).Range.Text = tempZ
        Table5.Cell(2, 3).Range.Text = "Ohm"
        Table5.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        Table5.Borders.OutsideColor = Word.WdColor.wdColorBlack
        Table5.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
        Table5.Borders.InsideColor = Word.WdColor.wdColorBlack
        Table5.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
        oPara5.Format.SpaceAfter = 21
        oPara5.Range.InsertParagraphAfter()
        oPara5.Range.InsertBreak(Word.WdBreakType.wdPageBreak)

        Me.Close()
        MessageBox.Show("Export to Word Complete", "Distance Relay Calculation",
                            MessageBoxButtons.OK)
    End Sub
End Class