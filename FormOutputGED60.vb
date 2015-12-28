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
        Dim oPara1 As Word.Paragraph, oPara2 As Word.Paragraph, oPara3 As Word.Paragraph
        Dim oPara4 As Word.Paragraph, oPara5 As Word.Paragraph, oPara6 As Word.Paragraph
        Dim oPara7 As Word.Paragraph, oPara8 As Word.Paragraph, oPara9 As Word.Paragraph
        Dim oPara10 As Word.Paragraph, oPara11 As Word.Paragraph, oPara12 As Word.Paragraph
        Dim oPara13 As Word.Paragraph, oPara14 As Word.Paragraph, oPara15 As Word.Paragraph
        Dim oPara16 As Word.Paragraph, oPara17 As Word.Paragraph, oPara18 As Word.Paragraph
        Dim oPara19 As Word.Paragraph, oPara20 As Word.Paragraph, oPara21 As Word.Paragraph
        Dim oPara22 As Word.Paragraph, oPara23 As Word.Paragraph, oPara24 As Word.Paragraph
        Dim oPara25 As Word.Paragraph, oPara26 As Word.Paragraph, oPara27 As Word.Paragraph
        Dim oPara28 As Word.Paragraph, oPara29 As Word.Paragraph, oPara30 As Word.Paragraph
        Dim oPara31 As Word.Paragraph, oPara32 As Word.Paragraph, oPara33 As Word.Paragraph
        Dim oPara34 As Word.Paragraph, oPara35 As Word.Paragraph, oPara36 As Word.Paragraph
        Dim oPara37 As Word.Paragraph, oPara38 As Word.Paragraph, oPara39 As Word.Paragraph

        'Start Word and open the document template.
        oWord = CreateObject("Word.Application")
        oWord.Visible = True
        oDoc = oWord.Documents.Add

        oPara1 = oDoc.Content.Paragraphs.Add
        oPara1.Range.Text = "Location    :    " & tempA
        oPara1.Range.Font.Bold = False
        oPara1.Format.SpaceAfter = 8
        oPara1.Range.InsertParagraphAfter()

        oPara2 = oDoc.Content.Paragraphs.Add
        oPara2.Range.Text = "Line Bay To :    " & tempB
        oPara2.Range.Font.Bold = False
        oPara2.Format.SpaceAfter = 8
        oPara2.Range.InsertParagraphAfter()

        oPara3 = oDoc.Content.Paragraphs.Add
        oPara3.Range.Text = "CT Ratio    :    " & tempC & "    A"
        oPara3.Range.Font.Bold = False
        oPara3.Format.SpaceAfter = 8
        oPara3.Range.InsertParagraphAfter()

        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
End Class