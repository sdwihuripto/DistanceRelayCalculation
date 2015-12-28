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
        tempC = c
        tempD = d
        tempE = e
        tempF = f
        tempG = g
        tempH = h
        tempI = i
        tempJ = j
        tempK = k
        tempL = l
        tempM = m
        tempN = n
        tempO = o
        tempP = p
        tempQ = q
        tempR = r
        tempS = s
        tempT = t
        tempU = u
        tempV = v
        tempW = w
        tempX = x
        tempY = y
        tempZ = z
        tempA1 = a1
        tempB1 = b1
        tempC1 = c1
        tempD1 = d1
        tempE1 = e1
        tempF1 = f1
        tempG1 = g1
        tempH1 = h1
        tempI1 = i1
        tempJ1 = j1
        tempK1 = k1
        tempL1 = l1
        tempM1 = m1
        tempN1 = n1
        tempO1 = o1
        tempP1 = p1
        tempQ1 = q1
        tempR1 = r1
        tempS1 = s1
        tempT1 = t1
        tempU1 = u1
        tempV1 = v1
        tempW1 = w1
        tempX1 = x1
        tempY1 = y1
        tempZ1 = z1

        OutSum1.Text = a
        OutSum2.Text = b
        OutSum3.Text = c
        OutSum4.Text = d
        OutSum5.Text = e
        OutSum6.Text = f
        OutSum7.Text = g
        OutSum8.Text = h
        OutSum9.Text = i
        Out1A.Text = j
        Out1B.Text = k
        Out1C.Text = l
        Out1D.Text = m
        Out2A.Text = n
        Out2B.Text = o
        Out2C.Text = p
        Out2D.Text = q
        Out2E.Text = r
        Out2F.Text = s
        Out2G.Text = t
        Out2H.Text = u
        Out2I.Text = v
        Out3A.Text = w
        Out3B.Text = x
        Out3C.Text = y
        Out3D.Text = z
        Out4A.Text = a1
        Out4B.Text = b1
        Out4C.Text = c1
        Out4D.Text = d1
        Out4E.Text = e1
        Out4F.Text = f1
        Out4G.Text = g1
        Out4H.Text = h1
        Out4I.Text = i1
        Out5A.Text = j1
        Out5B.Text = k1
        Out5C.Text = l1
        Out5D.Text = m1
        Out6A.Text = n1
        Out6B.Text = o1
        Out6C.Text = p1
        Out6D.Text = q1
        Out6E.Text = r1
        Out6F.Text = s1
        Out6G.Text = t1
        Out6H.Text = u1
        Out6I.Text = v1
        Out7A.Text = w1
        Out7B.Text = x1
        Out7C.Text = y1
        Out7D.Text = z1
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