Imports System.Drawing
Imports System.Threading
Imports FarPoint.Win.Spread


Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        CombinePDFFiles()
    End Sub

    Public Sub CombinePDFFiles()
        Dim Spread As FarPoint.Win.Spread.FpSpread = New FarPoint.Win.Spread.FpSpread()
        Spread.BringToFront()

        Spread.Size = New System.Drawing.Size(600, 300)
        Spread.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Dim coll As FarPoint.Win.Spread.SheetViewCollection
        coll = Spread.Sheets
        coll.Add(FpSpread1.Sheets(0))
        coll.Add(FpSpread2.Sheets(0))

        Dim printset As FarPoint.Win.Spread.PrintInfo = New FarPoint.Win.Spread.PrintInfo()
        printset.PrintToPdf = True
        printset.PdfWriteMode = PdfWriteMode.Append

        printset.PdfFileName = "D:\TEST\testPDF_" & Format(Now, "yyyyMMddHHmmss") & ".pdf"

        Me.Controls.Add(Spread)
        Spread.Sheets(0).PrintInfo = printset
        Spread.Sheets(1).PrintInfo = printset
        Spread.PrintSheet(-1)

    End Sub



    Public Sub CombinePDFFiles1()
        Dim printset As FarPoint.Win.Spread.PrintInfo = New FarPoint.Win.Spread.PrintInfo()
        printset.PrintToPdf = True
        printset.PdfWriteMode = PdfWriteMode.Append
        printset.PdfFileName = "D:\TEST\testPDF_" & Format(Now, "yyyyMMddHHmmss") & ".pdf"

        FpSpread1.Sheets(0).PrintInfo = printset
        FpSpread1.Sheets(1).PrintInfo = printset
        FpSpread1.Sheets(2).PrintInfo = printset

        FpSpread1.PrintSheet(-1)
    End Sub

End Class
