Imports Excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Office

Public Class Form1
    Dim MATRIX_B(53, 53)
    Dim MATRIX_S(53, 53)
    Dim MATRIX_J(53, 53)
    Dim eApp As excel.Application
    Dim eBook As Excel.Workbook
    Dim eSheet As Excel.Worksheet
    Dim eCell As Excel.Range
    Dim eCellArray As System.Array
    Dim D As String
    Dim ROW, COL, BIN, FIN, SIN


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.ShowDialog()
        D = OpenFileDialog1.FileName
        eApp = New Excel.Application
        eBook = eApp.Workbooks.Open(D)
        eSheet = eBook.Worksheets(1)
        eCell = eSheet.UsedRange
        eCellArray = eCell.Value
        eApp.Application.Quit()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ROW = UBound(eCellArray, 1)
        COL = UBound(eCellArray, 2)
        For k = 1 To COL
            Dim First(ROW) As String
            For i = 1 To ROW
                First(i) = eCellArray(i, k)
            Next i
            Dim FI As Array
            FI = First.Distinct().ToArray
            FIN = UBound(FI) - 2
            For l = 1 To COL - k
                Dim Second(ROW) As String
                For i = 1 To ROW
                    Second(i) = eCellArray(i, k + l)
                Next i
                Dim SI As Array
                SI = Second.Distinct().ToArray
                SIN = UBound(SI) - 2
                Dim ALL(UBound(FI) + UBound(SI) + 2)
                FI.CopyTo(ALL, 0)
                SI.CopyTo(ALL, UBound(FI) + 1)
                Dim Unique As Array
                Unique = ALL.Distinct().ToArray
                BIN = (UBound(ALL) - 6) - (UBound(Unique) - 4)
                MATRIX_B(k + l, k) = BIN
                MATRIX_S(k + l, k) = 2 * BIN / ((2 * BIN) + FIN + SIN)
                MATRIX_J(k + l, k) = BIN / (BIN + FIN + SIN)
                MATRIX_B(k, k + l) = BIN
                MATRIX_S(k, k + l) = 2 * BIN / ((2 * BIN) + FIN + SIN)
                MATRIX_J(k, k + l) = BIN / (BIN + FIN + SIN)
                Array.Clear(Second, 0, UBound(Second) - 1)
                Array.Clear(SI, 0, UBound(SI) - 1)
                Array.Clear(ALL, 0, UBound(ALL) - 1)
            Next l
            Array.Clear(First, 0, UBound(First) - 1)
            Array.Clear(FI, 0, UBound(FI) - 1)
        Next k
        RichTextBox1.SelectedText = MATRIX_B(1, 2)
        RichTextBox1.SelectedText = vbCrLf
        RichTextBox1.SelectedText = MATRIX_S(1, 2)
        RichTextBox1.SelectedText = vbCrLf
        RichTextBox1.SelectedText = MATRIX_J(1, 2)
     

    End Sub
End Class
