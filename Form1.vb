Imports System.IO
Imports System.Net.Security
Imports System.Security.Cryptography
Imports System.Text
Imports System.Windows.Forms.LinkLabel

Public Class Form1
    Dim dates As List(Of String) = New List(Of String)()
    Public lines As List(Of String) = New List(Of String)()
    Public parsedData As Dictionary(Of String, List(Of List(Of String))) = New Dictionary(Of String, List(Of List(Of String)))()
    Public parsedLines As List(Of List(Of String)) = New List(Of List(Of String))
    Dim patternObject As Form2 = New Form2
    Dim ColumnData As detailsData1 = New detailsData1
    Public patternLines As List(Of List(Of String)) = New List(Of List(Of String))()
    Private Sub open_Click(sender As Object, e As EventArgs) Handles open.Click
        openFile()

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        AddHandler DataGridView1.CellPainting, AddressOf CellPaint
        AddHandler DataGridView1.Sorted, AddressOf ColumnSorted
    End Sub

    Private Sub ColumnSorted(sender As Object, e As EventArgs)
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            DataGridView1.Rows.Item(i).Height = 100
        Next
    End Sub


    Private Sub CellPaint(sender As Object, e As DataGridViewCellPaintingEventArgs)
        If e.ColumnIndex < 1 Then
            Return
        End If

        If e.RowIndex.Equals(-1) Then
            Return
        End If

        Dim pen As Pen = New Pen(Brushes.Black, 0.5)
        Dim halfWidth As Integer = e.CellBounds.Width
        Dim halfHeight As Integer = e.CellBounds.Height / 3

        Dim item As Object = e.Value
        If TypeOf item Is String Then
            Dim str As String = CType(item, String)
            Dim items As String() = str.Split("//")

            If Not (items.Length.Equals(3)) Then
                Return
            End If
            e.Graphics.FillRectangle(Brushes.White, New Rectangle(e.CellBounds.Left, e.CellBounds.Top, e.CellBounds.Width, e.CellBounds.Height))
            e.Graphics.DrawRectangle(pen, New Rectangle(e.CellBounds.Left, e.CellBounds.Top, e.CellBounds.Width, halfHeight))
            e.Graphics.DrawRectangle(pen, New Rectangle(e.CellBounds.Left, e.CellBounds.Top + halfHeight, e.CellBounds.Width, halfHeight))
            e.Graphics.DrawRectangle(pen, New Rectangle(e.CellBounds.Left, e.CellBounds.Top + 2 * halfHeight, e.CellBounds.Width, halfHeight))
            e.Graphics.DrawString(items(0), New Font(FontFamily.GenericSerif, 12), Brushes.Black, New PointF(e.CellBounds.Left, e.CellBounds.Top + 10))
            e.Graphics.DrawString(items(1), New Font(FontFamily.GenericSerif, 12), Brushes.Black, New PointF(e.CellBounds.Left, e.CellBounds.Top + halfHeight + 10))
            e.Graphics.DrawString(items(2), New Font(FontFamily.GenericSerif, 12), Brushes.Black, New PointF(e.CellBounds.Left, e.CellBounds.Top + 2 * halfHeight + 10))
            e.Handled = True

        End If
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        If DateTimePicker1.Created Then
            Dim table As DataTable = New DataTable()

            Dim headers As List(Of String) = New List(Of String)()

            For Each d As String In dates

                If d.Equals("") Then
                    headers.Add("")
                    Continue For
                End If

                Dim parsedDate As Date = Date.Parse(d)
                If parsedDate >= DateTimePicker1.Value.Date Then
                    headers.Add(d)
                End If
            Next

            For Each column As String In headers
                table.Columns.Add(column)
            Next

            For Each line As List(Of String) In parsedLines
                Dim items As List(Of Object) = New List(Of Object)()
                Dim currentTime As String = line(3)

                If currentTime.Equals("Start Time") Then
                    Continue For
                End If

                items.Add(currentTime)
                For i As Integer = 1 To headers.Count - 1
                    Dim item As List(Of String) = parsedData(headers(i)).Where(Function(x) As Boolean
                                                                                   Return x(3).Equals(currentTime)
                                                                               End Function).FirstOrDefault()
                    If item Is Nothing Then
                        items.Add("")
                    Else
                        items.Add(item(6) + "//" + item(3) + "//" + item(4))
                    End If

                Next
                table.Rows.Add(items.ToArray())
            Next

            DataGridView1.Columns.Clear()
            DataGridView1.DataSource = table

            For i As Integer = 0 To parsedLines.Count - 1
                DataGridView1.Rows.Item(i).Height = 100
            Next

        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Form2.Show()
        If Form2.Button1.DialogResult.OK Then
            openFile()
        End If
    End Sub

    Public Sub openFile()

        Using ofd As OpenFileDialog = New OpenFileDialog() With {.Filter = "CSV files|*.csv"}

            If patternObject.pattern = Nothing Then
                    ColumnData.columnOne = 2
                    ColumnData.columnTwo = 3
                    ColumnData.columnThree = 4
                    ColumnData.columnFour = 6
                    ColumnData.columnFive = 7
                ElseIf patternObject.pattern.Equals("Pattern A") Then
                    ColumnData.columnOne = 2
                    ColumnData.columnTwo = 3
                    ColumnData.columnThree = 4
                    ColumnData.columnFour = 6
                    ColumnData.columnFive = 7
                ElseIf patternObject.pattern.Equals("Pattern B") Then
                    ColumnData.columnOne = 3
                    ColumnData.columnTwo = 4
                    ColumnData.columnThree = 5
                    ColumnData.columnFour = 9
                    ColumnData.columnFive = 10
                ElseIf patternObject.pattern.Equals("Pattern C") Then
                    ColumnData.columnOne = 1
                    ColumnData.columnTwo = 2
                    ColumnData.columnThree = 3
                    ColumnData.columnFour = 8
                End If


            If ofd.ShowDialog() = DialogResult.OK Then
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)
                parsedData.Clear()
                parsedLines.Clear()
                lines.Clear()
                dates.Clear()
                Dim table As DataTable = New DataTable()
                Dim items As List(Of Object) = New List(Of Object)()
                dates = New List(Of String)()
                lines = File.ReadAllLines(ofd.FileName, Encoding.GetEncoding("Shift-JIS")).ToList()

                dates.Add("")
                'For i As Integer = 0 To lines.Count - 1
                '    If Not i = ColumnData.columnOne Or Not i = ColumnData.columnTwo Or Not i = ColumnData.columnThree Or Not i = ColumnData.columnFour Or Not i = ColumnData.columnFive Then
                '        lines.RemoveAt(i)
                '    End If
                'Next
                Dim data As String()

                For i As Integer = 0 To lines.Count - 1
                    data = lines(i).Split(",")
                    Dim newData As List(Of String) = New List(Of String)()
                    Dim currentDate As String = data(2)
                    If currentDate.Equals("Program Date") Then
                        Continue For
                    End If
                    For val As Integer = 0 To data.Count - 1
                        If val = ColumnData.columnOne Or val = ColumnData.columnTwo Or val = ColumnData.columnThree Or val = ColumnData.columnFour Or val = ColumnData.columnFive Then
                            If data(val).Contains(data(3)) Then
                                newData.Insert(1, data(val))
                            Else
                                newData.Add(data(val))
                            End If
                        End If
                    Next
                    parsedLines.Add(newData)
                    If Not dates.Contains(currentDate) Then
                        dates.Add(currentDate)
                    End If
                Next


                'Adding dates to the data table
                For Each d As String In dates
                    table.Columns.Add(d)
                Next
                For i As Integer = 0 To dates.Count - 1
                    parsedData.Add(dates(i), New List(Of List(Of String)))
                Next

                For Each line As List(Of String) In parsedLines
                    Dim currentDate As String = line(0)
                    parsedData(currentDate).Add(line)
                Next


                'Adding rows
                For Each line As List(Of String) In parsedLines
                    items = New List(Of Object)()
                    Dim currentTime As String = line(1)

                    If currentTime.Equals("Start Time") Then
                        Continue For
                    End If

                    items.Add(currentTime)
                    For i As Integer = 2 To dates.Count - 1
                        Dim item As List(Of String) = parsedData(dates(i)).Where(Function(x) As Boolean
                                                                                     Return x(1).Equals(currentTime)
                                                                                 End Function).FirstOrDefault()
                        If item Is Nothing Then
                            Continue For
                        End If

                        items.Add(item(3) + "//" + item(1) + "//" + item(2))

                    Next
                    table.Rows.Add(items.ToArray())
                Next



                DataGridView1.DataSource = table

                For i As Integer = 0 To parsedLines.Count - 1
                    DataGridView1.Rows.Item(i).Height = 100
                Next

            End If
        End Using
    End Sub

End Class