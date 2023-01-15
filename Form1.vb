Imports System.IO
Imports System.Net.Security
Imports System.Security.Cryptography
Imports System.Text
Imports System.Windows.Forms.LinkLabel

Public Class Form1
    Dim dates As List(Of String) = New List(Of String)()
    Private lines As List(Of String) = New List(Of String)()
    Private parsedData As Dictionary(Of String, List(Of List(Of String))) = New Dictionary(Of String, List(Of List(Of String)))()
    Private parsedLines As List(Of List(Of String)) = New List(Of List(Of String))
    Private patternObject As Form2 = New Form2
    Private ColumnData As detailsData1 = New detailsData1
    Private patternLines As List(Of List(Of String)) = New List(Of List(Of String))()
    Private currentlyOpennedFile As String = Nothing


    Private Sub open_Click(sender As Object, e As EventArgs) Handles open.Click
        Using ofd As OpenFileDialog = New OpenFileDialog() With {.Filter = "CSV files|*.csv"}
            If ofd.ShowDialog() = DialogResult.OK Then
                currentlyOpennedFile = ofd.FileName
                loadData(ofd.FileName, DateTime.Now)
            End If
        End Using

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        AddHandler DataGridView1.CellPainting, AddressOf CellPaint
        AddHandler DataGridView1.Sorted, AddressOf ColumnSorted

        DataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
        DataGridView1.ColumnHeadersHeight = 60
        DataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing

    End Sub

    Private Sub ColumnSorted(sender As Object, e As EventArgs)
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            DataGridView1.Rows.Item(i).Height = 100
        Next
    End Sub


    Private Sub CellPaint(sender As Object, e As DataGridViewCellPaintingEventArgs)

        If e.RowIndex.Equals(-1) Then

            If e.ColumnIndex = 0 Then
                e.Graphics.FillRectangle(New SolidBrush(Color.FromArgb(68, 115, 197)), New Rectangle(e.CellBounds.Left, e.CellBounds.Top, e.CellBounds.Width, e.CellBounds.Height))
                Dim polygonPen As Pen = New Pen(Color.White)
                Dim margin As Single = 15
                Dim points As Point() = {New Point(e.CellBounds.Left + margin, e.CellBounds.Top + margin), New Point(e.CellBounds.Right - margin, e.CellBounds.Bottom - margin)}
                e.Graphics.DrawPolygon(polygonPen, points)
                e.Handled = True
                Return
            Else
                Return
            End If

        End If

            If e.ColumnIndex < 1 Then
            Dim rect As Rectangle = New Rectangle(e.CellBounds.Left, e.CellBounds.Top, e.CellBounds.Width, e.CellBounds.Height)
            e.Graphics.FillRectangle(New SolidBrush(Color.FromArgb(68, 115, 197)), rect)
            e.Graphics.DrawRectangle(New Pen(Color.White), rect)
            Dim mid As Single = e.CellBounds.Height / 2 + e.CellBounds.Top - 10
            Dim leftMargin As Single = 20
            e.Graphics.DrawString(e.Value.ToString(), New Font(FontFamily.GenericSerif, 12), Brushes.White, New PointF(e.CellBounds.Left + leftMargin, mid))
            e.Handled = True
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
            loadData(currentlyOpennedFile, DateTimePicker1.Value)
        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Form2.Show()
        If Form2.DialogResult = DialogResult.OK Then
            loadData(currentlyOpennedFile, DateTime.Now)
        End If
    End Sub

    Public Sub loadData(fileName As String, startDate As Date)

        startDate = New Date(startDate.Year, startDate.Month, startDate.Day)

        If patternObject.pattern = Nothing Then
            ColumnData.DateColumn = 2
            ColumnData.StartTimeColumn = 3
            ColumnData.TotalTimeColumn = 4
            ColumnData.TitleColumn = 6
            ColumnData.DetailsColumn = 7
        ElseIf patternObject.pattern.Equals("Pattern A") Then
            ColumnData.DateColumn = 2
            ColumnData.StartTimeColumn = 3
            ColumnData.TotalTimeColumn = 4
            ColumnData.TitleColumn = 6
            ColumnData.DetailsColumn = 7
        ElseIf patternObject.pattern.Equals("Pattern B") Then
            ColumnData.DateColumn = 3
            ColumnData.StartTimeColumn = 4
            ColumnData.TotalTimeColumn = 5
            ColumnData.TitleColumn = 9
            ColumnData.DetailsColumn = 10
        ElseIf patternObject.pattern.Equals("Pattern C") Then
            ColumnData.DateColumn = 1
            ColumnData.StartTimeColumn = 2
            ColumnData.TotalTimeColumn = 3
            ColumnData.TitleColumn = 8
        End If


        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)
        parsedData.Clear()
        parsedLines.Clear()
        lines.Clear()
        dates.Clear()
        Dim initialDates As List(Of Date) = New List(Of Date)
        Dim table As DataTable = New DataTable()
        Dim items As List(Of Object) = New List(Of Object)()
        dates = New List(Of String)()
        lines = File.ReadAllLines(fileName, Encoding.GetEncoding("Shift-JIS")).ToList()

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
                If val = ColumnData.DateColumn Or val = ColumnData.StartTimeColumn Or val = ColumnData.TotalTimeColumn Or val = ColumnData.TitleColumn Or val = ColumnData.DetailsColumn Then
                    If val = ColumnData.StartTimeColumn Then
                        newData.Insert(1, data(val))
                    ElseIf val = ColumnData.DateColumn Then
                        newData.Insert(0, data(val))
                    Else
                        newData.Add(data(val))
                    End If
                End If
            Next
            parsedLines.Add(newData)
            Dim d As Date = Date.Parse(currentDate)
            If Not initialDates.Contains(d) Then
                initialDates.Add(d)
            End If
        Next

        'Adding dates to the data table
        For Each d As Date In initialDates

            If d >= startDate Then
                dates.Add(d.ToString("yyyy/MM/dd"))
            End If
        Next

        For Each d As String In dates
            table.Columns.Add(d)
        Next

        For i As Integer = 0 To dates.Count - 1
            parsedData.Add(dates(i), New List(Of List(Of String)))
        Next

        For Each line As List(Of String) In parsedLines
            Dim currentDate As String = line(0)
            If parsedData.ContainsKey(currentDate) Then
                parsedData(currentDate).Add(line)
            End If
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

        For i As Integer = 1 To DataGridView1.Columns.Count - 1
            DataGridView1.Columns(i).Width = 250
        Next

    End Sub

End Class