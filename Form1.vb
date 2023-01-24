Imports System.IO
Imports System.Net.Security
Imports System.Text

Public Class Form1
    Dim dates As List(Of String) = New List(Of String)()
    Private lines As List(Of String) = New List(Of String)()
    Private parsedData As Dictionary(Of String, Dictionary(Of Integer, List(Of List(Of String)))) = New Dictionary(Of String, Dictionary(Of Integer, List(Of List(Of String))))
    Private parsedLines As List(Of List(Of String)) = New List(Of List(Of String))
    Private patternObject As Form2 = Nothing
    Private ColumnData As detailsData1 = New detailsData1
    Private parsedLineNew As Dictionary(Of String, List(Of String)) = New Dictionary(Of String, List(Of String))
    Private currentPattern As String = "Pattern A"
    Public currentlyOpennedFile As String = Nothing

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

        DataGridView1.Rows.Item(e.RowIndex).Height = 350

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

        Dim item As Object = e.Value
        If TypeOf item Is String Then
            Dim str As String = CType(item, String)
            Dim records As String() = str.Split("\n")

            If records.Length <= 1 Then
                Return
            End If

            Dim heightForEachRecord As Integer = CType(Math.Ceiling((e.CellBounds.Height * 1.0) / (records.Length - 1)), Integer)
            Dim startPoint As Integer = 0
            Dim halfHeight As Integer = heightForEachRecord / 2
            e.Graphics.FillRectangle(Brushes.White, New Rectangle(e.CellBounds.Left, e.CellBounds.Top, e.CellBounds.Width, e.CellBounds.Height))
            For Each record As String In records

                If record = "" Then
                    Continue For
                End If

                Dim items As String() = record.Split("//")
                If Not (items.Length.Equals(3)) Then
                    Return
                End If
                e.Graphics.DrawRectangle(pen, New Rectangle(e.CellBounds.Left, e.CellBounds.Top + startPoint, e.CellBounds.Width, halfHeight))
                e.Graphics.DrawRectangle(pen, New Rectangle(e.CellBounds.Left, e.CellBounds.Top + startPoint + halfHeight, e.CellBounds.Width, halfHeight))
                e.Graphics.DrawString(items(0), New Font(FontFamily.GenericSerif, 12), Brushes.Black, New PointF(e.CellBounds.Left, e.CellBounds.Top + 5 + startPoint))
                Dim startTime As DateTime = DateTime.Parse(items(1))
                Dim duration As DateTime = DateTime.Parse(items(2))
                Dim endTime As DateTime = startTime.AddMinutes(duration.Minute)
                If duration.Minute = 0 Then
                    endTime = endTime.AddHours(duration.Hour)
                End If
                e.Graphics.DrawString($"{items(1)} ~ {endTime.Hour}:{ConvertNumber(endTime.Minute)}:{ConvertNumber(endTime.Second)} ({items(2)})", New Font(FontFamily.GenericSerif, 12), Brushes.Black, New PointF(e.CellBounds.Left, e.CellBounds.Top + halfHeight + 5 + startPoint))
                e.Handled = True
                startPoint += heightForEachRecord
            Next

        End If
    End Sub

    Private Function ConvertNumber(number As Integer) As String
        If number = 0 Then
            Return "00"
        ElseIf number < 10 Then
            Return $"0{number}"
        Else
            Return number.ToString()
        End If
    End Function

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        If DateTimePicker1.Created Then
            loadData(currentlyOpennedFile, DateTimePicker1.Value)
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        patternObject = New Form2(currentPattern)
        If patternObject.ShowDialog() = DialogResult.OK Then
            currentPattern = patternObject.pattern
            currentlyOpennedFile = patternObject.currentlyOpennedFile
        End If
    End Sub

    Public Sub loadData(fileName As String, startDate As Date)

        If String.IsNullOrEmpty(fileName) Then
            Return
        End If

        startDate = New Date(startDate.Year, startDate.Month, startDate.Day)

        If currentPattern = Nothing Then
            ColumnData.DateColumn = 2
            ColumnData.StartTimeColumn = 3
            ColumnData.TotalTimeColumn = 4
            ColumnData.TitleColumn = 6
            ColumnData.DetailsColumn = 7
        ElseIf currentPattern.Equals("Pattern A") Then
            ColumnData.DateColumn = 2
            ColumnData.StartTimeColumn = 3
            ColumnData.TotalTimeColumn = 4
            ColumnData.TitleColumn = 6
            ColumnData.DetailsColumn = 7
        ElseIf currentPattern.Equals("Pattern B") Then
            ColumnData.DateColumn = 3
            ColumnData.StartTimeColumn = 4
            ColumnData.TotalTimeColumn = 5
            ColumnData.TitleColumn = 9
            ColumnData.DetailsColumn = 10
        ElseIf currentPattern.Equals("Pattern C") Then
            ColumnData.DateColumn = 1
            ColumnData.StartTimeColumn = 2
            ColumnData.TotalTimeColumn = 3
            ColumnData.TitleColumn = 8
            ColumnData.DetailsColumn = 9
        End If


        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)
        parsedData.Clear()
        parsedLines.Clear()
        parsedLineNew.Clear()
        lines.Clear()
        dates.Clear()
        Dim initialDates As List(Of Date) = New List(Of Date)
        Dim table As DataTable = New DataTable()
        Dim items As List(Of Object) = New List(Of Object)()
        dates = New List(Of String)()
        lines = File.ReadAllLines(fileName, Encoding.GetEncoding("Shift-JIS")).ToList()

        dates.Add("")
        Dim data As String()

        Dim checkDate As Date = DateTime.Now.AddDays(-90)
        For i As Integer = 0 To lines.Count - 1
            data = lines(i).Split(",")
            Dim newData As List(Of String) = New List(Of String)()
            Dim currentDate As String = data(ColumnData.DateColumn)
            If currentDate.Equals("Program Date") Then
                Continue For
            End If
            If currentDate < checkDate Then
                Continue For
            End If
            Dim d As Date = Date.Parse(currentDate)
            currentDate = d.ToString("dd/MM/yyyy")

            For val As Integer = 0 To data.Count - 1
                If val = ColumnData.DateColumn Or val = ColumnData.StartTimeColumn Or val = ColumnData.TotalTimeColumn Or val = ColumnData.TitleColumn Or val = ColumnData.DetailsColumn Then
                    If val = ColumnData.StartTimeColumn Then
                        newData.Insert(1, data(val))
                    ElseIf val = ColumnData.DateColumn Then
                        newData.Insert(0, currentDate)
                    Else
                        newData.Add(data(val))
                    End If
                End If
            Next

            parsedLines.Add(newData)
            If Not initialDates.Contains(d) Then
                initialDates.Add(d)
            End If
        Next

        'Adding dates to the data table
        For Each d As Date In initialDates
            If d >= startDate Then
                dates.Add(d.ToString("dd/MM/yyyy"))
            End If
        Next

        For Each d As String In dates
            table.Columns.Add(d)
        Next

        'Dim previousHour = 0
        'For Each line As List(Of String) In parsedLines
        '    Dim time As DateTime = DateTime.Parse(line(1))
        '    If Not (previousHour = time.Hour) Then
        '        parsedLineNew.Add(line)
        '        previousHour = time.Hour
        '    End If
        'Next

        For i As Integer = 0 To dates.Count - 1
            parsedData.Add(dates(i), New Dictionary(Of Integer, List(Of List(Of String))))
        Next

        Dim allTimeEnteries As List(Of Integer) = New List(Of Integer)()
        Dim previousTime As String = ""
        For Each line As List(Of String) In parsedLines
            Dim currentDate As String = line(0)
            If parsedData.ContainsKey(currentDate) Then
                Dim currentTime = DateTime.Parse(line(1))
                If Not parsedData(currentDate).ContainsKey(currentTime.Hour) Then
                    parsedData(currentDate).Add(currentTime.Hour, New List(Of List(Of String)))
                End If

                If Not allTimeEnteries.Contains(currentTime.Hour) Then
                    allTimeEnteries.Add(currentTime.Hour)
                End If

                If line(1) = previousTime Then
                    Continue For
                End If
                previousTime = line(1)
                'Adding line to the current time
                parsedData(currentDate)(currentTime.Hour).Add(line)
            End If
        Next

        'Adding rows
        For Each hour As Integer In allTimeEnteries
            Dim dataItems As List(Of Object) = New List(Of Object)()

            For Each d As String In dates
                If parsedData(d).ContainsKey(hour) Then
                    Dim str As String = ""
                    For Each line As List(Of String) In parsedData(d)(hour)
                        str += line(3) + "//" + line(1) + "//" + line(2) + "\n"
                    Next
                    str = str.Trim("\n")
                    dataItems.Add(str)
                Else
                    If d = "" Then
                        dataItems.Add($"{hour}:00:00")
                    Else
                        dataItems.Add("")
                    End If
                End If
            Next
            table.Rows.Add(dataItems.ToArray())
        Next
        If Not table.Rows.Count = 0 Then
            DataGridView1.DataSource = table
        End If

        For i As Integer = 0 To DataGridView1.RowCount - 1
            DataGridView1.Rows.Item(i).Height = 350
        Next

        For i As Integer = 1 To DataGridView1.Columns.Count - 1
            DataGridView1.Columns(i).Width = 250
        Next

    End Sub

End Class