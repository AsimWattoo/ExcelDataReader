Imports System.IO
Imports exc = Microsoft.Office.Interop.Excel

Public Class DataDisplayForm

    Private currentFile As String = ""
    Private OccurrenceDate As DateTime
    Private OccurrenceTime As DateTime

    Public Sub SetData(occurenceDate As DateTime, occurenceTime As DateTime, title As String, description As String, totalTime As String, file As String)
        currentFile = file
        Me.OccurrenceDate = occurenceDate
        Me.OccurrenceTime = occurenceTime
        lblTitle.Text = title
        lblDes.Text = description
        lblDate.Text = occurenceDate.ToString("dd/MM/yyyy")
        lblStart.Text = $"{ConvertNumber(occurenceTime.Hour)}:{ConvertNumber(occurenceTime.Minute)}"
        lblTotal.Text = totalTime
        Dim duration As DateTime = DateTime.Parse(totalTime)
        Dim endTime As DateTime = occurenceTime.AddMinutes(duration.Minute)
        If duration.Minute = 0 Then
            endTime = endTime.AddHours(duration.Hour)
        End If
        lblEnd.Text = $"{ConvertNumber(endTime.Hour)}:{ConvertNumber(endTime.Minute)}"
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

    Private Sub btnOpen_Click(sender As Object, e As EventArgs) Handles btnOpen.Click
        Process.Start("C:\Users\asimw\Downloads\Media Player\Media Player\Source Code\MFH Player_Code\bin\Debug\MFH Player.exe", $"{OccurrenceDate.ToString("dd-MM-yyyy")} {OccurrenceTime.ToString("HH:mm:ss")}")
    End Sub
End Class