Public Class Form2
    Public Property pattern As ApplicationMode
    Public currentlyOpennedFile As String = Nothing

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles PatternComboBox.SelectedIndexChanged
        Dim patternText As String = PatternComboBox.Text
        If patternText = GetPatternName(ApplicationMode.PatternA) Then
            pattern = ApplicationMode.PatternA
        ElseIf patternText = GetPatternName(ApplicationMode.PatternB) Then
            pattern = ApplicationMode.PatternB
        Else
            pattern = ApplicationMode.PatternC
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Public Sub New(currentPattern As ApplicationMode)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        pattern = currentPattern
        Dim selectedIndex = -1
        If currentPattern = ApplicationMode.PatternA Then
            selectedIndex = 0
        ElseIf currentPattern = ApplicationMode.PatternB Then
            selectedIndex = 1
        ElseIf currentPattern = ApplicationMode.PatternC Then
            selectedIndex = 2
        End If
        PatternComboBox.SelectedIndex = selectedIndex
    End Sub

    Private Sub Openbtn_Click(sender As Object, e As EventArgs) Handles Openbtn.Click
        Using ofd As OpenFileDialog = New OpenFileDialog() With {.Filter = "CSV files|*.csv"}
            If ofd.ShowDialog() = DialogResult.OK Then
                currentlyOpennedFile = ofd.FileName
            End If
        End Using
    End Sub

    Private Function GetPatternName(mode As ApplicationMode) As String
        If mode = ApplicationMode.PatternA Then
            Return "NHK"
        ElseIf mode = ApplicationMode.PatternB Then
            Return "Sumitomo"
        Else
            Return "Miharu"
        End If
    End Function
End Class