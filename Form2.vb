Public Class Form2
    Public Property pattern As String
    Public currentlyOpennedFile As String = Nothing

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles PatternComboBox.SelectedIndexChanged
        pattern = PatternComboBox.Text
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Public Sub New(currentPattern As String)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        pattern = currentPattern
        Dim selectedIndex = -1
        If currentPattern = "Pattern A" Then
            selectedIndex = 0
        ElseIf currentPattern = "Pattern B" Then
            selectedIndex = 1
        ElseIf currentPattern = "Pattern C" Then
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
End Class