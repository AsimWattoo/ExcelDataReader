Public Class Form2
    Public Property pattern As String
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles PatternComboBox.SelectedIndexChanged
        pattern = PatternComboBox.Text
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

End Class