Public Class FormMenu

    Private Sub CMMFToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CMMFToolStripMenuItem.Click
        Dim myform As New FormCMMF
        myform.Show()
    End Sub


    Private Sub FormMenu_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.Text = String.Format("PriceCMMF - Server: {0} Database: {1}", DataAccess.GetHostName, DataAccess.GetDataBaseName)
    End Sub
End Class
