Public Class FormMenu

    Private Sub CMMFToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CMMFToolStripMenuItem.Click
        Dim myform As New FormCMMF
        myform.Show()
    End Sub


    Private Sub FormMenu_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.Text = String.Format("PriceCMMF - Server: {0} Database: {1}", DataAccess.GetHostName, DataAccess.GetDataBaseName)
    End Sub

    Private Sub ReportingBySupplierToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReportingBySupplierToolStripMenuItem.Click
        Dim myform As New FormRBS
        myform.Show()
    End Sub

    Private Sub ReportingBySBUToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReportingBySBUToolStripMenuItem.Click
        Dim myform As New FormRSBUFamily
        myform.Show()
    End Sub

    Private Sub ReportingBySSMSPMToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReportingBySSMSPMToolStripMenuItem.Click
        Dim myform As New FormRSSMSupplier
        myform.Show()
    End Sub
End Class
