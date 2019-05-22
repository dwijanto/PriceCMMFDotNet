Imports System.Windows.Forms

Public Class DialogHelper
    Dim bs As New BindingSource
    Public Filter As String
    Public Sub New(bs As BindingSource)
        InitializeComponent()
        Me.bs = bs
        DataGridView1.AutoGenerateColumns = False
        DataGridView1.DataSource = bs
    End Sub


    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub



    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Dim myfilter As String = String.Empty
        myfilter = String.Format(Filter, TextBox1.Text)
        bs.Filter = myfilter
    End Sub
End Class
