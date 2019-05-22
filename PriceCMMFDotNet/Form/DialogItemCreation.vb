Imports System.Windows.Forms
Imports System.Threading

Public Class DialogItemCreation
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    'Dim myController As New CMMFController
    Dim myController As CMMFController
    Dim VendorBS As New BindingSource
    Private FileName As String = String.Empty
    Dim DRV As DataRowView

    Public Property NameOfRequester As String
    Public Property DateOfRequest As Date?
    Public Property TypeOfRequest As String
    Public Property VendorCode As Long
    Public Property Supplier As String
    Public Property NewProject As Boolean
    Public Property PlatformProject As Boolean
    Public Property DelegatedItem As Boolean

    Public Sub New(ByVal myController As CMMFController)

        InitializeComponent()
        Me.myController = myController
    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If Not myThread.IsAlive Then
            If Me.validate Then
                DRV = myController.BS.Current
                Dim savedialog1 As New SaveFileDialog
                savedialog1.FileName = String.Format("{0}-{1}-{2:yyyyMMdd}", ComboBox1.Text, DRV.Row.Item("cmmf"), Date.Today)
                savedialog1.DefaultExt = "xlsx"
                'Filename : Vendor Name + CMMF + Today Date short
                If savedialog1.ShowDialog = Windows.Forms.DialogResult.OK Then

                    NameOfRequester = TextBox1.Text
                    DateOfRequest = DateTimePicker1.Value.Date
                    TypeOfRequest = ComboBox1.Text
                    Supplier = TextBox2.Text
                    NewProject = RadioButton1.Checked
                    PlatformProject = RadioButton4.Checked
                    DelegatedItem = RadioButton5.Checked
                    FileName = savedialog1.FileName


                    myThread = New Thread(AddressOf DoGenerateExcel)
                    myThread.Start()
                End If
            Else

            End If
            
        Else

        End If
        'Me.DialogResult = System.Windows.Forms.DialogResult.OK
        'Me.Close()
    End Sub
    Private Overloads Function validate() As Boolean
        Dim myret As Boolean = True
        ErrorProvider1.SetError(TextBox1, "")
        If TextBox1.Text = "" Then
            ErrorProvider1.SetError(TextBox1, "Name of Requester cannot be blank.")
            myret = False
        End If
        ErrorProvider1.SetError(ComboBox1, "")
        If IsNothing(ComboBox1.SelectedItem) Then
            ErrorProvider1.SetError(ComboBox1, "Please select from the list.")
            myret = False
        End If
        ErrorProvider1.SetError(TextBox2, "")
        If TextBox2.Text = "" Then
            ErrorProvider1.SetError(TextBox2, "Supplier cannot be blank.")
            myret = False
        End If

        Return myret
    End Function
    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub


    Private Sub loadData()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""            
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Sub DoWork()
        VendorBS.DataSource = myController.Model.getVendorTable
        ProgressReport(4, "Binding Data")
    End Sub

    Public Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        Try
            If Me.InvokeRequired Then
                Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
                Me.Invoke(d, New Object() {id, message})
            Else
                Select Case id
                    Case 1
                        ToolStripStatusLabel1.Text = message
                    Case 4
                        ComboBox1.SelectedIndex = 0
                    Case 5
                        ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                    Case 6
                        ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                End Select
            End If
        Catch ex As Exception
            Debug.Print(ex.Message)
        End Try

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim myhelper As DialogHelper = New DialogHelper(VendorBS)

        myhelper.DataGridView1.Columns(0).HeaderText = "Supplier"
        myhelper.DataGridView1.Columns(0).DataPropertyName = "vendorcodename"
        myhelper.DataGridView1.Columns(0).Width = "300"
        myhelper.Filter = "[vendorcodename] like '%{0}%'"
        For i = 1 To myhelper.DataGridView1.Columns.Count - 1
            myhelper.DataGridView1.Columns(i).Visible = False
        Next
        If myhelper.ShowDialog = Windows.Forms.DialogResult.OK Then
            Dim drv As DataRowView = VendorBS.Current
            Debug.Print(drv.Row.Item("vendorcode"))
            TextBox2.Text = drv.Row.Item("vendorname")
            VendorCode = drv.Row.Item("vendorcode")
        End If
    End Sub

    Private Sub DialogItemCreation_Load(sender As Object, e As EventArgs) Handles Me.Load
        loadData()
    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged, RadioButton4.CheckedChanged
        If RadioButton4.Checked Then
            RadioButton5.Checked = False
            RadioButton6.Checked = True
        End If
        RadioButton5.Enabled = Not RadioButton4.Checked
        RadioButton6.Enabled = Not RadioButton4.Checked
    End Sub

    Private Sub DoGenerateExcel()
        Dim myItemCreation = New ItemCreation(Me, FileName, DRV)
        myItemCreation.GenerateExcel()
        If CheckBox1.Checked Then
            Process.Start(FileName)
        End If

    End Sub

End Class
