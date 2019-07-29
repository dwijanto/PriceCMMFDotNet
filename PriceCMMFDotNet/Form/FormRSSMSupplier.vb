Imports System.Threading
Public Class FormRSSMSupplier
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim DS As DataSet
    Dim PriceCMMFSSMBS As BindingSource
    Dim PriceCMMFSupplierBS As BindingSource
    Dim Title As String = String.Empty
    Dim SqlstrData As String = String.Empty
    Dim SqlstrPhoto As String = String.Empty
    Dim FileNameFullPath As String = String.Empty
    Dim GenerateReport1 As New GenerateReport()
   
    Dim myCheck As Boolean

    Private Sub FormRSBUFamily_Load(sender As Object, e As EventArgs) Handles Me.Load
        LoadData()
    End Sub

    Sub DoWork()
        Dim myModel As New CMMFModel
        DS = New DataSet
        If myModel.LoadDataSSMSupplier(DS) Then
            ProgressReport(4, "Initialize")
        End If

    End Sub

    Private Sub LoadData()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
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
                        ComboBox2.SelectedIndex = 0
                        PriceCMMFSSMBS = New BindingSource
                        PriceCMMFSSMBS.DataSource = DS.Tables(0)
                        PriceCMMFSupplierBS = New BindingSource
                        PriceCMMFSupplierBS.DataSource = PriceCMMFSSMBS
                        PriceCMMFSupplierBS.DataMember = "hdrel-SSM"

                        ComboBox1.DataSource = PriceCMMFSSMBS
                        ComboBox1.DisplayMember = "officersebname"
                        ComboBox1.ValueMember = "ssmid"

                        CheckedListBox1.DataSource = PriceCMMFSupplierBS
                        CheckedListBox1.DisplayMember = "shortname"
                        CheckedListBox1.ValueMember = "vendorcode"

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
        Dim SupplierCriteria As String = String.Empty
        Dim SupplierList As String = String.Empty
        Dim SSMCriteria As String = String.Empty
        Dim SSMIDCriteria As String = String.Empty
        If Me.validate Then
            Title = ComboBox1.Text
            Dim eolst As String = String.Empty
            Dim eolstPhoto As String = String.Empty
            If RadioButton1.Checked Then
                eolst = ""
                eolstPhoto = ""
            ElseIf RadioButton2.Checked Then
                eolst = " and not (eol)"
                eolstPhoto = " and not (c.eol)"
            ElseIf RadioButton3.Checked Then
                eolst = " and eol "
                eolstPhoto = " and c.eol"
            End If
            SSMCriteria = String.Format(" where ssm = '{0}'", ComboBox1.Text.Trim)
            SSMIDCriteria = String.Format(" where ssmid = {0}", DirectCast(ComboBox1.SelectedItem, DataRowView).Row.Item("ssmid"))
            For i = 1 To CheckedListBox1.Items.Count - 1
                If CheckedListBox1.GetItemCheckState(i) = System.Windows.Forms.CheckState.Checked Then
                    myCheck = True

                    SupplierList = SupplierList & IIf(SupplierList = "", "", " or ") & String.Format("shortname = '{0}'", DirectCast(CheckedListBox1.Items(i), DataRowView).Row.Item("shortname"))
                End If
            Next
            SupplierCriteria = IIf(SupplierList = "", "", " and (") & SupplierList & IIf(SupplierList = "", "", ")")

            If CheckBox3.Checked Then
                SqlstrData = String.Format("select * from getpricecmmfwithrange({0}) as (sbuname character(30), familyname character(50), ssm character(30),spm character(30), shortname  character(50), projectname character(100), rangename character(50), picture unknown, cmmf bigint, materialcode character(100),description character(255),range character varying, brandname character(30), countries character(25), voltage character(50), power character(20),pricedummy unknown,datedummy unknown, ""spm price"" numeric , ""spm price date"" date,""spm price hist1"" numeric, ""date hist1"" date,""spm price hist2"" numeric, ""date hist2"" date,""spm price hist3"" numeric,  ""date hist3"" date, ""spm price hist4"" numeric," & _
                """date hist4"" date, remarks text,leadtime integer, qty20 integer, qty40 integer, qty40hq integer, moq integer, loadingname character(30), typeofitem character(30), netprice real, amort real, contractno character(50), length real, width real, height real, lengthbox real, widthbox real, heightbox real, weightwo real, weightwi real, grossweight numeric,vendorcode bigint,eol  boolean) {1} {2} {3} ORDER BY sbuname,familyname,projectname,rangename,cmmf", ComboBox2.Text, SSMCriteria, SupplierCriteria, eolst)

            Else
                SqlstrData = String.Format("select * from getpricecmmf({0}) as (sbuname character(30), familyname character(50), ssm character(30),spm character(30), shortname  character(50), projectname character(100), rangename character(50), picture unknown, cmmf bigint, materialcode character(100),description character(255), brandname character(30), countries character(25), voltage character(50), power character(20),pricedummy unknown,datedummy unknown, ""spm price"" numeric , ""spm price date"" date,""spm price hist1"" numeric, ""date hist1"" date,""spm price hist2"" numeric, ""date hist2"" date,""spm price hist3"" numeric,  ""date hist3"" date, ""spm price hist4"" numeric," & _
                         """date hist4"" date, remarks text,leadtime integer, qty20 integer, qty40 integer, qty40hq integer, moq integer, loadingname character(30), typeofitem character(30), netprice real, amort real, contractno character(50), length real, width real, height real, lengthbox real, widthbox real, heightbox real, weightwo real, weightwi real, grossweight numeric,vendorcode bigint,eol  boolean) {1} {2} {3}  ORDER BY sbuname,familyname,projectname,rangename,cmmf", ComboBox2.Text, SSMCriteria, SupplierCriteria, eolst)
            End If

            SqlstrPhoto = String.Format("Select p.projectname,r.rangename,r.imagepath,c.cmmf" & _
            " FROM ( SELECT DISTINCT pl.cmmf, pl.vendorcode" & _
            " FROM pricelist pl" & _
            " GROUP BY pl.cmmf, pl.vendorcode" & _
            " ORDER BY pl.cmmf, pl.vendorcode) pcp" & _
            " LEFT JOIN vendor v ON v.vendorcode = pcp.vendorcode" & _
            " LEFT JOIN cmmf c1 ON c1.cmmf = pcp.cmmf" & _
            " Left join pccmmf c on c.cmmf = pcp.cmmf" & _
            " LEFT JOIN pcrange r ON r.pcrangeid = c.pcrangeid" & _
            " LEFT JOIN pcproject p ON p.pcprojectid = r.pcprojectid" & _
            " LEFT JOIN family f ON f.familyid = c1.comfam" & _
            " left join sbu on sbu.sbuid = f.sbuid" & _
            " LEFT JOIN ctpricelistsupplierr ct ON ct.cmmf = pcp.cmmf AND ct.vendorcode = pcp.vendorcode" & _
            " {0} {1} {2} and c1.plnt = {3} ORDER BY sbuname,familyname,p.projectname,r.rangename;", SSMIDCriteria, SupplierCriteria, eolst, ComboBox2.Text)

            Dim SaveFileDialog1 As New SaveFileDialog
            SaveFileDialog1.FileName = String.Format("{0}-{1}PriceList{2:yyyyMMdd}", ComboBox1.Text.Trim, ComboBox2.Text, Date.Today)
            SaveFileDialog1.DefaultExt = "xlsx"
            If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                FileNameFullPath = SaveFileDialog1.FileName
                If Not myThread.IsAlive Then
                    ToolStripStatusLabel1.Text = ""
                    myThread = New Thread(AddressOf DoReport)
                    myThread.SetApartmentState(ApartmentState.STA)
                    myThread.Start()
                Else
                    MessageBox.Show("Please wait until the current process is finished.")
                End If
            End If
        End If

    End Sub
    Public Overloads Function validate() As Boolean
        Dim myret As Boolean = True
        ErrorProvider1.SetError(ComboBox1, "")
        If IsNothing(ComboBox1.SelectedItem) Then
            ErrorProvider1.SetError(ComboBox1, "Please select from the list")
            myret = False
        End If
        Return myret
    End Function

    Private Sub DoReport()
        Dim SqlStr As String = SqlstrData
        Dim myCallBackReport As GenerateReportDelegate
        Dim myCallBackRawData As GenerateReportDelegate
        Dim myReport As GenerateReport
        Dim Template As String = String.Empty
        GenerateReport1.sqlstrphoto = SqlstrPhoto
        GenerateReport1.Title = Title
        GenerateReport1.Parent = Me
        If CheckBox3.Checked Then

            myCallBackReport = AddressOf GenerateReport1.GenerateReportWithRange
            myCallBackRawData = AddressOf GenerateReport1.FormatRawDataWithRange
            Template = "\Template\MyTEMPLATE With Range.xltx"
        Else
            myCallBackReport = AddressOf GenerateReport1.GenerateReportNoRange
            myCallBackRawData = AddressOf GenerateReport1.FormatRawData
            Template = "\Template\MyNewTemplate.xltx"
        End If
        myReport = New GenerateReport(Me, SqlStr, FileNameFullPath, myCallBackReport, myCallBackRawData, template:=Template, Location:="A7", FieldNames:=False, sqlstrphoto:=SqlstrPhoto, title:=Title)

        myReport.Run()
    End Sub

    Private Sub CheckedListBox1_ItemCheck(sender As Object, e As ItemCheckEventArgs) Handles CheckedListBox1.ItemCheck
        ErrorProvider1.SetError(ComboBox1, "")
    End Sub
End Class