Imports System.Threading
Imports System.Text

Public Class FormCMMF
    Dim myController As CMMFController
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    'Dim myColumnFilter = {"c.cmmf", "c.materialcode", "p.projectname", "ssm.officersebname", "spm.officersebname"}
    Dim myColumnFilter = {"cmmf", "materialcode", "projectname", "ssm", "spm"}
    Dim criteria As String = String.Empty
    'Dim DRV As DataRowView
    Dim PriceCMMFStatusList As New List(Of PriceCMMFStatus)
    Dim PriceCMMFStatusBS As New BindingSource
    Dim PriceCMMFSBUBS As New BindingSource
    Dim PriceCMMFFamilyBS As New BindingSource
    Dim PriceCMMFSSMBS As New BindingSource
    Dim PriceCMMFSPMBS As New BindingSource
    Dim PriceCMMFProjectBS As New BindingSource
    Dim PriceCMMFRangeBS As New BindingSource

    Dim PriceCMMFBrandBS As New BindingSource
    Dim PriceCMMFLoadingBS As New BindingSource
    Dim PriceCMMFPurchasingGRPBS As New BindingSource

    Dim myFilter As String = String.Empty

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        ToolStripComboBox1.ComboBox.SelectedIndex = 0

        PriceCMMFStatusList.Add(New PriceCMMFStatus(True, "End Of Life"))
        PriceCMMFStatusList.Add(New PriceCMMFStatus(False, "Active"))     
        PriceCMMFStatusBS.DataSource = PriceCMMFStatusList
        
    End Sub

    Private Function getCriteria() As String
        Dim sb As New StringBuilder
        If ToolStripComboBox1.ComboBox.SelectedIndex = 0 Then
            sb.Append(String.Format("where {0} = {1} order by {0}", myColumnFilter(ToolStripComboBox1.ComboBox.SelectedIndex), IIf(IsNumeric(ToolStripTextBox1.Text), ToolStripTextBox1.Text, 0)))
        Else
            sb.Append(String.Format("where upper({0}) like '%{1}%' order by {0}", myColumnFilter(ToolStripComboBox1.ComboBox.SelectedIndex), ToolStripTextBox1.Text.ToUpper))
        End If

        Return sb.ToString
    End Function

    Sub DoWork()
        myController = New CMMFController
        

        RemoveHandler myController.PositionChangeEventHandler, AddressOf BSPriceChange
        AddHandler myController.PositionChangeEventHandler, AddressOf BSPriceChange

        RemoveHandler myController.PricePositionChangeEventHandler, AddressOf BSAgrementChange
        AddHandler myController.PricePositionChangeEventHandler, AddressOf BSAgrementChange
    

        Try
            ProgressReport(1, "Loading...Please wait.")
            'ProgressReport(9, "Clear DataGridView2")
            ClearBinding()
            criteria = ""
            If myController.loaddata(criteria) Then
               
                ProgressReport(4, "Init Data")

            End If
            ProgressReport(1, String.Format("Loading...Done. Records {0}", myController.BS.Count))
        Catch ex As Exception
            ProgressReport(1, ex.Message)
        End Try
    End Sub

    Private Sub RefreshMYInterface()
        DataGridView1.Invalidate()
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

                        'myController.PriceCMMFSBUBS = New BindingSource
                        'myController.PriceCMMFSBUBS.DataSource = myController.DS.Tables(1)
                        'myController.PriceCMMFFamilyBS = New BindingSource

                        'myController.PriceCMMFFamilyBS.DataSource = myController.PriceCMMFSBUBS 'myController.DS.Tables(2)
                        'myController.PriceCMMFFamilyBS.DataMember = "hdrel-SBU"

                        'myController.PriceCMMFSSMBS = New BindingSource
                        'myController.PriceCMMFSSMBS.DataSource = myController.DS.Tables(3)
                        'myController.PriceCMMFSPMBS = New BindingSource
                        'myController.PriceCMMFSPMBS.DataSource = myController.DS.Tables(4)
                        ''PriceCMMFSPMBS.DataMember = "hdrel-SSM"


                        'myController.PriceCMMFProjectBS = New BindingSource
                        'myController.PriceCMMFProjectBS.DataSource = myController.DS.Tables(5)
                        'myController.PriceCMMFRangeBS = New BindingSource
                        'myController.PriceCMMFRangeBS.DataSource = myController.PriceCMMFProjectBS 'myController.DS.Tables(6)
                        'myController.PriceCMMFRangeBS.DataMember = "hdrel"

                        'myController.PriceCMMFBrandBS = New BindingSource
                        'myController.PriceCMMFBrandBS.DataSource = myController.DS.Tables(7)
                        'myController.PriceCMMFLoadingBS = New BindingSource
                        'myController.PriceCMMFLoadingBS.DataSource = myController.DS.Tables(8)
                        'myController.PriceCMMFPurchasingGRPBS = New BindingSource
                        'myController.PriceCMMFPurchasingGRPBS.DataSource = myController.DS.Tables(9)
                        myController.BS = New BindingSource
                        'myController.BS.DataSource = myController.DS.Tables(0)


                        PriceCMMFSBUBS = New BindingSource
                        PriceCMMFSBUBS.DataSource = myController.DS.Tables(1)
                        PriceCMMFFamilyBS = New BindingSource

                        PriceCMMFFamilyBS.DataSource = PriceCMMFSBUBS 'DS.Tables(2)
                        PriceCMMFFamilyBS.DataMember = "hdrel-SBU"

                        PriceCMMFSSMBS = New BindingSource
                        PriceCMMFSSMBS.DataSource = myController.DS.Tables(3)
                        PriceCMMFSPMBS = New BindingSource
                        PriceCMMFSPMBS.DataSource = myController.DS.Tables(4)
                        'PriceCMMFSPMBS.DataMember = "hdrel-SSM"


                        PriceCMMFProjectBS = New BindingSource
                        PriceCMMFProjectBS.DataSource = myController.DS.Tables(5)
                        PriceCMMFRangeBS = New BindingSource
                        PriceCMMFRangeBS.DataSource = PriceCMMFProjectBS 'DS.Tables(6)
                        PriceCMMFRangeBS.DataMember = "hdrel"

                        PriceCMMFBrandBS = New BindingSource
                        PriceCMMFBrandBS.DataSource = myController.DS.Tables(7)
                        PriceCMMFLoadingBS = New BindingSource
                        PriceCMMFLoadingBS.DataSource = myController.DS.Tables(8)
                        PriceCMMFPurchasingGRPBS = New BindingSource
                        PriceCMMFPurchasingGRPBS.DataSource = myController.DS.Tables(9)
                        myController.BS = New BindingSource
                        myController.BS.DataSource = myController.DS.Tables(0)

                        BindingData()
                        myController.BS.Filter = myFilter 'Filter must be applied after binding data

                    Case 8
                        If Not IsNothing(myController.BS.Current) Then
                            Dim drv As DataRowView = myController.BS.Current
                            If Not IsDBNull(drv.Row.Item("cmmf")) Then
                                myController.Model.GetPriceList(drv.Row.Item("cmmf"))
                                bindingdata2()
                            Else
                                clearBindingData2()
                                clearBindingData3()
                            End If
                        Else
                            clearBindingData2()
                            clearBindingData3()
                        End If
                    Case 9
                        If Not IsNothing(myController.BS.Current) Then
                            If Not IsNothing(myController.Model.BSPricelist.Current) Then
                                Dim drv As DataRowView = myController.Model.BSPricelist.Current
                                bindingdata3()
                                Dim myfilter As Long
                                If IsDBNull(drv.Row.Item("agreement")) Then
                                    myfilter = 0
                                Else
                                    myfilter = drv.Row.Item("agreement")
                                End If
                                myController.Model.BSAgreementList.Filter = String.Format("[agreement] = {0}", myfilter)
                            End If
                        Else
                        End If

                    Case 11
                        ClearBindingControls()
                End Select
            End If
        Catch ex As Exception
            Debug.Print(ex.Message)
        End Try
        
    End Sub

    Private Sub ToolStripButton6_Click(sender As Object, e As EventArgs) Handles ToolStripButton6.Click
       
        LoadData()
    End Sub

    Private Sub LoadData()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            criteria = getCriteria()
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Private Sub ToolStripComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ToolStripComboBox1.SelectedIndexChanged, ToolStripTextBox1.TextChanged
        'LoadData()
        ApplyFilter()
       

    End Sub
    Private Sub BSAgrementChange()
        ProgressReport(9, "Init Agreement")
    End Sub

    Private Sub BSPriceChange()
        ProgressReport(8, "Init Price")
    End Sub

    Private Sub BindingData()
        TextBox1.Clear()
        TextBox29.Clear()
        Label41.Text = ""

        TextBox1.DataBindings.Clear()
        TextBox2.DataBindings.Clear()
        TextBox3.DataBindings.Clear()
        TextBox4.DataBindings.Clear()
        TextBox5.DataBindings.Clear()
        TextBox6.DataBindings.Clear()
        TextBox7.DataBindings.Clear()
        TextBox8.DataBindings.Clear()
        TextBox9.DataBindings.Clear()
        TextBox10.DataBindings.Clear()
        TextBox11.DataBindings.Clear()
        TextBox12.DataBindings.Clear()
        TextBox13.DataBindings.Clear()
        TextBox14.DataBindings.Clear()
        TextBox15.DataBindings.Clear()
        TextBox16.DataBindings.Clear()
        TextBox17.DataBindings.Clear()
        TextBox18.DataBindings.Clear()
        TextBox19.DataBindings.Clear()
        TextBox20.DataBindings.Clear()
        TextBox21.DataBindings.Clear()
        TextBox22.DataBindings.Clear()
        TextBox23.DataBindings.Clear()
        TextBox24.DataBindings.Clear()
        TextBox25.DataBindings.Clear()
        TextBox26.DataBindings.Clear()
        TextBox27.DataBindings.Clear()
        TextBox28.DataBindings.Clear()
        TextBox29.DataBindings.Clear()
        Label41.DataBindings.Clear()
        ComboBox1.DataBindings.Clear()
        ComboBox2.DataBindings.Clear()
        ComboBox3.DataBindings.Clear()
        ComboBox4.DataBindings.Clear()
        ComboBox5.DataBindings.Clear()
        ComboBox6.DataBindings.Clear()
        ComboBox7.DataBindings.Clear()
        ComboBox8.DataBindings.Clear()
        ComboBox9.DataBindings.Clear()
        ComboBox10.DataBindings.Clear()

        ComboBox1.DisplayMember = "sbuname"
        ComboBox1.ValueMember = "sbuid"
        ComboBox1.DataSource = PriceCMMFSBUBS

        ComboBox2.DisplayMember = "familyname"
        ComboBox2.ValueMember = "familyid"
        ComboBox2.DataSource = PriceCMMFFamilyBS

        ComboBox3.DisplayMember = "ssmname"
        ComboBox3.ValueMember = "ssmid"
        ComboBox3.DataSource = PriceCMMFSSMBS

        ComboBox4.DisplayMember = "pmname"
        ComboBox4.ValueMember = "pmid"
        ComboBox4.DataSource = PriceCMMFSPMBS

        ComboBox5.DisplayMember = "projectname"
        ComboBox5.ValueMember = "pcprojectid"
        ComboBox5.DataSource = PriceCMMFProjectBS

        ComboBox6.DisplayMember = "rangename"
        ComboBox6.ValueMember = "pcrangeid"
        ComboBox6.DataSource = PriceCMMFRangeBS

        ComboBox7.DisplayMember = "brandname"
        ComboBox7.ValueMember = "brandid"
        ComboBox7.DataSource = PriceCMMFBrandBS

        ComboBox8.DisplayMember = "loadingname"
        ComboBox8.ValueMember = "loadingid"
        ComboBox8.DataSource = PriceCMMFLoadingBS

        ComboBox9.DisplayMember = "typeofitem"
        ComboBox9.ValueMember = "pgid"
        ComboBox9.DataSource = PriceCMMFPurchasingGRPBS

        ComboBox10.DisplayMember = "StatusDesc"
        ComboBox10.ValueMember = "Status"
        ComboBox10.DataSource = PriceCMMFStatusBS

        DataGridView1.AutoGenerateColumns = False
        DataGridView1.DataSource = myController.BS
        TextBox1.DataBindings.Add(New Binding("Text", myController.BS, "cmmf", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox2.DataBindings.Add(New Binding("Text", myController.BS, "description", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox3.DataBindings.Add(New Binding("Text", myController.BS, "countries", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox4.DataBindings.Add(New Binding("Text", myController.BS, "voltage", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox5.DataBindings.Add(New Binding("Text", myController.BS, "power", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox6.DataBindings.Add(New Binding("Text", myController.BS, "remarks", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox7.DataBindings.Add(New Binding("Text", myController.BS, "leadtime", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox8.DataBindings.Add(New Binding("Text", myController.BS, "qty20", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox9.DataBindings.Add(New Binding("Text", myController.BS, "qty40", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox10.DataBindings.Add(New Binding("Text", myController.BS, "qty40hq", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox11.DataBindings.Add(New Binding("Text", myController.BS, "moq", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox12.DataBindings.Add(New Binding("Text", myController.BS, "contractno", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox13.DataBindings.Add(New Binding("Text", myController.BS, "length", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox14.DataBindings.Add(New Binding("Text", myController.BS, "width", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox15.DataBindings.Add(New Binding("Text", myController.BS, "height", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox16.DataBindings.Add(New Binding("Text", myController.BS, "lengthbox", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox17.DataBindings.Add(New Binding("Text", myController.BS, "widthbox", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox18.DataBindings.Add(New Binding("Text", myController.BS, "heightbox", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox19.DataBindings.Add(New Binding("Text", myController.BS, "weightwo", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox20.DataBindings.Add(New Binding("Text", myController.BS, "weightwi", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox21.DataBindings.Add(New Binding("Text", myController.BS, "grossweight", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox22.DataBindings.Add(New Binding("Text", myController.BS, "nettweight", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox23.DataBindings.Add(New Binding("Text", myController.BS, "pcspercartoon", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox24.DataBindings.Add(New Binding("Text", myController.BS, "sppet", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox25.DataBindings.Add(New Binding("Text", myController.BS, "stcseb", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox26.DataBindings.Add(New Binding("Text", myController.BS, "stcsup", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox27.DataBindings.Add(New Binding("Text", myController.BS, "srdc", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox28.DataBindings.Add(New Binding("Text", myController.BS, "spps", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox29.DataBindings.Add(New Binding("Text", myController.BS, "materialcode", True, DataSourceUpdateMode.OnPropertyChanged))
        Label41.DataBindings.Add(New Binding("Text", myController.BS, "cmmf", True, DataSourceUpdateMode.OnPropertyChanged))

        ComboBox1.DataBindings.Add(New Binding("SelectedValue", myController.BS, "sbuid", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox2.DataBindings.Add(New Binding("SelectedValue", myController.BS, "familyid", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox3.DataBindings.Add(New Binding("SelectedValue", myController.BS, "ssmid", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox4.DataBindings.Add(New Binding("SelectedValue", myController.BS, "spmid", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox5.DataBindings.Add(New Binding("SelectedValue", myController.BS, "pcprojectid", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox6.DataBindings.Add(New Binding("SelectedValue", myController.BS, "pcrangeid", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox7.DataBindings.Add(New Binding("SelectedValue", myController.BS, "brandid", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox8.DataBindings.Add(New Binding("SelectedValue", myController.BS, "loadingid", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox9.DataBindings.Add(New Binding("SelectedValue", myController.BS, "pgid", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox10.DataBindings.Add(New Binding("SelectedValue", myController.BS, "eol", True, DataSourceUpdateMode.OnPropertyChanged))
    End Sub

    Private Sub clearBindingData2()
        TextBox30.Clear()
        TextBox32.Clear()
        TextBox33.Clear()
        TextBox34.Clear()
        TextBox35.Clear()
        TextBox36.Clear()
        TextBox37.Clear()
        TextBox38.Clear()
        TextBox39.Clear()
        TextBox40.Clear()
        TextBox41.Clear()
        TextBox42.Clear()
        TextBox43.Clear()
        TextBox44.Clear()
        TextBox30.DataBindings.Clear()
        TextBox32.DataBindings.Clear()
        TextBox33.DataBindings.Clear()
        TextBox34.DataBindings.Clear()
        TextBox35.DataBindings.Clear()
        TextBox36.DataBindings.Clear()
        TextBox37.DataBindings.Clear()
        TextBox38.DataBindings.Clear()
        TextBox39.DataBindings.Clear()
        TextBox40.DataBindings.Clear()
        TextBox41.DataBindings.Clear()
        TextBox42.DataBindings.Clear()
        TextBox43.DataBindings.Clear()
        TextBox44.DataBindings.Clear()
        DataGridView2.DataSource = Nothing       
    End Sub

    Private Sub bindingdata2()
        clearBindingData2()

        DataGridView2.AutoGenerateColumns = False
        DataGridView2.DataSource = myController.Model.BSPriceList

        TextBox30.DataBindings.Add(New Binding("Text", myController.Model.BSPriceList, "allcomments", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox32.DataBindings.Add(New Binding("Text", myController.Model.BSPriceList, "price", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0.00"))
        TextBox33.DataBindings.Add(New Binding("Text", myController.Model.BSPriceList, "validamort", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0.00"))
        TextBox34.DataBindings.Add(New Binding("Text", myController.Model.BSPriceList, "fobamort", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0.00"))
        TextBox35.DataBindings.Add(New Binding("Text", myController.Model.BSPriceList, "mydate", True, DataSourceUpdateMode.OnPropertyChanged, "", "dd-MMM-yyyy"))
        TextBox36.DataBindings.Add(New Binding("Text", myController.Model.BSPriceList, "agreement", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox37.DataBindings.Add(New Binding("Text", myController.Model.BSPricelist, "closingdate", True, DataSourceUpdateMode.OnPropertyChanged, "", "dd-MMM-yyyy"))
        TextBox38.DataBindings.Add(New Binding("Text", myController.Model.BSPriceList, "numbercmmf", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox39.DataBindings.Add(New Binding("Text", myController.Model.BSPriceList, "status", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox40.DataBindings.Add(New Binding("Text", myController.Model.BSPriceList, "deliveredqty", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0"))
        TextBox41.DataBindings.Add(New Binding("Text", myController.Model.BSPriceList, "agqty", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0"))
        TextBox42.DataBindings.Add(New Binding("Text", myController.Model.BSPriceList, "totalqty", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0"))
        TextBox43.DataBindings.Add(New Binding("Text", myController.Model.BSPriceList, "startdate", True, DataSourceUpdateMode.OnPropertyChanged, "", "dd-MMM-yyyy"))
        TextBox44.DataBindings.Add(New Binding("Text", myController.Model.BSPriceList, "enddate", True, DataSourceUpdateMode.OnPropertyChanged, "", "dd-MMM-yyyy"))

    End Sub

    Private Sub ClearBinding()
        ProgressReport(11, "Clear binding")
    End Sub

    Private Sub ClearBindingControls()
        'Clear binding first before adding. to avoid multiple binding.

        For Each ControlTC In TabControl1.Controls
            For Each Control In ControlTC.Controls
                If TypeOf (Control) Is GroupBox Then
                    For Each ControlGB In Control.Controls
                        If TypeOf (ControlGB) Is TextBox Then
                            Dim textbox = DirectCast(ControlGB, TextBox)
                            textbox.Clear()
                            textbox.DataBindings.Clear()
                        ElseIf TypeOf (ControlGB) Is ComboBox Then
                            Dim Combobox = DirectCast(ControlGB, ComboBox)
                            Combobox.DataBindings.Clear()
                        ElseIf TypeOf (ControlGB) Is ListBox Then
                            Dim listBox = DirectCast(ControlGB, ListBox)
                            listBox.Items.Clear()
                            listBox.DataBindings.Clear()
                        End If
                    Next
                ElseIf TypeOf (Control) Is DataGridView Then
                    Dim dgv = DirectCast(Control, DataGridView)
                    dgv.DataSource = Nothing
                End If
            Next
        Next
    End Sub


    Private Sub FormCMMF_Load(sender As Object, e As EventArgs) Handles Me.Load
        'Make sure the Control ready for binding
        For Each myTab As TabPage In TabControl1.Controls
            myTab.Show()
        Next
        LoadData()
    End Sub

    Private Sub bindingdata3()
        DataGridView3.AutoGenerateColumns = False
        DataGridView3.DataSource = myController.Model.BSAgreementList
    End Sub

    Private Sub clearBindingData3()
        DataGridView3.DataSource = Nothing
    End Sub

    Private Sub ToolStripButton5_Click(sender As Object, e As EventArgs) Handles ToolStripButton5.Click
        Me.Validate()
        PriceCMMFProjectBS.EndEdit()
        PriceCMMFRangeBS.EndEdit()
        
        myController.save()

    End Sub


    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click, Button2.Click
        Dim bs As New BindingSource
        bs.DataSource = myController.DS.Tables(2).Copy()
        Dim myhelper As New DialogHelper(bs)
        myhelper.Filter = "sbuname like '%{0}%' or familyname like '%{0}%'"
        myhelper.Column1.HeaderText = "SBU"
        myhelper.Column1.DataPropertyName = "sbuname"
        myhelper.Column1.Width = 250

        myhelper.Column2.HeaderText = "Family"
        myhelper.Column2.DataPropertyName = "familyname"
        myhelper.Column2.Width = 250

        If myhelper.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Dim drvhelper As DataRowView = bs.Current
            Dim drv As DataRowView = myController.BS.Current
            drv.Row.Item("sbuid") = drvhelper.Row.Item("sbuid")
            drv.Row.Item("sbuname") = drvhelper.Row.Item("sbuname")
            drv.Row.Item("familyid") = drvhelper.Row.Item("familyid")
            drv.Row.Item("familyname") = drvhelper.Row.Item("familyname")

            Dim myPos = PriceCMMFSBUBS.Find("sbuid", drvhelper.Row.Item("sbuid"))
            ComboBox1.SelectedIndex = myPos
            myPos = PriceCMMFFamilyBS.Find("familyid", drvhelper.Row.Item("familyid"))
            ComboBox2.SelectedIndex = myPos
            InvalidateMe()
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click, Button3.Click
        Dim bs As New BindingSource
        bs.DataSource = myController.DS.Tables(4).Copy()
        Dim myhelper As New DialogHelper(bs)
        myhelper.Filter = "ssmname like '%{0}%' or pmname like '%{0}%'"
        myhelper.Column1.HeaderText = "SSM Name"
        myhelper.Column1.DataPropertyName = "ssmname"
        myhelper.Column1.Width = 250

        myhelper.Column2.HeaderText = "PM Name"
        myhelper.Column2.DataPropertyName = "pmname"
        myhelper.Column2.Width = 250

        myhelper.Column3.Visible = False
        myhelper.Column4.Visible = False

        If myhelper.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Dim drvhelper As DataRowView = bs.Current
            Dim drv As DataRowView = myController.BS.Current
            drv.Row.Item("ssmid") = drvhelper.Row.Item("ssmid")
            drv.Row.Item("ssm") = drvhelper.Row.Item("ssmname")
            drv.Row.Item("spmid") = drvhelper.Row.Item("pmid")
            drv.Row.Item("spm") = drvhelper.Row.Item("pmname")

            Dim myPos = PriceCMMFSSMBS.Find("ssmid", drvhelper.Row.Item("ssmid"))
            ComboBox3.SelectedIndex = myPos
            myPos = PriceCMMFSPMBS.Find("pmid", drvhelper.Row.Item("pmid"))
            ComboBox4.SelectedIndex = myPos
            InvalidateMe()
        End If
    End Sub
   
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click, Button4.Click
        Dim bs As New BindingSource
        bs.DataSource = myController.DS.Tables(10).Copy()
        Dim myhelper As New DialogHelper(bs)
        myhelper.Filter = "projectname like '%{0}%' or rangename like '%{0}%'"
        myhelper.Column1.HeaderText = "Project Name"
        myhelper.Column1.DataPropertyName = "projectname"
        myhelper.Column1.Width = 250

        myhelper.Column2.HeaderText = "Range Name"
        myhelper.Column2.DataPropertyName = "rangename"
        myhelper.Column2.Width = 250

        myhelper.Column3.HeaderText = "SSM"
        myhelper.Column3.DataPropertyName = "ssm"
        myhelper.Column3.Width = 250

        myhelper.Column4.HeaderText = "SPM"
        myhelper.Column4.DataPropertyName = "spm"
        myhelper.Column4.Width = 250


        If myhelper.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Dim drvhelper As DataRowView = bs.Current
            Dim drv As DataRowView = myController.BS.Current
            drv.Row.Item("pcrangeid") = drvhelper.Row.Item("pcrangeid")
            'Find ProjectCombobox
            Dim myPos = PriceCMMFProjectBS.Find("pcprojectid", drvhelper.Row.Item("pcprojectid"))
            ComboBox5.SelectedIndex = myPos
            'Find RangeCombobox
            myPos = PriceCMMFRangeBS.Find("pcrangeid", drvhelper.Row.Item("pcrangeid"))
            ComboBox6.SelectedIndex = myPos

            drv.Row.Item("projectname") = drvhelper.Row.Item("projectname")
            InvalidateMe()
        End If
    End Sub



    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        Dim DRV = myController.BS.AddNew()
        DRV.Row.Item("eol") = False
    End Sub

    Private Sub DataGridView1_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError

    End Sub

    Private Sub InvalidateMe()
        DataGridView1.Invalidate()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged, TextBox29.TextChanged
        InvalidateMe()
    End Sub


    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged

    End Sub

    Private Sub ComboBox5_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles ComboBox5.Validating, ComboBox6.Validating
        Dim cb As ComboBox = DirectCast(sender, ComboBox)
        If IsNothing(cb.SelectedItem) Then
            Dim currDrv = myController.BS.Current
            Select Case cb.Name
                Case "ComboBox5" 'Project
                    Dim ProjectName As String = ComboBox5.Text
                    Dim drv As DataRowView = PriceCMMFProjectBS.AddNew
                    drv.Row.Item("projectname") = ProjectName
                    currDrv.Row.Item("projectname") = drv.Row.Item("projectname")
                    drv.EndEdit()
                    Dim mypos = PriceCMMFProjectBS.Find("pcprojectid", drv.Row.Item("pcprojectid"))
                    ComboBox5.SelectedIndex = mypos

                    InvalidateMe()

                Case "ComboBox6" 'Range
                    Dim RangeName As String = ComboBox6.Text

                    Dim drv As DataRowView = PriceCMMFRangeBS.AddNew
                    drv.Row.Item("rangename") = RangeName
                    drv.EndEdit()
                    currDrv.Row.Item("rangename") = drv.Row.Item("rangename")
                    Dim mypos = PriceCMMFRangeBS.Find("pcrangeid", drv.Row.Item("pcrangeid"))
                    ComboBox6.SelectedIndex = mypos

                    InvalidateMe()
            End Select
        End If
    End Sub

    Private Sub DeleteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteToolStripMenuItem.Click

        Dim cb As ToolStripMenuItem = DirectCast(sender, ToolStripMenuItem)
        Dim mycombo = cb.GetCurrentParent
        Dim abc As ContextMenuStrip = DirectCast(mycombo, ContextMenuStrip)
        Dim bcd As ComboBox = DirectCast(abc.SourceControl, ComboBox)
        If Not IsNothing(bcd.SelectedItem) Then
            Dim currDrv = myController.BS.Current
            Select Case bcd.Name
                Case "ComboBox5" 'Project
                    PriceCMMFProjectBS.RemoveCurrent()
                    currDrv.Row.Item("projectname") = ""
                Case "ComboBox6" 'Range
                    PriceCMMFRangeBS.RemoveCurrent()
            End Select
            InvalidateMe()
        End If
    End Sub


   
    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click
        If Not IsNothing(myController.GetCurrentRecord) Then
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView1.SelectedRows
                    myController.RemoveAt(drv.Index)
                Next
            End If
            ApplyFilter()
        End If
    End Sub

    Private Sub ApplyFilter()
        If Not myThread.IsAlive Then
            Dim sb As New StringBuilder
            myFilter = ""
            If ToolStripComboBox1.ComboBox.SelectedIndex = 0 Then
                sb.Append(String.Format("[{0}] = {1}", myColumnFilter(ToolStripComboBox1.ComboBox.SelectedIndex), IIf(IsNumeric(ToolStripTextBox1.Text), ToolStripTextBox1.Text, 0)))
            Else
                sb.Append(String.Format("[{0}] like '%{1}%'", myColumnFilter(ToolStripComboBox1.ComboBox.SelectedIndex), ToolStripTextBox1.Text.ToUpper))
            End If
            If ToolStripTextBox1.Text <> "" Then
                myFilter = sb.ToString
            End If
            If Not IsNothing(myController) Then
                myController.ApplyFilter = myFilter
                ProgressReport(1, String.Format("Filter...Done. Records {0}", myController.BS.Count))
            End If
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If

    End Sub


End Class

Public Class PriceCMMFStatus
    Property Status As Boolean
    Property StatusDesc As String

    Public Sub New(status As Boolean, statusdesc As String)
        Me.Status = status
        Me.StatusDesc = statusdesc
    End Sub

    Public Overrides Function ToString() As String
        Return StatusDesc.ToString
    End Function
End Class