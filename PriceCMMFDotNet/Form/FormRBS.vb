Imports System.Threading
'Imports Microsoft.Office.Interop
'Imports System.IO

Public Class FormRBS
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim VendorBS As BindingSource
    Private FileNameFullPath As String = String.Empty
    Dim myCriteria As String = String.Empty
    Dim sqlstrData As String = String.Empty
    Dim SqlStrPhoto As String = String.Empty
    Dim GenerateReport1 As New GenerateReport()
    Dim Title As String = String.Empty
    Private Sub FormRBS_Load(sender As Object, e As EventArgs) Handles Me.Load
        LoadData()
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

    Private Sub DoWork()
        Dim myModel As New CMMFModel
        VendorBS = New BindingSource
        VendorBS.DataSource = myModel.getVendorShortNameTable ' myModel.getVendorTable
        ProgressReport(4, "Initialize")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim myhelper As DialogHelper = New DialogHelper(VendorBS)

        myhelper.DataGridView1.Columns(0).HeaderText = "Supplier"
        myhelper.DataGridView1.Columns(0).DataPropertyName = "shortname"
        myhelper.DataGridView1.Columns(0).Width = "300"
        myhelper.Filter = "[shortname] like '%{0}%'"
        For i = 1 To myhelper.DataGridView1.Columns.Count - 1
            myhelper.DataGridView1.Columns(i).Visible = False
        Next
        If myhelper.ShowDialog = Windows.Forms.DialogResult.OK Then
            Dim drv As DataRowView = VendorBS.Current
            TextBox1.Text = drv.Row.Item("shortname")
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
        If Me.validate Then
            Title = TextBox1.Text
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

            If CheckBox3.Checked Then
                sqlstrData = String.Format("select * from getpricecmmfwithrange({0}) as (sbuname character(30), familyname character(50), ssm character(30),spm character(30), shortname  character(50), projectname character(100), rangename character(50), picture unknown, cmmf bigint, materialcode character(100),description character(255),range character varying, brandname character(30), countries character(25), voltage character(50), power character(20),pricedummy unknown,datedummy unknown, ""spm price"" numeric , ""spm price date"" date,""spm price hist1"" numeric, ""date hist1"" date,""spm price hist2"" numeric, ""date hist2"" date,""spm price hist3"" numeric,  ""date hist3"" date, ""spm price hist4"" numeric," &
                                           """date hist4"" date, remarks text,leadtime integer, qty20 integer, qty40 integer, qty40hq integer, moq integer, loadingname character(30), typeofitem character(30), netprice real, amort real, contractno character(50), length real, width real, height real, lengthbox real, widthbox real, heightbox real, weightwo real, weightwi real, grossweight numeric,vendorcode bigint,eol  boolean)" &
                                           " where shortname = '{1}' {2}  order by sbuname,familyname,projectname,rangename,cmmf", ComboBox2.Text, TextBox1.Text, eolst)

            Else
                sqlstrData = String.Format("select * from getpricecmmf({0}) as (sbuname character(30), familyname character(50), ssm character(30),spm character(30), shortname  character(50), projectname character(100), rangename character(50), picture unknown, cmmf bigint, materialcode character(100),description character(255), brandname character(30), countries character(25), voltage character(50), power character(20),pricedummy unknown,datedummy unknown, ""spm price"" numeric , ""spm price date"" date,""spm price hist1"" numeric, ""date hist1"" date,""spm price hist2"" numeric, ""date hist2"" date,""spm price hist3"" numeric,  ""date hist3"" date, ""spm price hist4"" numeric," &
                                           """date hist4"" date, remarks text,leadtime integer, qty20 integer, qty40 integer, qty40hq integer, moq integer, loadingname character(30), typeofitem character(30), netprice real, amort real, contractno character(50), length real, width real, height real, lengthbox real, widthbox real, heightbox real, weightwo real, weightwi real, grossweight numeric,vendorcode bigint,eol  boolean)" &
                                           " where shortname = '{1}' {2} order by sbuname,familyname,projectname,rangename,cmmf", ComboBox2.Text, TextBox1.Text, eolst)
            End If

            SqlStrPhoto = String.Format("Select p.projectname,r.rangename,r.imagepath,c.cmmf" &
            " FROM ( SELECT DISTINCT pl.cmmf, pl.vendorcode" &
            " FROM pricelist pl" &
            " GROUP BY pl.cmmf, pl.vendorcode" &
            " ORDER BY pl.cmmf, pl.vendorcode) pcp" & _
            " LEFT JOIN vendor v ON v.vendorcode = pcp.vendorcode" &
            " LEFT JOIN cmmf c1 ON c1.cmmf = pcp.cmmf" &
            " Left join pccmmf c on c.cmmf = pcp.cmmf" &
            " LEFT JOIN pcrange r ON r.pcrangeid = c.pcrangeid" &
            " LEFT JOIN pcproject p ON p.pcprojectid = r.pcprojectid" &
            " LEFT JOIN family f ON f.familyid = c1.comfam" &
            " left join sbu on sbu.sbuid = f.sbuid" &
            " LEFT JOIN ctpricelistsupplierr ct ON ct.cmmf = pcp.cmmf AND ct.vendorcode = pcp.vendorcode" &
            " where v.shortname = '{0}' {1} and c1.plnt = {2}" &
            " ORDER BY sbuname,familyname,p.projectname,r.rangename", TextBox1.Text, eolstPhoto, ComboBox2.Text)

            Dim SaveFileDialog1 As New SaveFileDialog
            SaveFileDialog1.FileName = String.Format("{0}-{1}PriceList{2:yyyyMMdd}", TextBox1.Text.Trim, ComboBox2.Text, Date.Today)
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
        ErrorProvider1.SetError(TextBox1, "")
        If TextBox1.Text = "" Then
            ErrorProvider1.SetError(TextBox1, "Supplier cannot be blank.")
            myret = False
        End If
        Return myret
    End Function

    Private Sub DoReport()
        Dim SqlStr As String = sqlstrData
        Dim myCallBackReport As GenerateReportDelegate
        Dim myCallBackRawData As GenerateReportDelegate
        Dim myReport As GenerateReport
        Dim Template As String = String.Empty
        GenerateReport1.sqlstrphoto = SqlStrPhoto
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
        myReport = New GenerateReport(Me, SqlStr, FileNameFullPath, myCallBackReport, myCallBackRawData, template:=Template, Location:="A7", FieldNames:=False, sqlstrphoto:=SqlStrPhoto, title:=Title)

        myreport.Run()
    End Sub

    'Sub GenerateReportNoRange(ByRef sender As Object, ByRef e As System.EventArgs)

    '    Dim GenerateReport1 = New GenerateReport
    '    Dim oWb = DirectCast(sender, Excel.Workbook)
    '    Dim oxl As Excel.Application = oWb.Parent
    '    Dim oBC = oWb.Worksheets(1)
    '    Dim oWF As Excel.WorksheetFunction
    '    Dim myImagePath = My.Settings.MyImagePath
    '    Dim i As Integer
    '    Application.DoEvents()

    '    oBC.Cells.EntireColumn.AutoFit()
    '    ProgressReport(1, "Data Formatting...")
    '    oBC.Range("H:H").ColumnWidth = 20
    '    oBC.Range("Q:Q").NumberFormat = "dd-MMM-yyyy"
    '    oBC.Range("A7:AP1000").VerticalAlignment = Excel.Constants.xlTop
    '    oBC.Range("M:M").ColumnWidth = 30
    '    oBC.Range("P:AA").ColumnWidth = 11
    '    oBC.Range("R:R").ColumnWidth = 8
    '    oBC.Range("T:T").ColumnWidth = 8
    '    oBC.Range("V:V").ColumnWidth = 8
    '    oBC.Range("X:X").ColumnWidth = 8
    '    oBC.Range("Z:Z").ColumnWidth = 8
    '    oBC.Range("G:G").ColumnWidth = 15
    '    oBC.Cells.EntireColumn.WrapText = True

    '    ProgressReport(1, "Put Picture(s)...")

    '    'assign Photo
    '    Dim DT As DataTable = DataAccess.GetDataTable(SqlStrPhoto, CommandType.Text)
    '    Dim BS As New BindingSource
    '    BS.DataSource = DT


    '    'rsTmp.MoveFirst
    '    Dim iRow As Long
    '    Dim Span As Double
    '    Dim j As Integer
    '    Dim Pict As Object
    '    Dim myfirstRow As Long
    '    Dim mylastRow As Long

    '    Dim myWD As Double
    '    Dim ganjil As Long
    '    Dim reccount As Long

    '    Dim myfilename As String
    '    iRow = 7
    '    oWF = oxl.WorksheetFunction
    '    oBC.Rows("7:" & DT.Rows.Count + 6).Interior.ColorIndex = 36


    '    Call GenerateReport1.MakeBorder(oBC, 7, 1, DT.Rows.Count + 6, 48, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous)
    '    Call GenerateReport1.MakeInsideBorder(oBC, 7, 1, DT.Rows.Count + 6, 48, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous)


    '    For i = 0 To DT.Rows.Count - 1
    '        ProgressReport(1, String.Format("Put Picture(s).. {0}/{1}", reccount, DT.Rows.Count - 1))
    '        If IsDBNull(DT.Rows(i).Item("projectname")) AndAlso IsDBNull(DT.Rows(i).Item("rangename")) Then
    '        Else
    '            If GenerateReport1.FirstOf(BS, 1) Then
    '                myfirstRow = iRow
    '            End If
    '            If GenerateReport1.LastOf(BS, 1) Then
    '                mylastRow = iRow
    '                ganjil = ganjil + 1
    '                If ganjil Mod 2 = 0 Then
    '                    oBC.Rows(myfirstRow & ":" & mylastRow).Interior.ColorIndex = 40
    '                End If
    '                If Not (IsDBNull(DT.Rows(i).Item("imagepath"))) Then
    '                    myfilename = myImagePath & "\" & Path.GetFileName(Trim(DT.Rows(i).Item("imagepath")))
    '                    If File.Exists(myfilename) Then
    '                        If oBC.Range("H" & myfirstRow & ":H" & mylastRow).Height < 100 Then
    '                            'Reserve height for cell
    '                            If myfirstRow = mylastRow Then
    '                                oBC.Cells(myfirstRow, 2).RowHeight = 100
    '                            Else
    '                                Span = (100 - oBC.Range("H" & myfirstRow & ":H" & mylastRow).Height) / (mylastRow + 1 - myfirstRow)
    '                                For j = 0 To (mylastRow - myfirstRow)
    '                                    If j = mylastRow - myfirstRow Then
    '                                        Span = 100 - oBC.Range("H" & myfirstRow & ":H" & mylastRow).Height
    '                                    End If
    '                                    oBC.Cells(myfirstRow + j, 2).RowHeight = oBC.Cells(myfirstRow + j, 2).RowHeight + Span
    '                                Next j
    '                            End If
    '                        End If
    '                        oBC.Range("H" & myfirstRow & ":H" & mylastRow).Merge()
    '                        'Print
    '                        Dim PictHT As Single, PictWD As Single
    '                        Dim RngHT As Single, RngWD As Single
    '                        Dim myScale As Double, MyL As Double, MyT As Double
    '                        ProgressReport(1, String.Format("Put Picture(s)...{0}/{1} {2}", reccount, DT.Rows.Count - 1, myfilename))

    '                        Pict = oBC.Pictures.Insert(myfilename)

    '                        Pict.Placement = 1

    '                        'do pic scale
    '                        Application.DoEvents()

    '                        PictHT = Pict.Height
    '                        PictWD = Pict.Width

    '                        'Size of Template Picture
    '                        RngHT = 98
    '                        RngWD = 98

    '                        myScale = oWF.Min(RngHT / PictHT, RngWD / PictWD)

    '                        With Pict
    '                            .Width = PictWD * myScale
    '                            .Height = PictHT * myScale

    '                        End With
    '                        'Put Picture Location
    '                        With oBC.Range("H" & myfirstRow)
    '                            MyT = .Top
    '                            MyL = .Left
    '                        End With

    '                        'Center Pict H & V
    '                        If PictHT > PictWD Then
    '                            MyL = MyL + (RngWD - Pict.Width) / 2
    '                        Else
    '                            MyT = MyT + (RngHT - Pict.Height) / 2
    '                        End If
    '                        myWD = 0

    '                        myWD = (oBC.Range("B" & iRow).Offset(0, 1).Left) - (oBC.Range("B" & iRow).Left) - 2

    '                        With Pict
    '                            .Top = MyT + 1
    '                            .Left = MyL + 2
    '                        End With
    '                        Pict = Nothing
    '                    End If
    '                End If
    '            End If

    '        End If

    '        'End If
    '        iRow = iRow + 1
    '        reccount = reccount + 1

    '    Next

    '    ProgressReport(1, "AutoFit ....")
    '    Application.DoEvents()
    '    oBC.Cells.EntireColumn.AutoFit()
    '    ProgressReport(1, "Interior Pattern ....")
    '    oBC.Range(oBC.Columns("AW:AW"), oBC.Columns("AW:AW").End(Excel.XlDirection.xlToRight)).Interior.Pattern = Excel.Constants.xlNone
    '    ProgressReport(1, "Set Range ....")

    '    oBC.Range("H:H").ColumnWidth = 20
    '    oBC.Range("Q:Q").NumberFormat = "dd-MMM-yyyy"
    '    oBC.Range("A7:AP1000").VerticalAlignment = Excel.Constants.xlTop
    '    oBC.Range("M:M").ColumnWidth = 30

    '    oBC.Range("M:M").ColumnWidth = 30
    '    oBC.Range("P:AA").ColumnWidth = 11
    '    oBC.Range("R:R").ColumnWidth = 8
    '    oBC.Range("T:T").ColumnWidth = 8
    '    oBC.Range("V:V").ColumnWidth = 8
    '    oBC.Range("X:X").ColumnWidth = 8
    '    oBC.Range("Z:Z").ColumnWidth = 8
    '    oBC.Range("G:G").ColumnWidth = 15
    '    ProgressReport(1, "Wrap Text ....")
    '    oBC.Cells.EntireColumn.WrapText = True
    '    oBC.Range("G1:I1").WrapText = False
    '    oBC.Cells(1, 8) = Date.Today
    '    oBC.Cells(1, 9) = TextBox1.Text
    '    oWb.Worksheets(1).Select()
    '    oxl.DisplayAlerts = False
    'End Sub

    'Sub FormatRawData()

    'End Sub

    'Private Sub FormatRawDataWithRange(ByRef sender As Object, ByRef e As System.EventArgs)
    'End Sub

    'Private Sub GenerateReportWithRange(ByRef sender As Object, ByRef e As System.EventArgs)
    '    Dim GenerateReport1 = New GenerateReport
    '    Dim oWb = DirectCast(sender, Excel.Workbook)
    '    Dim oxl As Excel.Application = oWb.Parent
    '    Dim oBC = oWb.Worksheets(1)
    '    Dim oWF As Excel.WorksheetFunction
    '    Dim myImagePath = My.Settings.MyImagePath

    '    oBC.Cells.EntireColumn.AutoFit()
    '    ProgressReport(1, "Data Formatting..")
    '    oBC.Range("H:H").ColumnWidth = 20
    '    oBC.Range("R:R").NumberFormat = "dd-MMM-yyyy"
    '    oBC.Range("A7:AV1000").VerticalAlignment = Excel.Constants.xlTop
    '    oBC.Range("N:N").ColumnWidth = 30


    '    oBC.Range("Q:AB").ColumnWidth = 11
    '    oBC.Range("S:S").ColumnWidth = 8
    '    oBC.Range("U:U").ColumnWidth = 8
    '    oBC.Range("W:W").ColumnWidth = 8
    '    oBC.Range("Y:Y").ColumnWidth = 8
    '    oBC.Range("AA:AA").ColumnWidth = 8
    '    oBC.Range("G:G").ColumnWidth = 15
    '    oBC.Cells.EntireColumn.WrapText = True
    '    ProgressReport(1, "Put Picture(s)...")


    '    'assign Photo
    '    Dim DT As DataTable = DataAccess.GetDataTable(SqlStrPhoto, CommandType.Text)
    '    Dim BS As New BindingSource
    '    BS.DataSource = DT


    '    Dim iRow As Long
    '    Dim Span As Double
    '    Dim j As Integer
    '    Dim Pict As Object
    '    Dim myfirstRow As Long
    '    Dim mylastRow As Long

    '    'Dim myi As Long
    '    Dim myWD As Double
    '    Dim ganjil As Long
    '    Dim reccount As Long

    '    Dim myfilename As String
    '    iRow = 7
    '    oWF = oxl.WorksheetFunction
    '    oBC.Rows("7:" & DT.Rows.Count + 6).Interior.ColorIndex = 36
    '    Call GenerateReport1.MakeBorder(oBC, 7, 1, DT.Rows.Count + 6, 49, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous)
    '    Call GenerateReport1.MakeInsideBorder(oBC, 7, 1, DT.Rows.Count + 6, 49, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous)



    '    ' While Not rsTmp.EOF
    '    For i = 0 To DT.Rows.Count - 1
    '        'If Not (IsNull(rstmp!imagepath)) Then
    '        'myform.StatusBar1.Panels(1).Text = "Put Picture(s)..." & reccount & "/" & rsTmp.RecordCount
    '        ProgressReport(1, String.Format("Put Picture(s).. {0}/{1}", reccount, DT.Rows.Count - 1))
    '        If IsDBNull(DT.Rows(i).Item("projectname")) AndAlso IsDBNull(DT.Rows(i).Item("rangename")) Then

    '        Else
    '            If GenerateReport1.FirstOf(BS, 1) Then
    '                myfirstRow = iRow
    '            End If
    '            If GenerateReport1.LastOf(BS, 1) Then
    '                mylastRow = iRow
    '                ganjil = ganjil + 1
    '                If ganjil Mod 2 = 0 Then
    '                    oBC.Rows(myfirstRow & ":" & mylastRow).Interior.ColorIndex = 40
    '                End If
    '                If Not (IsDBNull(DT.Rows(i).Item("imagepath"))) Then

    '                    myfilename = myImagePath & "\" & Path.GetFileName(Trim(DT.Rows(i).Item("imagepath")))

    '                    If File.Exists(myfilename) Then
    '                        If oBC.Range("H" & myfirstRow & ":H" & mylastRow).Height < 100 Then
    '                            'Reserve height for cell
    '                            If myfirstRow = mylastRow Then
    '                                oBC.Cells(myfirstRow, 2).RowHeight = 100
    '                            Else
    '                                Span = (100 - oBC.Range("H" & myfirstRow & ":H" & mylastRow).Height) / (mylastRow + 1 - myfirstRow)
    '                                For j = 0 To (mylastRow - myfirstRow)
    '                                    If j = mylastRow - myfirstRow Then
    '                                        Span = 100 - oBC.Range("H" & myfirstRow & ":H" & mylastRow).Height
    '                                    End If
    '                                    oBC.Cells(myfirstRow + j, 2).RowHeight = oBC.Cells(myfirstRow + j, 2).RowHeight + Span
    '                                Next j
    '                            End If
    '                        End If
    '                        oBC.Range("H" & myfirstRow & ":H" & mylastRow).Merge()
    '                        'Print
    '                        Dim PictHT As Single, PictWD As Single
    '                        Dim RngHT As Single, RngWD As Single
    '                        Dim myScale As Double, MyL As Double, MyT As Double
    '                        ProgressReport(1, String.Format("Put Picture(s)...{0}/{1} {2}", reccount, DT.Rows.Count - 1, myfilename))
    '                        'myform.StatusBar1.Panels(1).Text = "Put Picture(s)..." & reccount & "/" & rsTmp.RecordCount & myfilename 'rsTmp!imagepath

    '                        'oBC.Pictures.Insert (rsTmp!imagepath)
    '                        Pict = oBC.Pictures.Insert(myfilename)

    '                        'Pict = oBC.Pictures(oBC.Pictures.Count)
    '                        Pict.Placement = 1

    '                        'do pic scale
    '                        Application.DoEvents()

    '                        PictHT = Pict.Height
    '                        PictWD = Pict.Width

    '                        'Size of Template Picture
    '                        RngHT = 98
    '                        RngWD = 98

    '                        myScale = oWF.Min(RngHT / PictHT, RngWD / PictWD)

    '                        With Pict
    '                            .Width = PictWD * myScale
    '                            .Height = PictHT * myScale

    '                        End With
    '                        'Put Picture Location
    '                        With oBC.Range("H" & myfirstRow)
    '                            MyT = .Top
    '                            MyL = .Left
    '                        End With

    '                        'Center Pict H & V
    '                        If PictHT > PictWD Then
    '                            MyL = MyL + (RngWD - Pict.Width) / 2
    '                        Else
    '                            MyT = MyT + (RngHT - Pict.Height) / 2
    '                        End If
    '                        myWD = 0

    '                        myWD = (oBC.Range("B" & iRow).Offset(0, 1).Left) - (oBC.Range("B" & iRow).Left) - 2

    '                        With Pict
    '                            .Top = MyT + 1
    '                            .Left = MyL + 2
    '                        End With
    '                        Pict = Nothing
    '                    End If
    '                End If
    '            End If
    '        End If
    '        'If DT.Rows(i).Item("projectname") <> "" And DT.Rows(i).Item("rangename") <> "" Then

    '        'End If
    '        iRow = iRow + 1
    '        reccount = reccount + 1
    '    Next
    '    ' End While
    '    'End If
    '    ProgressReport(1, "AutoFit ....")
    '    Application.DoEvents()
    '    oBC.Cells.EntireColumn.AutoFit()
    '    ProgressReport(1, "Interior Pattern ....")
    '    oBC.Range(oBC.Columns("AX:AX"), oBC.Columns("AX:AX").End(Excel.XlDirection.xlToRight)).Interior.Pattern = Excel.Constants.xlNone
    '    ProgressReport(1, "Set Range ....")

    '    oBC.Range("H:H").ColumnWidth = 20
    '    oBC.Range("R:R").NumberFormat = "dd-MMM-yyyy"
    '    oBC.Range("A7:AP1000").VerticalAlignment = Excel.Constants.xlTop
    '    oBC.Range("N:N").ColumnWidth = 30
    '    oBC.Range("Q:AB").ColumnWidth = 11
    '    oBC.Range("S:S").ColumnWidth = 8
    '    oBC.Range("U:U").ColumnWidth = 8
    '    oBC.Range("W:W").ColumnWidth = 8
    '    oBC.Range("Y:Y").ColumnWidth = 8
    '    oBC.Range("AA:AA").ColumnWidth = 8
    '    oBC.Range("G:G").ColumnWidth = 15
    '    ProgressReport(1, "Wrap Text ....")
    '    'myform.StatusBar1.Panels(1).Text = "Wrap Text ...."
    '    oBC.Cells.EntireColumn.WrapText = True
    '    oBC.Range("G1:I1").WrapText = False
    '    oBC.Cells(1, 8) = Date.Today
    '    oBC.Cells(1, 9) = TextBox1.Text
    '    oWb.Worksheets(1).Select()
    '    oxl.DisplayAlerts = False
    '    oBC.PageSetup.PrintArea = "$A$1:$AJ" & DT.Rows.Count + 6

    'End Sub

End Class