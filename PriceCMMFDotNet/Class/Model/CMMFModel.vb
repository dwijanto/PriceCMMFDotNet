Imports System.Text

Public Class CMMFModel
    Implements IModel

    Public WithEvents BSPricelist As BindingSource
    Public Event PositionChangedEventhandler()

    Public BSAgreementList As BindingSource

    Public ReadOnly Property FilterField
        Get
            Return ""
        End Get
    End Property

    Public ReadOnly Property TableName As String Implements IModel.tablename
        Get
            Return "pccmmf"
        End Get
    End Property

    Public ReadOnly Property SortField As String Implements IModel.sortField
        Get
            Return "cmmf"
        End Get
    End Property

    Public Function SyncFamilyId() As Long
        Dim sqlstr = "update pccmmf set familyid =  foo.comfam from (select comfam,cmmf from cmmf where not comfam isnull) foo where pccmmf.cmmf = foo.cmmf; " &
                     "update pcproject set familyid = foo.familyid from " &
                     "(select pp.pcprojectid,pc.familyid from pccmmf pc " &
                     " left join pcrange pr on pr.pcrangeid = pc.pcrangeid" &
                     " left join pcproject pp on pp.pcprojectid = pr.pcprojectid " &
                     " where not pp.pcprojectid isnull and not pc.familyid isnull ) foo " &
                     " where pcproject.pcprojectid = foo.pcprojectid;"
        Return DataAccess.ExecuteNonQuery(sqlstr, CommandType.Text)
    End Function

    Private Function GetSqlstrSUBFamily() As String
        Dim sb As New StringBuilder
        sb.Append("select '--Select From List--' as sbuname,0 as sbuid union all (Select s.sbuname::text,s.sbuid from sbu  s where pcmmf order by s.sbuname);")                                                                         'SBU
        sb.Append("select '--Select From List--' as familyname,0 as familyid,0 as sbuid,'--Select From List--'::text as sbuname union all (Select distinct f.familyname::text,f.familyid,f.sbuid,s.sbuname::text from family f left join sbu s on s.sbuid = f.sbuid  where s.pcmmf order by f.familyname::text);")                                  'Family
        Return sb.ToString
    End Function

    Private Function GetSqlstrSSMSupplier() As String
        Dim sb As New StringBuilder
        sb.Append("with ssm as (select distinct ssmid from pcproject) select ssm.ssmid,o.officersebname  from ssm left join officerseb o on o.ofsebid = ssm.ssmid order by officersebname;")                                                                         'SBU
        sb.Append("select distinct v.shortname::character varying,v.vendorcode,ssmid from pcproject p left join pcrange r on r.pcprojectid = p.pcprojectid left join pccmmf c on c.pcrangeid = r.pcrangeid left join pricelist pl on pl.cmmf = c.cmmf left join vendor v on v.vendorcode = pl.vendorcode where not shortname isnull order by shortname;")                                  'Family
        Return sb.ToString
    End Function

    Private Function GetSqlstr(ByVal criteria) As String
        Dim sb As New StringBuilder
        sb.Append(String.Format("SELECT  r.pcprojectid,p.projectname::text, c.cmmf, c.materialcode::text, ssm.ofsebid as ssmid,spm.ofsebid as spmid,ssm.officersebname::text AS ssm, spm.officersebname::text AS spm,sbu.sbuname::text,sbu.sbuid," &
               " c.familyid,f.familyname::text, c.pcrangeid,  r.rangename::text, NULL::unknown AS picture,  c.description::text,c.brandid, b.brandname::text, c.countries::text, c.voltage, c.power, " &
               " c.leadtime, c.qty20, c.qty40, c.qty40hq, c.moq,c.loadingcode::text, l.loadingname,c.pgid, pg.typeofitem, c.netprice,  c.contractno, c.length," &
               " c.width, c.height, c.lengthbox, c.widthbox, c.heightbox, c.weightwo, c.weightwi, c.nettweight,c.grossweight, remarks," &
               " c.sppet,c.stcseb,c.stcsup,c.srdc,c.eol,c.pcspercartoon,c.spps, c.amort, l.loadinggroup,spm.mrpcontrollercode,pg.purchasinggroup FROM pccmmf c " &
               " LEFT JOIN pcrange r ON r.pcrangeid = c.pcrangeid " &
               " LEFT JOIN pcproject p ON p.pcprojectid = r.pcprojectid " &
               " LEFT JOIN family f ON f.familyid = c.familyid " &
               " LEFT JOIN sbu ON sbu.sbuid = f.sbuid " &
               " LEFT JOIN officerseb ssm ON ssm.ofsebid = p.ssmid" &
               " LEFT JOIN officerseb spm ON spm.ofsebid = p.spmid" &
               " LEFT JOIN brand b ON b.brandid = c.brandid" &
               " LEFT JOIN loading l ON l.loadingcode = c.loadingcode" &
               " LEFT JOIN purchasinggroup pg ON pg.pgid = c.pgid {0} order by c.cmmf;", criteria))
        sb.Append("select '--Select From List--' as sbuname,0 as sbuid union all (Select s.sbuname::text,s.sbuid from sbu  s where pcmmf order by s.sbuname);")                                                                         'SBU
        sb.Append("select '--Select From List--' as familyname,0 as familyid,0 as sbuid,'--Select From List--'::text as sbuname union all (Select distinct f.familyname::text,f.familyid,f.sbuid,s.sbuname::text from family f left join sbu s on s.sbuid = f.sbuid  where s.pcmmf order by f.familyname::text);")                                  'Family
        sb.Append("select '--Select From List--' as ssmname,0 as ssmid union all (Select o.officersebname::text,o.ofsebid from officerseb o where o.parent = 0 and teamtitleid <> 18  and o.isactive order by o.officersebname);")                                   'SSM
        sb.Append("select '--Select From List--' as ssmname,0 as ssmid,'--Select From List--' as pmname,0 as pmid union all (Select o.officersebname::text,o.ofsebid, p.officersebname::text,p.ofsebid from officerseb o left join officerseb p on p.parent = o.ofsebid where o.parent = 0 and not p.ofsebid isnull and o.isactive order by o.officersebname);")    'SPM
        sb.Append("select '--Select From List / Type for new--' as projectname,0 as pcprojectid union all (select p.projectname::text,p.pcprojectid from pcproject p  order by p.projectname);")
        sb.Append("select '--Select From List / Type for new--' as rangename,0 as pcrangeid,0 as pcprojectid  union all (select r.rangename::text,r.pcrangeid,r.pcprojectid from pcrange r  order by  r.rangename);")

        sb.Append("select '--Select From List--' as brandname,0 as brandid union all (Select brandname::text ,brandid from brand order by brandname);")
        sb.Append("select '--Select From List--' as loadingname,null::text as loadingcode,null as loadinggroup union all (Select loadingname::text,loadingcode,loadinggroup from loading where not loadinggroup isnull order by loadingname);")
        sb.Append("select '--Select From List--' as typeofitem,null::integer as pgid,null::text as purchasinggroup union all (Select typeofitem::text,pgid,purchasinggroup from purchasinggroup where pccmmf order by typeofitem::text);")
        sb.Append("select pp.projectname::text,pr.rangename::text,ssm.officersebname as ssm,spm.officersebname as spm, pp.pcprojectid,pr.pcrangeid,pp.ssmid,pp.spmid from pcproject pp" &
                  " left join officerseb ssm on ssm.ofsebid = pp.ssmid left join officerseb spm on spm.ofsebid = pp.spmid left join pcrange pr on pr.pcprojectid = pp.pcprojectid" &
                  " where  not( rangename isnull and projectname isnull) order by projectname,ssmid,spmid,pcprojectid") 'Project Helper
        Return sb.ToString
    End Function

    Private Function GetSqlstr() As String
        Dim sb As New StringBuilder
        sb.Append(String.Format("SELECT  p.projectname::text, c.cmmf, c.materialcode::text, ssm.officersebname::text AS ssm, spm.officersebname::text AS spm,sbu.sbuname,sbu.sbuid," &
               " f.familyname,   r.rangename, NULL::unknown AS picture,  c.description, b.brandname, c.countries, c.voltage, c.power, " &
               " c.leadtime, c.qty20, c.qty40, c.qty40hq, c.moq, l.loadingname, pg.typeofitem, c.netprice,  c.contractno, c.length," &
               " c.width, c.height, c.lengthbox, c.widthbox, c.heightbox, c.weightwo, c.weightwi, c.nettweight,c.grossweight, remarks," &
               " c.sppet,c.stcseb,c.stcsup,c.srdc,c.eol,c.pcspercartoon,c.spps FROM pccmmf c " &
               " LEFT JOIN pcrange r ON r.pcrangeid = c.pcrangeid " &
               " LEFT JOIN pcproject p ON p.pcprojectid = r.pcprojectid " &
               " LEFT JOIN family f ON f.familyid = c.familyid " &
               " LEFT JOIN sbu ON sbu.sbuid = f.sbuid " &
               " LEFT JOIN officerseb ssm ON ssm.ofsebid = p.ssmid" &
               " LEFT JOIN officerseb spm ON spm.ofsebid = p.spmid" &
               " LEFT JOIN brand b ON b.brandid = c.brandid" &
               " LEFT JOIN loading l ON l.loadingcode = c.loadingcode" &
               " LEFT JOIN purchasinggroup pg ON pg.pgid = c.pgid order by pc.cmmf;"))
        Return sb.ToString
    End Function

    Public Function LoadData(ByRef DS As DataSet) As Boolean Implements IModel.LoadData
        Dim sqlstr = GetSqlstr("")
        DS = DataAccess.GetDataSet(sqlstr, CommandType.Text, Nothing)
        DS.Tables(0).TableName = TableName

        Return True
    End Function

    Public Function LoadDataSSMSupplier(ByRef ds As DataSet) As Boolean
        Dim sqlstr = GetSqlstrSSMSupplier()
        ds = DataAccess.GetDataSet(sqlstr, CommandType.Text, Nothing)
        Dim rel As DataRelation
        Dim hcol As DataColumn
        Dim dcol As DataColumn
        'create relation SBU Family
        hcol = ds.Tables(0).Columns("ssmid") 'id in table header
        dcol = ds.Tables(1).Columns("ssmid") 'headerid in table detail
        rel = New DataRelation("hdrel-SSM", hcol, dcol)
        ds.Relations.Add(rel)
        Return True
    End Function

    Public Function LoadDataSBUFamily(ByRef ds As DataSet) As Boolean
        Dim sqlstr = GetSqlstrSUBFamily()
        ds = DataAccess.GetDataSet(sqlstr, CommandType.Text, Nothing)
        Dim rel As DataRelation
        Dim hcol As DataColumn
        Dim dcol As DataColumn
        'create relation SBU Family
        hcol = ds.Tables(0).Columns("sbuid") 'id in table header
        dcol = ds.Tables(1).Columns("sbuid") 'headerid in table detail
        rel = New DataRelation("hdrel-SBU", hcol, dcol)
        ds.Relations.Add(rel)
        Return True
    End Function

    Public Function LoadData(ByRef DS As DataSet, ByVal criteria As String) As Boolean
        Dim sqlstr = GetSqlstr(criteria)
        DS = DataAccess.GetDataSet(sqlstr, CommandType.Text, Nothing)
        DS.Tables(0).TableName = TableName
        Dim rel As DataRelation
        Dim hcol As DataColumn
        Dim dcol As DataColumn

        'create relation PCCMMF with PCRange
        hcol = DS.Tables(6).Columns("pcrangeid") 'id in table header
        dcol = DS.Tables(0).Columns("pcrangeid") 'headerid in table vendordoc
        rel = New DataRelation("hdrel-PCRange", hcol, dcol)
        DS.Relations.Add(rel)

        'Create relation PCCMMF with PCProject
        hcol = DS.Tables(5).Columns("pcprojectid") 'id in table header
        dcol = DS.Tables(0).Columns("pcprojectid") 'headerid in table vendordoc
        rel = New DataRelation("hdrel-PCProject", hcol, dcol)
        DS.Relations.Add(rel)

        'create relation Project and Pc Range
        hcol = DS.Tables(5).Columns("pcprojectid") 'PCProject
        dcol = DS.Tables(6).Columns("pcprojectid") 'PCRange
        rel = New DataRelation("hdrel", hcol, dcol)
        DS.Relations.Add(rel)

        'create relation SBU Family
        hcol = DS.Tables(1).Columns("sbuid") 'id in table header
        dcol = DS.Tables(2).Columns("sbuid") 'headerid in table vendordoc
        rel = New DataRelation("hdrel-SBU", hcol, dcol)
        DS.Relations.Add(rel)

        'create relation SSM PM
        hcol = DS.Tables(3).Columns("ssmid") 'id in table header
        dcol = DS.Tables(4).Columns("ssmid") 'headerid in table vendordoc
        rel = New DataRelation("hdrel-SSM", hcol, dcol)
        DS.Relations.Add(rel)

        Return True
    End Function

    Public Function getVendorTable() As DataTable
        Dim ds As New DataSet
        Dim sb As New StringBuilder
        sb.Append("Select distinct vendorname::text,v.vendorcode , v.vendorcode::text || ' - ' || v.vendorname::text as vendorcodename from pricelist pc" &
                  " left join vendor v on v.vendorcode = pc.vendorcode" &
                  " Union select vendorname::text,v.vendorcode,v.vendorcode::text || ' - ' || v.vendorname::text  from vendor v order by vendorname;")

        ds = DataAccess.GetDataSet(sb.ToString, CommandType.Text, Nothing)
        Return ds.Tables(0)
    End Function

    Public Function getVendorShortNameTable() As DataTable
        Dim ds As New DataSet
        Dim sb As New StringBuilder
        sb.Append("select shortname from vendor v " &
                  " inner join (select distinct vendorcode from pricelist) as pv on pv.vendorcode = v.vendorcode " &
                  " where not (shortname isnull) order by shortname")
        ds = DataAccess.GetDataSet(sb.ToString, CommandType.Text, Nothing)
        Return ds.Tables(0)
    End Function

    Public Function GetPriceList(ByVal criteria As String) As Boolean
        Dim ds As New DataSet
        BSPriceList = New BindingSource
        BSAgreementList = New BindingSource
        Dim sb As New StringBuilder
        sb.Append(String.Format("with pcd as ( select pcd.vendorcode,pcd.comment,pcd.cmmf,pcd.validon,phd.status,phd.reasonid,phd.negotiateddate " &
                                " from pricechangedtl pcd left join pricechangehd phd on phd.pricechangehdid = pcd.pricechangehdid) " &
                                " Select v.shortname::text, pl.validfrom as mydate,pl.amount::real/perunit::real as price,pl.cmmf,pc.comments,ag.agreement," &
                                " ag.value as amort,ag.closingdate,ag.status,ag.cmmf as numbercmmf,ag.agqty,ag.totalqty,ag.value as amort," &
                                " ag.deliveredqty,ag.startdate,ag.enddate,pcd.negotiateddate::character varying,pcd.comment,pcr.reasonname," &
                                " format('%s%s%s%s',case when pc.comments isnull then null else pc.comments || ' ' end,case when pcd.comment isnull then null else pcd.comment || ' ' end,case when pcd.negotiateddate isnull then null else format('[Negotiated date : %s]',to_char(pcd.negotiateddate,'FMYYYY-Mon-dd')) end,case when pcr.reasonname isnull then null else format('[Reason : %s]',pcr.reasonname) end)  as allcomments," &
                                " case when ag.enddate >= current_date then ag.value else 0 end as validamort," &
                                " case when ag.enddate >= current_date then pl.amount::real/perunit::real + ag.value else pl.amount::real/perunit::real  end as fobamort" &
                                " from pricelist pl " &
                                " left join pricecomment pc on pc.cmmf = pl.cmmf and pc.effectivedate = pl.validfrom" &
                                " left join vendor v on v.vendorcode = pl.vendorcode " &
                                " left join agreementcmmf ac on ac.material = pl.cmmf " &
                                " left join agvalue ag on ag.agreement = ac.agreement and ag.vendorcode = pl.vendorcode" &
                                " left join pcd  on pcd.vendorcode = v.vendorcode and pcd.cmmf = pl.cmmf and pcd.validon = pl.validfrom and (pcd.status = 5 or pcd.status = 7)" &
                                " left join pricechangereason pcr on pcr.id = pcd.reasonid where  pl.cmmf= {0} " &
                                " order by v.shortname, pl.validfrom desc;", criteria))
        sb.Append(String.Format("select distinct material,agreement from agreementtx agtx order by material;"))
        ds = DataAccess.GetDataSet(sb.ToString, CommandType.Text, Nothing)
        BSPriceList.DataSource = ds.Tables(0)
        BSAgreementList.DataSource = ds.Tables(1)
        Return True
    End Function

    Public Function save(ByVal obj As Object, ByVal mye As ContentBaseEventArgs) As Boolean Implements IModel.save
        Dim myret As Boolean = False
        Dim factory = DataAccess.factory
        Dim mytransaction As IDbTransaction
        Using conn As IDbConnection = factory.CreateConnection
            conn.Open()
            mytransaction = conn.BeginTransaction
            Dim dataadapter = factory.CreateAdapter
            Dim sqlstr As String = String.Empty

            sqlstr = "sp_insertpcproject"
            dataadapter.InsertCommand = factory.CreateCommand(sqlstr, conn)
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "projectname", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "ssmid", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "spmid", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "familyid", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "pcprojectid", ParameterDirection.InputOutput))
            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "sp_updatepcproject"
            dataadapter.UpdateCommand = factory.CreateCommand(sqlstr, conn)
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "pcprojectid", DataRowVersion.Original))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "projectname", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "ssmid", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "spmid", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "familyid", DataRowVersion.Current))
            dataadapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "sp_deletepcproject"
            dataadapter.DeleteCommand = factory.CreateCommand(sqlstr, conn)
            dataadapter.DeleteCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "pcprojectid", DataRowVersion.Original))           
            dataadapter.DeleteCommand.CommandType = CommandType.StoredProcedure

            dataadapter.InsertCommand.Transaction = mytransaction
            dataadapter.UpdateCommand.Transaction = mytransaction
            dataadapter.DeleteCommand.Transaction = mytransaction

            mye.ra = factory.Update(mye.dataset.Tables(5))
            'Table Range

            sqlstr = "sp_insertpcrange"
            dataadapter.InsertCommand = factory.CreateCommand(sqlstr, conn)
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "rangename", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "pcprojectid", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "pcrangeid", ParameterDirection.InputOutput))
            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "sp_updatepcrange"
            dataadapter.UpdateCommand = factory.CreateCommand(sqlstr, conn)
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "pcrangeid", DataRowVersion.Original))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "rangename", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "pcprojectid", DataRowVersion.Current))
            dataadapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "sp_deletepcrange"
            dataadapter.DeleteCommand = factory.CreateCommand(sqlstr, conn)
            dataadapter.DeleteCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "pcrangeid", DataRowVersion.Original))
            dataadapter.DeleteCommand.CommandType = CommandType.StoredProcedure

            dataadapter.InsertCommand.Transaction = mytransaction
            dataadapter.UpdateCommand.Transaction = mytransaction
            dataadapter.DeleteCommand.Transaction = mytransaction

            mye.ra = factory.Update(mye.dataset.Tables(6))

            'Update
            sqlstr = "sp_updatepccmmf"
            dataadapter.UpdateCommand = factory.CreateCommand(sqlstr, conn)
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "cmmf", DataRowVersion.Original))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "cmmf", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "pcrangeid", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "materialcode", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "description", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "brandid", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "countries", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "power", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Boolean, 0, "sap", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Decimal, 0, "netprice", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Decimal, 0, "amort", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "contractno", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "leadtime", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "qty20", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "qty40", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "qty40hq", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "moq", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Decimal, 0, "length", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Decimal, 0, "width", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Decimal, 0, "height", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Decimal, 0, "nettweight", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "customercode", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "colorid", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "pgid", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Decimal, 0, "lengthbox", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Decimal, 0, "widthbox", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Decimal, 0, "heightbox", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Decimal, 0, "grossweight", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Decimal, 0, "weightwo", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Decimal, 0, "weightwi", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "loadingcode", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "voltage", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "remarks", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Decimal, 0, "sppet", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Decimal, 0, "stcseb", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Decimal, 0, "stcsup", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Decimal, 0, "srdc", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Boolean, 0, "eol", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "familyid", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "pcspercartoon", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Decimal, 0, "spps", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "pcprojectid", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "ssmid", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "spmid", DataRowVersion.Current))

            dataadapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            'Insert
            sqlstr = "sp_insertpccmmf"
            dataadapter.InsertCommand = factory.CreateCommand(sqlstr, conn)                        
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "pcrangeid", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "materialcode", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "description", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "brandid", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "countries", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "power", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Boolean, 0, "sap", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Decimal, 0, "netprice", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Decimal, 0, "amort", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "contractno", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "leadtime", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "qty20", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "qty40", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "qty40hq", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "moq", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Decimal, 0, "length", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Decimal, 0, "width", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Decimal, 0, "height", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Decimal, 0, "nettweight", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "customercode", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "colorid", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "pgid", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Decimal, 0, "lengthbox", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Decimal, 0, "widthbox", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Decimal, 0, "heightbox", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Decimal, 0, "grossweight", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Decimal, 0, "weightwo", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Decimal, 0, "weightwi", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "loadingcode", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "voltage", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "remarks", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Decimal, 0, "sppet", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Decimal, 0, "stcseb", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Decimal, 0, "stcsup", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Decimal, 0, "srdc", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Boolean, 0, "eol", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "familyid", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "pcspercartoon", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Decimal, 0, "spps", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "pcprojectid", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "ssmid", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "spmid", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "cmmf", ParameterDirection.InputOutput))
            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            'Delete
            sqlstr = "sp_deletepccmmf"
            dataadapter.DeleteCommand = factory.CreateCommand(sqlstr, conn)
            dataadapter.DeleteCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "cmmf"))
            dataadapter.DeleteCommand.Parameters(0).Direction = ParameterDirection.Input
            dataadapter.DeleteCommand.CommandType = CommandType.StoredProcedure

            dataadapter.InsertCommand.Transaction = mytransaction
            dataadapter.UpdateCommand.Transaction = mytransaction
            dataadapter.DeleteCommand.Transaction = mytransaction

            mye.ra = factory.Update(mye.dataset.Tables(TableName))
            mytransaction.Commit()
            myret = True
        End Using
        Return myret
    End Function

    Private Sub BSPriceList_PositionChanged(sender As Object, e As EventArgs) Handles BSPriceList.PositionChanged

        RaiseEvent PositionChangedEventhandler()
    End Sub
End Class
