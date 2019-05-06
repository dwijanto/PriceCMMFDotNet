'-- Table: sales._user

'-- DROP TABLE sales._user;

'CREATE TABLE sales._user
'(
'  id bigserial NOT NULL,
'  userid character varying,
'  username character varying,
'  email character varying,
'  isadmin boolean NOT NULL DEFAULT false,
'  isactive boolean NOT NULL DEFAULT true,
'  CONSTRAINT userpk PRIMARY KEY (id)
')
'WITH (
'  OIDS=FALSE
');
'ALTER TABLE sales._user
'  OWNER TO postgres;
'GRANT ALL ON TABLE sales._user TO postgres;
'GRANT ALL ON TABLE sales._user TO public;

'-- Index: sales.useridx1

'-- DROP INDEX sales.useridx1;

'CREATE UNIQUE INDEX useridx1
'  ON sales._user
'  USING btree
'  (userid COLLATE pg_catalog."default");


Public Class UserModel
    Implements IModel

    Public ReadOnly Property FilterField
        Get
            Return "[userid] like '%{0}%' or [username] like '%{0}%' or [email] like '%{0}%'"
        End Get
    End Property

    Public ReadOnly Property TableName As String Implements IModel.tablename
        Get
            Return "sales._user"
        End Get
    End Property

    Public ReadOnly Property SortField As String Implements IModel.sortField
        Get
            Return "username"
        End Get
    End Property

    Public Function LoadData(ByRef DS As DataSet) As Boolean Implements IModel.LoadData
        Dim sqlstr = String.Format("select u.* from {0} u order by {1}", TableName, SortField)
        DS = DataAccess.GetDataSet(sqlstr, CommandType.Text, Nothing)
        DS.Tables(0).TableName = TableName
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
            'Update
            Dim sqlstr = "sales.sp_update_user"
            dataadapter.UpdateCommand = factory.CreateCommand(sqlstr, conn)
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id"))
            dataadapter.UpdateCommand.Parameters(0).SourceVersion = DataRowVersion.Original
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "userid"))
            dataadapter.UpdateCommand.Parameters(1).SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "username"))
            dataadapter.UpdateCommand.Parameters(2).SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "email"))
            dataadapter.UpdateCommand.Parameters(3).SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "isactive"))
            dataadapter.UpdateCommand.Parameters(4).SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "sales.sp_insert_user"
            dataadapter.InsertCommand = factory.CreateCommand(sqlstr, conn)
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "userid"))
            dataadapter.InsertCommand.Parameters(0).SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "username"))
            dataadapter.InsertCommand.Parameters(1).SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "email"))
            dataadapter.InsertCommand.Parameters(2).SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "isactive"))
            dataadapter.InsertCommand.Parameters(3).SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id"))
            dataadapter.InsertCommand.Parameters(4).Direction = ParameterDirection.InputOutput
            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "sales.sp_delete_user"
            dataadapter.DeleteCommand = factory.CreateCommand(sqlstr, conn)
            dataadapter.DeleteCommand.Parameters.Add(factory.CreateParameter("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id"))
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

End Class
