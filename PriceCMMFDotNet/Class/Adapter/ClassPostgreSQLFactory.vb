﻿Imports Npgsql
Public Class PostgreSQLFactory
    Inherits DbFactory




    Private myadapter As NpgsqlDataAdapter

    Dim conninfo As New NpgsqlConnectionStringBuilder

    Public Overrides ReadOnly Property GetHostName As String
        Get
            conninfo.ConnectionString = ConnectionString
            Return conninfo.Host
        End Get
    End Property
    Public Overrides ReadOnly Property GetDataBaseName As String
        Get
            conninfo.ConnectionString = ConnectionString
            Return conninfo.Database
        End Get
    End Property

    Public Overrides ReadOnly Property ConnectionString As String
        Get
            'Return String.Format("host=hon13-0002L;port=5432;database=LogisticDb20150120;CommandTimeout=10000;TimeOut=1024;Userid=admin;Password=admin;")
            'Return String.Format("host=localhost;port=5432;database=LogisticDb20150120;CommandTimeout=10000;TimeOut=1024;Userid=admin;Password=admin;")
            Return String.Format(My.Settings.Conn1)
        End Get
    End Property

    Public Overloads Overrides Function CreateAdapter() As IDbDataAdapter
        myadapter = New NpgsqlDataAdapter
        AddHandler myadapter.RowUpdated, AddressOf onRowUpdate
        Return myadapter
    End Function

    Public Overrides Function Update(tablename As Object) As Integer
        ' Return myadapter.Update(DirectCast(tablename, DataTable))
        Return myadapter.Update(tablename)
    End Function

    Public Overrides Function CreateAdapter(commandText As String) As IDbDataAdapter
        Return New NpgsqlDataAdapter(New NpgsqlCommand(commandText))
    End Function

    Public Overrides Function CreateCommand(commandText As String, connection As IDbConnection) As IDbCommand
        Return New NpgsqlCommand(commandText, connection)
    End Function

    Public Overrides Function CreateConnection() As IDbConnection
        Dim myret As IDbConnection
        myret = New NpgsqlConnection(ConnectionString)
        Return myret
    End Function

    Public Overloads Overrides Function CreateParameter() As IDbDataParameter
        Return New NpgsqlParameter
    End Function

    Public Overloads Overrides Function CreateParameter(name As String, value As Object) As IDbDataParameter
        Return New NpgsqlParameter(name, value)
    End Function

    Public Overloads Overrides Function CreateParameter(name As String, type As DbType, size As Integer) As IDbDataParameter
        Dim parm As NpgsqlParameter = New NpgsqlParameter
        parm.ParameterName = name
        parm.DbType = type
        parm.Size = size
        Return parm

    End Function
    Public Overloads Overrides Function CreateParameter(name As String, type As DbType, size As Integer, sourceColumn As String) As IDbDataParameter
        Dim parm As NpgsqlParameter = New NpgsqlParameter
        parm.ParameterName = name
        parm.DbType = type
        parm.Size = size
        parm.SourceColumn = sourceColumn
        Return parm
    End Function
    Public Overrides Function GetParameterValue(parameter As Object) As Object
        Debug.Assert(parameter <> Nothing)
        If (parameter Is Nothing) Then Return Nothing
        Return CType(parameter, NpgsqlParameter).Value
    End Function

    Private Sub onRowUpdate(sender As Object, e As NpgsqlRowUpdatedEventArgs)
        If e.StatementType = StatementType.Insert Or e.StatementType = StatementType.Update Then
            If e.Status <> UpdateStatus.ErrorsOccurred Then
                e.Status = UpdateStatus.SkipCurrentRow
            End If
        End If
    End Sub



    Public Overloads Overrides Function CreateParameter(name As String, type As DbType, size As Integer, sourcecolumn As String, sourceversion As DataRowVersion) As IDbDataParameter
        Dim parm As NpgsqlParameter = New NpgsqlParameter
        parm.ParameterName = name
        parm.DbType = type
        parm.Size = size
        parm.SourceColumn = sourcecolumn
        parm.SourceVersion = sourceversion
        Return parm
    End Function

    Public Overloads Overrides Function CreateParameter(name As String, type As DbType, size As Integer, sourcecolumn As String, paramInputOutput As ParameterDirection) As IDbDataParameter
        Dim parm As NpgsqlParameter = New NpgsqlParameter
        parm.ParameterName = name
        parm.DbType = type
        parm.Size = size
        parm.SourceColumn = sourcecolumn
        parm.Direction = paramInputOutput
        Return parm
    End Function
End Class
