﻿Public Class CMMFController   
    Implements IController
    Implements IToolbarAction

    Public Model As New CMMFModel
    Public DS As DataSet

    Public WithEvents BS As BindingSource
    Public Event PositionChangeEventHandler()
    Public Event PricePositionChangeEventHandler()

    'Public PriceCMMFSBUBS As New BindingSource
    'Public PriceCMMFFamilyBS As New BindingSource
    'Public PriceCMMFSSMBS As New BindingSource
    'Public PriceCMMFSPMBS As New BindingSource
    'Public PriceCMMFProjectBS As New BindingSource
    'Public PriceCMMFRangeBS As New BindingSource

    'Public PriceCMMFBrandBS As New BindingSource
    'Public PriceCMMFLoadingBS As New BindingSource
    'Public PriceCMMFPurchasingGRPBS As New BindingSource
    Public Sub New()
        MyBase.New()
    End Sub

    Public ReadOnly Property GetTable As DataTable Implements IController.GetTable
        Get
            Return DS.Tables(Model.TableName).copy()
        End Get
    End Property

    Public ReadOnly Property GetBindingSource As BindingSource
        Get
            Dim BS As New BindingSource
            BS.DataSource = GetTable
            BS.Sort = Model.SortField
            Return BS
        End Get
    End Property

    Public Function loaddata() As Boolean Implements IController.loaddata
        Dim myret As Boolean = False
        Model = New CMMFModel
        DS = New DataSet
        If Model.LoadData(DS) Then
            Dim pk(0) As DataColumn
            pk(0) = DS.Tables(0).Columns("cmmf")
            DS.Tables(0).PrimaryKey = pk
            BS = New BindingSource
            BS.DataSource = DS.Tables(0)

           

            myret = True
        End If
        Return myret
    End Function
    Public Function loaddata(ByVal criteria As String) As Boolean
        Dim myret As Boolean = False
        Model = New CMMFModel

        RemoveHandler Model.PositionChangedEventhandler, AddressOf BSAgrementChange
        AddHandler Model.PositionChangedEventhandler, AddressOf BSAgrementChange

        DS = New DataSet
        Model.BSPriceList = New BindingSource        
        If Model.LoadData(DS, criteria) Then
            Dim pk(0) As DataColumn
            pk(0) = DS.Tables(0).Columns("cmmf")
            DS.Tables(0).PrimaryKey = pk

            Dim pk1(0) As DataColumn
            pk1(0) = DS.Tables(5).Columns("pcprojectid")
            DS.Tables(5).PrimaryKey = pk1
            DS.Tables(5).Columns("pcprojectid").AutoIncrement = True
            DS.Tables(5).Columns("pcprojectid").AutoIncrementSeed = -1
            DS.Tables(5).Columns("pcprojectid").AutoIncrementStep = -1


            Dim pk2(0) As DataColumn
            pk2(0) = DS.Tables(6).Columns("pcrangeid")
            DS.Tables(6).PrimaryKey = pk2
            DS.Tables(6).Columns("pcrangeid").AutoIncrement = True
            DS.Tables(6).Columns("pcrangeid").AutoIncrementSeed = -1
            DS.Tables(6).Columns("pcrangeid").AutoIncrementStep = -1

            myret = True
            
        End If
        Return myret
    End Function

    Public Function save() As Boolean Implements IController.save
        Dim myret As Boolean = False

        BS.EndEdit()

        Dim ds2 As DataSet = DS.GetChanges()
        If Not IsNothing(ds2) Then
            Dim mymessage As String = String.Empty
            Dim ra As Integer
            Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
            Try
                If save(mye) Then
                    DS.Merge(ds2)
                    'Don't use DS.AcceptChanges. Use the statement below.
                    For Each mytable As DataTable In ds2.Tables
                        If mytable.Rows.Count > 0 Then                           
                            DS.Tables(mytable.TableName).AcceptChanges()
                        End If
                    Next                    
                    MessageBox.Show("Saved.")
                    myret = True
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                DS.Merge(ds2)
            End Try
        End If

        Return myret
    End Function

    Public Function Save(ByVal mye As ContentBaseEventArgs) As Boolean Implements IToolbarAction.Save
        Dim myret As Boolean = False
        If Model.save(Me, mye) Then
            myret = True
        End If
        Return myret
    End Function

    Public Property ApplyFilter As String Implements IToolbarAction.ApplyFilter
        Get
            Return BS.Filter
        End Get
        Set(ByVal value As String)
            BS.Filter = String.Format(value)
        End Set
    End Property
    Public Function GetCurrentRecord() As DataRowView Implements IToolbarAction.GetCurrentRecord
        Return BS.Current
    End Function

    Public Function GetNewRecord() As DataRowView Implements IToolbarAction.GetNewRecord
        Return BS.AddNew
    End Function

    Public Sub RemoveAt(value As Integer) Implements IToolbarAction.RemoveAt
        BS.RemoveAt(value)
    End Sub


    Private Sub BS_PositionChanged(sender As Object, e As EventArgs) Handles BS.PositionChanged
        RaiseEvent PositionChangeEventHandler()
    End Sub

    Private Sub BSAgrementChange()
        RaiseEvent PricePositionChangeEventHandler()
    End Sub


End Class